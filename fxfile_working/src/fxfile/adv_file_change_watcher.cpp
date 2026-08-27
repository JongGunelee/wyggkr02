//
// Copyright (c) 2001-2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "adv_file_change_watcher.h"

#include "functors.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace
{
const DWORD kIdleTime = 20;
// ReadDirectoryChangesW documents 64 KiB as the largest portable buffer for
// network redirected directories.  A real byte buffer is required because
// FILE_NOTIFY_INFORMATION is variable length.
const xpr_size_t kBufferSize = 64 * 1024;
const xpr_sint64_t kNotifyDebounceTime = 120;
const xpr_size_t kMaxRawNotifyQueue = 512;
const xpr_size_t kMaxNotifyPerWatch = 128;
const xpr_size_t kFileNotifyHeaderSize = FIELD_OFFSET(FILE_NOTIFY_INFORMATION, FileName);
const DWORD kCancelCompletionTimeout = 2000;
} // namespace anonymous

struct AdvFileChangeWatcher::AdvWatchOverlapped : public OVERLAPPED
{
    xpr_byte_t      mBuffer[kBufferSize];
    DriveWatchItem *mDriveWatchItem;
    xpr::string     mOldFileName;
    xpr::string     mRootPath;
};

enum AdvFileChangeWatcher::TaskCommand
{
    TaskCommandNone,
    TaskCommandRegister,
    TaskCommandModify,
    TaskCommandUnregister,
    TaskCommandUnregisterAll,
};

struct AdvFileChangeWatcher::Task
{
    Task(void)
        : mTaskCommand(TaskCommandNone)
        , mOldAdvWatchId(InvalidAdvWatchId)
        , mAdvWatchItem(XPR_NULL)
    {
    }

    ~Task(void)
    {
        XPR_SAFE_DELETE(mAdvWatchItem);
    }

    TaskCommand   mTaskCommand;
    AdvWatchId    mOldAdvWatchId;
    AdvWatchItem *mAdvWatchItem;
};

AdvFileChangeWatcher::AdvWatchItem::AdvWatchItem(void)
{
    clear();
    mAdvWatchId = 0;
}

AdvFileChangeWatcher::AdvWatchItem::~AdvWatchItem(void)
{
    clear();
}

void AdvFileChangeWatcher::AdvWatchItem::clear(void)
{
    mHwnd = XPR_NULL;
    mMsg = XPR_NULL;
    mPath.clear();
    mSubPath = XPR_FALSE;
    mParam = XPR_NULL;
}

AdvFileChangeWatcher::AdvWatchItem& AdvFileChangeWatcher::AdvWatchItem::operator = (const AdvWatchItem &aAdvWatchItem)
{
    mHwnd       = aAdvWatchItem.mHwnd;
    mMsg        = aAdvWatchItem.mMsg;
    mPath       = aAdvWatchItem.mPath;
    mSubPath    = aAdvWatchItem.mSubPath;
    mParam      = aAdvWatchItem.mParam;
    mAdvWatchId = aAdvWatchItem.mAdvWatchId;

    return *this;
}

xpr_bool_t AdvFileChangeWatcher::AdvWatchItem::isValidate(void)
{
    if (XPR_IS_NULL(mHwnd) || mMsg == 0 || mPath.empty())
        return XPR_FALSE;

    if (IsExistFile(mPath) == XPR_FALSE || IsFileSystemFolder(mPath) == XPR_FALSE)
        return XPR_FALSE;

    if (mPath.length() > 3 && *mPath.rbegin() == XPR_STRING_LITERAL('\\'))
        mPath.erase(mPath.length()-1);

    return XPR_TRUE;
}

class AdvFileChangeWatcher::DriveWatchItem
{
public:
    typedef std::list<AdvWatchOverlapped *> RetiredOverlappedList;

    DriveWatchItem(void)
        : mDirectory(XPR_NULL)
        , mCompletionPort(XPR_NULL)
        , mOverlapped(XPR_NULL)
        , mWatchSubtree(XPR_FALSE)
        , mIoPending(XPR_FALSE)
        , mStopping(XPR_FALSE)
    {
    }

    ~DriveWatchItem(void)
    {
        mStopping = XPR_TRUE;
        cancelPendingIo();
        unregisterAllWatches();

        CLOSE_HANDLE(mDirectory);
        CLOSE_HANDLE(mCompletionPort);

        XPR_SAFE_DELETE(mOverlapped);
        // Unacknowledged cancellation records deliberately remain allocated.
        // The kernel may still write the terminal OVERLAPPED status after the
        // directory handle closes; releasing them here would be a UAF.  Normal
        // late completions are reclaimed by releaseRetiredOverlapped().
        mRetiredOverlappedList.clear();
    }

    AdvWatchId registerWatch(const xpr::string &aRootPath, AdvWatchItem *aNewAdvWatchItem)
    {
        if (aRootPath.empty() == XPR_TRUE || XPR_IS_NULL(aNewAdvWatchItem))
            return InvalidAdvWatchId;

        if (XPR_IS_NULL(mDirectory))
        {
            mRootPath = aRootPath;
            mWatchSubtree = aNewAdvWatchItem->mSubPath;

            mDirectory = ::CreateFile(
                mRootPath.c_str(),
                FILE_LIST_DIRECTORY,
                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                XPR_NULL,
                OPEN_EXISTING,
                FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED,
                XPR_NULL);

            if (mDirectory == INVALID_HANDLE_VALUE)
            {
                mDirectory = XPR_NULL;
                return InvalidAdvWatchId;
            }

            mCompletionPort = ::CreateIoCompletionPort(
                mDirectory,
                mCompletionPort,
                reinterpret_cast<ULONG_PTR>(this),
                0);

            if (XPR_IS_NULL(mCompletionPort))
            {
                CLOSE_HANDLE(mDirectory);
                return InvalidAdvWatchId;
            }
        }

        aNewAdvWatchItem->mAdvWatchId = (AdvWatchId)aNewAdvWatchItem;

        mWatchList.push_back(aNewAdvWatchItem);
        mIdWatchMap[aNewAdvWatchItem->mAdvWatchId] = aNewAdvWatchItem;

        // Upgrade a shared exact-directory request if a later subscriber asks
        // for its subtree.  cancelPendingIo drains the old completion packet
        // before the OVERLAPPED storage is reused.
        if (XPR_IS_TRUE(aNewAdvWatchItem->mSubPath) &&
            XPR_IS_FALSE(mWatchSubtree))
        {
            cancelPendingIo();
            mWatchSubtree = XPR_TRUE;
            readDirectoryChanges();
        }

        return aNewAdvWatchItem->mAdvWatchId;
    }

    void unregisterWatch(AdvWatchId aAdvWatchId)
    {
        IdWatchMap::iterator sIterator = mIdWatchMap.find(aAdvWatchId);
        if (sIterator == mIdWatchMap.end())
            return;

        AdvWatchItem *sAdvWatchItem = sIterator->second;

        mWatchList.remove(sAdvWatchItem);
        mIdWatchMap.erase(sIterator);

        XPR_SAFE_DELETE(sAdvWatchItem);
    }

    void unregisterAllWatches(void)
    {
        WatchList::iterator sIterator;
        AdvWatchItem *sAdvWatchItem;

        sIterator = mWatchList.begin();
        for (; sIterator != mWatchList.end(); ++sIterator)
        {
            sAdvWatchItem = *sIterator;
            XPR_SAFE_DELETE(sAdvWatchItem);
        }

        mWatchList.clear();
        mIdWatchMap.clear();
    }

    AdvWatchItem *getRegisteredItem(AdvWatchId aAdvWatchId)
    {
        IdWatchMap::iterator sIterator = mIdWatchMap.find(aAdvWatchId);
        if (sIterator == mIdWatchMap.end())
            return XPR_NULL;

        return sIterator->second;
    }

    xpr_size_t getRegisteredCount(void)
    {
        return mWatchList.size();
    }

    void cancelPendingIo(void)
    {
        if (XPR_IS_FALSE(mIoPending) || XPR_IS_NULL(mDirectory) ||
            XPR_IS_NULL(mOverlapped))
            return;

        // IOCP requests must be drained from the completion port.
        // GetOverlappedResult(..., TRUE) with hEvent == NULL waits on the
        // directory handle and can block forever even after CancelIoEx queued
        // ERROR_OPERATION_ABORTED.  That formed a self-deadlock while the UI
        // thread joined this monitor thread during process shutdown.
        ::CancelIoEx(mDirectory, mOverlapped);

        const ULONGLONG sDeadline = ::GetTickCount64() + kCancelCompletionTimeout;
        while (XPR_IS_TRUE(mIoPending) && XPR_IS_NOT_NULL(mCompletionPort))
        {
            DWORD sCompletedBytes = 0;
            ULONG_PTR sCompletionKey = 0;
            LPOVERLAPPED sCompletedOverlapped = XPR_NULL;
            const ULONGLONG sNow = ::GetTickCount64();
            const DWORD sWait = sNow < sDeadline
                ? static_cast<DWORD>(sDeadline - sNow)
                : 0;

            ::GetQueuedCompletionStatus(
                mCompletionPort, &sCompletedBytes, &sCompletionKey,
                &sCompletedOverlapped, sWait);

            if (sCompletedOverlapped == mOverlapped)
            {
                mIoPending = XPR_FALSE;
                break;
            }

            if (XPR_IS_NOT_NULL(sCompletedOverlapped))
                releaseRetiredOverlapped(sCompletedOverlapped);

            if (::GetTickCount64() >= sDeadline)
                break;
        }

        if (XPR_IS_TRUE(mIoPending))
        {
            // Never reuse or free storage the kernel may still own.  A late
            // packet is reclaimed while the watcher remains alive; otherwise
            // this bounded 64 KiB record is left to process teardown.
            mRetiredOverlappedList.push_back(mOverlapped);
            mOverlapped = XPR_NULL;
            mIoPending = XPR_FALSE;
        }
    }

    xpr_bool_t releaseRetiredOverlapped(LPOVERLAPPED aOverlapped)
    {
        RetiredOverlappedList::iterator sIterator =
            mRetiredOverlappedList.begin();
        for (; sIterator != mRetiredOverlappedList.end(); ++sIterator)
        {
            if (*sIterator != aOverlapped)
                continue;

            AdvWatchOverlapped *sRetiredOverlapped = *sIterator;
            mRetiredOverlappedList.erase(sIterator);
            XPR_SAFE_DELETE(sRetiredOverlapped);
            return XPR_TRUE;
        }

        return XPR_FALSE;
    }

    xpr_bool_t readDirectoryChanges(void)
    {
        if (XPR_IS_TRUE(mStopping) || XPR_IS_TRUE(mIoPending) ||
            XPR_IS_NULL(mDirectory))
            return XPR_FALSE;

        if (XPR_IS_NULL(mOverlapped))
        {
            mOverlapped = new AdvWatchOverlapped;
            if (XPR_IS_NULL(mOverlapped))
                return XPR_FALSE;

            ::ZeroMemory(static_cast<OVERLAPPED *>(mOverlapped), sizeof(OVERLAPPED));
            mOverlapped->hEvent          = XPR_NULL;
            mOverlapped->mDriveWatchItem = this;
            mOverlapped->mRootPath       = mRootPath;
        }
        else
        {
            ::ZeroMemory(static_cast<OVERLAPPED *>(mOverlapped), sizeof(OVERLAPPED));
        }

        // if notify filter contains 'FILE_NOTIFY_CHANGE_LAST_ACCESS',
        // then it don't contain because occur too many events.

        const DWORD sNotifyFilter = 
            FILE_NOTIFY_CHANGE_FILE_NAME   |
            FILE_NOTIFY_CHANGE_DIR_NAME    |
            FILE_NOTIFY_CHANGE_ATTRIBUTES  |
            FILE_NOTIFY_CHANGE_SIZE        |
            FILE_NOTIFY_CHANGE_LAST_WRITE  |
            //FILE_NOTIFY_CHANGE_LAST_ACCESS |
            FILE_NOTIFY_CHANGE_CREATION;//    |
        //FILE_NOTIFY_CHANGE_SECURITY;

        DWORD sBufferSize = static_cast<DWORD>(kBufferSize);

        xpr_bool_t sResult = ::ReadDirectoryChangesW(
            mDirectory,
            mOverlapped->mBuffer,
            sBufferSize, 
            mWatchSubtree,
            sNotifyFilter,
            XPR_NULL,
            mOverlapped,
            XPR_NULL);

        if (XPR_IS_TRUE(sResult))
            mIoPending = XPR_TRUE;

        return sResult;
    }

    xpr_bool_t getQueuedCompletionStatus(DWORD aMilliseconds,
                                         LPOVERLAPPED &aOverlapped,
                                         DWORD &aNumberOfBytes,
                                         DWORD &aError)
    {
        aOverlapped = XPR_NULL;
        aNumberOfBytes = 0;
        aError = ERROR_SUCCESS;

        ULONG_PTR sCompletionKey = 0;

        xpr_bool_t sResult = ::GetQueuedCompletionStatus(
            mCompletionPort,
            &aNumberOfBytes,
            &sCompletionKey,
            &aOverlapped,
            aMilliseconds);

        if (XPR_IS_FALSE(sResult))
        {
            aError = ::GetLastError();
            if (XPR_IS_NULL(aOverlapped))
                return XPR_FALSE;
        }

        if (aOverlapped == mOverlapped)
            mIoPending = XPR_FALSE;
        else if (releaseRetiredOverlapped(aOverlapped))
        {
            aOverlapped = XPR_NULL;
            return XPR_FALSE;
        }

        return XPR_IS_NOT_NULL(aOverlapped);
    }

    void extractFileName(const FILE_NOTIFY_INFORMATION *aFileNotifyInfo, xpr::string &aFileName)
    {
        xpr_tchar_t sFileName[XPR_MAX_PATH + 1] = {0};
        // FileNameLength is already expressed in bytes (and is not NUL
        // terminated).  Multiplying it by sizeof(wchar_t) read past the IOCP
        // completion buffer on every notification.
        xpr_size_t sInputBytes = aFileNotifyInfo->FileNameLength;
        xpr_size_t sOutputBytes = XPR_MAX_PATH * sizeof(xpr_tchar_t);
        XPR_UTF16_TO_TCS(aFileNotifyInfo->FileName, &sInputBytes, sFileName, &sOutputBytes);
        const xpr_size_t sOutputChars =
            (std::min)(sOutputBytes / sizeof(xpr_tchar_t),
                       static_cast<xpr_size_t>(XPR_MAX_PATH));
        sFileName[sOutputChars] = 0;

        aFileName = sFileName;
    }

    void appendUpdateDirNotifies(NotifyList &aNotifyList)
    {
        const xpr_sint64_t sNow = xpr::timer_ms();
        WatchList::const_iterator sIterator = mWatchList.begin();
        for (; sIterator != mWatchList.end(); ++sIterator)
        {
            AdvWatchItem *sWatchItem = *sIterator;
            if (XPR_IS_NULL(sWatchItem))
                continue;

            NotifyInfo *sNotifyInfo = new NotifyInfo;
            if (XPR_IS_NULL(sNotifyInfo))
                continue;

            sNotifyInfo->mHwnd        = sWatchItem->mHwnd;
            sNotifyInfo->mMsg         = sWatchItem->mMsg;
            sNotifyInfo->mAdvWatchId  = sWatchItem->mAdvWatchId;
            sNotifyInfo->mEvent       = EventUpdateDir;
            sNotifyInfo->mDir         = sWatchItem->mPath;
            sNotifyInfo->mNotifyCount = 1;
            sNotifyInfo->mTime        = sNow;
            aNotifyList.push_back(sNotifyInfo);
        }
    }

    void OnCompletionRoutine(LPOVERLAPPED aOverlapped,
                             DWORD aNumberOfBytes,
                             DWORD aError,
                             NotifyList &aNotifyList)
    {
        AdvFileChangeWatcher::AdvWatchOverlapped *sAdvWatchOverlapped =
            reinterpret_cast<AdvFileChangeWatcher::AdvWatchOverlapped *>(aOverlapped);
        if (sAdvWatchOverlapped != mOverlapped)
            return;

        if (XPR_IS_TRUE(mStopping) || aError == ERROR_OPERATION_ABORTED)
            return;

        NotifyList sParsedNotifies;
        xpr_bool_t sValid = XPR_TRUE;

        // Zero bytes is the documented overflow/lost-change signal.  Never
        // inspect stale buffer contents in that case.
        if (aError != ERROR_SUCCESS || aNumberOfBytes == 0 ||
            aNumberOfBytes > kBufferSize)
        {
            sValid = XPR_FALSE;
        }
        else
        {
            xpr_size_t sOffset = 0;
            while (sOffset < aNumberOfBytes)
            {
                const xpr_size_t sRemaining = aNumberOfBytes - sOffset;
                if (sRemaining < kFileNotifyHeaderSize)
                {
                    sValid = XPR_FALSE;
                    break;
                }

                FILE_NOTIFY_INFORMATION *sFileNotifyInformation =
                    reinterpret_cast<PFILE_NOTIFY_INFORMATION>(
                        &sAdvWatchOverlapped->mBuffer[sOffset]);

                if ((sFileNotifyInformation->FileNameLength % sizeof(WCHAR)) != 0 ||
                    sFileNotifyInformation->FileNameLength >
                        sRemaining - kFileNotifyHeaderSize)
                {
                    sValid = XPR_FALSE;
                    break;
                }

                OnFileChanged(sAdvWatchOverlapped,
                              sFileNotifyInformation,
                              sParsedNotifies);

                const DWORD sNextOffset = sFileNotifyInformation->NextEntryOffset;
                if (sNextOffset == 0)
                    break;

                const xpr_size_t sEntryBytes =
                    kFileNotifyHeaderSize +
                    sFileNotifyInformation->FileNameLength;
                if (sNextOffset < sEntryBytes ||
                    (sNextOffset % sizeof(DWORD)) != 0 ||
                    sNextOffset > sRemaining)
                {
                    sValid = XPR_FALSE;
                    break;
                }

                sOffset += sNextOffset;
            }
        }

        if (XPR_IS_FALSE(sValid))
        {
            while (sParsedNotifies.empty() == false)
            {
                NotifyInfo *sNotifyInfo = sParsedNotifies.front();
                sParsedNotifies.pop_front();
                XPR_SAFE_DELETE(sNotifyInfo);
            }
            mOverlapped->mOldFileName.clear();
            appendUpdateDirNotifies(aNotifyList);
        }
        else
        {
            aNotifyList.splice(aNotifyList.end(), sParsedNotifies);
        }

        // Rearm only after the completed buffer has been fully consumed.  One
        // DriveWatchItem owns exactly one in-flight OVERLAPPED request.
        if (XPR_IS_FALSE(mStopping) && XPR_IS_FALSE(readDirectoryChanges()))
            appendUpdateDirNotifies(aNotifyList);
    }

    void OnFileChanged(AdvWatchOverlapped *aAdvWatchOverlapped, FILE_NOTIFY_INFORMATION *aFileNotifyInformation, NotifyList &aNotifyList)
    {
        if (XPR_IS_NULL(aAdvWatchOverlapped) || XPR_IS_NULL(aFileNotifyInformation))
            return;

        if (aFileNotifyInformation->Action == FILE_ACTION_RENAMED_OLD_NAME)
        {
            extractFileName(aFileNotifyInformation, aAdvWatchOverlapped->mOldFileName);
            return;
        }

        Event sEvent = EventNone;

        switch (aFileNotifyInformation->Action)
        {
        case FILE_ACTION_ADDED:            sEvent = EventCreated;  break;
        case FILE_ACTION_MODIFIED:         sEvent = EventModified; break;
        case FILE_ACTION_REMOVED:          sEvent = EventDeleted;  break;
        case FILE_ACTION_RENAMED_NEW_NAME: sEvent = EventRenamed;  break;

        default:
            return;
        }

        const xpr::string &sRootPath = aAdvWatchOverlapped->mRootPath;

        xpr::string sFileName;
        extractFileName(aFileNotifyInformation, sFileName);

        xpr::string sDir = sRootPath;
        xpr_size_t sFind = sFileName.rfind(XPR_STRING_LITERAL('\\'));
        if (sFind != xpr::string::npos)
        {
            if (sDir.empty() == false &&
                *sDir.rbegin() != XPR_STRING_LITERAL('\\'))
                sDir += XPR_STRING_LITERAL('\\');
            sDir += sFileName.substr(0, sFind);
            sFileName.erase(0, sFind+1);
        }

        DriveWatchItem *sDriveWatchItem = aAdvWatchOverlapped->mDriveWatchItem;
        if (XPR_IS_NULL(sDriveWatchItem))
            return;

        xpr::string sOldDir = sRootPath;
        xpr::string sOldFileName;
        if (sEvent == EventRenamed)
        {
            sFind = aAdvWatchOverlapped->mOldFileName.rfind(XPR_STRING_LITERAL('\\'));
            if (sFind != xpr::string::npos)
            {
                if (sOldDir.empty() == false &&
                    *sOldDir.rbegin() != XPR_STRING_LITERAL('\\'))
                    sOldDir += XPR_STRING_LITERAL('\\');
                sOldDir += aAdvWatchOverlapped->mOldFileName.substr(0, sFind);
                sOldFileName = aAdvWatchOverlapped->mOldFileName.substr(sFind+1);
            }
            else
            {
                sOldFileName = aAdvWatchOverlapped->mOldFileName;
            }
        }

        const xpr_sint64_t sNow = xpr::timer_ms();
        DriveWatchItem::WatchList::iterator sIterator =
            sDriveWatchItem->mWatchList.begin();
        for (; sIterator != sDriveWatchItem->mWatchList.end(); ++sIterator)
        {
            AdvWatchItem *sTargetAdvWatchItem = *sIterator;
            if (XPR_IS_NULL(sTargetAdvWatchItem))
                continue;

            const xpr_size_t sWatchLength = sTargetAdvWatchItem->mPath.length();
            const xpr_bool_t sExactDirectory =
                (_tcsicmp(sDir.c_str(), sTargetAdvWatchItem->mPath.c_str()) == 0) ?
                XPR_TRUE : XPR_FALSE;
            const xpr_bool_t sSubDirectory =
                (XPR_IS_TRUE(sTargetAdvWatchItem->mSubPath) &&
                 sDir.length() > sWatchLength &&
                 _tcsnicmp(sDir.c_str(), sTargetAdvWatchItem->mPath.c_str(),
                           sWatchLength) == 0 &&
                 sDir[sWatchLength] == XPR_STRING_LITERAL('\\')) ?
                XPR_TRUE : XPR_FALSE;
            if (XPR_IS_FALSE(sExactDirectory) && XPR_IS_FALSE(sSubDirectory))
                continue;

            NotifyInfo *sNotifyInfo = new NotifyInfo;
            if (XPR_IS_NULL(sNotifyInfo))
                continue;

            sNotifyInfo->mHwnd        = sTargetAdvWatchItem->mHwnd;
            sNotifyInfo->mMsg         = sTargetAdvWatchItem->mMsg;
            sNotifyInfo->mAdvWatchId  = sTargetAdvWatchItem->mAdvWatchId;
            sNotifyInfo->mEvent       = sEvent;
            sNotifyInfo->mDir         = sDir;
            sNotifyInfo->mFileName    = sFileName;
            sNotifyInfo->mTime        = sNow;
            sNotifyInfo->mNotifyCount = 1;

            if (sEvent == EventRenamed)
            {
                // A missing OLD_NAME means the pair was split/lost.  A single
                // bounded refresh is safer than inventing an old path.
                if (sOldFileName.empty())
                {
                    sNotifyInfo->mEvent = EventUpdateDir;
                    sNotifyInfo->mFileName.clear();
                }
                else
                {
                    sNotifyInfo->mOldDir = sOldDir;
                    sNotifyInfo->mOldFileName = sOldFileName;
                }
            }

            aNotifyList.push_back(sNotifyInfo);
        }

        aAdvWatchOverlapped->mOldFileName.clear();
    }

public:
    xpr::string         mRootPath;
    HANDLE              mDirectory;
    HANDLE              mCompletionPort;
    AdvWatchOverlapped *mOverlapped;
    xpr_bool_t          mWatchSubtree;
    xpr_bool_t          mIoPending;
    xpr_bool_t          mStopping;
    RetiredOverlappedList mRetiredOverlappedList;

    typedef std::list<AdvWatchItem *> WatchList;
    typedef std::map<AdvWatchId, AdvWatchItem *> IdWatchMap;

    WatchList           mWatchList;
    IdWatchMap          mIdWatchMap;
};

//
// AdvFileChangeWatcher
//
AdvFileChangeWatcher::AdvFileChangeWatcher(void)
    : mTaskEvent(XPR_NULL)
    , mTaskAddEvent(XPR_NULL)
    , mNotifyEvent(XPR_NULL)
{
}

AdvFileChangeWatcher::~AdvFileChangeWatcher(void)
{
    destroy();

    CLOSE_HANDLE(mTaskEvent);
    CLOSE_HANDLE(mTaskAddEvent);
    CLOSE_HANDLE(mNotifyEvent);
}

xpr_bool_t AdvFileChangeWatcher::create(void)
{
    CLOSE_HANDLE(mTaskEvent);
    CLOSE_HANDLE(mTaskAddEvent);
    CLOSE_HANDLE(mNotifyEvent);

    mTaskEvent    = ::CreateEvent(XPR_NULL, XPR_TRUE, XPR_FALSE, XPR_NULL);
    mTaskAddEvent = ::CreateEvent(XPR_NULL, XPR_TRUE, XPR_FALSE, XPR_NULL);
    mNotifyEvent  = ::CreateEvent(XPR_NULL, XPR_TRUE, XPR_FALSE, XPR_NULL);

    if (XPR_IS_NULL(mTaskEvent) || XPR_IS_NULL(mTaskAddEvent) || XPR_IS_NULL(mNotifyEvent))
    {
        destroy();
        return XPR_FALSE;
    }

    xpr_rcode_t sRcode;

    sRcode = mMonitorThread.start(dynamic_cast<xpr::Thread::Runnable *>(this));
    if (XPR_RCODE_IS_NOT_SUCCESS(sRcode))
    {
        destroy();
        return XPR_FALSE;
    }

    sRcode = mNotifyThread.start(dynamic_cast<xpr::Thread::Runnable *>(this));
    if (XPR_RCODE_IS_NOT_SUCCESS(sRcode))
    {
        destroy();
        return XPR_FALSE;
    }

    return XPR_TRUE;
}

void AdvFileChangeWatcher::destroy(void)
{
    // Stop the producer first.  Wake both finite-wait loops so shutdown does
    // not depend on their polling interval.  DriveWatchItem destruction on the
    // monitor thread cancels and drains its own OVERLAPPED request.
    mMonitorThread.stop();
    if (XPR_IS_NOT_NULL(mTaskEvent))
        ::SetEvent(mTaskEvent);
    mMonitorThread.join();

    mNotifyThread.stop();
    if (XPR_IS_NOT_NULL(mNotifyEvent))
        ::SetEvent(mNotifyEvent);
    mNotifyThread.join();

    clearTasks(mTaskList);
    clearNotifies(mNotifyList);
    clearNotifies(mNotifyMap);

    CLOSE_HANDLE(mTaskEvent);
    CLOSE_HANDLE(mTaskAddEvent);
    CLOSE_HANDLE(mNotifyEvent);
}

AdvFileChangeWatcher::AdvWatchId AdvFileChangeWatcher::registerWatch(AdvWatchItem *aAdvWatchItem)
{
    if (XPR_IS_NULL(aAdvWatchItem))
        return InvalidAdvWatchId;

    if (aAdvWatchItem->isValidate() == XPR_FALSE)
        return InvalidAdvWatchId;

    if (mMonitorThread.isRunning() == XPR_FALSE && create() == XPR_FALSE)
        return InvalidAdvWatchId;

    {
        xpr::MutexGuard sLockGuard(mTaskMutex);

        Task *sTask = new Task;
        sTask->mTaskCommand  = TaskCommandRegister;
        sTask->mAdvWatchItem = aAdvWatchItem;

        aAdvWatchItem->mAdvWatchId = (AdvWatchId)aAdvWatchItem;

        mTaskList.push_back(sTask);

        ::SetEvent(mTaskEvent);
    }

    return aAdvWatchItem->mAdvWatchId;
}

AdvFileChangeWatcher::AdvWatchId AdvFileChangeWatcher::modifyWatch(AdvWatchId aOldAdvWatchId, AdvWatchItem *aAdvWatchItem)
{
    if (XPR_IS_NULL(aAdvWatchItem))
        return InvalidAdvWatchId;

    if (aOldAdvWatchId == InvalidAdvWatchId)
        return registerWatch(aAdvWatchItem);

    if (aAdvWatchItem->isValidate() == XPR_FALSE)
    {
        unregisterWatch(aOldAdvWatchId);
        return InvalidAdvWatchId;
    }

    if (mMonitorThread.isRunning() == XPR_FALSE)
        return InvalidAdvWatchId;

    {
        xpr::MutexGuard sLockGuard(mTaskMutex);

        aAdvWatchItem->mAdvWatchId = (AdvWatchId)aAdvWatchItem;

        // Navigation can replace a pane's folder faster than the monitor
        // thread consumes commands.  Collapse a not-yet-processed A->B->C
        // chain into A->C instead of growing an unbounded task list.
        TaskList::reverse_iterator sPendingIterator = mTaskList.rbegin();
        for (; sPendingIterator != mTaskList.rend(); ++sPendingIterator)
        {
            Task *sPendingTask = *sPendingIterator;
            if (XPR_IS_NOT_NULL(sPendingTask) &&
                XPR_IS_NOT_NULL(sPendingTask->mAdvWatchItem) &&
                sPendingTask->mAdvWatchItem->mAdvWatchId == aOldAdvWatchId &&
                (sPendingTask->mTaskCommand == TaskCommandRegister ||
                 sPendingTask->mTaskCommand == TaskCommandModify))
            {
                XPR_SAFE_DELETE(sPendingTask->mAdvWatchItem);
                sPendingTask->mAdvWatchItem = aAdvWatchItem;
                ::SetEvent(mTaskEvent);
                return aAdvWatchItem->mAdvWatchId;
            }
        }

        Task *sTask = new Task;
        sTask->mTaskCommand   = TaskCommandModify;
        sTask->mOldAdvWatchId = aOldAdvWatchId;
        sTask->mAdvWatchItem  = aAdvWatchItem;

        mTaskList.push_back(sTask);

        ::SetEvent(mTaskEvent);
    }

    return aAdvWatchItem->mAdvWatchId;
}

void AdvFileChangeWatcher::unregisterWatch(AdvWatchId aAdvWatchId)
{
    if (aAdvWatchId == InvalidAdvWatchId)
        return;

    if (mMonitorThread.isRunning() == XPR_FALSE)
        return;

    {
        xpr::MutexGuard sLockGuard(mTaskMutex);

        TaskList::iterator sPendingIterator = mTaskList.begin();
        for (; sPendingIterator != mTaskList.end(); ++sPendingIterator)
        {
            Task *sPendingTask = *sPendingIterator;
            if (XPR_IS_NULL(sPendingTask) ||
                XPR_IS_NULL(sPendingTask->mAdvWatchItem) ||
                sPendingTask->mAdvWatchItem->mAdvWatchId != aAdvWatchId)
                continue;

            if (sPendingTask->mTaskCommand == TaskCommandRegister)
            {
                XPR_SAFE_DELETE(sPendingTask);
                mTaskList.erase(sPendingIterator);
                ::SetEvent(mTaskEvent);
                return;
            }

            if (sPendingTask->mTaskCommand == TaskCommandModify)
            {
                XPR_SAFE_DELETE(sPendingTask->mAdvWatchItem);
                sPendingTask->mTaskCommand = TaskCommandUnregister;
                ::SetEvent(mTaskEvent);
                return;
            }
        }

        Task *sTask = new Task;
        sTask->mTaskCommand   = TaskCommandUnregister;
        sTask->mOldAdvWatchId = aAdvWatchId;

        mTaskList.push_back(sTask);

        ::SetEvent(mTaskEvent);
    }
}

void AdvFileChangeWatcher::unregisterAllWatches(void)
{
    if (!mMonitorThread.isRunning())
        return;

    {
        xpr::MutexGuard sLockGuard(mTaskMutex);

        clearTasks(mTaskList);

        Task *sTask = new Task;
        sTask->mTaskCommand = TaskCommandUnregisterAll;

        mTaskList.push_back(sTask);

        ::SetEvent(mTaskEvent);
    }
}

xpr_bool_t AdvFileChangeWatcher::getRootPath(const xpr::string &aPath, xpr::string &aRootPath)
{
    aRootPath.clear();
    if (aPath.empty())
        return XPR_FALSE;

    // The legacy watcher opened C:\ or the server root recursively and then
    // filtered every event in user space.  Long sessions therefore accumulated
    // unrelated activity from the entire drive.  Each DriveWatchItem now owns
    // the exact directory requested by an ExplorerCtrl pane.
    aRootPath = aPath;
    while (aRootPath.length() > 3 &&
           *aRootPath.rbegin() == XPR_STRING_LITERAL('\\'))
        aRootPath.erase(aRootPath.length() - 1);

    return XPR_TRUE;
}

xpr_sint_t AdvFileChangeWatcher::runThread(xpr::Thread &aThread)
{
    Thread &sThread = (Thread &)aThread;

    if (aThread == mMonitorThread)
        return runMonitorThread(sThread);

    if (aThread == mNotifyThread)
        return runNotifyThread(sThread);

    return 0;
}

xpr_sint_t AdvFileChangeWatcher::runMonitorThread(Thread &aThread)
{
    DriveWatchMap::iterator sIterator;
    DriveWatchItem *sDriveWatchItem;
    LPOVERLAPPED sOverlapped;
    DWORD sNumberOfBytes;
    DWORD sCompletionError;
    TaskList sTaskList;
    NotifyList sNotifyList;

    DWORD sResult;

    while (mMonitorThread.isStop() == XPR_FALSE)
    {
        sResult = ::WaitForSingleObject(mTaskEvent, kIdleTime);
        if (mMonitorThread.isStop() == XPR_TRUE)
            break;

        if (sResult == WAIT_OBJECT_0)
        {
            {
                xpr::MutexGuard sLockGuard(mTaskMutex);

                mTaskList.swap(sTaskList);
                ::ResetEvent(mTaskEvent);
            }

            if (sTaskList.empty() == false)
            {
                processTasks(sTaskList);
                clearTasks(sTaskList);
            }
        }
        else if (sResult == WAIT_TIMEOUT)
        {
            sIterator = mDriveWatchMap.begin();
            for (; sIterator != mDriveWatchMap.end(); ++sIterator)
            {
                if (mMonitorThread.isStop() == XPR_TRUE)
                    break;

                sDriveWatchItem = sIterator->second;
                if (XPR_IS_NULL(sDriveWatchItem))
                    continue;

                if (::WaitForSingleObject(mTaskEvent, 0) == WAIT_OBJECT_0)
                    break;

                if (XPR_IS_TRUE(sDriveWatchItem->getQueuedCompletionStatus(
                        10, sOverlapped, sNumberOfBytes, sCompletionError)))
                {
                    sDriveWatchItem->OnCompletionRoutine(
                        sOverlapped, sNumberOfBytes, sCompletionError,
                        sNotifyList);

                    if (sNotifyList.empty() == false)
                        queueNotifies(sNotifyList);
                }
            }
        }
    }

    unregisterAllTasks();

    {
        xpr::MutexGuard sLockGuard(mTaskMutex);
        clearTasks(sTaskList);
    }

    {
        xpr::MutexGuard sLockGuard(mNotifyMutex);
        clearNotifies(mNotifyList);
    }

    return 0;
}

void AdvFileChangeWatcher::processTasks(TaskList &aTaskList)
{
    TaskList::iterator sIterator;

    sIterator = aTaskList.begin();
    for (; sIterator != aTaskList.end(); ++sIterator)
    {
        Task *sTask = *sIterator;

        switch (sTask->mTaskCommand)
        {
        case TaskCommandRegister:      registerTask(*sTask);   break;
        case TaskCommandModify:        modifyTask(*sTask);     break;
        case TaskCommandUnregister:    unregisterTask(*sTask); break;
        case TaskCommandUnregisterAll: unregisterAllTasks();   break;
        }
    }
}

void AdvFileChangeWatcher::registerTask(Task &aTask)
{
    xpr::string sRootPath;
    if (!getRootPath(aTask.mAdvWatchItem->mPath, sRootPath))
        return;

    xpr_bool_t sResult = XPR_FALSE;
    DriveWatchItem *sDriveWatchItem = XPR_NULL;

    DriveWatchMap::iterator sIterator = mDriveWatchMap.find(sRootPath);
    if (sIterator == mDriveWatchMap.end())
    {
        sDriveWatchItem = new DriveWatchItem;
        if (XPR_IS_NOT_NULL(sDriveWatchItem))
            mDriveWatchMap[sRootPath] = sDriveWatchItem;
    }
    else
    {
        sDriveWatchItem = sIterator->second;
        if (XPR_IS_NULL(sDriveWatchItem))
            mDriveWatchMap.erase(sIterator);
    }

    if (XPR_IS_NOT_NULL(sDriveWatchItem))
    {
        AdvWatchId sNewAdvWatchId;
        sNewAdvWatchId = sDriveWatchItem->registerWatch(sRootPath, aTask.mAdvWatchItem);
        if (sNewAdvWatchId != InvalidAdvWatchId)
        {
            if (sDriveWatchItem->getRegisteredCount() == 1)
            {
                sDriveWatchItem->readDirectoryChanges();
            }

            sResult = XPR_TRUE;
        }
        else
        {
            XPR_SAFE_DELETE(sDriveWatchItem);
            mDriveWatchMap.erase(sRootPath);
            return;
        }
    }

    if (XPR_IS_FALSE(sResult))
        return;

    if (XPR_IS_NOT_NULL(sDriveWatchItem))
        mIdDriveWatchMap[aTask.mAdvWatchItem->mAdvWatchId] = sDriveWatchItem;

    aTask.mAdvWatchItem = XPR_NULL;
}

void AdvFileChangeWatcher::modifyTask(Task &aTask)
{
    DriveWatchItem *sOldDriveWatchItem = XPR_NULL;
    DriveWatchItem *sNewDriveWatchItem = XPR_NULL;

    IdDriveWatchMap::iterator sOldIdIterator = mIdDriveWatchMap.find(aTask.mOldAdvWatchId);
    if (sOldIdIterator != mIdDriveWatchMap.end())
        sOldDriveWatchItem = sOldIdIterator->second;

    AdvWatchId sNewAdvWatchId = InvalidAdvWatchId;

    xpr_bool_t sResult = aTask.mAdvWatchItem->isValidate();

    // compare
    if (XPR_IS_NOT_NULL(sOldDriveWatchItem))
    {
        AdvWatchItem *sOldAdvWatchItem = sOldDriveWatchItem->getRegisteredItem(aTask.mOldAdvWatchId);
        if (XPR_IS_NOT_NULL(sOldAdvWatchItem) && XPR_IS_NOT_NULL(aTask.mAdvWatchItem))
        {
            if (sOldAdvWatchItem->mPath == aTask.mAdvWatchItem->mPath)
            {
                if (sOldAdvWatchItem->mAdvWatchId != aTask.mAdvWatchItem->mAdvWatchId)
                {
                    // new item change
                    sOldDriveWatchItem->unregisterWatch(aTask.mOldAdvWatchId);
                    mIdDriveWatchMap.erase(sOldIdIterator);

                    xpr::string sRootPath;
                    if (getRootPath(aTask.mAdvWatchItem->mPath, sRootPath))
                    {
                        sOldDriveWatchItem->registerWatch(sRootPath, aTask.mAdvWatchItem);

                        mIdDriveWatchMap[aTask.mAdvWatchItem->mAdvWatchId] = sOldDriveWatchItem;

                        aTask.mAdvWatchItem = XPR_NULL;
                    }
                }

                return;
            }
        }
    }

    // old
    if (XPR_IS_NOT_NULL(sOldDriveWatchItem))
    {
        sOldDriveWatchItem->unregisterWatch(aTask.mOldAdvWatchId);
    }

    // new
    {
        xpr::string sRootPath;
        if (getRootPath(aTask.mAdvWatchItem->mPath, sRootPath))
        {
            DriveWatchMap::iterator sIterator = mDriveWatchMap.find(sRootPath);
            if (sIterator == mDriveWatchMap.end())
            {
                sNewDriveWatchItem = new DriveWatchItem;
                if (XPR_IS_NOT_NULL(sNewDriveWatchItem))
                    mDriveWatchMap[sRootPath] = sNewDriveWatchItem;
            }
            else
            {
                sNewDriveWatchItem = sIterator->second;
                if (XPR_IS_NULL(sNewDriveWatchItem))
                    mDriveWatchMap.erase(sIterator);
            }

            if (XPR_IS_NOT_NULL(sNewDriveWatchItem))
                sNewAdvWatchId = sNewDriveWatchItem->registerWatch(sRootPath, aTask.mAdvWatchItem);
            if (sNewAdvWatchId != InvalidAdvWatchId)
            {
                if (sNewDriveWatchItem->getRegisteredCount() == 1)
                {
                    sNewDriveWatchItem->readDirectoryChanges();
                }
            }
            else
            {
                XPR_SAFE_DELETE(sNewDriveWatchItem);
                mDriveWatchMap.erase(sRootPath);
            }
        }
    }

    if (XPR_IS_NOT_NULL(sOldDriveWatchItem))
    {
        xpr_size_t nOldRegisteredCount = sOldDriveWatchItem->getRegisteredCount();
        if (nOldRegisteredCount == 0)
        {
            mDriveWatchMap.erase(sOldDriveWatchItem->mRootPath);
            XPR_SAFE_DELETE(sOldDriveWatchItem);
        }
    }

    if (sOldIdIterator != mIdDriveWatchMap.end())
        mIdDriveWatchMap.erase(sOldIdIterator);

    if (XPR_IS_NOT_NULL(sNewDriveWatchItem) && sNewAdvWatchId != InvalidAdvWatchId)
        mIdDriveWatchMap[aTask.mAdvWatchItem->mAdvWatchId] = sNewDriveWatchItem;

    if (sNewAdvWatchId != InvalidAdvWatchId)
        aTask.mAdvWatchItem = XPR_NULL;
}

void AdvFileChangeWatcher::unregisterTask(Task &aTask)
{
    IdDriveWatchMap::iterator sIterator = mIdDriveWatchMap.find(aTask.mOldAdvWatchId);
    if (sIterator == mIdDriveWatchMap.end())
        return;

    DriveWatchItem *sDriveWatchItem = sIterator->second;
    if (XPR_IS_NULL(sDriveWatchItem))
    {
        mIdDriveWatchMap.erase(sIterator);
        return;
    }

    sDriveWatchItem->unregisterWatch(aTask.mOldAdvWatchId);

    xpr_size_t sRegisteredCount = sDriveWatchItem->getRegisteredCount();
    if (sRegisteredCount == 0)
    {
        mDriveWatchMap.erase(sDriveWatchItem->mRootPath);
        XPR_SAFE_DELETE(sDriveWatchItem);
    }

    mIdDriveWatchMap.erase(sIterator);
}

void AdvFileChangeWatcher::unregisterAllTasks(void)
{
    DriveWatchMap::iterator sIterator;
    DriveWatchItem *sDriveWatchItem;

    sIterator = mDriveWatchMap.begin();
    for (; sIterator != mDriveWatchMap.end(); ++sIterator)
    {
        sDriveWatchItem = sIterator->second;
        XPR_SAFE_DELETE(sDriveWatchItem);
    }

    mDriveWatchMap.clear();
    mIdDriveWatchMap.clear();
}

void AdvFileChangeWatcher::clearTasks(TaskList &aTaskList)
{
    Task *sTask;
    TaskList::iterator sIterator;

    sIterator = aTaskList.begin();
    for (; sIterator != aTaskList.end(); ++sIterator)
    {
        sTask = *sIterator;
        XPR_SAFE_DELETE(sTask);
    }

    aTaskList.clear();
}

void AdvFileChangeWatcher::queueNotifies(NotifyList &aNotifyList)
{
    if (aNotifyList.empty())
        return;

    xpr::MutexGuard sLockGuard(mNotifyMutex);

    while (aNotifyList.empty() == false)
    {
        NotifyInfo *sNotifyInfo = aNotifyList.front();
        aNotifyList.pop_front();
        if (XPR_IS_NULL(sNotifyInfo))
            continue;

        xpr_bool_t sMerged = XPR_FALSE;

        // Repeated writes to the same file are common (cloud sync, AV and
        // editors).  They carry no additional UI information within one burst.
        if (sNotifyInfo->mEvent == EventModified)
        {
            NotifyList::reverse_iterator sIterator = mNotifyList.rbegin();
            for (; sIterator != mNotifyList.rend(); ++sIterator)
            {
                NotifyInfo *sPending = *sIterator;
                if (XPR_IS_NOT_NULL(sPending) &&
                    sPending->mAdvWatchId == sNotifyInfo->mAdvWatchId &&
                    sPending->mEvent == EventModified &&
                    _tcsicmp(sPending->mDir.c_str(), sNotifyInfo->mDir.c_str()) == 0 &&
                    _tcsicmp(sPending->mFileName.c_str(),
                             sNotifyInfo->mFileName.c_str()) == 0)
                {
                    sPending->mNotifyCount += sNotifyInfo->mNotifyCount;
                    XPR_SAFE_DELETE(sNotifyInfo);
                    sMerged = XPR_TRUE;
                    break;
                }
            }
        }

        if (XPR_IS_TRUE(sMerged))
            continue;

        // An overflow summary supersedes all not-yet-posted exact events for
        // the same pane.  Likewise, once a summary exists, later raw events are
        // folded into it until the notify thread posts it.
        NotifyList::iterator sIterator = mNotifyList.begin();
        while (sIterator != mNotifyList.end())
        {
            NotifyInfo *sPending = *sIterator;
            if (XPR_IS_NULL(sPending) ||
                sPending->mAdvWatchId != sNotifyInfo->mAdvWatchId)
            {
                ++sIterator;
                continue;
            }

            if (sPending->mEvent == EventUpdateDir)
            {
                sPending->mNotifyCount += sNotifyInfo->mNotifyCount;
                XPR_SAFE_DELETE(sNotifyInfo);
                sMerged = XPR_TRUE;
                break;
            }

            if (sNotifyInfo->mEvent == EventUpdateDir ||
                mNotifyList.size() >= kMaxRawNotifyQueue)
            {
                sNotifyInfo->mEvent = EventUpdateDir;
                sNotifyInfo->mFileName.clear();
                sNotifyInfo->mOldDir.clear();
                sNotifyInfo->mOldFileName.clear();
                sNotifyInfo->mNotifyCount += sPending->mNotifyCount;
                XPR_SAFE_DELETE(sPending);
                sIterator = mNotifyList.erase(sIterator);
                continue;
            }

            ++sIterator;
        }

        if (XPR_IS_TRUE(sMerged))
            continue;

        if (mNotifyList.size() >= kMaxRawNotifyQueue)
        {
            sNotifyInfo->mEvent = EventUpdateDir;
            sNotifyInfo->mFileName.clear();
            sNotifyInfo->mOldDir.clear();
            sNotifyInfo->mOldFileName.clear();

            // The queue can only remain full here if it consists of other
            // watchers.  Collapse the oldest watcher's records too; silently
            // dropping one exact event would leave a stale row indefinitely.
            NotifyInfo *sOldest = mNotifyList.front();
            mNotifyList.pop_front();
            if (XPR_IS_NOT_NULL(sOldest))
            {
                NotifyList::iterator sOldIterator = mNotifyList.begin();
                while (sOldIterator != mNotifyList.end())
                {
                    NotifyInfo *sOldPending = *sOldIterator;
                    if (XPR_IS_NOT_NULL(sOldPending) &&
                        sOldPending->mAdvWatchId == sOldest->mAdvWatchId)
                    {
                        sOldest->mNotifyCount += sOldPending->mNotifyCount;
                        XPR_SAFE_DELETE(sOldPending);
                        sOldIterator = mNotifyList.erase(sOldIterator);
                    }
                    else
                    {
                        ++sOldIterator;
                    }
                }
                sOldest->mEvent = EventUpdateDir;
                sOldest->mFileName.clear();
                sOldest->mOldDir.clear();
                sOldest->mOldFileName.clear();
                mNotifyList.push_front(sOldest);
            }
        }

        mNotifyList.push_back(sNotifyInfo);
    }

    ::SetEvent(mNotifyEvent);
}

xpr_sint_t AdvFileChangeWatcher::runNotifyThread(Thread &aThread)
{
    NotifyList sNotifyList;

    while (mNotifyThread.isStop() == XPR_FALSE)
    {
        const DWORD sResult = ::WaitForSingleObject(mNotifyEvent, kIdleTime);
        if (mNotifyThread.isStop() == XPR_TRUE)
            break;

        if (sResult == WAIT_OBJECT_0)
        {
            {
                xpr::MutexGuard sLockGuard(mNotifyMutex);

                sNotifyList.swap(mNotifyList);
                ::ResetEvent(mNotifyEvent);
            }

            while (sNotifyList.empty() == false)
            {
                NotifyInfo *sNotifyInfo = sNotifyList.front();
                sNotifyList.pop_front();
                if (XPR_IS_NULL(sNotifyInfo))
                    continue;

                NotifyList &sPending = mNotifyMap[sNotifyInfo->mAdvWatchId];
                xpr_bool_t sMerged = XPR_FALSE;

                if (sPending.empty() == false &&
                    sPending.front()->mEvent == EventUpdateDir)
                {
                    sPending.front()->mNotifyCount += sNotifyInfo->mNotifyCount;
                    XPR_SAFE_DELETE(sNotifyInfo);
                    continue;
                }

                if (sNotifyInfo->mEvent == EventUpdateDir)
                {
                    while (sPending.empty() == false)
                    {
                        NotifyInfo *sOld = sPending.front();
                        sPending.pop_front();
                        sNotifyInfo->mNotifyCount += sOld->mNotifyCount;
                        XPR_SAFE_DELETE(sOld);
                    }
                    sPending.push_back(sNotifyInfo);
                    continue;
                }

                if (sNotifyInfo->mEvent == EventModified)
                {
                    NotifyList::reverse_iterator sIterator = sPending.rbegin();
                    for (; sIterator != sPending.rend(); ++sIterator)
                    {
                        NotifyInfo *sOld = *sIterator;
                        if (XPR_IS_NOT_NULL(sOld) &&
                            sOld->mEvent == EventModified &&
                            _tcsicmp(sOld->mDir.c_str(),
                                     sNotifyInfo->mDir.c_str()) == 0 &&
                            _tcsicmp(sOld->mFileName.c_str(),
                                     sNotifyInfo->mFileName.c_str()) == 0)
                        {
                            sOld->mNotifyCount += sNotifyInfo->mNotifyCount;
                            XPR_SAFE_DELETE(sNotifyInfo);
                            sMerged = XPR_TRUE;
                            break;
                        }
                    }
                }

                if (XPR_IS_FALSE(sMerged))
                    sPending.push_back(sNotifyInfo);

                if (sPending.size() > kMaxNotifyPerWatch)
                {
                    NotifyInfo *sSummary = sPending.back();
                    sPending.pop_back();
                    while (sPending.empty() == false)
                    {
                        NotifyInfo *sOld = sPending.front();
                        sPending.pop_front();
                        sSummary->mNotifyCount += sOld->mNotifyCount;
                        XPR_SAFE_DELETE(sOld);
                    }
                    sSummary->mEvent = EventUpdateDir;
                    sSummary->mFileName.clear();
                    sSummary->mOldDir.clear();
                    sSummary->mOldFileName.clear();
                    sPending.push_back(sSummary);
                }
            }
        }

        // Use the oldest event, not the newest one.  Continuous writers used
        // to postpone delivery forever and grow mNotifyMap without bound.
        const xpr_sint64_t sNow = xpr::timer_ms();
        NotifyMap::iterator sNotifyIterator = mNotifyMap.begin();
        while (sNotifyIterator != mNotifyMap.end())
        {
            NotifyList &sPending = sNotifyIterator->second;
            if (sPending.empty())
            {
                sNotifyIterator = mNotifyMap.erase(sNotifyIterator);
                continue;
            }

            NotifyInfo *sOldest = sPending.front();
            if (sNow - sOldest->mTime >= kNotifyDebounceTime)
            {
                NotifyList::iterator sIterator = sPending.begin();
                while (sIterator != sPending.end())
                {
                    NotifyInfo *sNotifyInfo = *sIterator;
                    if (XPR_IS_NOT_NULL(sNotifyInfo) &&
                        ::PostMessage(sNotifyInfo->mHwnd, sNotifyInfo->mMsg,
                                      reinterpret_cast<WPARAM>(sNotifyInfo),
                                      XPR_NULL) == XPR_TRUE)
                    {
                        sIterator = sPending.erase(sIterator);
                    }
                    else
                    {
                        ++sIterator;
                    }
                }

                // Failed PostMessage payloads remain owned by the watcher.
                clearNotifies(sPending);
            }

            if (sPending.empty())
                sNotifyIterator = mNotifyMap.erase(sNotifyIterator);
            else
                ++sNotifyIterator;
        }
    }

    clearNotifies(mNotifyMap);

    return 0;
}

void AdvFileChangeWatcher::clearNotifies(NotifyList &aNotifyList)
{
    clear(aNotifyList);
}

void AdvFileChangeWatcher::clearNotifies(NotifyMap &aNotifyMap)
{
    if (aNotifyMap.empty() == true)
        return;

    NotifyMap::iterator sIterator;

    sIterator = aNotifyMap.begin();
    for (; sIterator != aNotifyMap.end(); ++sIterator)
    {
        NotifyList &sNotifyList = sIterator->second;
        clearNotifies(sNotifyList);
    }

    aNotifyMap.clear();
}
} // namespace fxfile
