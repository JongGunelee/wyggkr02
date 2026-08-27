//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "file_operation_lock_store.h"
#include "multi_rename.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cmd
{
namespace
{
const xpr_size_t kMaxBackupCandidateCount = 1000;

xpr_bool_t isDestinationCollisionError(DWORD aError)
{
    return (aError == ERROR_FILE_EXISTS || aError == ERROR_ALREADY_EXISTS) ? XPR_TRUE : XPR_FALSE;
}

MultiRename::Result getMoveErrorResult(DWORD aError, DWORD aAttributes = INVALID_FILE_ATTRIBUTES)
{
    switch (aError)
    {
    case ERROR_FILE_NOT_FOUND:
    case ERROR_PATH_NOT_FOUND:
    case ERROR_INVALID_DRIVE:
    case ERROR_BAD_PATHNAME:
    case ERROR_BAD_NETPATH:
    case ERROR_BAD_NET_NAME:
        return MultiRename::ResultPathNotExist;

    case ERROR_FILENAME_EXCED_RANGE:
    case ERROR_BUFFER_OVERFLOW:
        return MultiRename::ResultExcessPathLength;

    case ERROR_SHARING_VIOLATION:
    case ERROR_LOCK_VIOLATION:
    case ERROR_USER_MAPPED_FILE:
        return MultiRename::ResultShared;

    case ERROR_ACCESS_DENIED:
    case ERROR_WRITE_PROTECT:
        if (aAttributes != INVALID_FILE_ATTRIBUTES &&
            XPR_TEST_BITS(aAttributes, FILE_ATTRIBUTE_READONLY))
        {
            return MultiRename::ResultReadOnly;
        }
        break;
    }

    return MultiRename::ResultUnknownError;
}
} // namespace

MultiRename::MultiRename(void)
    : mHwnd(XPR_NULL), mMsg(0)
    , mStatus(StatusNone)
    , mPreparedCount(0), mValidatedCount(0), mRenamedCount(0)
    , mInvalidItem(-1)
    // FlagReadOnlyRename used to have value zero, which made
    // XPR_TEST_BITS(flags, flag) true for every flags value.  There is no
    // batch-rename UI switch today, so keep that long-standing effective
    // behaviour as the default while making the flag itself a valid bit.
    , mFlags(FlagReadOnlyRename)
{
}

MultiRename::~MultiRename(void)
{
    stop();

    RenDeque::iterator sIterator;
    RenItem *sRenItem;

    sIterator = mRenDeque.begin();
    for (; sIterator != mRenDeque.end(); ++sIterator)
    {
        sRenItem = *sIterator;
        XPR_SAFE_DELETE(sRenItem);
    }

    mRenDeque.clear();
}

void MultiRename::setOwner(HWND aHwnd, xpr_uint_t aMsg)
{
    mHwnd = aHwnd;
    mMsg  = aMsg;
}

void MultiRename::setBackupName(const xpr_tchar_t *aBackup)
{
    if (XPR_IS_NOT_NULL(aBackup))
        mBackup = aBackup;
}

void MultiRename::setFlags(xpr_uint_t aFlags)
{
    mFlags = aFlags;
}

xpr_uint_t MultiRename::getFlags(void)
{
    return mFlags;
}

xpr_bool_t MultiRename::isFlag(xpr_uint_t aFlag)
{
    return XPR_TEST_BITS(mFlags, aFlag);
}

void MultiRename::addPath(const xpr_tchar_t *aDir, const xpr_tchar_t *aOld, const xpr_tchar_t *aNew)
{
    if (XPR_IS_NULL(aDir) || XPR_IS_NULL(aOld) || XPR_IS_NULL(aNew))
        return;

    RenItem *sRenItem = new RenItem;
    if (XPR_IS_NULL(sRenItem))
        return;

    sRenItem->mDir    = aDir;
    sRenItem->mOld    = aOld;
    sRenItem->mNew    = aNew;
    sRenItem->mResult = ResultNone;

    mRenDeque.push_back(sRenItem);
}

void MultiRename::addPath(const xpr::string &aDir, const xpr::string &aOld, const xpr::string &aNew)
{
    addPath(aDir.c_str(), aOld.c_str(), aNew.c_str());
}

xpr_bool_t MultiRename::start(void)
{
    {
        xpr::MutexGuard sLockGuard(mMutex);

        mInvalidItem    = -1;
        mPreparedCount  = 0;
        mValidatedCount = 0;
        mRenamedCount   = 0;
        mStatus         = StatusPreparing;
    }

    RenDeque::iterator sIterator = mRenDeque.begin();
    for (; sIterator != mRenDeque.end(); ++sIterator)
    {
        RenItem *sRenItem = *sIterator;
        if (XPR_IS_NOT_NULL(sRenItem))
            sRenItem->mResult = ResultNone;
    }

    xpr_rcode_t sRcode = mThread.start(dynamic_cast<xpr::Thread::Runnable *>(this));

    if (XPR_RCODE_IS_ERROR(sRcode))
    {
        xpr::MutexGuard sLockGuard(mMutex);
        mStatus = StatusNone;
    }

    return XPR_RCODE_IS_SUCCESS(sRcode);
}

void MultiRename::stop(void)
{
    mThread.stop();
    mThread.join();
}

xpr_sint_t MultiRename::runThread(xpr::Thread &aThread)
{
    xpr_bool_t sReadOnlyRename = isFlag(FlagReadOnlyRename);

    xpr::string sSrc;
    xpr::string sDst;
    xpr::string sTemp;
    xpr::string sTempFileName;
    DWORD sAttributes;

    xpr_sint_t sInvalidItem = -1;
    xpr_bool_t sInvalid = XPR_FALSE;
    xpr_sint_t sFirstRenameFailedItem = -1;
    xpr_bool_t sRenameFailed = XPR_FALSE;
    Status sStatus = StatusNone;

    xpr_size_t sLen;
    RenItem *sRenItem;
    RenItem *sRenItem2;
    RenDeque::iterator sIterator;
    RenDeque::iterator sIterator2;

    typedef std::tr1::unordered_multimap<xpr::string, RenItem *> HashPathMap;
    typedef std::pair<HashPathMap::iterator, HashPathMap::iterator> HashPathPairIterator;
    HashPathMap sHashOldPathMap;
    HashPathMap sHashNewPathMap;
    HashPathMap::iterator sHashPathIterator;
    HashPathPairIterator sPairRangeIterator;

    sIterator = mRenDeque.begin();
    for (; sIterator != mRenDeque.end(); ++sIterator)
    {
        sRenItem = *sIterator;
        if (XPR_IS_NULL(sRenItem))
            continue;

        sSrc = sRenItem->mDir + XPR_STRING_LITERAL('\\') + sRenItem->mOld;
        sDst = sRenItem->mDir + XPR_STRING_LITERAL('\\') + sRenItem->mNew;

        sSrc.upper_case();
        sDst.upper_case();

        sHashOldPathMap.insert(HashPathMap::value_type(sSrc, sRenItem));
        sHashNewPathMap.insert(HashPathMap::value_type(sDst, sRenItem));

        {
            xpr::MutexGuard sLockGuard(mMutex);
            mPreparedCount++;
        }
    }

    {
        xpr::MutexGuard sLockGuard(mMutex);
        mStatus = StatusValidating;
    }

    sIterator = mRenDeque.begin();
    for (; sIterator != mRenDeque.end(); ++sIterator)
    {
        sRenItem = *sIterator;
        if (XPR_IS_NULL(sRenItem))
            continue;

        if (mThread.isStop())
            break;

        sSrc = sRenItem->mDir + XPR_STRING_LITERAL('\\') + sRenItem->mOld;
        sDst = sRenItem->mDir + XPR_STRING_LITERAL('\\') + sRenItem->mNew;

        sLen = sDst.length();

        if (sLen == 0 || sLen >= XPR_MAX_PATH)
        {
            sRenItem->mResult = (sLen == 0) ? ResultEmptiedName : ResultExcessPathLength;
            sInvalidItem = (xpr_sint_t)std::distance(mRenDeque.begin(), sIterator);
            sInvalid = XPR_TRUE;
            break;
        }

        if (VerifyFileName(sRenItem->mNew) == XPR_FALSE)
        {
            sRenItem->mResult = ResultInvalidName;
            sInvalidItem = (xpr_sint_t)std::distance(mRenDeque.begin(), sIterator);
            sInvalid = XPR_TRUE;
            break;
        }

        sDst.upper_case();
        sPairRangeIterator = sHashNewPathMap.equal_range(sDst);

        sHashPathIterator = sPairRangeIterator.first;
        for (; sHashPathIterator != sPairRangeIterator.second; ++sHashPathIterator)
        {
            sRenItem2 = sHashPathIterator->second;
            if (XPR_IS_NULL(sRenItem2))
                continue;

            if (sRenItem != sRenItem2)
                break;
        }

        if (sHashPathIterator != sPairRangeIterator.second)
        {
            sRenItem->mResult = ResultEqualedName;
            sInvalidItem = (xpr_sint_t)std::distance(mRenDeque.begin(), sIterator);
            sInvalid = XPR_TRUE;
            break;
        }

        {
            xpr::MutexGuard sLockGuard(mMutex);
            mValidatedCount++;
        }
    }

    {
        xpr::MutexGuard sLockGuard(mMutex);
        mStatus = StatusRenaming;
    }

    if (mThread.isStop() == XPR_FALSE && XPR_IS_FALSE(sInvalid))
    {
        sIterator = mRenDeque.begin();
        for (; sIterator != mRenDeque.end(); ++sIterator)
        {
            sRenItem = *sIterator;
            if (XPR_IS_NULL(sRenItem))
                continue;

            if (mThread.isStop() == XPR_TRUE)
                break;

            sRenItem->mResult = ResultNone;

            sSrc = sRenItem->mDir + XPR_STRING_LITERAL('\\') + sRenItem->mOld;
            sDst = sRenItem->mDir + XPR_STRING_LITERAL('\\') + sRenItem->mNew;

            if (FileOperationLockStore::instance().affectsLockedPath(
                    std::wstring(sSrc.c_str())))
            {
                // Use the existing per-item read-only result so the batch UI
                // reports a protected item instead of silently renaming it.
                sRenItem->mResult = ResultReadOnly;
                if (XPR_IS_FALSE(sRenameFailed))
                {
                    sFirstRenameFailedItem =
                        (xpr_sint_t)std::distance(mRenDeque.begin(), sIterator);
                    sRenameFailed = XPR_TRUE;
                }
                xpr::MutexGuard sLockGuard(mMutex);
                mRenamedCount++;
                continue;
            }

            // notice : The file name must compare by case.
            if (_tcscmp(sSrc.c_str(), sDst.c_str()) != 0)
            {
                sAttributes = ::GetFileAttributes(sSrc.c_str());
                if (sAttributes == INVALID_FILE_ATTRIBUTES)
                {
                    sRenItem->mResult = getMoveErrorResult(::GetLastError());
                }
                else
                {
                    const xpr_bool_t sWasReadOnly =
                        XPR_TEST_BITS(sAttributes, FILE_ATTRIBUTE_READONLY) ? XPR_TRUE : XPR_FALSE;
                    xpr_bool_t sReadOnlyCleared = XPR_FALSE;

                    if (XPR_IS_TRUE(sWasReadOnly))
                    {
                        if (XPR_IS_FALSE(sReadOnlyRename))
                        {
                            sRenItem->mResult = ResultReadOnly;
                        }
                        else if (::SetFileAttributes(sSrc.c_str(), sAttributes & ~FILE_ATTRIBUTE_READONLY) == XPR_FALSE)
                        {
                            sRenItem->mResult = getMoveErrorResult(::GetLastError(), sAttributes);
                        }
                        else
                        {
                            sReadOnlyCleared = XPR_TRUE;
                        }
                    }

                    if (sRenItem->mResult == ResultNone)
                    {
                        if (::MoveFile(sSrc.c_str(), sDst.c_str()) == XPR_TRUE)
                        {
                            // Keep the worker model aligned with the physical
                            // path so a later partial failure can be retried.
                            sRenItem->mOld = sRenItem->mNew;
                            if (XPR_IS_TRUE(sReadOnlyCleared) &&
                                ::SetFileAttributes(sDst.c_str(), sAttributes) == XPR_FALSE)
                            {
                                sRenItem->mResult = getMoveErrorResult(::GetLastError(), sAttributes);
                            }
                        }
                        else
                        {
                            DWORD sDirectMoveError = ::GetLastError();

                            // A backup is only valid when the direct move failed
                            // because the destination already exists.  Access,
                            // sharing, source and path failures must not displace it.
                            if (XPR_IS_TRUE(isDestinationCollisionError(sDirectMoveError)))
                            {
                                xpr_bool_t sDestinationBackedUp = XPR_FALSE;
                                DWORD sBackupError = ERROR_ALREADY_EXISTS;
                                sTempFileName = sRenItem->mNew;

                                for (xpr_size_t sAttempt = 0;
                                     sAttempt < kMaxBackupCandidateCount && XPR_IS_FALSE(mThread.isStop());
                                     ++sAttempt)
                                {
                                    sTempFileName.insert(0, mBackup);

                                    sTemp  = sRenItem->mDir;
                                    sTemp += XPR_STRING_LITERAL('\\');
                                    sTemp += sTempFileName;

                                    if (sTemp.length() >= XPR_MAX_PATH)
                                    {
                                        sBackupError = ERROR_FILENAME_EXCED_RANGE;
                                        break;
                                    }

                                    if (::MoveFile(sDst.c_str(), sTemp.c_str()) == XPR_TRUE)
                                    {
                                        sDestinationBackedUp = XPR_TRUE;
                                        break;
                                    }

                                    sBackupError = ::GetLastError();
                                    if (XPR_IS_FALSE(isDestinationCollisionError(sBackupError)))
                                        break;
                                }

                                if (XPR_IS_TRUE(mThread.isStop()))
                                {
                                    if (XPR_IS_TRUE(sDestinationBackedUp) &&
                                        ::MoveFile(sTemp.c_str(), sDst.c_str()) == XPR_FALSE)
                                    {
                                        xpr::string sUpperDst = sDst;
                                        sUpperDst.upper_case();
                                        sPairRangeIterator = sHashOldPathMap.equal_range(sUpperDst);

                                        for (sHashPathIterator = sPairRangeIterator.first;
                                             sHashPathIterator != sPairRangeIterator.second;
                                             ++sHashPathIterator)
                                        {
                                            sRenItem2 = sHashPathIterator->second;
                                            if (XPR_IS_NOT_NULL(sRenItem2) && sRenItem != sRenItem2)
                                                sRenItem2->mOld = sTempFileName;
                                        }
                                    }

                                    sRenItem->mResult = ResultUnknownError;
                                }
                                else if (XPR_IS_FALSE(sDestinationBackedUp))
                                {
                                    sRenItem->mResult = getMoveErrorResult(sBackupError);
                                }
                                else if (::MoveFile(sSrc.c_str(), sDst.c_str()) == XPR_TRUE)
                                {
                                    sRenItem->mOld = sRenItem->mNew;
                                    // A pending item which formerly referred to
                                    // dst must now follow the displaced object.
                                    xpr::string sUpperDst = sDst;
                                    sUpperDst.upper_case();
                                    sPairRangeIterator = sHashOldPathMap.equal_range(sUpperDst);

                                    for (sHashPathIterator = sPairRangeIterator.first;
                                         sHashPathIterator != sPairRangeIterator.second;
                                         ++sHashPathIterator)
                                    {
                                        sRenItem2 = sHashPathIterator->second;
                                        if (XPR_IS_NOT_NULL(sRenItem2) && sRenItem != sRenItem2)
                                            sRenItem2->mOld = sTempFileName;
                                    }

                                    if (XPR_IS_TRUE(sReadOnlyCleared) &&
                                        ::SetFileAttributes(sDst.c_str(), sAttributes) == XPR_FALSE)
                                    {
                                        sRenItem->mResult = getMoveErrorResult(::GetLastError(), sAttributes);
                                    }
                                }
                                else
                                {
                                    const DWORD sSecondMoveError = ::GetLastError();

                                    // The requested rename did not happen.  Put
                                    // the displaced destination back.  If that
                                    // rollback itself fails, keep the pending
                                    // item's old name synchronized with reality.
                                    if (::MoveFile(sTemp.c_str(), sDst.c_str()) == XPR_FALSE)
                                    {
                                        xpr::string sUpperDst = sDst;
                                        sUpperDst.upper_case();
                                        sPairRangeIterator = sHashOldPathMap.equal_range(sUpperDst);

                                        for (sHashPathIterator = sPairRangeIterator.first;
                                             sHashPathIterator != sPairRangeIterator.second;
                                             ++sHashPathIterator)
                                        {
                                            sRenItem2 = sHashPathIterator->second;
                                            if (XPR_IS_NOT_NULL(sRenItem2) && sRenItem != sRenItem2)
                                                sRenItem2->mOld = sTempFileName;
                                        }
                                    }

                                    sRenItem->mResult = getMoveErrorResult(sSecondMoveError, sAttributes);
                                }
                            }
                            else
                            {
                                sRenItem->mResult = getMoveErrorResult(sDirectMoveError, sAttributes);
                            }
                        }
                    }

                    // Every failure path before a successful source move leaves
                    // the source at sSrc, so restore its original attributes.
                    if (XPR_IS_TRUE(sReadOnlyCleared) &&
                        ::GetFileAttributes(sSrc.c_str()) != INVALID_FILE_ATTRIBUTES)
                    {
                        if (::SetFileAttributes(sSrc.c_str(), sAttributes) == XPR_FALSE &&
                            sRenItem->mResult == ResultNone)
                        {
                            sRenItem->mResult = getMoveErrorResult(::GetLastError(), sAttributes);
                        }
                    }
                }
            }

            if (sRenItem->mResult == ResultNone)
                sRenItem->mResult = ResultSucceeded;

            if (sRenItem->mResult != ResultSucceeded &&
                XPR_IS_FALSE(sRenameFailed))
            {
                sFirstRenameFailedItem =
                    (xpr_sint_t)std::distance(mRenDeque.begin(), sIterator);
                sRenameFailed = XPR_TRUE;
            }

            {
                xpr::MutexGuard sLockGuard(mMutex);
                mRenamedCount++;
            }
        }
    }

    sHashOldPathMap.clear();
    sHashNewPathMap.clear();

    if (mThread.isStop() == XPR_TRUE)
    {
        sStatus = StatusStopped;
    }
    else
    {
        if (XPR_IS_TRUE(sInvalid))
            sStatus = StatusInvalid;
        else if (XPR_IS_TRUE(sRenameFailed))
        {
            // Do not report a partial or failed filesystem operation as a
            // completed batch. The dialog remains open on the first failure.
            sInvalidItem = sFirstRenameFailedItem;
            sStatus = StatusRenameFailed;
        }
        else
            sStatus = StatusRenameCompleted;
    }

    {
        xpr::MutexGuard sLockGuard(mMutex);
        mInvalidItem = sInvalidItem;
        mStatus = sStatus;
    }

    ::PostMessage(mHwnd, mMsg, (WPARAM)XPR_NULL, (LPARAM)XPR_NULL);

    return 0;
}

MultiRename::Status MultiRename::getStatus(xpr_size_t *aPreparedCount, xpr_size_t *aValidatedCount, xpr_size_t *aRenamedCount)
{
    xpr::MutexGuard sLockGuard(mMutex);

    if (aPreparedCount)  *aPreparedCount  = mPreparedCount;
    if (aValidatedCount) *aValidatedCount = mValidatedCount;
    if (aRenamedCount)   *aRenamedCount   = mRenamedCount;

    return mStatus;
}

MultiRename::Result MultiRename::getItemResult(xpr_size_t aIndex)
{
    if (!FXFILE_STL_IS_INDEXABLE(aIndex, mRenDeque))
        return ResultNone;

    return mRenDeque[aIndex]->mResult;
}

xpr_bool_t MultiRename::getItemOldName(xpr_size_t aIndex, xpr::string &aOldName)
{
    xpr::MutexGuard sLockGuard(mMutex);

    if (!FXFILE_STL_IS_INDEXABLE(aIndex, mRenDeque) ||
        XPR_IS_NULL(mRenDeque[aIndex]))
    {
        return XPR_FALSE;
    }

    aOldName = mRenDeque[aIndex]->mOld;
    return XPR_TRUE;
}

xpr_sint_t MultiRename::getInvalidItem(void)
{
    xpr::MutexGuard sLockGuard(mMutex);

    return mInvalidItem;
}
} // namespace cmd
} // namespace fxfile
