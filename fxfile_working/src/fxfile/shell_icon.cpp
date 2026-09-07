//
// Copyright (c) 2001-2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "shell_icon.h"

#include "sys_img_list.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
ShellIcon::ShellIcon(void)
    : mEvent(XPR_NULL)
    , mGeneration(1)
    , mHwnd(XPR_NULL), mMsg(0)
{
}

ShellIcon::~ShellIcon(void)
{
    stopThread();

    CLOSE_HANDLE(mEvent);

    clear();
}

void ShellIcon::clear(void)
{
    DWORD sWorkerThreadId = 0;

    {
        xpr::MutexGuard sLockGuard(mMutex);

        // Invalidate both queued work and the one request that may already be
        // inside a shell extension.  Each pane owns its own generation.
        ++mGeneration;
        if (mGeneration == 0)
            mGeneration = 1;

        IconDeque::iterator sIterator;
        AsyncIcon *sAsyncIcon;

        sIterator = mIconDeque.begin();
        for (; sIterator != mIconDeque.end(); ++sIterator)
        {
            sAsyncIcon = *sIterator;
            XPR_SAFE_DELETE(sAsyncIcon);
        }

        mIconDeque.clear();

        if (XPR_IS_NOT_NULL(mThread.getThreadHandle().mHandle))
            sWorkerThreadId = (DWORD)mThread.getThreadId();
    }

    // Never hold mMutex while asking COM to cancel: the worker must be able to
    // leave the shell call and acquire that mutex for its generation gate.
    if (sWorkerThreadId != 0)
        ::CoCancelCall(sWorkerThreadId, 0);
}

void ShellIcon::stopThread(void)
{
    mThread.stop();

    if (XPR_IS_NOT_NULL(mEvent))
        ::SetEvent(mEvent);

    // isRunning() remains false during the short start-to-entry window even
    // though a real thread handle already exists.  Use handle ownership for
    // lifecycle decisions so cancellation cannot miss that window.
    if (XPR_IS_NOT_NULL(mThread.getThreadHandle().mHandle))
        ::CoCancelCall((DWORD)mThread.getThreadId(), 0);

    mThread.join();

    {
        xpr::MutexGuard sLockGuard(mMutex);
        CLOSE_HANDLE(mEvent);
    }
}

void ShellIcon::setOwner(HWND aHwnd, xpr_uint_t aMsg)
{
    mHwnd = aHwnd;
    mMsg  = aMsg;
}

xpr_bool_t ShellIcon::getAsyncIcon(AsyncIcon *aAsyncIcon)
{
    if (XPR_IS_NULL(aAsyncIcon))
        return XPR_FALSE;

    if (XPR_IS_NULL(aAsyncIcon->mPidl) &&
        aAsyncIcon->mPath.empty() == XPR_TRUE)
        return XPR_FALSE;

    {
        xpr::MutexGuard sLockGuard(mMutex);

        // The native handle is valid immediately after start() returns, while
        // ThreadState is set later by threadEntry.  It is the race-free token
        // and also permits a stopped ShellIcon instance to restart.
        if (XPR_IS_NULL(mThread.getThreadHandle().mHandle))
        {
            CLOSE_HANDLE(mEvent);
            mEvent = ::CreateEvent(XPR_NULL, XPR_TRUE, XPR_FALSE, XPR_NULL);
            if (XPR_IS_NULL(mEvent))
                return XPR_FALSE;

            mThread.setPriority(3); // THREAD_PRIORITY_BELOW_NORMAL
            xpr_rcode_t sRcode = mThread.start(
                dynamic_cast<xpr::Thread::Runnable *>(this));
            if (XPR_RCODE_IS_NOT_SUCCESS(sRcode))
            {
                CLOSE_HANDLE(mEvent);
                return XPR_FALSE;
            }
        }

        // Task 063: Limit async queue depth to prevent UI-thread stall on
        // large folders (e.g., 1,000+ items).  When the queue is full the
        // caller receives XPR_FALSE; ExplorerCtrl resets mIconRequestIssued
        // so the item will be retried on the next LVN_GETDISPINFO.
        static const xpr_size_t kMaxQueueSize = 300;
        if (mIconDeque.size() >= kMaxQueueSize)
            return XPR_FALSE;

        aAsyncIcon->mGeneration = mGeneration;
        mIconDeque.push_back(aAsyncIcon);

        ::SetEvent(mEvent);
    }

    return XPR_TRUE;
}

xpr_sint_t ShellIcon::runThread(xpr::Thread &aThread)
{
    // important!!!
    HRESULT sOleResult = ::OleInitialize(XPR_NULL);
    const HRESULT sCancelResult = ::CoEnableCallCancellation(XPR_NULL);

    xpr_bool_t sResult;
    AsyncIcon *sAsyncIcon;

    DWORD sWait;

    while (mThread.isStop() == XPR_FALSE)
    {
        sWait = ::WaitForSingleObject(mEvent, 10);
        if (sWait != WAIT_OBJECT_0)
            continue;

        {
            xpr::MutexGuard sLockGuard(mMutex);

            if (mIconDeque.empty() == true)
            {
                ::ResetEvent(mEvent);
                continue;
            }

            sAsyncIcon = mIconDeque.front();
            mIconDeque.pop_front();

            if (XPR_IS_NULL(sAsyncIcon))
                continue;
        }

        // stopThread may race with dequeue before CoCancelCall can observe a
        // running COM call.  Do not enter a shell extension after stop.
        if (mThread.isStop() == XPR_TRUE)
        {
            XPR_SAFE_DELETE(sAsyncIcon);
            break;
        }

        // Rebind an absolute PIDL on this COM apartment.  Keeping the local
        // interface lifetime inside this block prevents cross-apartment shell
        // extension calls, a frequent source of long hangs at shutdown.
        LPSHELLFOLDER sShellFolder = XPR_NULL;
        LPCITEMIDLIST sChildPidl = XPR_NULL;
        if (XPR_IS_NOT_NULL(sAsyncIcon->mPidl))
        {
            ::SHBindToParent(sAsyncIcon->mPidl,
                             IID_IShellFolder,
                             reinterpret_cast<void **>(&sShellFolder),
                             &sChildPidl);
        }

        switch (sAsyncIcon->mType)
        {
        case TypeIcon:
            {
                // 1. Extract Icon from File
                // 2. Shell Icon from Path
                // 3. Shell Icon from Shell Folder & PIDL

                xpr_bool_t sLargeIcon   = XPR_TEST_BITS(sAsyncIcon->mFlags, FlagLargeIcon);
                xpr_bool_t sFastNetIcon = XPR_TEST_BITS(sAsyncIcon->mFlags, FlagFastNetIcon);

                sAsyncIcon->mResult.mIcon = ShellIcon::getIcon(
                    sAsyncIcon->mIconPath,
                    sAsyncIcon->mIconIndex,
                    sAsyncIcon->mPath,
                    sFastNetIcon,
                    sLargeIcon);

                if (XPR_IS_NULL(sAsyncIcon->mResult.mIcon))
                {
                    if (XPR_IS_NOT_NULL(sShellFolder) &&
                        XPR_IS_NOT_NULL(sChildPidl))
                    {
                        sAsyncIcon->mResult.mIcon = GetItemIcon(
                            sShellFolder,
                            const_cast<LPITEMIDLIST>(sChildPidl),
                            sLargeIcon);
                    }
                }
            }
            break;

        case TypeIconIndex:
            {
                if (sAsyncIcon->mPath.empty() == XPR_FALSE)
                    sAsyncIcon->mResult.mIconIndex = GetItemIconIndex(sAsyncIcon->mPath.c_str());
                else if (XPR_IS_NOT_NULL(sShellFolder) &&
                         XPR_IS_NOT_NULL(sChildPidl))
                    sAsyncIcon->mResult.mIconIndex = GetItemIconIndex(
                        sShellFolder,
                        const_cast<LPITEMIDLIST>(sChildPidl));
            }
            break;

        case TypeOverlayIndex:
            {
                if (sAsyncIcon->mPath.empty() == XPR_FALSE)
                    sAsyncIcon->mResult.mIconIndex = GetItemIconOverlayIndex(sAsyncIcon->mPath.c_str());
                else if (XPR_IS_NOT_NULL(sShellFolder) &&
                         XPR_IS_NOT_NULL(sChildPidl))
                    sAsyncIcon->mResult.mIconIndex = GetItemIconOverlayIndex(
                        sShellFolder,
                        const_cast<LPITEMIDLIST>(sChildPidl));
            }
            break;
        }

        COM_RELEASE(sShellFolder);

        xpr_bool_t sCurrentGeneration = XPR_FALSE;
        {
            xpr::MutexGuard sLockGuard(mMutex);
            sCurrentGeneration = (sAsyncIcon->mGeneration == mGeneration &&
                                  mThread.isStop() == XPR_FALSE) ? XPR_TRUE : XPR_FALSE;
        }

        // A pane may navigate while its worker is blocked in a shell extension.
        // Discard that eventual stale completion before it reaches the UI queue;
        // ExplorerCtrl::mCode remains the second, UI-side identity guard.
        if (XPR_IS_FALSE(sCurrentGeneration))
        {
            XPR_SAFE_DELETE(sAsyncIcon);
            continue;
        }

        sResult = ::PostMessage(mHwnd, mMsg, (WPARAM)sAsyncIcon, (LPARAM)XPR_NULL);
        if (XPR_IS_FALSE(sResult))
        {
            XPR_SAFE_DELETE(sAsyncIcon);
        }
    }

    if (SUCCEEDED(sCancelResult))
        ::CoDisableCallCancellation(XPR_NULL);
    if (SUCCEEDED(sOleResult))
        ::OleUninitialize();

    return 0;
}

HICON ShellIcon::getIcon(const xpr::string &aIconPath, xpr_sint_t aIconIndex, const xpr::string &aPath, xpr_bool_t aFastNetIcon, xpr_bool_t aLarge)
{
    HICON sIcon = XPR_NULL;

    if (aIconPath.empty() == XPR_FALSE)
    {
        sIcon = extractIcon(aIconPath, aIconIndex, aLarge);
        if (XPR_IS_NULL(sIcon))
        {
            if (aIconPath[0] == XPR_STRING_LITERAL('%'))
            {
                xpr::string sRealIconPath;
                GetEnvRealPath(aIconPath, sRealIconPath);

                sIcon = extractIcon(sRealIconPath, aIconIndex, aLarge);
            }
        }
    }

    if (XPR_IS_NULL(sIcon))
    {
        if (aPath.empty() == XPR_FALSE)
        {
            if (XPR_IS_TRUE(aFastNetIcon))
            {
                // network icon index of shell32.dll
                // ---------------------------------
                // network computer : 15
                // network folder   : 85

                if (IsNetItem(aPath.c_str()) == XPR_TRUE)
                {
                    if (aPath.rfind(XPR_STRING_LITERAL('\\')) < 2)
                    {
                        sIcon = extractIcon(XPR_STRING_LITERAL("%SystemRoot%\\system32\\SHELL32.dll"), 15, aLarge);
                    }
                    else
                    {
                        sIcon = extractIcon(XPR_STRING_LITERAL("%SystemRoot%\\system32\\SHELL32.dll"), 85, aLarge);
                    }
                }
            }

            if (XPR_IS_NULL(sIcon))
            {
                LPITEMIDLIST sFullPidl = Path2Pidl(aPath);
                if (XPR_IS_NOT_NULL(sFullPidl))
                {
                    sIcon = GetItemIcon(sFullPidl, XPR_FALSE, aLarge);
                    COM_FREE(sFullPidl);
                }
            }
        }
    }

    return sIcon;
}

HICON ShellIcon::extractIcon(const xpr::string &aIconPath, xpr_sint_t aIconIndex, xpr_bool_t aLarge)
{
    HICON sIcon = XPR_NULL;

    ::ExtractIconEx(
        aIconPath.c_str(),
        aIconIndex,
        XPR_IS_TRUE(aLarge) ? &sIcon : XPR_NULL,
        XPR_IS_TRUE(aLarge) ? XPR_NULL : &sIcon,
        1);

    return sIcon;
}
} // namespace fxfile
