//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "directory_enumeration_worker.h"

#include "../base/shell_enumerator.h"
#include "../base/pidl_win.h"

#include <process.h>
#include <mutex>
#include <new>
#include <set>

namespace fxfile
{
namespace
{
const xpr_size_t kEnumerationBatchSize = 128;
// The first enumerated Shell item can be filtered by the UI (for example a
// hidden+system desktop.ini). Publish a small finite burst item-by-item so a
// filtered first item cannot hold all visible content until final completion.
// Steady state remains bounded at 128 items and four outstanding batches.
const xpr_size_t kFirstEnumerationBatchSize = 1;
const xpr_size_t kInitialEnumerationBatchCount = 8;

std::mutex &batchRegistryMutex(void)
{
    // Enumeration workers are detached and can finish while the application
    // is unwinding process-wide C++ statics.  Keep these two tiny registry
    // objects alive until process teardown instead of risking a late worker
    // touching an already-destroyed function-local static.
    static std::mutex *sMutex = new std::mutex;
    return *sMutex;
}

std::set<DirectoryEnumerationWorker::Batch *> &batchRegistry(void)
{
    static std::set<DirectoryEnumerationWorker::Batch *> *sBatches =
        new std::set<DirectoryEnumerationWorker::Batch *>;
    return *sBatches;
}

void registerBatch(DirectoryEnumerationWorker::Batch *aBatch)
{
    std::lock_guard<std::mutex> sLock(batchRegistryMutex());
    batchRegistry().insert(aBatch);
}

DirectoryEnumerationWorker::Batch *unregisterBatch(
    DirectoryEnumerationWorker::Batch *aBatch)
{
    std::lock_guard<std::mutex> sLock(batchRegistryMutex());
    std::set<DirectoryEnumerationWorker::Batch *>::iterator sIt =
        batchRegistry().find(aBatch);
    if (sIt == batchRegistry().end())
        return XPR_NULL;
    DirectoryEnumerationWorker::Batch *sOwned = *sIt;
    batchRegistry().erase(sIt);
    return sOwned;
}

struct EnumerationContext
{
    HWND mOwnerHwnd;
    xpr_uint_t mMessage;
    LPITEMIDLIST mFolderFullPidl;
    xpr_sint_t mListType;
    xpr_sint_t mAttributes;
    xpr_uint_t mGeneration;
    xpr_uint_t mOwnerToken;
    HANDLE mCancelEvent;
    HANDLE mFlowSemaphore;
};

xpr_bool_t isCancelled(const EnumerationContext *aContext)
{
    return (::WaitForSingleObject(aContext->mCancelEvent, 0) == WAIT_OBJECT_0) ?
           XPR_TRUE : XPR_FALSE;
}

xpr_bool_t postBatch(EnumerationContext *aContext,
                     DirectoryEnumerationWorker::Batch *aBatch,
                     xpr_bool_t aFlowControlled = XPR_TRUE)
{
    if (XPR_IS_TRUE(aFlowControlled))
    {
        HANDLE sWaitHandles[2] = {aContext->mCancelEvent,
                                  aContext->mFlowSemaphore};
        const DWORD sWaitResult =
            ::WaitForMultipleObjects(2, sWaitHandles, FALSE, INFINITE);
        if (sWaitResult != WAIT_OBJECT_0 + 1)
        {
            DirectoryEnumerationWorker::destroyBatch(aBatch);
            return XPR_FALSE;
        }
        if (!::DuplicateHandle(::GetCurrentProcess(),
                               aContext->mFlowSemaphore,
                               ::GetCurrentProcess(), &aBatch->mFlowSemaphore,
                               0, FALSE, DUPLICATE_SAME_ACCESS))
        {
            ::ReleaseSemaphore(aContext->mFlowSemaphore, 1, XPR_NULL);
            DirectoryEnumerationWorker::destroyBatch(aBatch);
            return XPR_FALSE;
        }
    }

    // Register before posting.  The receiver claims ownership atomically;
    // pane shutdown can instead reclaim every still-unclaimed batch by its
    // owner token.  A stale pointer left in the Windows queue is harmless
    // because claimBatch validates it without dereferencing it.
    registerBatch(aBatch);
    if (XPR_IS_TRUE(isCancelled(aContext)) ||
        !::IsWindow(aContext->mOwnerHwnd) ||
        !::PostMessage(aContext->mOwnerHwnd, aContext->mMessage,
                       reinterpret_cast<WPARAM>(aBatch), 0))
    {
        DirectoryEnumerationWorker::destroyBatch(unregisterBatch(aBatch));
        return XPR_FALSE;
    }

    return XPR_TRUE;
}

unsigned __stdcall enumerationProc(void *aParameter)
{
    EnumerationContext *sContext =
        reinterpret_cast<EnumerationContext *>(aParameter);
    const ULONGLONG sStarted = ::GetTickCount64();
    xpr_bool_t sSucceeded = XPR_FALSE;
    xpr_bool_t sTransportSucceeded = XPR_TRUE;

    const HRESULT sComResult = ::CoInitializeEx(XPR_NULL, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(sComResult))
    {
        LPSHELLFOLDER sParentShellFolder = XPR_NULL;
        LPSHELLFOLDER sShellFolder = XPR_NULL;
        LPCITEMIDLIST sSimplePidl = XPR_NULL;

        if (XPR_IS_TRUE(base::Pidl::getSimplePidl(
                sContext->mFolderFullPidl, sParentShellFolder, sSimplePidl)) &&
            XPR_IS_TRUE(base::Pidl::getShellFolder(
                sParentShellFolder, sSimplePidl, sShellFolder)))
        {
            base::ShellEnumerator sEnumerator;
            if (XPR_IS_TRUE(sEnumerator.enumerate(
                    XPR_NULL, sShellFolder, sContext->mListType,
                    sContext->mAttributes)))
            {
                DirectoryEnumerationWorker::Batch *sBatch =
                    new DirectoryEnumerationWorker::Batch;
                sBatch->mGeneration = sContext->mGeneration;
                sBatch->mOwnerToken = sContext->mOwnerToken;
                xpr_size_t sInitialBatchCount = 0;

                LPITEMIDLIST sPidl = XPR_NULL;
                while (XPR_IS_FALSE(isCancelled(sContext)) &&
                       XPR_IS_TRUE(sEnumerator.next(&sPidl)))
                {
                    DirectoryEnumerationWorker::Item sItem;
                    sItem.mPidl = sPidl;
                    xpr_ulong_t sRequestedShellAttributes =
                        SFGAO_FILESYSTEM | SFGAO_FOLDER | SFGAO_CANRENAME |
                        SFGAO_CANCOPY | SFGAO_CANMOVE | SFGAO_CANDELETE |
                        SFGAO_LINK | SFGAO_GHOSTED;
                    const HRESULT sAttributeResult =
                        sShellFolder->GetAttributesOf(
                            1, (LPCITEMIDLIST *)&sPidl,
                            &sRequestedShellAttributes);
                    if (FAILED(sAttributeResult))
                        sRequestedShellAttributes = 0;
                    sItem.mShellAttributes = sRequestedShellAttributes;

                    WIN32_FIND_DATAW sFindData = {0};
                    const HRESULT sFindDataResult =
                        ::SHGetDataFromIDListW(
                            sShellFolder, sPidl, SHGDFIL_FINDDATA,
                            &sFindData, sizeof(sFindData));
                    if (SUCCEEDED(sFindDataResult))
                    {
                        sItem.mFileAttributes = sFindData.dwFileAttributes;
                        if ((sFindData.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) != 0)
                            sItem.mShellAttributes |= SFGAO_GHOSTED;
                        if ((sFindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
                            sItem.mShellAttributes |= SFGAO_FOLDER;
                        else
                            sItem.mShellAttributes &= ~SFGAO_FOLDER;

                        // This is the exact filesystem leaf name and avoids a
                        // second Shell callback on the UI thread for every
                        // item. cAlternateFileName is intentionally ignored.
                        if (sFindData.cFileName[0] != L'\0')
                            sItem.mName = sFindData.cFileName;
                    }
                    sItem.mHasKnownMetadata =
                        (SUCCEEDED(sAttributeResult) &&
                         SUCCEEDED(sFindDataResult)) ? XPR_TRUE : XPR_FALSE;

                    if (sItem.mName.empty())
                    {
                        STRRET sName = {0};
                        wchar_t sNameBuffer[XPR_MAX_PATH + 1] = {0};
                        if (SUCCEEDED(sShellFolder->GetDisplayNameOf(
                                sPidl, SHGDN_INFOLDER, &sName)) &&
                            SUCCEEDED(::StrRetToBufW(&sName, sPidl, sNameBuffer,
                                                    _countof(sNameBuffer))))
                            sItem.mName = sNameBuffer;
                    }

                    sBatch->mItems.push_back(sItem);
                    sPidl = XPR_NULL;

                    const xpr_size_t sBatchSize =
                        sInitialBatchCount < kInitialEnumerationBatchCount ?
                        kFirstEnumerationBatchSize : kEnumerationBatchSize;
                    if (sBatch->mItems.size() >= sBatchSize)
                    {
                        if (XPR_IS_FALSE(postBatch(sContext, sBatch)))
                        {
                            sTransportSucceeded = XPR_FALSE;
                            sBatch = XPR_NULL;
                            break;
                        }
                        ++sInitialBatchCount;
                        sBatch = new DirectoryEnumerationWorker::Batch;
                        sBatch->mGeneration = sContext->mGeneration;
                        sBatch->mOwnerToken = sContext->mOwnerToken;
                    }
                }

                if (XPR_IS_NOT_NULL(sBatch))
                {
                    if (XPR_IS_FALSE(sBatch->mItems.empty()))
                    {
                        if (XPR_IS_FALSE(postBatch(sContext, sBatch)))
                            sTransportSucceeded = XPR_FALSE;
                    }
                    else
                        delete sBatch;
                }

                sSucceeded = XPR_IS_FALSE(isCancelled(sContext)) &&
                             XPR_IS_TRUE(sTransportSucceeded);
            }
        }

        COM_RELEASE(sShellFolder);
        COM_RELEASE(sParentShellFolder);
        ::CoUninitialize();
    }

    if (XPR_IS_FALSE(isCancelled(sContext)))
    {
        DirectoryEnumerationWorker::Batch *sCompletion =
            new DirectoryEnumerationWorker::Batch;
        sCompletion->mGeneration = sContext->mGeneration;
        sCompletion->mOwnerToken = sContext->mOwnerToken;
        sCompletion->mComplete = XPR_TRUE;
        sCompletion->mSucceeded = sSucceeded;
        sCompletion->mEnumerationMilliseconds =
            static_cast<xpr_uint64_t>(::GetTickCount64() - sStarted);
        postBatch(sContext, sCompletion, XPR_FALSE);
    }

    COM_FREE(sContext->mFolderFullPidl);
    CLOSE_HANDLE(sContext->mCancelEvent);
    CLOSE_HANDLE(sContext->mFlowSemaphore);
    delete sContext;
    ::_endthreadex(0);
    return 0;
}
} // namespace anonymous

DirectoryEnumerationWorker::Batch::Batch(void)
    : mGeneration(0)
    , mOwnerToken(0)
    , mComplete(XPR_FALSE)
    , mSucceeded(XPR_FALSE)
    , mEnumerationMilliseconds(0)
    , mFlowSemaphore(XPR_NULL)
{
}

DirectoryEnumerationWorker::Item::Item(void)
    : mPidl(XPR_NULL)
    , mShellAttributes(0)
    , mFileAttributes(0)
    , mHasKnownMetadata(XPR_FALSE)
{
}

xpr_bool_t DirectoryEnumerationWorker::start(
    HWND aOwnerHwnd,
    xpr_uint_t aMessage,
    LPCITEMIDLIST aFolderFullPidl,
    xpr_sint_t aListType,
    xpr_sint_t aAttributes,
    xpr_uint_t aGeneration,
    xpr_uint_t aOwnerToken,
    HANDLE aCancelEvent)
{
    if (!::IsWindow(aOwnerHwnd) || XPR_IS_NULL(aFolderFullPidl) ||
        XPR_IS_NULL(aCancelEvent))
        return XPR_FALSE;

    EnumerationContext *sContext = new (std::nothrow) EnumerationContext;
    if (XPR_IS_NULL(sContext))
        return XPR_FALSE;
    sContext->mOwnerHwnd = aOwnerHwnd;
    sContext->mMessage = aMessage;
    sContext->mFolderFullPidl = base::Pidl::clone(aFolderFullPidl);
    sContext->mListType = aListType;
    sContext->mAttributes = aAttributes;
    sContext->mGeneration = aGeneration;
    sContext->mOwnerToken = aOwnerToken;
    sContext->mCancelEvent = XPR_NULL;
    sContext->mFlowSemaphore = ::CreateSemaphore(XPR_NULL, 4, 4, XPR_NULL);

    if (XPR_IS_NULL(sContext->mFolderFullPidl) ||
        XPR_IS_NULL(sContext->mFlowSemaphore) ||
        !::DuplicateHandle(::GetCurrentProcess(), aCancelEvent,
                           ::GetCurrentProcess(), &sContext->mCancelEvent,
                           SYNCHRONIZE, FALSE, 0))
    {
        COM_FREE(sContext->mFolderFullPidl);
        CLOSE_HANDLE(sContext->mFlowSemaphore);
        delete sContext;
        return XPR_FALSE;
    }

    unsigned sThreadId = 0;
    HANDLE sThread = reinterpret_cast<HANDLE>(::_beginthreadex(
        XPR_NULL, 0, enumerationProc, sContext, 0, &sThreadId));
    if (XPR_IS_NULL(sThread))
    {
        COM_FREE(sContext->mFolderFullPidl);
        CLOSE_HANDLE(sContext->mCancelEvent);
        CLOSE_HANDLE(sContext->mFlowSemaphore);
        delete sContext;
        return XPR_FALSE;
    }

    ::CloseHandle(sThread);
    return XPR_TRUE;
}

void DirectoryEnumerationWorker::destroyBatch(Batch *aBatch)
{
    if (XPR_IS_NULL(aBatch))
        return;

    for (std::vector<Item>::iterator sIt = aBatch->mItems.begin();
         sIt != aBatch->mItems.end(); ++sIt)
        COM_FREE(sIt->mPidl);
    if (XPR_IS_NOT_NULL(aBatch->mFlowSemaphore))
    {
        ::ReleaseSemaphore(aBatch->mFlowSemaphore, 1, XPR_NULL);
        CLOSE_HANDLE(aBatch->mFlowSemaphore);
    }
    delete aBatch;
}

DirectoryEnumerationWorker::Batch *DirectoryEnumerationWorker::claimBatch(
    Batch *aBatch)
{
    if (XPR_IS_NULL(aBatch))
        return XPR_NULL;
    return unregisterBatch(aBatch);
}

void DirectoryEnumerationWorker::cleanupOwnerBatches(
    xpr_uint_t aOwnerToken)
{
    std::vector<Batch *> sOwnedBatches;
    {
        std::lock_guard<std::mutex> sLock(batchRegistryMutex());
        std::set<Batch *>::iterator sIt = batchRegistry().begin();
        while (sIt != batchRegistry().end())
        {
            Batch *sBatch = *sIt;
            if (XPR_IS_NOT_NULL(sBatch) &&
                sBatch->mOwnerToken == aOwnerToken)
            {
                sOwnedBatches.push_back(sBatch);
                sIt = batchRegistry().erase(sIt);
            }
            else
                ++sIt;
        }
    }
    for (std::vector<Batch *>::iterator sIt = sOwnedBatches.begin();
         sIt != sOwnedBatches.end(); ++sIt)
        destroyBatch(*sIt);
}
} // namespace fxfile
