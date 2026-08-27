//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "file_op_thread.h"
#include "adaptive_file_operation.h"
#include "modern_shell_file_operation.h"
#include "file_operation_lock_store.h"
#include "main_frame.h"
#include "explorer_ctrl.h"

#include "option.h"

#include "file_op_undo.h"
#include "winapi_ex.h"

#include <set>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace
{
//
// timer id
//
enum
{
    TM_ID_THREAD_END = 1,
};

//
// user defined message
//
enum
{
    WM_POST_END = WM_USER+100,
};
} // namespace anonymous

xpr_sint_t FileOpThread::mRefCount      = 0;
xpr_bool_t FileOpThread::mCompleteFlash = XPR_FALSE;

FileOpThread::FileOpThread(void)
    : mRefIndex(mRefCount++)
    , mShFileOpStruct(XPR_NULL)
    , mSelectItem(XPR_FALSE)
    , mNotifyHwnd(XPR_NULL), mMsg(0)
    , mThread(XPR_NULL), mThreadId(0), mStopEvent(XPR_NULL)
    , mUndo(XPR_TRUE)
    , mOperationSucceeded(XPR_FALSE)
{
    CreateEx(0, AfxRegisterWndClass(CS_GLOBALCLASS), XPR_STRING_LITERAL(""), 0,0,0,0,0,AfxGetMainWnd()->m_hWnd,0);
}

FileOpThread::~FileOpThread(void)
{
    CLOSE_HANDLE(mThread);
    CLOSE_HANDLE(mStopEvent);
    mRefCount--;
}

BEGIN_MESSAGE_MAP(FileOpThread, CWnd)
    ON_WM_TIMER()
    ON_WM_CLOSE()
    ON_MESSAGE(WM_POST_END, OnPostEnd)
END_MESSAGE_MAP()

void FileOpThread::setCompleteFlash(xpr_bool_t aCompleteFlash)
{
    mCompleteFlash = aCompleteFlash;
}

void FileOpThread::setUndo(xpr_bool_t aUndo)
{
    mUndo = aUndo;
}

void FileOpThread::start(xpr_bool_t   aCopy,
                         xpr_tchar_t *aSource,
                         xpr_sint_t   aSourceCount,
                         xpr_tchar_t *aTarget,
                         xpr_bool_t   aDuplicate,
                         xpr_bool_t   aSelectedItem,
                         HWND         aNotifyHwnd,
                         xpr_uint_t   aMsg)
{
    xpr_uint_t sFunc = 0;
    WORD sFlags = 0;
    // Drop Effect
    // 1-Copy, 2-Move, 3-Link, 5-Copy+Link
    if (XPR_IS_FALSE(aCopy))
        sFunc = FO_MOVE;
    else
    {
        sFunc = FO_COPY;
        CString sCompare;
        AfxExtractSubString(sCompare, aSource, 0, '\0');
        sCompare = sCompare.Left(sCompare.ReverseFind('\\'));
        if (_tcscmp(sCompare, aTarget) == 0)
            sFlags |= FOF_RENAMEONCOLLISION;
    }

    // windows vista bug
    //if (aSourceCount > 1)
    //    sFlags |= FOF_MULTIDESTFILES;

    if (XPR_IS_TRUE(aDuplicate))
        sFlags |= FOF_RENAMEONCOLLISION;

    SHFILEOPSTRUCT *sShFileOpStruct = new SHFILEOPSTRUCT;
    memset(sShFileOpStruct, 0, sizeof(SHFILEOPSTRUCT));
    sShFileOpStruct->hwnd   = AfxGetMainWnd()->GetSafeHwnd();
    sShFileOpStruct->wFunc  = sFunc;
    sShFileOpStruct->fFlags = sFlags;
    sShFileOpStruct->pFrom  = aSource;
    sShFileOpStruct->pTo    = aTarget;

    start(sShFileOpStruct, aMsg >= WM_USER, aNotifyHwnd, aMsg);
}

void FileOpThread::start(SHFILEOPSTRUCT *aShFileOpStruct,
                         xpr_bool_t      aSelectItem,
                         HWND            aNotifyHwnd,
                         xpr_uint_t      aMsg)
{
    if (XPR_IS_NOT_NULL(mThread))
        return;

    mShFileOpStruct = aShFileOpStruct;
    mSelectItem     = aSelectItem;
    mNotifyHwnd     = aNotifyHwnd;
    mMsg            = aMsg;

    if ((XPR_IS_TRUE(mUndo)) || (XPR_IS_TRUE(mSelectItem) && XPR_IS_NOT_NULL(mNotifyHwnd) && mMsg > 0))
        mShFileOpStruct->fFlags |= FOF_WANTMAPPINGHANDLE;

    mStopEvent = ::CreateEvent(XPR_NULL, XPR_TRUE, XPR_FALSE, XPR_NULL);
    if (XPR_IS_NULL(mStopEvent))
    {
        mShFileOpStruct->fAnyOperationsAborted = XPR_TRUE;
        PostMessage(WM_CLOSE);
        return;
    }

    mThread = (HANDLE)::_beginthreadex(XPR_NULL, 0, FileOpProc, this, 0, &mThreadId);
    if (XPR_IS_NULL(mThread))
    {
        mShFileOpStruct->fAnyOperationsAborted = XPR_TRUE;
        ::SetEvent(mStopEvent);
        PostMessage(WM_CLOSE);
    }
}

unsigned __stdcall FileOpThread::FileOpProc(LPVOID lpParam)
{
    FileOpThread *sFileOpThread = (FileOpThread *)lpParam;

    if (XPR_IS_NOT_NULL(sFileOpThread))
        sFileOpThread->OnFileOp();

    ::_endthreadex(0);
    return 0;
}

//
// undo
//   - copy, move, rename: post-processing
//   - delete: pre-processing
//
void FileOpThread::OnFileOp(void)
{
    const HRESULT sComResult = ::CoInitializeEx(XPR_NULL, COINIT_APARTMENTTHREADED);

    std::wstring sBlockedPath;
    if (FileOperationLockStore::instance().isOperationBlocked(
            mShFileOpStruct, sBlockedPath))
    {
        std::wstring sMessage = L"FxFile 작업 잠금으로 보호된 경로입니다.\n\n";
        sMessage += sBlockedPath;
        sMessage += L"\n\n편집 > 파일·폴더 잠금 관리에서 잠금을 해제한 뒤 다시 시도하십시오.";
        ::MessageBoxW(mShFileOpStruct->hwnd, sMessage.c_str(),
                      L"FxFile 파일·폴더 잠금", MB_OK | MB_ICONWARNING |
                      MB_TASKMODAL);
        mOperationSucceeded = XPR_FALSE;
        mShFileOpStruct->fAnyOperationsAborted = XPR_TRUE;
        ::SetEvent(mStopEvent);
        if (SUCCEEDED(sComResult))
            ::CoUninitialize();
        PostMessage(WM_POST_END);
        return;
    }

    if (XPR_IS_TRUE(mUndo) && mShFileOpStruct->wFunc == FO_DELETE && XPR_TEST_BITS(mShFileOpStruct->fFlags, FOF_ALLOWUNDO))
    {
        FileOpUndo aFileOpUndo;
        aFileOpUndo.addOperation(mShFileOpStruct);
    }

    captureOperationSources();

    DWORD sAdaptiveError = ERROR_SUCCESS;
    const AdaptiveFileOperation::Result sAdaptiveResult =
        AdaptiveFileOperation::tryExecute(mShFileOpStruct, &sAdaptiveError);

    if (sAdaptiveResult == AdaptiveFileOperation::ResultNotApplicable)
    {
        HRESULT sModernError = S_OK;
        const ModernShellFileOperation::Result sModernResult =
            ModernShellFileOperation::tryExecute(mShFileOpStruct,
                                                 &sModernError);
        if (sModernResult == ModernShellFileOperation::ResultNotApplicable)
        {
            // Compatibility-only path for collision name mappings and
            // legacy multi-destination semantics.
            const int sShellResult = ::SHFileOperation(mShFileOpStruct);
            mOperationSucceeded = (sShellResult == 0 &&
                                   XPR_IS_FALSE(mShFileOpStruct->fAnyOperationsAborted));
        }
        else
        {
            mOperationSucceeded =
                (sModernResult == ModernShellFileOperation::ResultSucceeded);
            mShFileOpStruct->fAnyOperationsAborted =
                (sModernResult == ModernShellFileOperation::ResultCancelled);
        }
    }
    else
    {
        mOperationSucceeded = (sAdaptiveResult == AdaptiveFileOperation::ResultSucceeded);
        mShFileOpStruct->fAnyOperationsAborted =
            (sAdaptiveResult == AdaptiveFileOperation::ResultCancelled);
    }

    // All potentially blocking filesystem probes belong to the worker.  The
    // WM_POST_END handler runs on the main UI thread and must only apply this
    // immutable snapshot.
    captureOperationResult();

    ::SetEvent(mStopEvent);

    if (SUCCEEDED(sComResult))
        ::CoUninitialize();

    PostMessage(WM_POST_END);
}

void FileOpThread::captureOperationSources(void)
{
    mSourceSnapshots.clear();
    if (XPR_IS_NULL(mShFileOpStruct) || XPR_IS_NULL(mShFileOpStruct->pFrom))
        return;

    const xpr_tchar_t *sSource = mShFileOpStruct->pFrom;
    while (*sSource != XPR_STRING_LITERAL('\0'))
    {
        SourceSnapshot sSnapshot;
        sSnapshot.mPath = sSource;
        const DWORD sAttributes = ::GetFileAttributes(sSource);
        sSnapshot.mDirectory =
            (sAttributes != INVALID_FILE_ATTRIBUTES &&
             XPR_TEST_BITS(sAttributes, FILE_ATTRIBUTE_DIRECTORY)) ?
            XPR_TRUE : XPR_FALSE;
        mSourceSnapshots.push_back(sSnapshot);
        sSource += _tcslen(sSource) + 1;
    }
}

void FileOpThread::captureOperationResult(void)
{
    mResultSnapshots.clear();
    if (XPR_IS_NULL(mShFileOpStruct) || mSourceSnapshots.empty())
        return;

    const xpr_bool_t sExactTarget =
        !XPR_TEST_BITS(mShFileOpStruct->fFlags,
                       FOF_RENAMEONCOLLISION | FOF_MULTIDESTFILES);
    const xpr_tchar_t *sTarget =
        (mShFileOpStruct->wFunc == FO_DELETE) ? XPR_NULL :
        mShFileOpStruct->pTo;

    mResultSnapshots.reserve(mSourceSnapshots.size());
    for (std::vector<SourceSnapshot>::const_iterator sIt =
             mSourceSnapshots.begin();
         sIt != mSourceSnapshots.end(); ++sIt)
    {
        ResultSnapshot sSnapshot;
        sSnapshot.mSourcePath = sIt->mPath;
        sSnapshot.mDirectory = sIt->mDirectory;
        sSnapshot.mTargetExists = XPR_FALSE;

        const DWORD sSourceAttributes =
            ::GetFileAttributes(sIt->mPath.c_str());
        const DWORD sSourceError =
            (sSourceAttributes == INVALID_FILE_ATTRIBUTES) ?
            ::GetLastError() : ERROR_SUCCESS;
        sSnapshot.mSourceGone =
            (sSourceAttributes == INVALID_FILE_ATTRIBUTES &&
             (sSourceError == ERROR_FILE_NOT_FOUND ||
              sSourceError == ERROR_PATH_NOT_FOUND)) ?
            XPR_TRUE : XPR_FALSE;

        if ((mShFileOpStruct->wFunc == FO_COPY ||
             mShFileOpStruct->wFunc == FO_MOVE) &&
            XPR_IS_TRUE(sExactTarget) && XPR_IS_NOT_NULL(sTarget))
        {
            const xpr_tchar_t *sLeaf =
                _tcsrchr(sIt->mPath.c_str(), XPR_STRING_LITERAL('\\'));
            sLeaf = XPR_IS_NOT_NULL(sLeaf) ? sLeaf + 1 :
                                             sIt->mPath.c_str();
            sSnapshot.mTargetPath = sTarget;
            if (!sSnapshot.mTargetPath.empty() &&
                sSnapshot.mTargetPath[sSnapshot.mTargetPath.length() - 1] !=
                    XPR_STRING_LITERAL('\\'))
                sSnapshot.mTargetPath += XPR_STRING_LITERAL('\\');
            sSnapshot.mTargetPath += sLeaf;
            sSnapshot.mTargetExists =
                (::GetFileAttributes(sSnapshot.mTargetPath.c_str()) !=
                 INVALID_FILE_ATTRIBUTES) ? XPR_TRUE : XPR_FALSE;
        }

        mResultSnapshots.push_back(sSnapshot);
    }
}

void FileOpThread::reconcileOperationResult(void)
{
    if (XPR_IS_NULL(mShFileOpStruct) || mResultSnapshots.empty())
        return;

    FileOperationReconcileItems sItems;
    sItems.reserve(mResultSnapshots.size());

    for (std::vector<ResultSnapshot>::const_iterator sIt =
             mResultSnapshots.begin();
         sIt != mResultSnapshots.end(); ++sIt)
    {
        FileOperationReconcileItem sItem;
        sItem.mSourcePath = sIt->mSourcePath;
        sItem.mTargetPath = sIt->mTargetPath;
        sItem.mDirectory = sIt->mDirectory;
        sItem.mSourceGone = sIt->mSourceGone;
        sItem.mTargetExists = sIt->mTargetExists;
        sItems.push_back(sItem);
    }

    // Exact UI reconciliation may create PIDLs and sort the view.  Keep this
    // bounded so a large operation cannot monopolise the message pump.
    const xpr_size_t kExactUiLimit = 64;
    MainFrame *sMainFrame = dynamic_cast<MainFrame *>(AfxGetMainWnd());
    if (sItems.size() <= kExactUiLimit && XPR_IS_NOT_NULL(sMainFrame))
    {
        for (xpr_sint_t i = 0; i < MAX_VIEW_SPLIT; ++i)
        {
            ExplorerCtrl *sExplorerCtrl = sMainFrame->getExplorerCtrl(i);
            if (XPR_IS_NOT_NULL(sExplorerCtrl) &&
                ::IsWindow(sExplorerCtrl->GetSafeHwnd()))
                sExplorerCtrl->reconcileFileOperationItems(
                    mShFileOpStruct->wFunc, sItems);
        }
    }

    // Avoid a shell-notification storm.  Large jobs emit one UPDATEDIR per
    // affected parent instead of thousands of item events echoed to all panes.
    const xpr_size_t kExactNotificationLimit = 64;
    if (sItems.size() <= kExactNotificationLimit)
    {
        for (FileOperationReconcileItems::const_iterator sIt = sItems.begin();
             sIt != sItems.end(); ++sIt)
        {
            if (mShFileOpStruct->wFunc == FO_MOVE &&
                XPR_IS_TRUE(sIt->mSourceGone) &&
                XPR_IS_TRUE(sIt->mTargetExists))
            {
                const LONG sEvent = XPR_IS_TRUE(sIt->mDirectory) ?
                                    SHCNE_RENAMEFOLDER : SHCNE_RENAMEITEM;
                ::SHChangeNotify(sEvent, SHCNF_PATH | SHCNF_FLUSHNOWAIT,
                                 sIt->mSourcePath.c_str(),
                                 sIt->mTargetPath.c_str());
            }
            else
            {
                if ((mShFileOpStruct->wFunc == FO_DELETE ||
                     mShFileOpStruct->wFunc == FO_MOVE) &&
                    XPR_IS_TRUE(sIt->mSourceGone))
                {
                    const LONG sEvent = XPR_IS_TRUE(sIt->mDirectory) ?
                                        SHCNE_RMDIR : SHCNE_DELETE;
                    ::SHChangeNotify(sEvent,
                                     SHCNF_PATH | SHCNF_FLUSHNOWAIT,
                                     sIt->mSourcePath.c_str(), XPR_NULL);
                }
                if (mShFileOpStruct->wFunc == FO_COPY &&
                    XPR_IS_TRUE(sIt->mTargetExists))
                {
                    const LONG sEvent = XPR_IS_TRUE(sIt->mDirectory) ?
                                        SHCNE_MKDIR : SHCNE_CREATE;
                    ::SHChangeNotify(sEvent,
                                     SHCNF_PATH | SHCNF_FLUSHNOWAIT,
                                     sIt->mTargetPath.c_str(), XPR_NULL);
                }
            }
        }
    }
    else
    {
        std::set<xpr::string> sChangedDirectories;
        for (FileOperationReconcileItems::const_iterator sIt = sItems.begin();
             sIt != sItems.end(); ++sIt)
        {
            if (XPR_IS_TRUE(sIt->mSourceGone))
            {
                xpr::string sParent(sIt->mSourcePath);
                const xpr_size_t sSplit =
                    sParent.find_last_of(XPR_STRING_LITERAL('\\'));
                if (sSplit != xpr::string::npos)
                    sParent.resize(sSplit == 2 ? 3 : sSplit);
                sChangedDirectories.insert(sParent);
            }
            if (XPR_IS_TRUE(sIt->mTargetExists))
            {
                xpr::string sParent(sIt->mTargetPath);
                const xpr_size_t sSplit =
                    sParent.find_last_of(XPR_STRING_LITERAL('\\'));
                if (sSplit != xpr::string::npos)
                    sParent.resize(sSplit == 2 ? 3 : sSplit);
                sChangedDirectories.insert(sParent);
            }
        }
        for (std::set<xpr::string>::const_iterator sIt =
                 sChangedDirectories.begin();
             sIt != sChangedDirectories.end(); ++sIt)
        {
            ::SHChangeNotify(SHCNE_UPDATEDIR,
                             SHCNF_PATH | SHCNF_FLUSHNOWAIT,
                             sIt->c_str(), XPR_NULL);
        }
    }
}

LRESULT FileOpThread::OnPostEnd(WPARAM, LPARAM)
{
    SetTimer(TM_ID_THREAD_END, 10, XPR_NULL);

    reconcileOperationResult();

    // [bug patched] 2007/08/15, 2007/10/27
    // If file operation is copy or move, then this program hide behind foreground. I prevented it.
    // And, if file operation completed on background, taskbar flush like FlushWindow API function.
    xpr_bool_t sSetForegournd = XPR_FALSE;

    HWND sHwnd;
    HWND sParentHwnd;

    sHwnd = ::GetForegroundWindow();
    while (true)
    {
        sParentHwnd = ::GetParent(sHwnd);
        if (XPR_IS_NULL(sParentHwnd))
            break;

        sHwnd = sParentHwnd;
    }

    sSetForegournd = (sHwnd == mShFileOpStruct->hwnd) || (sHwnd == m_hWnd);

    if (XPR_IS_TRUE(sSetForegournd))
    {
        SetForceForegroundWindow(mShFileOpStruct->hwnd);
    }
    else
    {
        if (XPR_IS_TRUE(mCompleteFlash))
        {
            FLASHWINFO sFlashWInfo = {0};
            sFlashWInfo.cbSize  = sizeof(sFlashWInfo);
            sFlashWInfo.dwFlags = FLASHW_TRAY;
            sFlashWInfo.hwnd    = mShFileOpStruct->hwnd;
            sFlashWInfo.uCount  = 3;
            ::FlashWindowEx(&sFlashWInfo);
        }
    }

    return 0;
}

void FileOpThread::OnTimer(UINT_PTR nIDEvent) 
{
    if (nIDEvent == TM_ID_THREAD_END)
    {
        DWORD sResult;
        DWORD sExitCode = -1;
        xpr_bool_t sSucceeded;

        sResult    = ::WaitForSingleObject(mStopEvent, 0);
        sSucceeded = ::GetExitCodeThread(mThread, &sExitCode);
        if (sResult == WAIT_OBJECT_0 && sExitCode == 0 && XPR_IS_TRUE(sSucceeded))
        {
            KillTimer(TM_ID_THREAD_END);

            if (XPR_IS_TRUE(mOperationSucceeded) &&
                XPR_IS_NOT_NULL(mNotifyHwnd) && mMsg > 0 && XPR_IS_TRUE(mSelectItem))
            {
                PasteSelItems sPasteSelItems = {0};
                sPasteSelItems.mSource           = mShFileOpStruct->pFrom;
                sPasteSelItems.mTarget           = mShFileOpStruct->pTo;
                sPasteSelItems.mHandleToMappings = (HandleToMappings *)mShFileOpStruct->hNameMappings;

                DWORD_PTR sNotifyResult = 0;
                ::SendMessageTimeout(mNotifyHwnd, mMsg,
                                     (WPARAM)&sPasteSelItems, 0,
                                     SMTO_ABORTIFHUNG | SMTO_BLOCK,
                                     2000, &sNotifyResult);
            }

            if (XPR_IS_TRUE(mOperationSucceeded) &&
                XPR_IS_TRUE(mUndo) && mShFileOpStruct->wFunc != FO_DELETE)
            {
                FileOpUndo aFileOpUndo;
                aFileOpUndo.addOperation(mShFileOpStruct);
            }

            if (XPR_IS_NOT_NULL(mShFileOpStruct->hNameMappings))
            {
                ::SHFreeNameMappings(mShFileOpStruct->hNameMappings);
                mShFileOpStruct->hNameMappings = XPR_NULL;
            }

            CWnd::OnTimer(nIDEvent);

            DestroyWindow();
            return;
        }
    }

    CWnd::OnTimer(nIDEvent);
}

void FileOpThread::OnClose(void) 
{
    DestroyWindow();
}

xpr_bool_t FileOpThread::DestroyWindow(void)
{
    if (XPR_IS_NOT_NULL(mThread))
    {
        DWORD sExitCode = STILL_ACTIVE;
        if (::GetExitCodeThread(mThread, &sExitCode) == XPR_FALSE ||
            sExitCode == STILL_ACTIVE)
        {
            return XPR_FALSE;
        }
    }

    if (XPR_IS_NOT_NULL(mShFileOpStruct))
    {
        xpr_tchar_t *sSource = (xpr_tchar_t *)mShFileOpStruct->pFrom;
        xpr_tchar_t *sTarget = (xpr_tchar_t *)mShFileOpStruct->pTo;
        XPR_SAFE_DELETE_ARRAY(sSource);
        XPR_SAFE_DELETE_ARRAY(sTarget);
        XPR_SAFE_DELETE(mShFileOpStruct);
    }

    CLOSE_HANDLE(mThread);
    CLOSE_HANDLE(mStopEvent);

    return CWnd::DestroyWindow();
}

void FileOpThread::PostNcDestroy(void) 
{
    delete this;
}
} // namespace fxfile
