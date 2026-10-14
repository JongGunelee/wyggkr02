//
// Copyright (c) 2001-2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "rename_helper.h"
#include "file_op_undo.h"
#include "file_operation_lock_store.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
RenameHelper::RenameHelper(RenameStyle aRenameStyle)
    : mRenameStyle(aRenameStyle)
    , mSelItem0(0)
    , mSelItem1(-1)
    , mShellFolder(XPR_NULL)
    , mPidl(XPR_NULL)
    , mFullPidl(XPR_NULL)
    , mNewPidl(XPR_NULL)
    , mNewFullPidl(XPR_NULL)
    , mShellAttributes(0)
    , mCachedIconIndex(-1)
{
}

RenameHelper::~RenameHelper(void)
{
    COM_RELEASE(mShellFolder);
    COM_FREE(mPidl);
    COM_FREE(mFullPidl);
    COM_FREE(mNewPidl);
    COM_FREE(mNewFullPidl);
}

void RenameHelper::setRenameStyle(RenameStyle eRenameStyle)
{
    mRenameStyle = eRenameStyle;
}

void RenameHelper::setItem(LPSHELLFOLDER aShellFolder, LPCITEMIDLIST aPidl, LPITEMIDLIST aFullPidl, xpr_ulong_t aShellAttributes,
                           const xpr_tchar_t *aCachedPath, xpr_sint_t aCachedIconIndex)
{
    // A modal editor pumps folder-change messages.  The row can be removed
    // while the dialog is open, so borrowed row COM/PIDL pointers are unsafe.
    if (XPR_IS_NOT_NULL(aShellFolder)) aShellFolder->AddRef();
    LPITEMIDLIST sPidl = fxfile::base::Pidl::clone(aPidl);
    LPITEMIDLIST sFullPidl = fxfile::base::Pidl::clone(aFullPidl);
    COM_RELEASE(mShellFolder);
    COM_FREE(mPidl);
    COM_FREE(mFullPidl);
    COM_FREE(mNewPidl);
    COM_FREE(mNewFullPidl);
    mShellFolder     = aShellFolder;
    mPidl            = sPidl;
    mFullPidl        = sFullPidl;
    mShellAttributes = aShellAttributes;
    mCachedIconIndex = aCachedIconIndex;

    mEditName.clear();
    mSourcePath.clear();
    mKeepExt.clear();
    mParsingExt.clear();
    mSelItem0 = 0;
    mSelItem1 = -1;

    xpr_bool_t sFileSystem = XPR_TEST_BITS(mShellAttributes, SFGAO_FILESYSTEM);
    xpr_bool_t sFolder     = XPR_TEST_BITS(mShellAttributes, SFGAO_FOLDER);

    if (XPR_IS_TRUE(sFileSystem))
    {
        if (XPR_IS_NOT_NULL(aCachedPath) && aCachedPath[0] != 0)
            mSourcePath = aCachedPath;
        else
        {
            xpr_tchar_t sPath[XPR_MAX_PATH + 1] = {0};
            if (::SHGetPathFromIDList(mFullPidl, sPath)) mSourcePath = sPath;
        }
    }

    // Ordinary filesystem names are already in the pane model.  Opening F2
    // must not query display-name/overlay providers again (AV/cloud latency).
    const xpr_bool_t sPlainFileSystem = XPR_IS_TRUE(sFileSystem) &&
        mSourcePath.length() > 3 && XPR_TEST_NONE_BITS(mShellAttributes, SFGAO_LINK);
    if (XPR_IS_TRUE(sPlainFileSystem))
        mEditName = mSourcePath.substr(mSourcePath.rfind(XPR_STRING_LITERAL('\\')) + 1);
    else if (XPR_IS_TRUE(sFileSystem) && XPR_IS_FALSE(sFolder))
        GetName(mShellFolder, mPidl, SHGDN_INFOLDER | SHGDN_FORPARSING, mEditName);
    else
        GetName(mShellFolder, mPidl, SHGDN_INFOLDER | SHGDN_FOREDITING, mEditName);

    if (XPR_IS_FALSE(sFolder))
    {
        if (XPR_IS_TRUE(sFileSystem) && mRenameStyle != RenameStyleNone)
        {
            xpr_tchar_t sFullName[XPR_MAX_PATH + 1] = {0};
            if (XPR_IS_TRUE(sPlainFileSystem))
                _tcscpy(sFullName, mEditName.c_str());
            else
                GetName(mShellFolder, mPidl, SHGDN_INFOLDER | SHGDN_FORPARSING, sFullName);

            if (_tcsicmp(mEditName.c_str(), sFullName) == 0)
            {
                xpr_tchar_t *sExt = (xpr_tchar_t *)GetFileExt(mEditName.c_str());
                if (XPR_IS_NOT_NULL(sExt))
                {
                    if (mRenameStyle == RenameStyleKeepExt)
                    {
                        mKeepExt = sExt;
                        sExt[0] = '\0';
                    }
                    else if (mRenameStyle == RenameStyleSelExceptForExt)
                    {
                        mSelItem1 = (xpr_sint_t)(sExt - mEditName.c_str());
                    }
                }
            }
        }
    }
}

const xpr_tchar_t *RenameHelper::getEditName(void)
{
    return mEditName.c_str();
}

void RenameHelper::getEditSel(xpr_sint_t *aSelItem0, xpr_sint_t *aSelItem1)
{
    if (XPR_IS_NOT_NULL(aSelItem0)) *aSelItem0 = mSelItem0;
    if (XPR_IS_NOT_NULL(aSelItem1)) *aSelItem1 = mSelItem1;
}

xpr_sint_t RenameHelper::getIconIndex(void)
{
    // Use a resolved list icon only.  Cosmetic provider work is never a
    // prerequisite for showing a responsive rename dialog.
    return mCachedIconIndex;
}

RenameHelper::Result RenameHelper::rename(HWND aHwnd, const xpr_tchar_t *aNewName)
{
    if (XPR_IS_NULL(aNewName))
        return ResultInvalidParameter;

    mNewName = aNewName;

    if (mNewName.empty() == XPR_TRUE)
        return ResultEmptiedName;

    xpr_size_t nNewNameLen = mNewName.length();
    if (mNewName[nNewNameLen-1] == '.')
        mNewName.erase(nNewNameLen-1);

    mNewName += mKeepExt.c_str();

    xpr_size_t sInputBytes;
    xpr_size_t sOutputBytes;
    xpr_wchar_t sUtf16NewName[XPR_MAX_PATH + 1] = {0};
    sInputBytes = mNewName.length() * sizeof(xpr_tchar_t);
    sOutputBytes = XPR_MAX_PATH * sizeof(xpr_wchar_t);
    XPR_TCS_TO_UTF16(mNewName.c_str(), &sInputBytes, sUtf16NewName, &sOutputBytes);
    sUtf16NewName[sOutputBytes / sizeof(xpr_wchar_t)] = 0;

    xpr::string sOldPath;
    xpr::string sNewPath;
    sOldPath = mSourcePath;
    if (sOldPath.empty()) GetName(mShellFolder, mPidl, SHGDN_FORPARSING, sOldPath);
    sNewPath = sOldPath;

    if (FileOperationLockStore::instance().affectsLockedPath(
            std::wstring(sOldPath.c_str())))
    {
        CWnd::FromHandle(aHwnd)->MessageBox(
            XPR_STRING_LITERAL("FxFile 작업 잠금으로 보호된 항목입니다.\n\n편집 > 파일·폴더 잠금 관리에서 잠금을 해제한 뒤 다시 시도하십시오."),
            XPR_STRING_LITERAL("FxFile 파일·폴더 잠금"),
            MB_OK | MB_ICONWARNING);
        return ResultUserCancel;
    }

    xpr_size_t sFind = sNewPath.rfind(XPR_STRING_LITERAL('\\'));
    if (sFind != xpr::string::npos)
    {
        sNewPath.erase(sFind+1);

        sNewPath += mNewName;
        sNewPath += mParsingExt;
    }

    // An unchanged label is a successful no-op, not a failed filesystem
    // move.  Avoid sending it through MoveFile/SetNameOf, which can produce a
    // misleading collision and reopen the editor indefinitely.
    if (_tcscmp(sOldPath.c_str(), sNewPath.c_str()) == 0)
    {
        mNewPidl = fxfile::base::Pidl::clone(mPidl);
        mNewFullPidl = fxfile::base::Pidl::clone(mFullPidl);
        return (XPR_IS_NOT_NULL(mNewPidl) &&
                XPR_IS_NOT_NULL(mNewFullPidl)) ?
               ResultSucceeded : ResultUnknownError;
    }

    HRESULT sHResult = E_FAIL;

    xpr_bool_t sFileSystem =  XPR_TEST_BITS(mShellAttributes, SFGAO_FILESYSTEM);
    xpr_bool_t sFile       = !XPR_TEST_BITS(mShellAttributes, SFGAO_FOLDER);

    // [remark]
    // Some file name can rename using MoveFile API function, but it can not rename using IShellFolder::SetNameOf method.
    // This result code of IShellFolder::SetNameOf is 0x800704c7.
    // This bug is confirmed to WinXP sp2 or higher
    // ex) 'test (test,1981) - test.ogg' -> '@test (test,1981) - test.ogg'

    xpr_bool_t sMoveFile = XPR_FALSE;
    if (XPR_IS_TRUE(sFileSystem))
    {
        if (VerifyFileName(mNewName) == XPR_FALSE)
        {
            const xpr_tchar_t *sMsg = gApp.loadString(XPR_STRING_LITERAL("popup.rename.msg.invalid_filename"));
            CWnd::FromHandle(aHwnd)->MessageBox(sMsg, XPR_NULL, MB_OK | MB_ICONSTOP);
            return ResultInvalidName;
        }

        xpr_sint_t sMsgId;
        DWORD sFileAttributes = ::GetFileAttributes(sOldPath.c_str());

        if (sOldPath.length() <= 3)
        {
            xpr_tchar_t sMsg[XPR_MAX_PATH * 3 + 1] = {0};
            _stprintf(sMsg, gApp.loadFormatString(XPR_STRING_LITERAL("popup.rename.msg.confirm_drive_rename"), XPR_STRING_LITERAL("%c,%s,%s")), sOldPath[0], mEditName.c_str(), mNewName.c_str());
            sMsgId = ::MessageBox(aHwnd, sMsg, gApp.loadString(XPR_STRING_LITERAL("popup.rename.msg.title")), MB_YESNO | MB_ICONQUESTION);
            if (sMsgId == IDNO)
                return ResultUserCancel;
        }
        else if (sFileAttributes != INVALID_FILE_ATTRIBUTES &&
                 (XPR_TEST_BITS(sFileAttributes, FILE_ATTRIBUTE_SYSTEM) || XPR_TEST_BITS(sFileAttributes, FILE_ATTRIBUTE_READONLY)))
        {
            xpr::string sAttributes;
            if (XPR_TEST_BITS(sFileAttributes, FILE_ATTRIBUTE_READONLY))
                sAttributes = gApp.loadString(XPR_STRING_LITERAL("popup.rename.msg.confirm_file_rename.read_only_file"));
            else
                sAttributes = gApp.loadString(XPR_STRING_LITERAL("popup.rename.msg.confirm_file_rename.system_file"));

            xpr::string sOldFileName = sOldPath.substr(sOldPath.rfind(XPR_STRING_LITERAL('\\'))+1);
            xpr::string sNewFileName = sNewPath.substr(sNewPath.rfind(XPR_STRING_LITERAL('\\'))+1);

            xpr_tchar_t sMsg[XPR_MAX_PATH * 3 + 1] = {0};
            _stprintf(sMsg, gApp.loadFormatString(XPR_STRING_LITERAL("popup.rename.msg.confirm_file_rename"), XPR_STRING_LITERAL("%s,%s,%s")), sOldFileName.c_str(), sAttributes.c_str(), sNewFileName.c_str());
            sMsgId = ::MessageBox(aHwnd, sMsg, gApp.loadString(XPR_STRING_LITERAL("popup.rename.msg.title")), MB_YESNO | MB_ICONQUESTION);
            if (sMsgId == IDNO)
                return ResultUserCancel;
        }

        if (::MoveFile(sOldPath.c_str(), sNewPath.c_str()) == XPR_TRUE)
        {
            LPITEMIDLIST sFullPidl = fxfile::base::Pidl::create(sNewPath.c_str());
            if (XPR_IS_NOT_NULL(sFullPidl))
            {
                LPSHELLFOLDER sShellFolder = XPR_NULL;
                LPCITEMIDLIST sPidl        = XPR_NULL;
                fxfile::base::Pidl::getSimplePidl(sFullPidl, sShellFolder, sPidl);

                mNewPidl = fxfile::base::Pidl::clone(sPidl);

                COM_RELEASE(sShellFolder);
                COM_FREE(sFullPidl);
            }

            sMoveFile = XPR_TRUE;
            sHResult = S_OK;
        }
    }

    if (FAILED(sHResult))
    {
        xpr_uint_t sFlags = SHGDN_INFOLDER;
        if (XPR_IS_TRUE(sFileSystem) && XPR_IS_TRUE(sFile))
            sFlags |= SHGDN_FORPARSING;

        sHResult = mShellFolder->SetNameOf(aHwnd, mPidl, sUtf16NewName, sFlags, &mNewPidl);
    }

    if (FAILED(sHResult))
        return ResultUnknownError;

    mNewFullPidl = fxfile::base::Pidl::getFullPidl(mShellFolder, mNewPidl);

    xpr_slong_t sEventId = SHCNE_RENAMEITEM;
    xpr_uint_t sFlags = SHCNF_IDLIST;
    // SHCNF_IDLIST requires desktop-relative (absolute) PIDLs, not the
    // child identifiers supplied to IShellFolder::SetNameOf.
    const void *sItem1 = mFullPidl;
    const void *sItem2 = mNewFullPidl;

    if (XPR_IS_FALSE(sFile))
        sEventId = SHCNE_RENAMEFOLDER;

    if (XPR_IS_TRUE(sMoveFile))
    {
        sFlags = SHCNF_PATH;
        sItem1 = sOldPath.c_str();
        sItem2 = sNewPath.c_str();
    }

    // Deliver promptly, but never wait for every external Shell listener to
    // finish on the editor/UI thread.  Pane watchers reconcile independently.
    sFlags |= SHCNF_FLUSHNOWAIT;
    if (XPR_IS_NOT_NULL(sItem1) && XPR_IS_NOT_NULL(sItem2))
        ::SHChangeNotify(sEventId, sFlags, sItem1, sItem2);

    if (XPR_IS_TRUE(sFileSystem))
    {
        FileOpUndo sFileOpUndo;
        sFileOpUndo.addRename(sOldPath, sNewPath);
    }

    return ResultSucceeded;
}

LPITEMIDLIST RenameHelper::getNewItemIDList(xpr_bool_t aNull)
{
    LPITEMIDLIST sNewPidl = mNewPidl;

    if (XPR_IS_TRUE(aNull))
        mNewPidl = XPR_NULL;

    return sNewPidl;
}

const xpr_tchar_t *RenameHelper::getNewName(void)
{
    return mNewName.c_str();
}
} // namespace fxfile
