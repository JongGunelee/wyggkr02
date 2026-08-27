//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "context_menu.h"
#include "option.h"
#include "file_scrap.h"
#include "archive_engine.h"
#include "archive_progress_dlg.h"
#include "main_frame.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#pragma warning(disable:4244)
#pragma warning(disable:4312)

namespace fxfile
{
namespace
{
const xpr_tchar_t PROP_CONTEXT_MENU[] = XPR_STRING_LITERAL("OldWndProc");
} // namespace anonymous

xpr_bool_t ContextMenu::mShowFileScrapMenu = XPR_FALSE;
xpr_bool_t ContextMenu::mAnimationMenu     = XPR_FALSE;

ContextMenu::ContextMenu(HWND aHwnd)
    : mHwnd(aHwnd)
    , mShellFolder(XPR_NULL), mPidls(XPR_NULL), mCount(0)
    , mMenu(XPR_NULL)
    , mContextMenu(XPR_NULL), mContextMenu2(XPR_NULL), mContextMenu3(XPR_NULL)
    , mFirstId(CMID_DEF_ID_FIRST)
    , mSubclassData(XPR_NULL)
{
}

ContextMenu::~ContextMenu(void)
{
    destroy();
}

void ContextMenu::setFileScrapMenu(xpr_bool_t aShowFileScrapMenu)
{
    mShowFileScrapMenu = aShowFileScrapMenu;
}

void ContextMenu::setAnimationMenu(xpr_bool_t aAnimationMenu)
{
    mAnimationMenu = aAnimationMenu;
}

xpr_bool_t ContextMenu::init(LPSHELLFOLDER aShellFolder, LPCITEMIDLIST *aPidls, xpr_uint_t aCount)
{
    if (XPR_IS_NULL(aShellFolder) || XPR_IS_NULL(aPidls) || aCount == 0)
        return XPR_FALSE;

    mShellFolder = aShellFolder;
    mPidls = aPidls;
    mCount = aCount;

    HRESULT sComResult;
    sComResult = mShellFolder->GetUIObjectOf(mHwnd, mCount, (LPCITEMIDLIST *)mPidls, IID_IContextMenu, XPR_NULL, (LPVOID *)&mContextMenu);
    if (FAILED(sComResult))
        return XPR_FALSE;

    sComResult = mContextMenu->QueryInterface(IID_IContextMenu3, (LPVOID *)&mContextMenu3);
    if (FAILED(sComResult))
    {
        mContextMenu->QueryInterface(IID_IContextMenu2, (LPVOID *)&mContextMenu2);
    }

    return XPR_TRUE;
}

xpr_bool_t ContextMenu::init(LPSHELLFOLDER aShellFolder)
{
    if (XPR_IS_NULL(aShellFolder))
        return XPR_FALSE;

    mShellFolder = aShellFolder;

    HRESULT sComResult;
    sComResult = mShellFolder->CreateViewObject(mHwnd, IID_IContextMenu, (LPVOID*)&mContextMenu);
    if (FAILED(sComResult))
        return XPR_FALSE;

    sComResult = mContextMenu->QueryInterface(IID_IContextMenu3, (LPVOID *)&mContextMenu3);
    if (FAILED(sComResult))
    {
        mContextMenu->QueryInterface(IID_IContextMenu2, (LPVOID *)&mContextMenu2);
    }

    return XPR_TRUE;
}

LRESULT CALLBACK ContextMenu::subclassWndProc(HWND aWnd, xpr_uint_t aMsg, WPARAM wParam, LPARAM lParam)
{
    LPSUBCLASSDATA sSubclassData = (LPSUBCLASSDATA)::GetProp(aWnd, PROP_CONTEXT_MENU);
    if (XPR_IS_NOT_NULL(sSubclassData))
    {
        LPCONTEXTMENU2 sContextMenu2 = sSubclassData->mContextMenu2;
        LPCONTEXTMENU3 sContextMenu3 = sSubclassData->mContextMenu3;

        switch (aMsg)
        {
        case WM_MENUCHAR:
            if (XPR_IS_NOT_NULL(sContextMenu3))
            {
                LRESULT lResult = 0;
                sContextMenu3->HandleMenuMsg2(aMsg, wParam, lParam, &lResult);
                return lResult;
            }

        case WM_INITMENUPOPUP:
        case WM_DRAWITEM:
        case WM_MEASUREITEM:
            if (XPR_IS_NOT_NULL(sContextMenu3))
            {
                LRESULT lResult = 0;
                sContextMenu3->HandleMenuMsg2(aMsg, wParam, lParam, &lResult);
            }
            else if (XPR_IS_NOT_NULL(sContextMenu2))
            {
                if (SUCCEEDED(sContextMenu2->HandleMenuMsg(aMsg, wParam, lParam)))
                    return XPR_FALSE;
            }

            return (aMsg == WM_INITMENUPOPUP ? 0 : TRUE);
        }

        return CallWindowProc(sSubclassData->mOldWndProc, aWnd, aMsg, wParam, lParam);
    }

    return DefWindowProc(aWnd, aMsg, wParam, lParam);
}

void ContextMenu::getInterface(LPCONTEXTMENU *aContextMenu, LPCONTEXTMENU2 *aContextMenu2) const
{
    *aContextMenu  = mContextMenu;
    *aContextMenu2 = mContextMenu2;
}

xpr_bool_t ContextMenu::getMenu(CMenu *aMenu, xpr_uint_t aFirstId, xpr_uint_t aQueryFlags)
{
    if (XPR_IS_NULL(aMenu) || IsWindow(mHwnd) == XPR_FALSE)
        return XPR_FALSE;

    HRESULT sComResult = E_FAIL;
    if (XPR_IS_NOT_NULL(mContextMenu3))
    {
        sComResult = mContextMenu3->QueryContextMenu(aMenu->m_hMenu, 0, aFirstId, aFirstId + 0x7FFF, aQueryFlags);
    }
    else if (XPR_IS_NOT_NULL(mContextMenu2))
    {
        sComResult = mContextMenu2->QueryContextMenu(aMenu->m_hMenu, 0, aFirstId, aFirstId + 0x7FFF, aQueryFlags);
    }

    if (FAILED(sComResult))
    {
        sComResult = mContextMenu->QueryContextMenu(aMenu->m_hMenu, 0, aFirstId, aFirstId + 0x7FFF, aQueryFlags);
        COM_RELEASE(mContextMenu2);
        COM_RELEASE(mContextMenu3);

        if (FAILED(sComResult))
        {
            return XPR_FALSE;
        }
    }

    if (XPR_IS_NOT_NULL(mContextMenu2) || XPR_IS_NOT_NULL(mContextMenu3))
    {
        int nIndex = 
#if defined(XPR_CFG_COMPILER_64BIT)
            GWLP_WNDPROC;
#else
            GWL_WNDPROC;
#endif

        mSubclassData = new SUBCLASSDATA;
        mSubclassData->mContextMenu2 = mContextMenu2;
        mSubclassData->mContextMenu3 = mContextMenu3;
        mSubclassData->mOldWndProc   = (WNDPROC)::SetWindowLongPtr(mHwnd, nIndex, (LONG_PTR)subclassWndProc);
        ::SetProp(mHwnd, PROP_CONTEXT_MENU, (HANDLE)mSubclassData);
    }

    if (XPR_IS_NOT_NULL(mPidls) && mCount > 0)
    {
        // Place '7-Zip 내장 압축 관리 (&7)' at the very top (index 0) of the context menu
        xpr_sint_t nInsert = 0;

        HMENU hZipMenu = ::CreatePopupMenu();
        if (hZipMenu != NULL)
        {
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_COMPRESS_7Z, _T("7z으로 압축하기 (최고 압축률 .7z)"));
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_COMPRESS_ZIP, _T("Zip으로 압축하기 (표준 호환 .zip)"));
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_COMPRESS_CUSTOM, _T("압축 설정하여 압축하기..."));
            ::AppendMenu(hZipMenu, MF_SEPARATOR, 0, NULL);
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_EXTRACT_HERE, _T("여기에 압축 풀기"));
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_EXTRACT_TO_FOLDER, _T("각각의 폴더에 압축 풀기"));
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_EXTRACT_CUSTOM, _T("폴더를 지정하여 압축 풀기..."));
            ::AppendMenu(hZipMenu, MF_SEPARATOR, 0, NULL);
            ::AppendMenu(hZipMenu, MF_STRING, aFirstId + CMID_7ZIP_OPEN_ARCHIVE, _T("압축 파일 열어보기 (가상 폴더 탐색)"));

            ::InsertMenu(aMenu->m_hMenu, nInsert++, MF_BYPOSITION | MF_POPUP, (UINT_PTR)hZipMenu, _T("7-Zip 내장 압축 관리 (&7)"));
            ::InsertMenu(aMenu->m_hMenu, nInsert++, MF_BYPOSITION | MF_SEPARATOR, 0, XPR_NULL);
        }

        if (XPR_IS_TRUE(mShowFileScrapMenu))
        {
            xpr_sint_t nScrapInsert = aMenu->GetMenuItemCount();
            if (nScrapInsert >= 2) nScrapInsert -= 2;
            ::InsertMenu(aMenu->m_hMenu, nScrapInsert, MF_BYPOSITION, aFirstId + CMID_FILE_SCRAP, gApp.loadString(XPR_STRING_LITERAL("cmd.shell.add_to_scrap")));
        }
    }

    return XPR_TRUE;
}

xpr_uint_t ContextMenu::getFirstId(void) const
{
    return mFirstId;
}

xpr_uint_t ContextMenu::trackPopupMenu(xpr_uint_t aFlags, LPPOINT aPoint, xpr_uint_t aQueryFlags)
{
    if (XPR_IS_FALSE(mAnimationMenu))
        aFlags |= 0x4000L;

    CMenu sMenu;
    if (sMenu.CreatePopupMenu() == XPR_FALSE)
        return -1;

    xpr_uint_t sId = -1;
    if (getMenu(&sMenu, CMID_DEF_ID_FIRST, aQueryFlags) == XPR_TRUE)
    {
        sId = ::TrackPopupMenuEx(sMenu.m_hMenu, aFlags, aPoint->x, aPoint->y, mHwnd, XPR_NULL);
        mMenu = sMenu.Detach();
    }

    return sId;
}

xpr_bool_t ContextMenu::getCommandVerb(xpr_uint_t aId, xpr_tchar_t *aVerb, xpr_sint_t aMaxVerbLength) const
{
    if (XPR_IS_NULL(mContextMenu))
        return XPR_FALSE;

    if (aId == 0xFFFFFFFF)
        return XPR_FALSE;

    HRESULT sComResult;
    sComResult = mContextMenu->GetCommandString(aId, GCS_VERB, XPR_NULL, (xpr_char_t *)aVerb, aMaxVerbLength);

    return SUCCEEDED(sComResult);
}

xpr_bool_t ContextMenu::invokeCommand(xpr_uint_t aId)
{
    if (XPR_IS_TRUE(mShowFileScrapMenu) && aId == CMID_FILE_SCRAP)
    {
        FileScrap &sFileScrap = FileScrap::instance();
        xpr_uint_t sGroupId = sFileScrap.getCurGroupId();

        sFileScrap.addItem(sGroupId, mShellFolder, mPidls, mCount);
        return XPR_TRUE;
    }

    // 7-Zip Command Handlers
    if (aId >= CMID_7ZIP_COMPRESS_7Z && aId <= CMID_7ZIP_OPEN_ARCHIVE)
    {
        std::vector<std::wstring> sPaths;
        for (xpr_uint_t i = 0; i < mCount; ++i)
        {
            xpr_tchar_t sBuf[MAX_PATH] = {0};
            if (fxfile::base::Pidl::getName(mShellFolder, mPidls[i], SHGDN_FORPARSING, sBuf, MAX_PATH))
            {
                if (sBuf[0] != 0)
                    sPaths.push_back(sBuf);
            }
        }

        if (sPaths.empty())
            return XPR_TRUE;

        if (aId == CMID_7ZIP_COMPRESS_7Z || aId == CMID_7ZIP_COMPRESS_ZIP || aId == CMID_7ZIP_COMPRESS_CUSTOM)
        {
            std::wstring sArchiveTarget;
            if (sPaths.size() == 1)
            {
                sArchiveTarget = sPaths[0] + (aId == CMID_7ZIP_COMPRESS_ZIP ? L".zip" : L".7z");
            }
            else
            {
                size_t sSlash = sPaths[0].find_last_of(L"\\/");
                std::wstring sParent = (sSlash != std::wstring::npos) ? sPaths[0].substr(0, sSlash) : L"";
                std::wstring sBaseName = (sSlash != std::wstring::npos) ? sPaths[0].substr(sSlash + 1) : L"Archive";
                sArchiveTarget = sParent + L"\\" + sBaseName + (aId == CMID_7ZIP_COMPRESS_ZIP ? L".zip" : L".7z");
            }

            archive::ArchiveFormat sFmt = (aId == CMID_7ZIP_COMPRESS_ZIP) ? archive::FormatZip : archive::Format7z;

            // Create high-quality native progress dialog
            archive::ArchiveProgressDialog sProgressDlg;
            std::wstring sDlgTitle = (sFmt == archive::FormatZip) ? L"7-Zip ZIP 파일 압축" : L"7-Zip 7z 파일 압축";
            std::wstring sMainInst = L"파일 및 폴더를 압축하는 중입니다...";
            sProgressDlg.create(mHwnd, sDlgTitle, sMainInst);

            // Progress callback
            auto sCallback = [&sProgressDlg](int aPercent, const std::wstring &aFile) -> bool {
                if (sProgressDlg.isCancelled())
                    return false;
                sProgressDlg.update(aPercent, aFile);
                return true;
            };

            bool sSuccess = archive::ArchiveEngine::createArchive(sArchiveTarget, sPaths, sFmt, 5, sCallback);
            sProgressDlg.close();

            if (sSuccess)
            {
                CString sMsg;
                sMsg.Format(_T("7-Zip 압축이 성공적으로 완료되었습니다:\n%s"), sArchiveTarget.c_str());
                ::MessageBox(mHwnd, sMsg, _T("7-Zip 내장 압축 관리"), MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                if (!archive::ArchiveEngine::is7zAvailable())
                {
                    ::MessageBox(mHwnd, _T("7-Zip이 설치되어 있지 않습니다.\n'C:\\Program Files\\7-Zip\\7z.exe'를 확인해 주십시오."),
                                 _T("7-Zip 내장 압축 관리"), MB_OK | MB_ICONWARNING);
                }
                else
                {
                    ::MessageBox(mHwnd, _T("압축 파일 생성 중 오류가 발생하였거나 취소되었습니다."), _T("7-Zip 내장 압축 관리"), MB_OK | MB_ICONWARNING);
                }
            }
            return XPR_TRUE;
        }
        else if (aId == CMID_7ZIP_EXTRACT_HERE || aId == CMID_7ZIP_EXTRACT_TO_FOLDER || aId == CMID_7ZIP_EXTRACT_CUSTOM)
        {
            bool sAnySuccess = false;

            // Create high-quality native progress dialog
            archive::ArchiveProgressDialog sProgressDlg;
            sProgressDlg.create(mHwnd, L"7-Zip 압축 풀기", L"압축 파일의 내용을 추출하는 중입니다...");

            auto sCallback = [&sProgressDlg](int aPercent, const std::wstring &aFile) -> bool {
                if (sProgressDlg.isCancelled())
                    return false;
                sProgressDlg.update(aPercent, aFile);
                return true;
            };

            for (size_t i = 0; i < sPaths.size(); ++i)
            {
                archive::ArchiveEngine sEngine;
                if (sEngine.openArchive(sPaths[i]))
                {
                    std::wstring sDestDir = sPaths[i];
                    if (aId == CMID_7ZIP_EXTRACT_TO_FOLDER)
                    {
                        size_t sDot = sDestDir.find_last_of(L'.');
                        if (sDot != std::wstring::npos)
                            sDestDir = sDestDir.substr(0, sDot);
                    }
                    else
                    {
                        size_t sSlash = sDestDir.find_last_of(L"\\/");
                        if (sSlash != std::wstring::npos)
                            sDestDir = sDestDir.substr(0, sSlash);
                    }
                    if (sEngine.extractAll(sDestDir, sCallback))
                        sAnySuccess = true;
                }
            }

            sProgressDlg.close();

            if (sAnySuccess)
            {
                ::MessageBox(mHwnd, _T("7-Zip 압축 풀기가 성공적으로 완료되었습니다."), _T("7-Zip 내장 압축 관리"), MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                ::MessageBox(mHwnd, _T("압축 해제 대상 파일이 아니거나 해제에 실패하였습니다."), _T("7-Zip 내장 압축 관리"), MB_OK | MB_ICONWARNING);
            }
            return XPR_TRUE;
        }
        else if (aId == CMID_7ZIP_OPEN_ARCHIVE)
        {
            return XPR_TRUE;
        }
    }

    xpr_tchar_t sVerb[0xff] = {0};
    getCommandVerb(aId, sVerb, 0xfe);

    xpr_bool_t sComResult = XPR_FALSE;

    if (XPR_IS_RANGE(0, aId, 0x7FFF - 1))
    {
        CMINVOKECOMMANDINFO sCmInvokeCommandInfo = {0};
        sCmInvokeCommandInfo.cbSize = sizeof(CMINVOKECOMMANDINFO);
        sCmInvokeCommandInfo.hwnd   = mHwnd;
        sCmInvokeCommandInfo.lpVerb = (const xpr_char_t *)MAKEINTRESOURCE(aId);
        sCmInvokeCommandInfo.nShow  = SW_SHOW;

        HRESULT sComResult;
        if (XPR_IS_NOT_NULL(mContextMenu3))
        {
            sComResult = mContextMenu3->InvokeCommand(&sCmInvokeCommandInfo);
        }
        else if (XPR_IS_NOT_NULL(mContextMenu2))
        {
            sComResult = mContextMenu2->InvokeCommand(&sCmInvokeCommandInfo);
        }
        else
        {
            sComResult = mContextMenu->InvokeCommand(&sCmInvokeCommandInfo);
        }

        if (SUCCEEDED(sComResult))
            sComResult = XPR_TRUE;
    }

    return sComResult;
}

xpr_bool_t ContextMenu::invokeCommand(const xpr_tchar_t *aVerb)
{
    if (XPR_IS_NULL(aVerb))
        return XPR_FALSE;

    xpr_char_t sVerb[0xff] = {0};
#ifdef XPR_CFG_UNICODE
    xpr_size_t sInputBytes = _tcslen(aVerb) * sizeof(xpr_tchar_t);
    xpr_size_t sOutputBytes = 0xfe;
    XPR_UTF16_TO_MBS(aVerb, &sInputBytes, sVerb, &sOutputBytes);
    sVerb[sOutputBytes / sizeof(xpr_char_t)] = 0;
#else
    strcpy(sVerb, aVerb);
#endif

    CMINVOKECOMMANDINFO sCmInvokeCommandInfo = {0};
    sCmInvokeCommandInfo.cbSize = sizeof(CMINVOKECOMMANDINFO);
    sCmInvokeCommandInfo.hwnd   = mHwnd;
    sCmInvokeCommandInfo.lpVerb = sVerb;
    sCmInvokeCommandInfo.nShow  = SW_SHOW;

    HRESULT sComResult;
    if (XPR_IS_NOT_NULL(mContextMenu3))
    {
        sComResult = mContextMenu3->InvokeCommand(&sCmInvokeCommandInfo);
    }
    else if (XPR_IS_NOT_NULL(mContextMenu2))
    {
        sComResult = mContextMenu2->InvokeCommand(&sCmInvokeCommandInfo);
    }
    else
    {
        sComResult = mContextMenu->InvokeCommand(&sCmInvokeCommandInfo);
    }

    return SUCCEEDED(sComResult) ? XPR_TRUE : XPR_FALSE;
}

void ContextMenu::destroy(void)
{
    if (XPR_IS_NOT_NULL(mSubclassData))
    {
        int nIndex = 
#if defined(XPR_CFG_COMPILER_64BIT)
            GWLP_WNDPROC;
#else
            GWL_WNDPROC;
#endif

        ::SetWindowLongPtr(mHwnd, nIndex, (LONG_PTR)mSubclassData->mOldWndProc);
        ::RemoveProp(mHwnd, PROP_CONTEXT_MENU);
        XPR_SAFE_DELETE(mSubclassData);
    }

    if (XPR_IS_NOT_NULL(mMenu))
    {
        ::DestroyMenu(mMenu);
        mMenu = XPR_NULL;
    }

    COM_RELEASE(mContextMenu);
    COM_RELEASE(mContextMenu2);
    COM_RELEASE(mContextMenu3);
}

xpr_bool_t ContextMenu::trackItemMenu(LPSHELLFOLDER  aShellFolder,
                                      LPCITEMIDLIST *aPidls,
                                      xpr_uint_t     aCount,
                                      LPPOINT        aPoint,
                                      xpr_uint_t     aFlags,
                                      HWND           aWnd,
                                      xpr_uint_t     aQueryFlags)
{
    if (XPR_IS_NULL(aShellFolder) || XPR_IS_NULL(aPidls) || aCount == 0 || XPR_IS_NULL(aPoint))
        return XPR_FALSE;

    xpr_bool_t sResult = XPR_FALSE;

    ContextMenu sContextMenu(aWnd);
    if (sContextMenu.init(aShellFolder, aPidls, aCount) == XPR_TRUE)
    {
        xpr_uint_t sId = sContextMenu.trackPopupMenu(aFlags, aPoint, aQueryFlags);
        if (sId != -1)
        {
            xpr_uint_t sCmdId = sId - sContextMenu.getFirstId();

            sResult = sContextMenu.invokeCommand(sCmdId);
        }
    }

    sContextMenu.destroy();

    return sResult;
}

xpr_bool_t ContextMenu::trackBackMenu(LPSHELLFOLDER aShellFolder, LPPOINT aPoint, xpr_uint_t aFlags, HWND aWnd)
{
    if (XPR_IS_NULL(aShellFolder) || XPR_IS_NULL(aPoint))
        return XPR_FALSE;

    xpr_bool_t sResult = XPR_FALSE;

    ContextMenu sContextMenu(aWnd);
    if (sContextMenu.init(aShellFolder) == XPR_TRUE)
    {
        xpr_uint_t sId = sContextMenu.trackPopupMenu(aFlags, aPoint);
        if (sId != -1)
        {
            xpr_uint_t sCmdId = sId - sContextMenu.getFirstId();

            sResult = sContextMenu.invokeCommand(sCmdId);
        }
    }

    sContextMenu.destroy();

    return sResult;
}

xpr_bool_t ContextMenu::invokeCommand(LPSHELLFOLDER aShellFolder, LPCITEMIDLIST *aPidls, xpr_uint_t aCount, xpr_uint_t aId, HWND aWnd)
{
    if (XPR_IS_NULL(aShellFolder) || XPR_IS_NULL(aPidls) || aCount == 0)
        return XPR_FALSE;

    xpr_bool_t sResult = XPR_FALSE;

    ContextMenu sContextMenu(aWnd);
    if (sContextMenu.init(aShellFolder, aPidls, aCount) == XPR_TRUE)
    {
        CMenu sMenu;
        if (sMenu.CreatePopupMenu() == XPR_TRUE)
        {
            if (sContextMenu.getMenu(&sMenu, CMID_DEF_ID_FIRST, CMF_EXPLORE) == XPR_TRUE)
                sResult = sContextMenu.invokeCommand(aId);
        }
    }
    sContextMenu.destroy();

    return sResult;
}

xpr_bool_t ContextMenu::invokeCommand(LPSHELLFOLDER aShellFolder, LPCITEMIDLIST *aPidls, xpr_uint_t aCount, const xpr_tchar_t *aVerb, HWND aWnd)
{
    if (XPR_IS_NULL(aShellFolder) || XPR_IS_NULL(aPidls) || aCount == 0 || XPR_IS_NULL(aVerb))
        return XPR_FALSE;

    xpr_bool_t sResult = XPR_FALSE;

    ContextMenu sContextMenu(aWnd);
    if (sContextMenu.init(aShellFolder, aPidls, aCount) == XPR_TRUE)
    {
        CMenu sMenu;
        if (sMenu.CreatePopupMenu() == XPR_TRUE)
        {
            if (sContextMenu.getMenu(&sMenu, CMID_DEF_ID_FIRST, CMF_EXPLORE) == XPR_TRUE)
                sResult = sContextMenu.invokeCommand(aVerb);
        }
    }
    sContextMenu.destroy();

    return sResult;
}
} // namespace fxfile
