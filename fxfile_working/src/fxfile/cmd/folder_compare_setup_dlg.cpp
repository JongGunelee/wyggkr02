//
// Copyright (c) 2012-2026 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.
//

#include "stdafx.h"
#include "folder_compare_setup_dlg.h"
#include "pidl_win.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cmd
{
BEGIN_MESSAGE_MAP(FolderCompareSetupDlg, CDialog)
    ON_BN_CLICKED(IDC_COMPARE_SETUP_BROWSE1, OnBrowsePath1)
    ON_BN_CLICKED(IDC_COMPARE_SETUP_BROWSE2, OnBrowsePath2)
    ON_CBN_SELCHANGE(IDC_COMPARE_SETUP_PANE1, OnSelchangePane1)
    ON_CBN_SELCHANGE(IDC_COMPARE_SETUP_PANE2, OnSelchangePane2)
    ON_BN_CLICKED(IDC_COMPARE_SETUP_SWAP,    OnSwapFolders)
END_MESSAGE_MAP()

FolderCompareSetupDlg::FolderCompareSetupDlg(CWnd *pParent)
    : super(FolderCompareSetupDlg::IDD, pParent)
    , mBySize(XPR_TRUE)
    , mByTime(XPR_TRUE)
    , mByContent(XPR_FALSE)
    , mByAttributes(XPR_FALSE)
    , mIncludeSubfolders(XPR_TRUE)
    , mExcludeFilter(XPR_STRING_LITERAL("*.tmp;*.bak;Thumbs.db;.git;Desktop.ini"))
{
}

FolderCompareSetupDlg::~FolderCompareSetupDlg(void)
{
}

void FolderCompareSetupDlg::setInitialPaths(const xpr_tchar_t *aPath1, const xpr_tchar_t *aPath2)
{
    if (aPath1) mPath1 = aPath1;
    if (aPath2) mPath2 = aPath2;
}

void FolderCompareSetupDlg::addPaneOption(xpr_sint_t aPaneIndex, const xpr_tchar_t *aPath)
{
    if (aPath && aPath[0] != 0)
    {
        mPaneOptions.push_back(std::make_pair(aPaneIndex, xpr::string(aPath)));
    }
}

void FolderCompareSetupDlg::DoDataExchange(CDataExchange* pDX)
{
    super::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMPARE_SETUP_PATH1,          mEditPath1);
    DDX_Control(pDX, IDC_COMPARE_SETUP_PATH2,          mEditPath2);
    DDX_Control(pDX, IDC_COMPARE_SETUP_PANE1,          mComboPane1);
    DDX_Control(pDX, IDC_COMPARE_SETUP_PANE2,          mComboPane2);
    DDX_Control(pDX, IDC_COMPARE_SETUP_EXCLUDE_FILTER, mEditExclude);
}

xpr_bool_t FolderCompareSetupDlg::OnInitDialog(void)
{
    super::OnInitDialog();

    mEditPath1.SetWindowText(mPath1.c_str());
    mEditPath2.SetWindowText(mPath2.c_str());

    CheckDlgButton(IDC_COMPARE_SETUP_CHK_SIZE,      mBySize ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(IDC_COMPARE_SETUP_CHK_TIME,      mByTime ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(IDC_COMPARE_SETUP_CHK_CONTENT,   mByContent ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(IDC_COMPARE_SETUP_CHK_ATTR,      mByAttributes ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(IDC_COMPARE_SETUP_CHK_SUBFOLDER, mIncludeSubfolders ? BST_CHECKED : BST_UNCHECKED);

    mEditExclude.SetWindowText(mExcludeFilter.c_str());

    // Pane 콤보박스 구성
    mComboPane1.AddString(XPR_STRING_LITERAL("-- 창 선택 --"));
    mComboPane2.AddString(XPR_STRING_LITERAL("-- 창 선택 --"));
    mComboPane1.SetCurSel(0);
    mComboPane2.SetCurSel(0);

    for (size_t i = 0; i < mPaneOptions.size(); ++i)
    {
        xpr_tchar_t sBuf[64];
        _sntprintf(sBuf, 64, XPR_STRING_LITERAL("창 #%d"), mPaneOptions[i].first);
        
        int idx1 = mComboPane1.AddString(sBuf);
        mComboPane1.SetItemData(idx1, (DWORD_PTR)i);

        int idx2 = mComboPane2.AddString(sBuf);
        mComboPane2.SetItemData(idx2, (DWORD_PTR)i);
    }

    return XPR_TRUE;
}

static xpr_sint_t CALLBACK CompareBrowseCallbackProc(HWND hwnd, xpr_uint_t uMsg, LPARAM lParam, LPARAM dwData)
{
    if (uMsg == BFFM_INITIALIZED && dwData != 0)
    {
        ::SendMessage(hwnd, BFFM_SETSELECTION, XPR_FALSE, dwData);
    }
    return 0;
}

void FolderCompareSetupDlg::browsePath(CEdit &aEdit)
{
    CString sCurrent;
    aEdit.GetWindowText(sCurrent);

    LPITEMIDLIST sOldFullPidl = XPR_NULL;
    if (!sCurrent.IsEmpty())
    {
        sOldFullPidl = fxfile::base::Pidl::create(sCurrent);
    }

    BROWSEINFO sBrowseInfo = {0};
    sBrowseInfo.hwndOwner = GetSafeHwnd();
    sBrowseInfo.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
    sBrowseInfo.lpszTitle = XPR_STRING_LITERAL("비교할 대상 폴더를 선택하세요:");
    sBrowseInfo.lpfn      = (BFFCALLBACK)CompareBrowseCallbackProc;
    sBrowseInfo.lParam    = (LPARAM)sOldFullPidl;

    LPITEMIDLIST sFullPidl = ::SHBrowseForFolder(&sBrowseInfo);
    if (sFullPidl != XPR_NULL)
    {
        xpr_tchar_t sPath[XPR_MAX_PATH + 1] = {0};
        ::SHGetPathFromIDList(sFullPidl, sPath);
        aEdit.SetWindowText(sPath);
        ::CoTaskMemFree(sFullPidl);
    }

    if (sOldFullPidl != XPR_NULL)
    {
        fxfile::base::Pidl::free(sOldFullPidl);
    }
}

void FolderCompareSetupDlg::OnBrowsePath1(void)
{
    browsePath(mEditPath1);
}

void FolderCompareSetupDlg::OnBrowsePath2(void)
{
    browsePath(mEditPath2);
}

void FolderCompareSetupDlg::OnSelchangePane1(void)
{
    int sel = mComboPane1.GetCurSel();
    if (sel > 0)
    {
        size_t optIdx = (size_t)mComboPane1.GetItemData(sel);
        if (optIdx < mPaneOptions.size())
        {
            mEditPath1.SetWindowText(mPaneOptions[optIdx].second.c_str());
        }
    }
}

void FolderCompareSetupDlg::OnSelchangePane2(void)
{
    int sel = mComboPane2.GetCurSel();
    if (sel > 0)
    {
        size_t optIdx = (size_t)mComboPane2.GetItemData(sel);
        if (optIdx < mPaneOptions.size())
        {
            mEditPath2.SetWindowText(mPaneOptions[optIdx].second.c_str());
        }
    }
}

void FolderCompareSetupDlg::OnSwapFolders(void)
{
    CString s1, s2;
    mEditPath1.GetWindowText(s1);
    mEditPath2.GetWindowText(s2);
    mEditPath1.SetWindowText(s2);
    mEditPath2.SetWindowText(s1);
}

void FolderCompareSetupDlg::OnOK(void)
{
    CString s1, s2, sEx;
    mEditPath1.GetWindowText(s1);
    mEditPath2.GetWindowText(s2);
    mEditExclude.GetWindowText(sEx);

    s1.Trim();
    s2.Trim();
    sEx.Trim();

    if (s1.IsEmpty() || s2.IsEmpty())
    {
        MessageBox(XPR_STRING_LITERAL("비교할 두 폴더 경로를 모두 지정해야 합니다."), XPR_STRING_LITERAL("폴더 비교 안내"), MB_OK | MB_ICONWARNING);
        return;
    }

    DWORD attr1 = ::GetFileAttributes(s1);
    if (attr1 == INVALID_FILE_ATTRIBUTES || !(attr1 & FILE_ATTRIBUTE_DIRECTORY))
    {
        CString msg;
        msg.Format(XPR_STRING_LITERAL("기준 폴더가 존재하지 않거나 폴더가 아닙니다:\n%s"), (LPCTSTR)s1);
        MessageBox(msg, XPR_STRING_LITERAL("폴더 비교 안내"), MB_OK | MB_ICONSTOP);
        return;
    }

    DWORD attr2 = ::GetFileAttributes(s2);
    if (attr2 == INVALID_FILE_ATTRIBUTES || !(attr2 & FILE_ATTRIBUTE_DIRECTORY))
    {
        CString msg;
        msg.Format(XPR_STRING_LITERAL("대상 폴더가 존재하지 않거나 폴더가 아닙니다:\n%s"), (LPCTSTR)s2);
        MessageBox(msg, XPR_STRING_LITERAL("폴더 비교 안내"), MB_OK | MB_ICONSTOP);
        return;
    }

    mPath1 = (LPCTSTR)s1;
    mPath2 = (LPCTSTR)s2;
    mExcludeFilter = (LPCTSTR)sEx;

    mBySize            = (IsDlgButtonChecked(IDC_COMPARE_SETUP_CHK_SIZE) == BST_CHECKED) ? XPR_TRUE : XPR_FALSE;
    mByTime            = (IsDlgButtonChecked(IDC_COMPARE_SETUP_CHK_TIME) == BST_CHECKED) ? XPR_TRUE : XPR_FALSE;
    mByContent         = (IsDlgButtonChecked(IDC_COMPARE_SETUP_CHK_CONTENT) == BST_CHECKED) ? XPR_TRUE : XPR_FALSE;
    mByAttributes      = (IsDlgButtonChecked(IDC_COMPARE_SETUP_CHK_ATTR) == BST_CHECKED) ? XPR_TRUE : XPR_FALSE;
    mIncludeSubfolders = (IsDlgButtonChecked(IDC_COMPARE_SETUP_CHK_SUBFOLDER) == BST_CHECKED) ? XPR_TRUE : XPR_FALSE;

    super::OnOK();
}

} // namespace cmd
} // namespace fxfile