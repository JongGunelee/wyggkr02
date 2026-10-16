//
// Copyright (c) 2012-2026 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.
//

#ifndef __FXFILE_FOLDER_COMPARE_REPORT_DLG_H__
#define __FXFILE_FOLDER_COMPARE_REPORT_DLG_H__ 1
#pragma once

#include "resource.h"
#include "gui/ResizingDialog.h"
#include "sync_dirs.h"
#include <vector>

namespace fxfile
{
class MainFrame;

namespace cmd
{
class FolderCompareReportDlg : public CResizingDialog
{
    typedef CResizingDialog super;

public:
    FolderCompareReportDlg(MainFrame *aMainFrame = XPR_NULL, CWnd *pParent = XPR_NULL);
    virtual ~FolderCompareReportDlg(void);

    enum { IDD = IDD_FOLDER_COMPARE_REPORT };

public:
    void setResult(SyncDirs *aSyncDirs, const xpr_tchar_t *aDir1, const xpr_tchar_t *aDir2);

protected:
    MainFrame  *mMainFrame;
    SyncDirs   *mSyncDirs;
    xpr::string mDir1;
    xpr::string mDir2;

    struct ItemViewData
    {
        SyncItem   *pItem;
        CString     statusText;
        CString     subPath;
        CString     size1Text;
        CString     time1Text;
        CString     size2Text;
        CString     time2Text;
        CString     diffReason;
        int         statusType; // 0: Equal, 1: Diff, 2: LeftOnly, 3: RightOnly
    };

    std::vector<ItemViewData> mAllItems;
    std::vector<size_t>       mFilteredIndices;

    size_t mCountEqual;
    size_t mCountDiff;
    size_t mCountLeftOnly;
    size_t mCountRightOnly;
    size_t mCountTotal;

    xpr_sint64_t mSizeEqual;
    xpr_sint64_t mSizeDiff;
    xpr_sint64_t mSizeLeftOnly;
    xpr_sint64_t mSizeRightOnly;

protected:
    CComboBox mComboFilter;
    CListCtrl mListCtrl;

protected:
    virtual void DoDataExchange(CDataExchange* pDX);
    virtual xpr_bool_t OnInitDialog(void);

    afx_msg void OnSelchangeFilter(void);
    afx_msg void OnBtnCopyReport(void);
    afx_msg void OnBtnSaveReport(void);
    afx_msg void OnBtnSelectInWindow(void);
    afx_msg void OnBtnSyncTool(void);

    DECLARE_MESSAGE_MAP()

private:
    void populateItems(void);
    void applyFilter(int aFilterMode);
    CString generateMarkdownReport(void) const;
    CString formatFileSize(xpr_sint64_t aSize) const;
    CString formatFileTime(const FILETIME &aFt) const;
};
} // namespace cmd
} // namespace fxfile

#endif // __FXFILE_FOLDER_COMPARE_REPORT_DLG_H__