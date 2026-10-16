//
// Copyright (c) 2012-2026 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.
//

#ifndef __FXFILE_FOLDER_COMPARE_SETUP_DLG_H__
#define __FXFILE_FOLDER_COMPARE_SETUP_DLG_H__ 1
#pragma once

#include "resource.h"
#include <vector>
#include <utility>

namespace fxfile
{
namespace cmd
{
class FolderCompareSetupDlg : public CDialog
{
    typedef CDialog super;

public:
    FolderCompareSetupDlg(CWnd *pParent = XPR_NULL);
    virtual ~FolderCompareSetupDlg(void);

    enum { IDD = IDD_FOLDER_COMPARE_SETUP };

public:
    void setInitialPaths(const xpr_tchar_t *aPath1, const xpr_tchar_t *aPath2);
    void addPaneOption(xpr_sint_t aPaneIndex, const xpr_tchar_t *aPath);

    const xpr::string &getPath1(void) const { return mPath1; }
    const xpr::string &getPath2(void) const { return mPath2; }

    xpr_bool_t isBySize(void) const { return mBySize; }
    xpr_bool_t isByTime(void) const { return mByTime; }
    xpr_bool_t isByContent(void) const { return mByContent; }
    xpr_bool_t isByAttributes(void) const { return mByAttributes; }
    xpr_bool_t isIncludeSubfolders(void) const { return mIncludeSubfolders; }
    const xpr::string &getExcludeFilter(void) const { return mExcludeFilter; }

protected:
    xpr::string mPath1;
    xpr::string mPath2;
    xpr_bool_t  mBySize;
    xpr_bool_t  mByTime;
    xpr_bool_t  mByContent;
    xpr_bool_t  mByAttributes;
    xpr_bool_t  mIncludeSubfolders;
    xpr::string mExcludeFilter;

    std::vector<std::pair<xpr_sint_t, xpr::string> > mPaneOptions;

protected:
    CEdit     mEditPath1;
    CEdit     mEditPath2;
    CComboBox mComboPane1;
    CComboBox mComboPane2;
    CEdit     mEditExclude;

protected:
    virtual void DoDataExchange(CDataExchange* pDX);
    virtual xpr_bool_t OnInitDialog(void);

    afx_msg void OnBrowsePath1(void);
    afx_msg void OnBrowsePath2(void);
    afx_msg void OnSelchangePane1(void);
    afx_msg void OnSelchangePane2(void);
    afx_msg void OnSwapFolders(void);
    virtual void OnOK(void);

    DECLARE_MESSAGE_MAP()

private:
    void browsePath(CEdit &aEdit);
};
} // namespace cmd
} // namespace fxfile

#endif // __FXFILE_FOLDER_COMPARE_SETUP_DLG_H__