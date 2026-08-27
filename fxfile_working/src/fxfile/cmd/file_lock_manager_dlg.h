// Safe file/folder lock manager dialog.
#ifndef __FXFILE_FILE_LOCK_MANAGER_DLG_H__
#define __FXFILE_FILE_LOCK_MANAGER_DLG_H__ 1
#pragma once

#include "gui/ResizingDialog.h"

#include <string>
#include <vector>

namespace fxfile
{
namespace cmd
{
class FileLockManagerDlg : public CResizingDialog
{
    typedef CResizingDialog super;

public:
    explicit FileLockManagerDlg(const std::vector<std::wstring> &aPaths);

protected:
    virtual BOOL OnInitDialog(void);
    virtual BOOL PreTranslateMessage(MSG *aMsg);
    void initializeToolTips(void);
    void refresh(void);
    bool selectedPath(std::wstring &aPath) const;
    void setReadOnly(bool aReadOnly);
    bool isProtected(const std::wstring &aPath) const;
    std::wstring describePath(const std::wstring &aPath) const;
    void refreshLockingProcesses(const std::wstring &aPath);

protected:
    afx_msg void OnRefresh(void);
    afx_msg void OnFxLock(void);
    afx_msg void OnFxUnlock(void);
    afx_msg void OnReadOnly(void);
    afx_msg void OnWritable(void);
    afx_msg void OnSecurity(void);
    DECLARE_MESSAGE_MAP()

private:
    std::vector<std::wstring> mPaths;
    CListCtrl mPathList;
    CListCtrl mProcessList;
    CToolTipCtrl mToolTipCtrl;
};
} // namespace cmd
} // namespace fxfile

#endif // __FXFILE_FILE_LOCK_MANAGER_DLG_H__
