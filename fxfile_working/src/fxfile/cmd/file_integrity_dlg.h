//
// SHA-256 based file integrity monitoring dialog.
//

#ifndef __FXFILE_FILE_INTEGRITY_DLG_H__
#define __FXFILE_FILE_INTEGRITY_DLG_H__ 1
#pragma once

#include <vector>

namespace fxfile
{
namespace cmd
{
class FileIntegrityDlg : public CDialog
{
    typedef CDialog super;

public:
    FileIntegrityDlg(void);
    virtual ~FileIntegrityDlg(void);

    void addPath(const xpr_tchar_t *aPath);

protected:
    struct Item
    {
        xpr::string mPath;
        xpr::string mBaselineHash;
        xpr::string mCurrentHash;
        ULONGLONG   mBaselineSize;
        ULONGLONG   mCurrentSize;
        FILETIME    mBaselineWriteTime;
        FILETIME    mCurrentWriteTime;
        xpr_bool_t  mBaselineValid;
        xpr_bool_t  mCurrentValid;

        Item(void);
    };

    typedef std::vector<Item> ItemVector;

protected:
    DECLARE_MESSAGE_MAP()
    virtual xpr_bool_t OnInitDialog(void);
    virtual void OnCancel(void);
    afx_msg void OnRefresh(void);
    afx_msg void OnResetBaseline(void);
    afx_msg void OnCopyReport(void);
    afx_msg void OnAutoMonitor(void);
    afx_msg void OnTimer(UINT_PTR aEventId);

    void scan(xpr_bool_t aResetBaseline);
    void updateRow(xpr_sint_t aIndex);
    void updateStatus(void);
    xpr::string formatFileTime(const FILETIME &aFileTime) const;
    xpr::string getDifference(xpr_sint_t aIndex) const;
    xpr::string getIntegrityStatus(const Item &aItem) const;

protected:
    ItemVector mItems;
    CListCtrl  mListCtrl;
    CProgressCtrl mProgressCtrl;
    xpr_bool_t mBusy;
};
} // namespace cmd
} // namespace fxfile

#endif // __FXFILE_FILE_INTEGRITY_DLG_H__
