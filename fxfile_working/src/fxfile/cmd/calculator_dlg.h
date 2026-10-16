//
// Lightweight four-function calculator dialog.
//

#ifndef __FXFILE_CALCULATOR_DLG_H__
#define __FXFILE_CALCULATOR_DLG_H__ 1
#pragma once

namespace fxfile
{
namespace cmd
{
class CalculatorDlg : public CDialog
{
    typedef CDialog super;

public:
    CalculatorDlg(void);
    virtual ~CalculatorDlg(void);

protected:
    DECLARE_MESSAGE_MAP()
    virtual xpr_bool_t OnInitDialog(void);
    virtual void OnOK(void);
    afx_msg void OnClear(void);
};
} // namespace cmd
} // namespace fxfile

#endif // __FXFILE_CALCULATOR_DLG_H__
