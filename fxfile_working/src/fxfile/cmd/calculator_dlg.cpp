//
// Lightweight four-function calculator dialog.
//

#include "stdafx.h"
#include "calculator_dlg.h"

#include "resource.h"

#include <float.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cmd
{
namespace
{
class ExpressionParser
{
public:
    enum Error
    {
        ErrorNone,
        ErrorInvalid,
        ErrorDivideByZero
    };

public:
    explicit ExpressionParser(const xpr_tchar_t *aText)
        : mCur(aText), mError(ErrorNone)
    {
    }

    xpr_bool_t parse(xpr_double_t &aValue)
    {
        skipSpaces();
        if (*mCur == 0)
        {
            mError = ErrorInvalid;
            return XPR_FALSE;
        }

        aValue = parseExpression();
        skipSpaces();

        if (mError != ErrorNone || *mCur != 0 || _finite(aValue) == 0)
        {
            if (mError == ErrorNone)
                mError = ErrorInvalid;
            return XPR_FALSE;
        }

        return XPR_TRUE;
    }

    Error getError(void) const
    {
        return mError;
    }

private:
    xpr_double_t parseExpression(void)
    {
        xpr_double_t sValue = parseTerm();

        while (mError == ErrorNone)
        {
            skipSpaces();
            xpr_tchar_t sOperator = *mCur;
            if (sOperator != XPR_STRING_LITERAL('+') && sOperator != XPR_STRING_LITERAL('-'))
                break;

            ++mCur;
            xpr_double_t sRight = parseTerm();
            sValue = (sOperator == XPR_STRING_LITERAL('+')) ? sValue + sRight : sValue - sRight;
        }

        return sValue;
    }

    xpr_double_t parseTerm(void)
    {
        xpr_double_t sValue = parseFactor();

        while (mError == ErrorNone)
        {
            skipSpaces();
            xpr_tchar_t sOperator = *mCur;
            if (sOperator != XPR_STRING_LITERAL('*') && sOperator != XPR_STRING_LITERAL('/'))
                break;

            ++mCur;
            xpr_double_t sRight = parseFactor();
            if (sOperator == XPR_STRING_LITERAL('/') && sRight == 0.0)
            {
                mError = ErrorDivideByZero;
                return 0.0;
            }

            sValue = (sOperator == XPR_STRING_LITERAL('*')) ? sValue * sRight : sValue / sRight;
        }

        return sValue;
    }

    xpr_double_t parseFactor(void)
    {
        skipSpaces();

        if (*mCur == XPR_STRING_LITERAL('+'))
        {
            ++mCur;
            return parseFactor();
        }

        if (*mCur == XPR_STRING_LITERAL('-'))
        {
            ++mCur;
            return -parseFactor();
        }

        if (*mCur == XPR_STRING_LITERAL('('))
        {
            ++mCur;
            xpr_double_t sValue = parseExpression();
            skipSpaces();
            if (*mCur != XPR_STRING_LITERAL(')'))
            {
                mError = ErrorInvalid;
                return 0.0;
            }

            ++mCur;
            return sValue;
        }

        xpr_tchar_t *sEnd = XPR_NULL;
        xpr_double_t sValue = _tcstod(mCur, &sEnd);
        if (sEnd == mCur)
        {
            mError = ErrorInvalid;
            return 0.0;
        }

        mCur = sEnd;
        return sValue;
    }

    void skipSpaces(void)
    {
        while (*mCur != 0 && _istspace(*mCur) != 0)
            ++mCur;
    }

private:
    const xpr_tchar_t *mCur;
    Error              mError;
};
} // anonymous namespace

CalculatorDlg::CalculatorDlg(void)
    : super(IDD_CALCULATOR, XPR_NULL)
{
}

CalculatorDlg::~CalculatorDlg(void)
{
}

BEGIN_MESSAGE_MAP(CalculatorDlg, super)
    ON_BN_CLICKED(IDC_CALCULATOR_CLEAR, OnClear)
END_MESSAGE_MAP()

xpr_bool_t CalculatorDlg::OnInitDialog(void)
{
    super::OnInitDialog();

    SetWindowText(gApp.loadString(XPR_STRING_LITERAL("popup.calculator.title")));
    SetDlgItemText(IDC_CALCULATOR_LABEL_EXPRESSION, gApp.loadString(XPR_STRING_LITERAL("popup.calculator.label.expression")));
    SetDlgItemText(IDC_CALCULATOR_LABEL_RESULT,     gApp.loadString(XPR_STRING_LITERAL("popup.calculator.label.result")));
    SetDlgItemText(IDC_CALCULATOR_HINT,             gApp.loadString(XPR_STRING_LITERAL("popup.calculator.hint")));
    SetDlgItemText(IDOK,                            gApp.loadString(XPR_STRING_LITERAL("popup.calculator.button.calculate")));
    SetDlgItemText(IDC_CALCULATOR_CLEAR,            gApp.loadString(XPR_STRING_LITERAL("popup.calculator.button.clear")));
    SetDlgItemText(IDCANCEL,                        gApp.loadString(XPR_STRING_LITERAL("popup.calculator.button.close")));

    GetDlgItem(IDC_CALCULATOR_EXPRESSION)->SetFocus();
    return XPR_FALSE;
}

void CalculatorDlg::OnOK(void)
{
    xpr_tchar_t sExpression[1024] = {0};
    GetDlgItemText(IDC_CALCULATOR_EXPRESSION, sExpression, XPR_COUNT_OF(sExpression));

    ExpressionParser sParser(sExpression);
    xpr_double_t sValue = 0.0;
    if (XPR_IS_FALSE(sParser.parse(sValue)))
    {
        const xpr_tchar_t *sStringId =
            (sParser.getError() == ExpressionParser::ErrorDivideByZero)
            ? XPR_STRING_LITERAL("popup.calculator.error.divide_by_zero")
            : XPR_STRING_LITERAL("popup.calculator.error.invalid");

        SetDlgItemText(IDC_CALCULATOR_RESULT, gApp.loadString(sStringId));
        ::MessageBeep(MB_ICONWARNING);
        return;
    }

    if (sValue == 0.0)
        sValue = 0.0;

    xpr_tchar_t sResult[128] = {0};
    _sntprintf_s(sResult, XPR_COUNT_OF(sResult), _TRUNCATE, XPR_STRING_LITERAL("%.15g"), sValue);
    SetDlgItemText(IDC_CALCULATOR_RESULT, sResult);
}

void CalculatorDlg::OnClear(void)
{
    SetDlgItemText(IDC_CALCULATOR_EXPRESSION, XPR_STRING_LITERAL(""));
    SetDlgItemText(IDC_CALCULATOR_RESULT,     XPR_STRING_LITERAL(""));
    GetDlgItem(IDC_CALCULATOR_EXPRESSION)->SetFocus();
}
} // namespace cmd
} // namespace fxfile
