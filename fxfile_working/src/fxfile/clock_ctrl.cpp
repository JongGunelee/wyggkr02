//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "clock_ctrl.h"
#include "option.h"
#include "option_manager.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
ClockCtrl::ClockCtrl(void)
    : mDragging(XPR_FALSE)
    , mDragStartPt(0, 0)
    , mDragStartLeft(0)
{
}

ClockCtrl::~ClockCtrl(void)
{
    if (mClockFont.m_hObject != XPR_NULL)
        mClockFont.DeleteObject();
}

BEGIN_MESSAGE_MAP(ClockCtrl, CStatic)
    ON_WM_PAINT()
    ON_WM_ERASEBKGND()
    ON_WM_SETCURSOR()
    ON_WM_LBUTTONDOWN()
    ON_WM_MOUSEMOVE()
    ON_WM_LBUTTONUP()
    ON_WM_CONTEXTMENU()
    ON_WM_RBUTTONUP()
END_MESSAGE_MAP()

void ClockCtrl::setText(const xpr_tchar_t *aText)
{
    if (aText == XPR_NULL)
        mClockText.Empty();
    else
        mClockText = aText;

    if (::IsWindow(m_hWnd))
        Invalidate(FALSE);
}

void ClockCtrl::updateFont(int aHeight)
{
    if (mClockFont.m_hObject != XPR_NULL)
        mClockFont.DeleteObject();

    LOGFONT sLf = {0};
    sLf.lfHeight = aHeight;
    sLf.lfWeight = FW_BOLD;
    sLf.lfCharSet = DEFAULT_CHARSET;
    _tcscpy(sLf.lfFaceName, _T("맑은 고딕"));

    mClockFont.CreateFontIndirect(&sLf);

    if (::IsWindow(m_hWnd))
        Invalidate(FALSE);
}

int ClockCtrl::calcMaxButtonRight(void) const
{
    CWnd *pParent = GetParent();
    if (pParent == XPR_NULL || !::IsWindow(pParent->m_hWnd))
        return 0;

    CToolBar *pToolBar = dynamic_cast<CToolBar*>(pParent);
    if (pToolBar == XPR_NULL)
        return 0;

    CToolBarCtrl &sToolBarCtrl = pToolBar->GetToolBarCtrl();
    int sBtnCount = sToolBarCtrl.GetButtonCount();

    int sMaxRight = 0;
    CRect sBtnRect;
    TBBUTTON sButton;

    for (int i = 0; i < sBtnCount; ++i)
    {
        if (sToolBarCtrl.GetButton(i, &sButton))
        {
            if (!(sButton.fsState & TBSTATE_HIDDEN))
            {
                pToolBar->GetItemRect(i, &sBtnRect);
                if (sBtnRect.right > sMaxRight)
                    sMaxRight = sBtnRect.right;
            }
        }
    }

    return sMaxRight;
}

BOOL ClockCtrl::OnEraseBkgnd(CDC *pDC)
{
    // Prevent flickering by drawing entire background in OnPaint
    return TRUE;
}

void ClockCtrl::OnPaint(void)
{
    CPaintDC dc(this);

    CRect rcClient;
    GetClientRect(&rcClient);
    if (rcClient.Width() <= 0 || rcClient.Height() <= 0)
        return;

    // Double buffering
    CDC memDC;
    CBitmap memBmp;
    memDC.CreateCompatibleDC(&dc);
    memBmp.CreateCompatibleBitmap(&dc, rcClient.Width(), rcClient.Height());
    CBitmap *pOldBmp = memDC.SelectObject(&memBmp);

    // Parent background fill
    memDC.FillSolidRect(&rcClient, ::GetSysColor(COLOR_BTNFACE));

    // Clock background box (rounded rectangle)
    CRect rcBox = rcClient;
    rcBox.DeflateRect(1, 1);

    CPen pen(PS_SOLID, 1, RGB(190, 198, 208));
    CBrush brush(RGB(240, 244, 248));
    CPen *pOldPen = memDC.SelectObject(&pen);
    CBrush *pOldBrush = memDC.SelectObject(&brush);

    memDC.RoundRect(&rcBox, CPoint(6, 6));

    memDC.SelectObject(pOldBrush);
    memDC.SelectObject(pOldPen);

    // Draw clock text
    CFont *pOldFont = XPR_NULL;
    if (mClockFont.m_hObject != XPR_NULL)
        pOldFont = memDC.SelectObject(&mClockFont);

    memDC.SetBkMode(TRANSPARENT);
    memDC.SetTextColor(RGB(20, 24, 30));
    memDC.DrawText(mClockText, &rcBox, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    if (pOldFont != XPR_NULL)
        memDC.SelectObject(pOldFont);

    // BitBlt to display
    dc.BitBlt(0, 0, rcClient.Width(), rcClient.Height(), &memDC, 0, 0, SRCCOPY);
    memDC.SelectObject(pOldBmp);
}

BOOL ClockCtrl::OnSetCursor(CWnd *pWnd, UINT nHitTest, UINT message)
{
    if (gOpt != XPR_NULL && !gOpt->mMain.mClockLocked)
    {
        ::SetCursor(::LoadCursor(NULL, IDC_SIZEALL));
        return TRUE;
    }

    return CStatic::OnSetCursor(pWnd, nHitTest, message);
}

void ClockCtrl::OnLButtonDown(UINT nFlags, CPoint point)
{
    if (gOpt != XPR_NULL && !gOpt->mMain.mClockLocked)
    {
        mDragging = XPR_TRUE;
        SetCapture();

        CPoint ptScreen = point;
        ClientToScreen(&ptScreen);
        mDragStartPt = ptScreen;

        CRect rc;
        GetWindowRect(&rc);
        if (GetParent() != XPR_NULL)
            GetParent()->ScreenToClient(&rc);
        mDragStartLeft = rc.left;
        return;
    }

    CStatic::OnLButtonDown(nFlags, point);
}

void ClockCtrl::OnMouseMove(UINT nFlags, CPoint point)
{
    if (mDragging && GetParent() != XPR_NULL)
    {
        CPoint ptScreen = point;
        ClientToScreen(&ptScreen);

        int dx = ptScreen.x - mDragStartPt.x;

        CRect rcParent;
        GetParent()->GetClientRect(&rcParent);

        int sMaxRight = calcMaxButtonRight();
        int minLeft = (sMaxRight > 0) ? (sMaxRight + 15) : 5;

        CRect rc;
        GetWindowRect(&rc);
        int maxLeft = max(minLeft, rcParent.Width() - rc.Width() - 5);

        int newLeft = max(minLeft, min(maxLeft, mDragStartLeft + dx));

        CRect rcCur;
        GetWindowRect(&rcCur);
        GetParent()->ScreenToClient(&rcCur);

        if (rcCur.left != newLeft)
        {
            MoveWindow(newLeft, rcCur.top, rcCur.Width(), rcCur.Height(), TRUE);
            if (gOpt != XPR_NULL)
                gOpt->mMain.mClockPosX = newLeft;
        }
        return;
    }

    CStatic::OnMouseMove(nFlags, point);
}

void ClockCtrl::OnLButtonUp(UINT nFlags, CPoint point)
{
    if (mDragging)
    {
        mDragging = XPR_FALSE;
        ReleaseCapture();

        OptionManager::instance().saveMainOption();
        return;
    }

    CStatic::OnLButtonUp(nFlags, point);
}

void ClockCtrl::OnContextMenu(CWnd *pWnd, CPoint point)
{
    CWnd *pParent = GetParent();
    if (pParent != XPR_NULL)
    {
        pParent->SendMessage(WM_CONTEXTMENU, (WPARAM)m_hWnd, MAKELPARAM(point.x, point.y));
    }
}

void ClockCtrl::OnRButtonUp(UINT nFlags, CPoint point)
{
    CPoint ptScreen = point;
    ClientToScreen(&ptScreen);
    OnContextMenu(this, ptScreen);
}

} // namespace fxfile
