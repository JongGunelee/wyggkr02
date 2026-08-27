//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_CLOCK_CTRL_H__
#define __FXFILE_CLOCK_CTRL_H__ 1
#pragma once

namespace fxfile
{
class ClockCtrl : public CStatic
{
public:
    ClockCtrl(void);
    virtual ~ClockCtrl(void);

public:
    void setText(const xpr_tchar_t *aText);
    void updateFont(int aHeight = -18);

protected:
    int calcMaxButtonRight(void) const;

protected:
    CFont       mClockFont;
    CString     mClockText;
    xpr_bool_t  mDragging;
    CPoint      mDragStartPt;
    int         mDragStartLeft;

protected:
    DECLARE_MESSAGE_MAP()
    afx_msg void OnPaint(void);
    afx_msg BOOL OnEraseBkgnd(CDC *pDC);
    afx_msg BOOL OnSetCursor(CWnd *pWnd, UINT nHitTest, UINT message);
    afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
    afx_msg void OnMouseMove(UINT nFlags, CPoint point);
    afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
    afx_msg void OnContextMenu(CWnd *pWnd, CPoint point);
    afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
};
} // namespace fxfile

#endif // __FXFILE_CLOCK_CTRL_H__
