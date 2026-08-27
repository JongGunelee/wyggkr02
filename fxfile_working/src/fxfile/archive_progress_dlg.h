//
// Copyright (c) 2026 FxFile Project. All rights reserved.
// Use of this source code is governed by a GPLv3 license.
//
// High-Quality Native Progress Dialog for 7-Zip Archive Operations
//

#ifndef __FXFILE_ARCHIVE_PROGRESS_DLG_H__
#define __FXFILE_ARCHIVE_PROGRESS_DLG_H__ 1
#pragma once

#include <windows.h>
#include <commctrl.h>
#include <string>

namespace fxfile
{
namespace archive
{

class ArchiveProgressDialog
{
public:
    ArchiveProgressDialog();
    ~ArchiveProgressDialog();

    bool create(HWND aParentHwnd, const std::wstring &aTitle, const std::wstring &aMainInstruction);
    void update(int aPercent, const std::wstring &aCurrentItem);
    void close();
    bool isCancelled() const { return mCancelled; }

    void pumpMessages();

private:
    static LRESULT CALLBACK wndProc(HWND aHwnd, UINT aMsg, WPARAM aWParam, LPARAM aLParam);

    HWND mHwnd;
    HWND mParentHwnd;
    HWND mProgressBarHwnd;
    HWND mTitleStaticHwnd;
    HWND mItemStaticHwnd;
    HWND mPercentStaticHwnd;
    HWND mCancelBtnHwnd;

    HFONT mFontBold;
    HFONT mFontNormal;

    bool mCancelled;
    bool mParentWasEnabled;
    int  mLastPercent;

    static const wchar_t *sClassName;
    static bool sClassRegistered;
};

} // namespace archive
} // namespace fxfile

#endif // __FXFILE_ARCHIVE_PROGRESS_DLG_H__
