//
// Copyright (c) 2026 FxFile Project. All rights reserved.
// Use of this source code is governed by a GPLv3 license.
//
// High-Quality Native Progress Dialog for 7-Zip Archive Operations
//

#include "stdafx.h"
#include "archive_progress_dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace archive
{

const wchar_t *ArchiveProgressDialog::sClassName = L"FxFile_ArchiveProgressDlgClass";
bool ArchiveProgressDialog::sClassRegistered = false;

ArchiveProgressDialog::ArchiveProgressDialog()
    : mHwnd(NULL)
    , mParentHwnd(NULL)
    , mProgressBarHwnd(NULL)
    , mTitleStaticHwnd(NULL)
    , mItemStaticHwnd(NULL)
    , mPercentStaticHwnd(NULL)
    , mCancelBtnHwnd(NULL)
    , mFontBold(NULL)
    , mFontNormal(NULL)
    , mCancelled(false)
    , mParentWasEnabled(true)
    , mLastPercent(-1)
{
}

ArchiveProgressDialog::~ArchiveProgressDialog()
{
    close();
}

LRESULT CALLBACK ArchiveProgressDialog::wndProc(HWND aHwnd, UINT aMsg, WPARAM aWParam, LPARAM aLParam)
{
    ArchiveProgressDialog *pThis = (ArchiveProgressDialog *)::GetWindowLongPtr(aHwnd, GWLP_USERDATA);

    switch (aMsg)
    {
    case WM_CREATE:
    {
        CREATESTRUCT *pCs = (CREATESTRUCT *)aLParam;
        ::SetWindowLongPtr(aHwnd, GWLP_USERDATA, (LONG_PTR)pCs->lpCreateParams);
        return 0;
    }
    case WM_COMMAND:
    {
        if (LOWORD(aWParam) == IDCANCEL && pThis)
        {
            pThis->mCancelled = true;
            if (pThis->mCancelBtnHwnd)
                ::EnableWindow(pThis->mCancelBtnHwnd, FALSE);
            if (pThis->mTitleStaticHwnd)
                ::SetWindowTextW(pThis->mTitleStaticHwnd, L"작업을 취소하는 중입니다...");
            return 0;
        }
        break;
    }
    case WM_CLOSE:
    {
        if (pThis)
        {
            pThis->mCancelled = true;
            if (pThis->mCancelBtnHwnd)
                ::EnableWindow(pThis->mCancelBtnHwnd, FALSE);
            if (pThis->mTitleStaticHwnd)
                ::SetWindowTextW(pThis->mTitleStaticHwnd, L"작업을 취소하는 중입니다...");
        }
        return 0; // Prevent window destruction until operation completes
    }
    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)aWParam;
        ::SetBkMode(hdcStatic, TRANSPARENT);
        return (LRESULT)::GetSysColorBrush(COLOR_BTNFACE);
    }
    }

    return ::DefWindowProc(aHwnd, aMsg, aWParam, aLParam);
}

bool ArchiveProgressDialog::create(HWND aParentHwnd, const std::wstring &aTitle, const std::wstring &aMainInstruction)
{
    mParentHwnd = aParentHwnd;
    mCancelled = false;
    mLastPercent = -1;

    // Disable parent window to prevent re-entrancy / crashes during file compression
    if (mParentHwnd && ::IsWindow(mParentHwnd))
    {
        mParentWasEnabled = ::IsWindowEnabled(mParentHwnd) ? true : false;
        ::EnableWindow(mParentHwnd, FALSE);
    }

    // Ensure common controls are initialized
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_PROGRESS_CLASS;
    ::InitCommonControlsEx(&icex);

    HINSTANCE hInstance = ::GetModuleHandle(NULL);

    if (!sClassRegistered)
    {
        WNDCLASSEXW wc = {0};
        wc.cbSize        = sizeof(WNDCLASSEXW);
        wc.style         = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc   = wndProc;
        wc.hInstance     = hInstance;
        wc.hCursor       = ::LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = sClassName;

        if (::RegisterClassExW(&wc))
            sClassRegistered = true;
    }

    // Create crisp typography fonts (Malgun Gothic / Segoe UI)
    LOGFONTW lf = {0};
    lf.lfHeight = -13;
    lf.lfWeight = FW_BOLD;
    wcscpy_s(lf.lfFaceName, L"맑은 고딕");
    mFontBold = ::CreateFontIndirectW(&lf);

    lf.lfHeight = -12;
    lf.lfWeight = FW_NORMAL;
    mFontNormal = ::CreateFontIndirectW(&lf);

    // Calculate window bounds with accurate client area
    int clientWidth = 480;
    int clientHeight = 160;

    RECT rcWnd = { 0, 0, clientWidth, clientHeight };
    DWORD dwStyle = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE;
    DWORD dwExStyle = WS_EX_DLGMODALFRAME | WS_EX_TOPMOST;
    ::AdjustWindowRectEx(&rcWnd, dwStyle, FALSE, dwExStyle);

    int totalWidth = rcWnd.right - rcWnd.left;
    int totalHeight = rcWnd.bottom - rcWnd.top;

    int posX = (::GetSystemMetrics(SM_CXSCREEN) - totalWidth) / 2;
    int posY = (::GetSystemMetrics(SM_CYSCREEN) - totalHeight) / 2;

    if (mParentHwnd && ::IsWindow(mParentHwnd))
    {
        RECT rcParent;
        ::GetWindowRect(mParentHwnd, &rcParent);
        posX = rcParent.left + ((rcParent.right - rcParent.left) - totalWidth) / 2;
        posY = rcParent.top + ((rcParent.bottom - rcParent.top) - totalHeight) / 2;
    }

    mHwnd = ::CreateWindowExW(
        dwExStyle,
        sClassName,
        aTitle.c_str(),
        dwStyle,
        posX, posY, totalWidth, totalHeight,
        mParentHwnd, NULL, hInstance, this);

    if (!mHwnd)
    {
        if (mParentHwnd && ::IsWindow(mParentHwnd) && mParentWasEnabled)
            ::EnableWindow(mParentHwnd, TRUE);
        return false;
    }

    // 1. Title / Main Instruction Static (Top)
    mTitleStaticHwnd = ::CreateWindowExW(
        0, L"STATIC", aMainInstruction.c_str(),
        WS_CHILD | WS_VISIBLE | SS_LEFT | SS_ENDELLIPSIS,
        24, 16, 432, 22,
        mHwnd, NULL, hInstance, NULL);
    if (mFontBold && mTitleStaticHwnd)
        ::SendMessage(mTitleStaticHwnd, WM_SETFONT, (WPARAM)mFontBold, TRUE);

    // 2. Current Item / File Path Static (Middle-Top)
    mItemStaticHwnd = ::CreateWindowExW(
        0, L"STATIC", L"준비 중...",
        WS_CHILD | WS_VISIBLE | SS_LEFT | SS_PATHELLIPSIS,
        24, 44, 432, 20,
        mHwnd, NULL, hInstance, NULL);
    if (mFontNormal && mItemStaticHwnd)
        ::SendMessage(mItemStaticHwnd, WM_SETFONT, (WPARAM)mFontNormal, TRUE);

    // 3. Native Progress Bar Control (Middle)
    mProgressBarHwnd = ::CreateWindowExW(
        0, PROGRESS_CLASSW, NULL,
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        24, 72, 432, 22,
        mHwnd, NULL, hInstance, NULL);
    if (mProgressBarHwnd)
    {
        ::SendMessage(mProgressBarHwnd, PBM_SETRANGE32, 0, 100);
        ::SendMessage(mProgressBarHwnd, PBM_SETPOS, 0, 0);
    }

    // 4. Percent Static Text (Bottom-Left)
    mPercentStaticHwnd = ::CreateWindowExW(
        0, L"STATIC", L"0%",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        24, 108, 100, 20,
        mHwnd, NULL, hInstance, NULL);
    if (mFontNormal && mPercentStaticHwnd)
        ::SendMessage(mPercentStaticHwnd, WM_SETFONT, (WPARAM)mFontNormal, TRUE);

    // 5. Cancel Button (Bottom-Right, completely separated and visible)
    mCancelBtnHwnd = ::CreateWindowExW(
        0, L"BUTTON", L"취소",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
        366, 106, 90, 28,
        mHwnd, (HMENU)IDCANCEL, hInstance, NULL);
    if (mFontNormal && mCancelBtnHwnd)
        ::SendMessage(mCancelBtnHwnd, WM_SETFONT, (WPARAM)mFontNormal, TRUE);

    ::UpdateWindow(mHwnd);
    return true;
}

void ArchiveProgressDialog::update(int aPercent, const std::wstring &aCurrentItem)
{
    if (!mHwnd || !::IsWindow(mHwnd))
        return;

    if (aPercent < 0) aPercent = 0;
    if (aPercent > 100) aPercent = 100;

    if (aPercent != mLastPercent)
    {
        mLastPercent = aPercent;
        if (mProgressBarHwnd)
            ::SendMessage(mProgressBarHwnd, PBM_SETPOS, (WPARAM)aPercent, 0);

        if (mPercentStaticHwnd)
        {
            wchar_t sBuf[32];
            _snwprintf_s(sBuf, _countof(sBuf), _TRUNCATE, L"%d%%", aPercent);
            ::SetWindowTextW(mPercentStaticHwnd, sBuf);
        }
    }

    if (!aCurrentItem.empty() && mItemStaticHwnd)
    {
        ::SetWindowTextW(mItemStaticHwnd, aCurrentItem.c_str());
    }

    pumpMessages();
}

void ArchiveProgressDialog::pumpMessages()
{
    if (!mHwnd || !::IsWindow(mHwnd))
        return;

    MSG msg;
    // Process only messages directed to the progress dialog to avoid main frame re-entrancy
    while (::PeekMessage(&msg, mHwnd, 0, 0, PM_REMOVE))
    {
        if (::IsDialogMessage(mHwnd, &msg))
            continue;
        ::TranslateMessage(&msg);
        ::DispatchMessage(&msg);
    }
}

void ArchiveProgressDialog::close()
{
    if (mHwnd && ::IsWindow(mHwnd))
    {
        ::DestroyWindow(mHwnd);
        mHwnd = NULL;
    }

    if (mFontBold)
    {
        ::DeleteObject(mFontBold);
        mFontBold = NULL;
    }

    if (mFontNormal)
    {
        ::DeleteObject(mFontNormal);
        mFontNormal = NULL;
    }

    // Re-enable parent window when dialog closes
    if (mParentHwnd && ::IsWindow(mParentHwnd) && mParentWasEnabled)
    {
        ::EnableWindow(mParentHwnd, TRUE);
        ::SetActiveWindow(mParentHwnd);
        ::SetForegroundWindow(mParentHwnd);
    }
}

} // namespace archive
} // namespace fxfile
