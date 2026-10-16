#define UNICODE
#define _UNICODE

#include <windows.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <vsstyle.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

namespace
{
const COLORREF kFocusColor = RGB(255, 255, 0);
const COLORREF kFocusTextColor = RGB(0, 0, 0);

enum ProbeMode
{
    ProbeCurrent = 0,
    ProbeThemeState,
    ProbeThemeStateAndFill,
    ProbeLegacyCellFill
};

HWND gLists[4] = { nullptr, nullptr, nullptr, nullptr };

bool SaveClientBitmap(HWND window, const wchar_t *path)
{
    RECT client = {};
    GetClientRect(window, &client);
    const int width = client.right - client.left;
    const int height = client.bottom - client.top;
    if (width <= 0 || height <= 0)
        return false;

    BITMAPINFO info = {};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    info.bmiHeader.biWidth = width;
    info.bmiHeader.biHeight = -height;
    info.bmiHeader.biPlanes = 1;
    info.bmiHeader.biBitCount = 32;
    info.bmiHeader.biCompression = BI_RGB;

    void *pixels = nullptr;
    HDC windowDc = GetDC(window);
    HDC memoryDc = CreateCompatibleDC(windowDc);
    HBITMAP bitmap = CreateDIBSection(windowDc, &info, DIB_RGB_COLORS, &pixels, nullptr, 0);
    HGDIOBJ oldBitmap = SelectObject(memoryDc, bitmap);
    PatBlt(memoryDc, 0, 0, width, height, WHITENESS);
    SendMessage(window, WM_PRINT, reinterpret_cast<WPARAM>(memoryDc),
        PRF_CLIENT | PRF_CHILDREN | PRF_ERASEBKGND);

    BITMAPFILEHEADER fileHeader = {};
    const DWORD pixelBytes = static_cast<DWORD>(width * height * 4);
    fileHeader.bfType = 0x4D42;
    fileHeader.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
    fileHeader.bfSize = fileHeader.bfOffBits + pixelBytes;

    HANDLE file = CreateFile(path, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    bool saved = false;
    if (file != INVALID_HANDLE_VALUE)
    {
        DWORD written = 0;
        saved = WriteFile(file, &fileHeader, sizeof(fileHeader), &written, nullptr) != FALSE &&
                WriteFile(file, &info.bmiHeader, sizeof(info.bmiHeader), &written, nullptr) != FALSE &&
                WriteFile(file, pixels, pixelBytes, &written, nullptr) != FALSE;
        CloseHandle(file);
    }

    SelectObject(memoryDc, oldBitmap);
    DeleteObject(bitmap);
    DeleteDC(memoryDc);
    ReleaseDC(window, windowDc);
    return saved;
}

void ApplyThemeNeutralState(NMLVCUSTOMDRAW *draw)
{
    draw->nmcd.uItemState &= ~CDIS_SELECTED;
    draw->iStateId = LISS_NORMAL;
    draw->clrFace = kFocusColor;
    draw->clrTextBk = kFocusColor;
    draw->clrText = kFocusTextColor;
}

LRESULT HandleCustomDraw(HWND list, ProbeMode mode, NMLVCUSTOMDRAW *draw)
{
    const DWORD stage = draw->nmcd.dwDrawStage;
    if (stage == CDDS_PREPAINT)
        return CDRF_NOTIFYITEMDRAW;

    const int item = static_cast<int>(draw->nmcd.dwItemSpec);
    const bool target = item == 2;

    if (stage == CDDS_ITEMPREPAINT)
    {
        if (target && (mode == ProbeThemeStateAndFill || mode == ProbeLegacyCellFill))
        {
            RECT bounds = {};
            const int boundsType = mode == ProbeThemeStateAndFill ? LVIR_BOUNDS : LVIR_SELECTBOUNDS;
            ListView_GetItemRect(list, item, &bounds, boundsType);
            HBRUSH brush = CreateSolidBrush(kFocusColor);
            FillRect(draw->nmcd.hdc, &bounds, brush);
            DeleteObject(brush);
        }

        return target ? CDRF_NOTIFYSUBITEMDRAW : CDRF_DODEFAULT;
    }

    if (stage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM) && target)
    {
        if (mode == ProbeLegacyCellFill && draw->iSubItem != 0)
            return CDRF_DODEFAULT;

        if (mode == ProbeCurrent)
        {
            draw->nmcd.uItemState &= ~CDIS_SELECTED;
            draw->clrTextBk = kFocusColor;
            draw->clrText = kFocusTextColor;
        }
        else
        {
            ApplyThemeNeutralState(draw);
        }

        return CDRF_NEWFONT;
    }

    return CDRF_DODEFAULT;
}

void AddListContent(HWND list)
{
    const wchar_t *headers[] = { L"Name", L"Size", L"Type", L"Modified" };
    const int widths[] = { 145, 80, 95, 115 };
    for (int column = 0; column < 4; ++column)
    {
        LVCOLUMN value = {};
        value.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        value.pszText = const_cast<wchar_t *>(headers[column]);
        value.cx = widths[column];
        value.iSubItem = column;
        ListView_InsertColumn(list, column, &value);
    }

    const wchar_t *names[] = { L"Alpha", L"Bravo", L"Selected row", L"Delta", L"Echo" };
    for (int row = 0; row < 5; ++row)
    {
        LVITEM value = {};
        value.mask = LVIF_TEXT;
        value.iItem = row;
        value.pszText = const_cast<wchar_t *>(names[row]);
        ListView_InsertItem(list, &value);
        ListView_SetItemText(list, row, 1, const_cast<wchar_t *>(L"1 KB"));
        ListView_SetItemText(list, row, 2, const_cast<wchar_t *>(L"File folder"));
        ListView_SetItemText(list, row, 3, const_cast<wchar_t *>(L"2026-08-29"));
    }

    ListView_SetItemState(list, 2, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetSelectionMark(list, 2);
}

LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        {
            const wchar_t *labels[] = {
                L"Current: clrTextBk + clear CDIS_SELECTED",
                L"Candidate: + LISS_NORMAL + clrFace",
                L"Candidate: explicit full-row fill",
                L"Candidate: legacy name-cell fill"
            };

            HFONT font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            for (int index = 0; index < 4; ++index)
            {
                const int left = 12 + ((index % 2) * 665);
                const int top = 12 + ((index / 2) * 235);
                HWND label = CreateWindowEx(0, L"STATIC", labels[index], WS_CHILD | WS_VISIBLE,
                    left, top, 650, 22, window, nullptr, nullptr, nullptr);
                SendMessage(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);

                gLists[index] = CreateWindowEx(WS_EX_CLIENTEDGE, WC_LISTVIEW, L"",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS,
                    left, top + 26, 650, 180, window,
                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(1001 + index)), nullptr, nullptr);
                SendMessage(gLists[index], WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
                DWORD extendedStyle = LVS_EX_DOUBLEBUFFER;
                if (index != ProbeLegacyCellFill)
                    extendedStyle |= LVS_EX_FULLROWSELECT;
                ListView_SetExtendedListViewStyle(gLists[index], extendedStyle);
                SetWindowTheme(gLists[index], L"Explorer", nullptr);
                AddListContent(gLists[index]);
            }

            HWND note = CreateWindowEx(0, L"STATIC",
                L"All lists use Explorer theme, LVS_SHOWSELALWAYS and RGB(255,255,0); the last list has LVS_EX_FULLROWSELECT disabled.",
                WS_CHILD | WS_VISIBLE, 12, 488, 1300, 22, window, nullptr, nullptr, nullptr);
            SendMessage(note, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);

            HWND focusSink = CreateWindowEx(0, L"BUTTON", L"Focus outside the lists",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP, 12, 518, 190, 30, window,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(1100)), nullptr, nullptr);
            SendMessage(focusSink, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
            SetFocus(focusSink);
            SetTimer(window, 1, 500, nullptr);
            return 0;
        }

    case WM_NOTIFY:
        {
            NMHDR *header = reinterpret_cast<NMHDR *>(lParam);
            if (header != nullptr && header->code == NM_CUSTOMDRAW)
            {
                for (int index = 0; index < 4; ++index)
                {
                    if (header->hwndFrom == gLists[index])
                    {
                        return HandleCustomDraw(gLists[index], static_cast<ProbeMode>(index),
                            reinterpret_cast<NMLVCUSTOMDRAW *>(lParam));
                    }
                }
            }
            break;
        }

    case WM_TIMER:
        if (wParam == 1)
        {
            KillTimer(window, 1);
            const bool saved = SaveClientBitmap(window,
                L"C:\\Users\\Public\\Documents\\ESTsoft\\CreatorTemp\\row_focus_visual_probe.bmp");
            SetWindowText(window, saved
                ? L"FxFile row-focus visual probe - captured"
                : L"FxFile row-focus visual probe - capture failed");
            return 0;
        }
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(window, message, wParam, lParam);
}
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int showCommand)
{
    INITCOMMONCONTROLSEX controls = { sizeof(controls), ICC_LISTVIEW_CLASSES };
    InitCommonControlsEx(&controls);

    WNDCLASS windowClass = {};
    windowClass.lpfnWndProc = WindowProc;
    windowClass.hInstance = instance;
    windowClass.hCursor = LoadCursor(nullptr, IDC_ARROW);
    windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    windowClass.lpszClassName = L"FxFileRowFocusVisualProbe";
    RegisterClass(&windowClass);

    HWND window = CreateWindowEx(0, windowClass.lpszClassName,
        L"FxFile row-focus visual probe", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 1360, 610,
        nullptr, nullptr, instance, nullptr);
    if (window == nullptr)
        return 1;

    ShowWindow(window, showCommand);
    UpdateWindow(window);

    MSG message = {};
    while (GetMessage(&message, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&message);
        DispatchMessage(&message);
    }

    return static_cast<int>(message.wParam);
}
