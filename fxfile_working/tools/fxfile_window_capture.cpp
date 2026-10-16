#define UNICODE
#define _UNICODE

#include <windows.h>
#include <string.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

namespace
{
const wchar_t *kOutputPath =
    L"C:\\Users\\Public\\Documents\\ESTsoft\\CreatorTemp\\fxfile_deployed_capture.bmp";

BOOL CALLBACK FindFxFileWindow(HWND window, LPARAM parameter)
{
    if (IsWindowVisible(window) == FALSE)
        return TRUE;

    wchar_t title[512] = { 0 };
    GetWindowText(window, title, static_cast<int>(sizeof(title) / sizeof(title[0])));
    const size_t length = wcslen(title);
    const wchar_t suffix[] = L" - fxfile";
    const size_t suffixLength = (sizeof(suffix) / sizeof(suffix[0])) - 1;
    if (length < suffixLength || _wcsicmp(title + length - suffixLength, suffix) != 0)
        return TRUE;

    *reinterpret_cast<HWND *>(parameter) = window;
    return FALSE;
}

bool WriteBitmap(HWND window)
{
    RECT bounds = {};
    if (GetWindowRect(window, &bounds) == FALSE)
        return false;

    const int width = bounds.right - bounds.left;
    const int height = bounds.bottom - bounds.top;
    if (width <= 0 || height <= 0)
        return false;

    BITMAPINFO bitmapInfo = {};
    bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmapInfo.bmiHeader.biWidth = width;
    bitmapInfo.bmiHeader.biHeight = -height;
    bitmapInfo.bmiHeader.biPlanes = 1;
    bitmapInfo.bmiHeader.biBitCount = 32;
    bitmapInfo.bmiHeader.biCompression = BI_RGB;

    void *pixels = nullptr;
    HDC windowDc = GetWindowDC(window);
    if (windowDc == nullptr)
        return false;

    HDC memoryDc = CreateCompatibleDC(windowDc);
    HBITMAP bitmap = CreateDIBSection(windowDc, &bitmapInfo, DIB_RGB_COLORS, &pixels, nullptr, 0);
    if (memoryDc == nullptr || bitmap == nullptr || pixels == nullptr)
    {
        if (bitmap != nullptr)
            DeleteObject(bitmap);
        if (memoryDc != nullptr)
            DeleteDC(memoryDc);
        ReleaseDC(window, windowDc);
        return false;
    }

    HGDIOBJ oldBitmap = SelectObject(memoryDc, bitmap);
    PatBlt(memoryDc, 0, 0, width, height, WHITENESS);

    // Prefer the already composed on-screen window pixels.  PrintWindow can
    // initiate a separate WM_PRINT paint path whose themed ListView selection
    // state differs from what the user is currently seeing.
    BOOL rendered = BitBlt(memoryDc, 0, 0, width, height, windowDc, 0, 0,
        SRCCOPY | CAPTUREBLT);
    if (rendered == FALSE)
    {
#ifndef PW_RENDERFULLCONTENT
#define PW_RENDERFULLCONTENT 0x00000002
#endif
        rendered = PrintWindow(window, memoryDc, PW_RENDERFULLCONTENT);
        if (rendered == FALSE)
        {
            DWORD_PTR ignored = 0;
            rendered = SendMessageTimeout(window, WM_PRINT, reinterpret_cast<WPARAM>(memoryDc),
                PRF_NONCLIENT | PRF_CLIENT | PRF_CHILDREN | PRF_ERASEBKGND,
                SMTO_ABORTIFHUNG | SMTO_BLOCK, 3000, &ignored) != 0;
        }
    }

    bool saved = false;
    if (rendered != FALSE)
    {
        BITMAPFILEHEADER fileHeader = {};
        const DWORD pixelBytes = static_cast<DWORD>(width * height * 4);
        fileHeader.bfType = 0x4D42;
        fileHeader.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
        fileHeader.bfSize = fileHeader.bfOffBits + pixelBytes;

        HANDLE file = CreateFile(kOutputPath, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
            FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file != INVALID_HANDLE_VALUE)
        {
            DWORD written = 0;
            saved = WriteFile(file, &fileHeader, sizeof(fileHeader), &written, nullptr) != FALSE &&
                    WriteFile(file, &bitmapInfo.bmiHeader, sizeof(bitmapInfo.bmiHeader), &written, nullptr) != FALSE &&
                    WriteFile(file, pixels, pixelBytes, &written, nullptr) != FALSE;
            CloseHandle(file);
        }
    }

    SelectObject(memoryDc, oldBitmap);
    DeleteObject(bitmap);
    DeleteDC(memoryDc);
    ReleaseDC(window, windowDc);
    return saved;
}
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    HWND fxfileWindow = nullptr;
    EnumWindows(FindFxFileWindow, reinterpret_cast<LPARAM>(&fxfileWindow));
    if (fxfileWindow == nullptr)
        return 2;

    return WriteBitmap(fxfileWindow) ? 0 : 3;
}
