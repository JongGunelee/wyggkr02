#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <shellapi.h>

#include <stdio.h>
#include <string>

#include "../src/fxfile/adaptive_file_operation.h"

namespace
{
std::wstring doubleNull(const wchar_t *aPath)
{
    std::wstring sPath(aPath);
    sPath.push_back(L'\0');
    return sPath;
}
}

int wmain(int aArgumentCount, wchar_t **aArguments)
{
    const bool sDelete = aArgumentCount == 3 &&
                         (_wcsicmp(aArguments[1], L"delete") == 0 ||
                          _wcsicmp(aArguments[1], L"trash") == 0);
    const bool sMulti = aArgumentCount >= 5 &&
                        _wcsicmp(aArguments[1], L"copymulti") == 0;
    if ((!sMulti && !sDelete && aArgumentCount != 4) ||
        (sMulti && aArgumentCount < 5))
    {
        fwprintf(stderr, L"usage: adaptive_file_operation_test <copy|move> <source> <target-dir>\n"
                         L"       adaptive_file_operation_test copymulti <target-dir> <source> [source...]\n"
                         L"       adaptive_file_operation_test <delete|trash> <source>\n");
        return 2;
    }

    const HRESULT sComResult = ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(sComResult))
        return 3;

    std::wstring sSource;
    std::wstring sTarget;
    if (sMulti)
    {
        sTarget = doubleNull(aArguments[2]);
        for (int i = 3; i < aArgumentCount; ++i)
        {
            sSource += aArguments[i];
            sSource.push_back(L'\0');
        }
        sSource.push_back(L'\0');
    }
    else if (!sDelete)
    {
        sSource = doubleNull(aArguments[2]);
        sTarget = doubleNull(aArguments[3]);
    }
    SHFILEOPSTRUCT sOperation = {0};
    sOperation.wFunc = sDelete ? FO_DELETE :
        ((_wcsicmp(aArguments[1], L"move") == 0) ? FO_MOVE : FO_COPY);
    if (sDelete)
        sSource = doubleNull(aArguments[2]);
    sOperation.pFrom = sSource.c_str();
    sOperation.pTo = sTarget.c_str();
    sOperation.fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI;
    if (_wcsicmp(aArguments[1], L"trash") == 0)
        sOperation.fFlags |= FOF_ALLOWUNDO;

    LARGE_INTEGER sFrequency = {0};
    LARGE_INTEGER sBegin = {0};
    LARGE_INTEGER sEnd = {0};
    ::QueryPerformanceFrequency(&sFrequency);
    ::QueryPerformanceCounter(&sBegin);
    DWORD sError = ERROR_SUCCESS;
    const fxfile::AdaptiveFileOperation::Result sResult =
        fxfile::AdaptiveFileOperation::tryExecute(&sOperation, &sError);
    ::QueryPerformanceCounter(&sEnd);
    ::CoUninitialize();

    const double sSeconds = static_cast<double>(sEnd.QuadPart - sBegin.QuadPart) /
                            static_cast<double>(sFrequency.QuadPart);
    wprintf(L"result=%d error=%lu seconds=%.6f\n",
            static_cast<int>(sResult), sError, sSeconds);
    return (sResult == fxfile::AdaptiveFileOperation::ResultSucceeded) ? 0 :
           (sResult == fxfile::AdaptiveFileOperation::ResultNotApplicable ? 4 : 1);
}
