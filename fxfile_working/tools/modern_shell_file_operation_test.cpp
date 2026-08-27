#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <shellapi.h>

#include <stdio.h>
#include <string>

#include "../src/fxfile/modern_shell_file_operation.h"

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
    if (aArgumentCount < 3 || aArgumentCount > 4)
    {
        fwprintf(stderr, L"usage: modern_shell_file_operation_test <copy|move> <source> <target-dir>\n"
                         L"       modern_shell_file_operation_test <delete|trash> <source>\n");
        return 2;
    }
    const bool sDelete = _wcsicmp(aArguments[1], L"delete") == 0 ||
                         _wcsicmp(aArguments[1], L"trash") == 0;
    if ((!sDelete && aArgumentCount != 4) ||
        (sDelete && aArgumentCount != 3))
        return 2;

    const HRESULT sCom = ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(sCom))
        return 3;
    std::wstring sSource = doubleNull(aArguments[2]);
    std::wstring sTarget = sDelete ? std::wstring() :
                                     doubleNull(aArguments[3]);
    SHFILEOPSTRUCT sOperation = {0};
    sOperation.wFunc = sDelete ? FO_DELETE :
        (_wcsicmp(aArguments[1], L"move") == 0 ? FO_MOVE : FO_COPY);
    sOperation.pFrom = sSource.c_str();
    sOperation.pTo = sDelete ? NULL : sTarget.c_str();
    sOperation.fFlags = FOF_NOCONFIRMATION | FOF_NOERRORUI;
    if (_wcsicmp(aArguments[1], L"trash") == 0)
        sOperation.fFlags |= FOF_ALLOWUNDO;

    HRESULT sError = S_OK;
    const fxfile::ModernShellFileOperation::Result sResult =
        fxfile::ModernShellFileOperation::tryExecute(&sOperation, &sError);
    ::CoUninitialize();
    wprintf(L"result=%d hr=0x%08lx\n", static_cast<int>(sResult), sError);
    return sResult == fxfile::ModernShellFileOperation::ResultSucceeded ? 0 :
           (sResult == fxfile::ModernShellFileOperation::ResultNotApplicable ? 4 : 1);
}
