//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#if defined(FXFILE_MODERN_SHELL_STANDALONE)
#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <shellapi.h>
#else
#include "stdafx.h"
#endif
#include "modern_shell_file_operation.h"

#include <string>
#include <vector>
#include <set>
#include <new>
#include <cwctype>

#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>

namespace fxfile
{
namespace
{
std::wstring normalizePath(const std::wstring &aPath)
{
    wchar_t sFull[MAX_PATH + 1] = {0};
    std::wstring sPath(aPath);
    if (::GetFullPathNameW(aPath.c_str(), MAX_PATH, sFull, NULL) > 0)
        sPath.assign(sFull);
    while (sPath.length() > 3 &&
           (sPath.back() == L'\\' || sPath.back() == L'/'))
        sPath.pop_back();
    for (size_t i = 0; i < sPath.length(); ++i)
        sPath[i] = static_cast<wchar_t>(::towlower(sPath[i]));
    return sPath;
}

bool isInside(const std::wstring &aPath, const std::wstring &aParent)
{
    const std::wstring sPath = normalizePath(aPath);
    const std::wstring sParent = normalizePath(aParent);
    return !sParent.empty() && sPath.length() >= sParent.length() &&
           sPath.compare(0, sParent.length(), sParent) == 0 &&
           (sPath.length() == sParent.length() ||
            sPath[sParent.length()] == L'\\');
}

bool isProtectedDeleteSource(const std::wstring &aPath)
{
    if (::PathIsRootW(aPath.c_str()))
        return true;

    // FxFile never automates deletion inside the Windows directory.  This is
    // deliberately stricter than an administrator token: callers can use the
    // Windows-owned security UI, but this engine will not bypass WRP/UAC.
    wchar_t sWindows[MAX_PATH + 1] = {0};
    if (::GetWindowsDirectoryW(sWindows, MAX_PATH) > 0 &&
        (isInside(aPath, sWindows) || isInside(sWindows, aPath)))
        return true;

    typedef BOOL (WINAPI *SfcIsFileProtectedProc)(HANDLE, LPCWSTR);
    HMODULE sSfc = ::LoadLibraryW(L"sfc.dll");
    if (sSfc == NULL)
        return false;
    SfcIsFileProtectedProc sIsProtected =
        reinterpret_cast<SfcIsFileProtectedProc>(
            ::GetProcAddress(sSfc, "SfcIsFileProtected"));
    const bool sProtected = sIsProtected != NULL &&
                            sIsProtected(NULL, aPath.c_str()) != FALSE;
    ::FreeLibrary(sSfc);
    return sProtected;
}

std::wstring shellPath(IShellItem *aItem)
{
    if (aItem == NULL)
        return std::wstring();
    PWSTR sName = NULL;
    if (FAILED(aItem->GetDisplayName(SIGDN_FILESYSPATH, &sName)) ||
        sName == NULL)
        return std::wstring();
    const std::wstring sPath(sName);
    ::CoTaskMemFree(sName);
    return sPath;
}

class OperationProgressSink : public IFileOperationProgressSink
{
public:
    OperationProgressSink(void) : mReferenceCount(1), mDelete(false) {}

    void setDelete(bool aDelete) { mDelete = aDelete; }
    const std::set<std::wstring> &completed(void) const { return mCompleted; }
    const std::set<std::wstring> &failed(void) const { return mFailed; }

    STDMETHODIMP QueryInterface(REFIID aIid, void **aObject)
    {
        if (aObject == NULL)
            return E_POINTER;
        *aObject = NULL;
        if (aIid == IID_IUnknown || aIid == IID_IFileOperationProgressSink)
        {
            *aObject = static_cast<IFileOperationProgressSink *>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef(void) { return ::InterlockedIncrement(&mReferenceCount); }
    STDMETHODIMP_(ULONG) Release(void)
    {
        const ULONG sCount = ::InterlockedDecrement(&mReferenceCount);
        if (sCount == 0)
            delete this;
        return sCount;
    }

    STDMETHODIMP StartOperations(void) { return S_OK; }
    STDMETHODIMP FinishOperations(HRESULT) { return S_OK; }
    STDMETHODIMP PreRenameItem(DWORD, IShellItem *, LPCWSTR) { return S_OK; }
    STDMETHODIMP PostRenameItem(DWORD, IShellItem *, LPCWSTR, HRESULT, IShellItem *) { return S_OK; }
    STDMETHODIMP PreMoveItem(DWORD, IShellItem *, IShellItem *, LPCWSTR) { return S_OK; }
    STDMETHODIMP PostMoveItem(DWORD, IShellItem *aItem, IShellItem *, LPCWSTR,
                              HRESULT aResult, IShellItem *)
    {
        record(aItem, aResult);
        return S_OK;
    }
    STDMETHODIMP PreCopyItem(DWORD, IShellItem *, IShellItem *, LPCWSTR) { return S_OK; }
    STDMETHODIMP PostCopyItem(DWORD, IShellItem *aItem, IShellItem *, LPCWSTR,
                              HRESULT aResult, IShellItem *)
    {
        record(aItem, aResult);
        return S_OK;
    }
    STDMETHODIMP PreDeleteItem(DWORD, IShellItem *) { return S_OK; }
    STDMETHODIMP PostDeleteItem(DWORD, IShellItem *aItem, HRESULT aResult,
                                IShellItem *)
    {
        if (!mDelete)
            return S_OK;
        record(aItem, aResult);
        return S_OK;
    }
    STDMETHODIMP PreNewItem(DWORD, IShellItem *, LPCWSTR) { return S_OK; }
    STDMETHODIMP PostNewItem(DWORD, IShellItem *, LPCWSTR, LPCWSTR, DWORD,
                             HRESULT, IShellItem *) { return S_OK; }
    STDMETHODIMP UpdateProgress(UINT, UINT) { return S_OK; }
    STDMETHODIMP ResetTimer(void) { return S_OK; }
    STDMETHODIMP PauseTimer(void) { return S_OK; }
    STDMETHODIMP ResumeTimer(void) { return S_OK; }

private:
    void record(IShellItem *aItem, HRESULT aResult)
    {
        const std::wstring sPath = normalizePath(shellPath(aItem));
        if (!sPath.empty())
            (SUCCEEDED(aResult) ? mCompleted : mFailed).insert(sPath);
    }

    volatile LONG mReferenceCount;
    bool mDelete;
    std::set<std::wstring> mCompleted;
    std::set<std::wstring> mFailed;
};

void showDeleteSummary(const SHFILEOPSTRUCT *aOperation,
                       const std::vector<std::wstring> &aSources,
                       const OperationProgressSink *aSink,
                       bool aCancelled)
{
    if (aOperation == NULL || aSink == NULL ||
        (aOperation->fFlags & FOF_NOERRORUI) != 0)
        return;
    const size_t sCompleted = aSink->completed().size();
    const size_t sFailed = aSink->failed().size();
    const size_t sRemaining = aSources.size() > sCompleted + sFailed ?
                              aSources.size() - sCompleted - sFailed : 0;
    wchar_t sMessage[512] = {0};
    _snwprintf_s(sMessage, _countof(sMessage), _TRUNCATE,
                 L"삭제 작업이 완료되지 않았습니다.\n\n"
                 L"완료: %Iu개\n실패: %Iu개\n처리되지 않음: %Iu개\n\n"
                 L"완료된 항목과 남은 항목을 혼동해 성공으로 보고하지 않습니다.",
                 sCompleted, sFailed, sRemaining);
    ::MessageBoxW(aOperation->hwnd, sMessage, L"FxFile 삭제 작업",
                  MB_OK | (aCancelled ? MB_ICONINFORMATION : MB_ICONERROR) |
                  MB_TASKMODAL);
}

bool parseSources(const wchar_t *aMultiString,
                  std::vector<std::wstring> &aSources)
{
    if (aMultiString == NULL)
        return false;

    const wchar_t *sPath = aMultiString;
    while (*sPath != L'\0')
    {
        std::wstring sSource(sPath);
        if (sSource.empty() || sSource.find_first_of(L"*?") != std::wstring::npos)
            return false;
        aSources.push_back(sSource);
        sPath += sSource.length() + 1;
    }
    return !aSources.empty();
}

std::wstring leafName(const std::wstring &aPath)
{
    std::wstring sPath(aPath);
    while (sPath.length() > 3 &&
           (sPath.back() == L'\\' || sPath.back() == L'/'))
        sPath.pop_back();
    const size_t sSlash = sPath.find_last_of(L"\\/");
    return sSlash == std::wstring::npos ? sPath : sPath.substr(sSlash + 1);
}

std::wstring joinPath(const std::wstring &aDirectory,
                      const std::wstring &aLeaf)
{
    std::wstring sPath(aDirectory);
    if (!sPath.empty() && sPath.back() != L'\\')
        sPath.push_back(L'\\');
    sPath += aLeaf;
    return sPath;
}

bool targetCollisionExists(const std::vector<std::wstring> &aSources,
                           const wchar_t *aTarget)
{
    if (aTarget == NULL)
        return true;
    const DWORD sAttributes = ::GetFileAttributesW(aTarget);
    if (sAttributes == INVALID_FILE_ATTRIBUTES ||
        (sAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        return true;

    for (size_t i = 0; i < aSources.size(); ++i)
    {
        const std::wstring sLeaf = leafName(aSources[i]);
        if (sLeaf.empty())
            return true;
        const std::wstring sCandidate = joinPath(aTarget, sLeaf);
        if (::GetFileAttributesW(sCandidate.c_str()) != INVALID_FILE_ATTRIBUTES)
            return true;
        const DWORD sError = ::GetLastError();
        if (sError != ERROR_FILE_NOT_FOUND && sError != ERROR_PATH_NOT_FOUND)
            return true;
    }
    return false;
}
} // namespace anonymous

ModernShellFileOperation::Result ModernShellFileOperation::tryExecute(
    SHFILEOPSTRUCT *aFileOperation,
    HRESULT *aError)
{
    if (aError != NULL)
        *aError = S_OK;
    if (aFileOperation == NULL || aFileOperation->pFrom == NULL)
        return ResultNotApplicable;
    if (aFileOperation->wFunc != FO_COPY &&
        aFileOperation->wFunc != FO_MOVE &&
        aFileOperation->wFunc != FO_DELETE)
        return ResultNotApplicable;
    if ((aFileOperation->fFlags &
         (FOF_RENAMEONCOLLISION | FOF_MULTIDESTFILES)) != 0)
        return ResultNotApplicable;

    std::vector<std::wstring> sSources;
    if (!parseSources(aFileOperation->pFrom, sSources))
        return ResultNotApplicable;
    if (aFileOperation->wFunc == FO_DELETE)
    {
        for (size_t i = 0; i < sSources.size(); ++i)
        {
            if (isProtectedDeleteSource(sSources[i]))
            {
                if (aError != NULL)
                    *aError = E_ACCESSDENIED;
                if ((aFileOperation->fFlags & FOF_NOERRORUI) == 0)
                    ::MessageBoxW(aFileOperation->hwnd,
                                  L"Windows 보호 경로는 FxFile에서 자동 삭제하지 않습니다.\n\n"
                                  L"편집 > 파일·폴더 잠금 관리 > Windows 보안에서 "
                                  L"Windows가 제공하는 권한 UI를 사용하십시오.",
                                  L"FxFile 보호 경계",
                                  MB_OK | MB_ICONWARNING | MB_TASKMODAL);
                return ResultFailed;
            }
        }
    }
    if ((aFileOperation->wFunc == FO_COPY ||
         aFileOperation->wFunc == FO_MOVE) &&
        targetCollisionExists(sSources, aFileOperation->pTo))
        return ResultNotApplicable;

    IFileOperation *sOperation = NULL;
    HRESULT sResult = ::CoCreateInstance(CLSID_FileOperation, NULL,
                                          CLSCTX_INPROC_SERVER,
                                          IID_PPV_ARGS(&sOperation));
    if (FAILED(sResult) || sOperation == NULL)
    {
        if (aError != NULL)
            *aError = sResult;
        return ResultNotApplicable;
    }

    sOperation->SetOwnerWindow(aFileOperation->hwnd);
    OperationProgressSink *sSink = new (std::nothrow) OperationProgressSink;
    DWORD sCookie = 0;
    if (sSink != NULL)
    {
        sSink->setDelete(aFileOperation->wFunc == FO_DELETE);
        if (FAILED(sOperation->Advise(sSink, &sCookie)))
            sCookie = 0;
    }
    FILEOP_FLAGS sFlags = static_cast<FILEOP_FLAGS>(
        aFileOperation->fFlags & ~FOF_WANTMAPPINGHANDLE);
    if (aFileOperation->wFunc == FO_DELETE &&
        (aFileOperation->fFlags & FOF_ALLOWUNDO) != 0)
        sFlags = static_cast<FILEOP_FLAGS>(sFlags | FOFX_RECYCLEONDELETE);
    sResult = sOperation->SetOperationFlags(sFlags);

    IShellItem *sDestination = NULL;
    if (SUCCEEDED(sResult) && aFileOperation->wFunc != FO_DELETE)
        sResult = ::SHCreateItemFromParsingName(aFileOperation->pTo, NULL,
                                                IID_PPV_ARGS(&sDestination));

    for (size_t i = 0; SUCCEEDED(sResult) && i < sSources.size(); ++i)
    {
        IShellItem *sSource = NULL;
        sResult = ::SHCreateItemFromParsingName(sSources[i].c_str(), NULL,
                                                IID_PPV_ARGS(&sSource));
        if (SUCCEEDED(sResult))
        {
            if (aFileOperation->wFunc == FO_COPY)
                sResult = sOperation->CopyItem(sSource, sDestination,
                                               NULL, NULL);
            else if (aFileOperation->wFunc == FO_MOVE)
                sResult = sOperation->MoveItem(sSource, sDestination,
                                               NULL, NULL);
            else
                sResult = sOperation->DeleteItem(sSource, NULL);
        }
        if (sSource != NULL)
            sSource->Release();
    }

    if (SUCCEEDED(sResult))
        sResult = sOperation->PerformOperations();

    BOOL sAborted = FALSE;
    if (SUCCEEDED(sResult))
        sResult = sOperation->GetAnyOperationsAborted(&sAborted);

    if (sCookie != 0)
        sOperation->Unadvise(sCookie);
    if (sDestination != NULL)
        sDestination->Release();
    sOperation->Release();

    const bool sCancelled = sAborted ||
        sResult == HRESULT_FROM_WIN32(ERROR_CANCELLED) ||
        sResult == HRESULT_FROM_WIN32(ERROR_REQUEST_ABORTED);
    if (sSink != NULL && !sSink->failed().empty() && SUCCEEDED(sResult))
        sResult = E_FAIL;
    if (aError != NULL)
        *aError = sResult;
    if (aFileOperation->wFunc == FO_DELETE &&
        (sCancelled || FAILED(sResult)))
        showDeleteSummary(aFileOperation, sSources, sSink, sCancelled);
    if (sSink != NULL)
        sSink->Release();
    if (sCancelled)
        return ResultCancelled;
    return SUCCEEDED(sResult) ? ResultSucceeded : ResultFailed;
}
} // namespace fxfile
