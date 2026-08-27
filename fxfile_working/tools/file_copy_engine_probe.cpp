#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <stdio.h>
#include <string>
#include <vector>
#include <thread>
#include <atomic>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")

namespace
{
std::wstring MakeDoubleNull(const wchar_t *path)
{
    std::wstring result(path);
    result.push_back(L'\0');
    return result;
}

HRESULT RunLegacy(const wchar_t *source, const wchar_t *target)
{
    std::wstring from = MakeDoubleNull(source);
    std::wstring to = MakeDoubleNull(target);
    SHFILEOPSTRUCT operation = {};
    operation.wFunc = FO_COPY;
    operation.pFrom = from.c_str();
    operation.pTo = to.c_str();
    operation.fFlags = FOF_SILENT | FOF_NOCONFIRMATION |
                       FOF_NOERRORUI | FOF_NOCONFIRMMKDIR;
    const int result = SHFileOperation(&operation);
    if (result != 0)
        return HRESULT_FROM_WIN32(result);
    return operation.fAnyOperationsAborted ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
}

HRESULT RunModern(const wchar_t *source, const wchar_t *target)
{
    IFileOperation *operation = nullptr;
    IShellItem *sourceItem = nullptr;
    IShellItem *targetItem = nullptr;

    HRESULT result = CoCreateInstance(CLSID_FileOperation, nullptr,
                                      CLSCTX_INPROC_SERVER,
                                      IID_PPV_ARGS(&operation));
    if (SUCCEEDED(result))
        result = operation->SetOperationFlags(FOF_SILENT | FOF_NOCONFIRMATION |
                                              FOF_NOERRORUI | FOF_NOCONFIRMMKDIR);
    if (SUCCEEDED(result))
        result = SHCreateItemFromParsingName(source, nullptr, IID_PPV_ARGS(&sourceItem));
    if (SUCCEEDED(result))
        result = SHCreateItemFromParsingName(target, nullptr, IID_PPV_ARGS(&targetItem));
    if (SUCCEEDED(result))
        result = operation->CopyItem(sourceItem, targetItem, nullptr, nullptr);
    if (SUCCEEDED(result))
        result = operation->PerformOperations();

    BOOL aborted = FALSE;
    if (SUCCEEDED(result) && SUCCEEDED(operation->GetAnyOperationsAborted(&aborted)) && aborted)
        result = HRESULT_FROM_WIN32(ERROR_CANCELLED);

    if (targetItem != nullptr) targetItem->Release();
    if (sourceItem != nullptr) sourceItem->Release();
    if (operation != nullptr) operation->Release();
    return result;
}

void EnumerateFiles(const std::wstring &directory, std::vector<std::wstring> &files)
{
    WIN32_FIND_DATA findData = {};
    const std::wstring pattern = directory + L"\\*";
    HANDLE find = FindFirstFile(pattern.c_str(), &findData);
    if (find == INVALID_HANDLE_VALUE)
        return;

    do
    {
        if (wcscmp(findData.cFileName, L".") == 0 || wcscmp(findData.cFileName, L"..") == 0)
            continue;
        if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
            files.push_back(directory + L"\\" + findData.cFileName);
    } while (FindNextFile(find, &findData));
    FindClose(find);
}

HRESULT CopyOneFile(const std::wstring &source, const std::wstring &target, bool noBuffering)
{
    COPYFILE2_EXTENDED_PARAMETERS parameters = {};
    parameters.dwSize = sizeof(parameters);
    parameters.dwCopyFlags = noBuffering ? COPY_FILE_NO_BUFFERING : 0;
    return CopyFile2(source.c_str(), target.c_str(), &parameters);
}

HRESULT RunCopyFile2(const wchar_t *sourceDirectory,
                     const wchar_t *targetDirectory,
                     int threadCount,
                     bool noBuffering)
{
    std::vector<std::wstring> files;
    EnumerateFiles(sourceDirectory, files);
    if (files.empty())
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    if (!CreateDirectory(targetDirectory, nullptr) && GetLastError() != ERROR_ALREADY_EXISTS)
        return HRESULT_FROM_WIN32(GetLastError());

    std::atomic<size_t> next(0);
    std::atomic<HRESULT> failure(S_OK);
    std::vector<std::thread> workers;
    const int actualThreads = (threadCount < 1) ? 1 : threadCount;
    for (int worker = 0; worker < actualThreads; ++worker)
    {
        workers.emplace_back([&]() {
            for (;;)
            {
                const size_t index = next.fetch_add(1);
                if (index >= files.size() || FAILED(failure.load()))
                    break;
                const wchar_t *leaf = PathFindFileName(files[index].c_str());
                const std::wstring target = std::wstring(targetDirectory) + L"\\" + leaf;
                const HRESULT result = CopyOneFile(files[index], target, noBuffering);
                if (FAILED(result))
                    failure.store(result);
            }
        });
    }
    for (auto &worker : workers)
        worker.join();
    return failure.load();
}
}

int wmain(int argc, wchar_t **argv)
{
    if (argc < 4)
    {
        fwprintf(stderr, L"usage: file_copy_engine_probe <legacy|modern|copyfile2|copyfile2mt|copyfile2large> <source> <target>\n");
        return 2;
    }

    const HRESULT comResult = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(comResult))
        return 3;

    LARGE_INTEGER frequency = {}, begin = {}, end = {};
    QueryPerformanceFrequency(&frequency);
    QueryPerformanceCounter(&begin);

    HRESULT result = E_INVALIDARG;
    if (_wcsicmp(argv[1], L"legacy") == 0)
        result = RunLegacy(argv[2], argv[3]);
    else if (_wcsicmp(argv[1], L"modern") == 0)
        result = RunModern(argv[2], argv[3]);
    else if (_wcsicmp(argv[1], L"copyfile2") == 0)
        result = RunCopyFile2(argv[2], argv[3], 1, false);
    else if (_wcsicmp(argv[1], L"copyfile2mt") == 0)
        result = RunCopyFile2(argv[2], argv[3], 4, false);
    else if (_wcsicmp(argv[1], L"copyfile2large") == 0)
        result = RunCopyFile2(argv[2], argv[3], 1, true);

    QueryPerformanceCounter(&end);
    CoUninitialize();
    const double seconds = static_cast<double>(end.QuadPart - begin.QuadPart) /
                           static_cast<double>(frequency.QuadPart);
    wprintf(L"engine=%s seconds=%.6f hresult=0x%08X\n",
            argv[1], seconds, static_cast<unsigned>(result));
    return FAILED(result) ? 1 : 0;
}
