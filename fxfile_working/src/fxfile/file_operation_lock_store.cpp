// FxFile-local, reversible file/folder operation locks.
#include "stdafx.h"
#include "file_operation_lock_store.h"

#include "conf_dir.h"

#include <algorithm>
#include <cwctype>

namespace fxfile
{
FileOperationLockStore::FileOperationLockStore(void)
    : mLoaded(false)
{
}

FileOperationLockStore &FileOperationLockStore::instance(void)
{
    static FileOperationLockStore sInstance;
    return sInstance;
}

std::wstring FileOperationLockStore::normalize(const std::wstring &aPath)
{
    if (aPath.empty())
        return std::wstring();
    wchar_t sFull[XPR_MAX_PATH + 1] = {0};
    std::wstring sPath = aPath;
    if (::GetFullPathNameW(aPath.c_str(), XPR_MAX_PATH, sFull, NULL) > 0)
        sPath.assign(sFull);
    while (sPath.length() > 3 &&
           (sPath.back() == L'\\' || sPath.back() == L'/'))
        sPath.pop_back();
    std::transform(sPath.begin(), sPath.end(), sPath.begin(),
                   [](wchar_t c) { return static_cast<wchar_t>(::towlower(c)); });
    return sPath;
}

bool FileOperationLockStore::isInside(const std::wstring &aPath,
                                      const std::wstring &aParent)
{
    if (aPath.length() < aParent.length() ||
        aPath.compare(0, aParent.length(), aParent) != 0)
        return false;
    return aPath.length() == aParent.length() ||
           aPath[aParent.length()] == L'\\';
}

std::wstring FileOperationLockStore::storagePath(void) const
{
    wchar_t sDir[XPR_MAX_PATH + 1] = {0};
    if (XPR_IS_FALSE(ConfDir::instance().getSaveDir(sDir, XPR_MAX_PATH)))
        return std::wstring();
    std::wstring sPath(sDir);
    if (!sPath.empty() && sPath.back() != L'\\')
        sPath.push_back(L'\\');
    sPath += L"fxfile-operation-locks.conf";
    return sPath;
}

void FileOperationLockStore::ensureLoaded(void)
{
    if (mLoaded)
        return;
    mLoaded = true;

    const std::wstring sPath = storagePath();
    HANDLE sFile = ::CreateFileW(sPath.c_str(), GENERIC_READ,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE |
                                  FILE_SHARE_DELETE,
                                  NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL,
                                  NULL);
    if (sFile == INVALID_HANDLE_VALUE)
        return;
    LARGE_INTEGER sSize = {0};
    if (!::GetFileSizeEx(sFile, &sSize) || sSize.QuadPart < 2 ||
        sSize.QuadPart > 4 * 1024 * 1024)
    {
        ::CloseHandle(sFile);
        return;
    }
    std::vector<wchar_t> sText(static_cast<size_t>(sSize.QuadPart / 2) + 1, 0);
    DWORD sRead = 0;
    if (!::ReadFile(sFile, &sText[0], static_cast<DWORD>(sSize.QuadPart),
                    &sRead, NULL))
    {
        ::CloseHandle(sFile);
        return;
    }
    ::CloseHandle(sFile);
    size_t sStart = (sText[0] == 0xfeff) ? 1 : 0;
    std::wstring sContents(&sText[sStart], sRead / sizeof(wchar_t) - sStart);
    size_t sPos = 0;
    while (sPos <= sContents.length())
    {
        const size_t sEnd = sContents.find_first_of(L"\r\n", sPos);
        const std::wstring sLine = normalize(sContents.substr(
            sPos, sEnd == std::wstring::npos ? std::wstring::npos : sEnd - sPos));
        if (!sLine.empty())
            mPaths.insert(sLine);
        if (sEnd == std::wstring::npos)
            break;
        sPos = sEnd + 1;
        while (sPos < sContents.length() &&
               (sContents[sPos] == L'\r' || sContents[sPos] == L'\n'))
            ++sPos;
    }
}

bool FileOperationLockStore::save(void)
{
    const std::wstring sPath = storagePath();
    if (sPath.empty())
        return false;
    wchar_t sTempSuffix[64] = {0};
    _snwprintf_s(sTempSuffix, _countof(sTempSuffix), _TRUNCATE,
                 L".tmp.%lu", ::GetCurrentProcessId());
    const std::wstring sTemp = sPath + sTempSuffix;

    HANDLE sFile = ::CreateFileW(sTemp.c_str(), GENERIC_WRITE, 0, NULL,
                                  CREATE_ALWAYS,
                                  FILE_ATTRIBUTE_TEMPORARY, NULL);
    if (sFile == INVALID_HANDLE_VALUE)
        return false;
    std::wstring sText(1, static_cast<wchar_t>(0xfeff));
    for (std::set<std::wstring>::const_iterator sIt = mPaths.begin();
         sIt != mPaths.end(); ++sIt)
    {
        sText += *sIt;
        sText += L"\r\n";
    }
    DWORD sWritten = 0;
    const DWORD sBytes = static_cast<DWORD>(sText.size() * sizeof(wchar_t));
    const bool sOk = ::WriteFile(sFile, sText.data(), sBytes, &sWritten, NULL) &&
                     sWritten == sBytes && ::FlushFileBuffers(sFile);
    ::CloseHandle(sFile);
    if (!sOk || !::MoveFileExW(sTemp.c_str(), sPath.c_str(),
                                MOVEFILE_REPLACE_EXISTING |
                                MOVEFILE_WRITE_THROUGH))
    {
        ::DeleteFileW(sTemp.c_str());
        return false;
    }
    return true;
}

bool FileOperationLockStore::isLocked(const std::wstring &aPath)
{
    std::lock_guard<std::mutex> sGuard(mMutex);
    ensureLoaded();
    const std::wstring sPath = normalize(aPath);
    for (std::set<std::wstring>::const_iterator sIt = mPaths.begin();
         sIt != mPaths.end(); ++sIt)
        if (isInside(sPath, *sIt))
            return true;
    return false;
}

bool FileOperationLockStore::affectsLockedPath(const std::wstring &aPath)
{
    std::lock_guard<std::mutex> sGuard(mMutex);
    ensureLoaded();
    const std::wstring sPath = normalize(aPath);
    for (std::set<std::wstring>::const_iterator sIt = mPaths.begin();
         sIt != mPaths.end(); ++sIt)
        if (isInside(sPath, *sIt) || isInside(*sIt, sPath))
            return true;
    return false;
}

bool FileOperationLockStore::hasLocks(void)
{
    std::lock_guard<std::mutex> sGuard(mMutex);
    ensureLoaded();
    return !mPaths.empty();
}

bool FileOperationLockStore::setLocked(
    const std::vector<std::wstring> &aPaths,
    bool aLocked)
{
    std::lock_guard<std::mutex> sGuard(mMutex);
    ensureLoaded();
    const std::set<std::wstring> sBefore = mPaths;
    for (size_t i = 0; i < aPaths.size(); ++i)
    {
        const std::wstring sPath = normalize(aPaths[i]);
        if (sPath.empty())
            continue;
        if (aLocked)
            mPaths.insert(sPath);
        else
            mPaths.erase(sPath);
    }
    if (save())
        return true;
    mPaths = sBefore;
    return false;
}

bool FileOperationLockStore::isOperationBlocked(
    const SHFILEOPSTRUCT *aOperation,
    std::wstring &aBlockedPath)
{
    if (aOperation == NULL)
        return false;
    if ((aOperation->wFunc == FO_COPY || aOperation->wFunc == FO_MOVE) &&
        aOperation->pTo != NULL && isLocked(aOperation->pTo))
    {
        aBlockedPath = aOperation->pTo;
        return true;
    }

    // A destination directory itself may be unlocked while an existing item
    // with the same leaf name is locked.  Check the effective destination as
    // well so collision/fallback paths cannot overwrite or merge into a lock.
    if ((aOperation->wFunc == FO_COPY || aOperation->wFunc == FO_MOVE) &&
        aOperation->pTo != NULL)
    {
        const wchar_t *sSource = aOperation->pFrom;
        while (sSource != NULL && *sSource != L'\0')
        {
            std::wstring sSourcePath(sSource);
            while (sSourcePath.length() > 3 &&
                   (sSourcePath.back() == L'\\' || sSourcePath.back() == L'/'))
                sSourcePath.pop_back();
            const size_t sSlash = sSourcePath.find_last_of(L"\\/");
            const std::wstring sLeaf = sSlash == std::wstring::npos ?
                                       sSourcePath : sSourcePath.substr(sSlash + 1);
            if (!sLeaf.empty())
            {
                std::wstring sCandidate(aOperation->pTo);
                if (!sCandidate.empty() && sCandidate.back() != L'\\')
                    sCandidate.push_back(L'\\');
                sCandidate += sLeaf;
                if (affectsLockedPath(sCandidate))
                {
                    aBlockedPath = sCandidate;
                    return true;
                }
            }
            sSource += wcslen(sSource) + 1;
        }
    }

    if (aOperation->wFunc == FO_COPY)
        return false;
    const wchar_t *sSource = aOperation->pFrom;
    while (sSource != NULL && *sSource != L'\0')
    {
        if (affectsLockedPath(sSource))
        {
            aBlockedPath = sSource;
            return true;
        }
        sSource += wcslen(sSource) + 1;
    }
    return false;
}
} // namespace fxfile
