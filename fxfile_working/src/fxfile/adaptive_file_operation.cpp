//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#if defined(FXFILE_ADAPTIVE_STANDALONE)
#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <shellapi.h>
#else
#include "stdafx.h"
#endif
#include "adaptive_file_operation.h"

#include <atomic>
#include <algorithm>
#include <cwctype>
#include <set>
#include <string>
#include <thread>
#include <vector>

#include <shobjidl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <winioctl.h>

namespace fxfile
{
namespace
{
const ULONGLONG kSmallFileLimit = 8ULL * 1024ULL * 1024ULL;
const size_t kManyFileThreshold = 32;
const size_t kMediumFileThreshold = 8;
const unsigned kMaximumWorkers = 4;
const size_t kRobocopyManyFileThreshold = 1000;
const size_t kRobocopyManyDirectoryThreshold = 128;
const ULONGLONG kRobocopyLargeTreeThreshold = 2ULL * 1024ULL * 1024ULL * 1024ULL;

struct FileJob
{
    std::wstring source;
    std::wstring target;
    ULONGLONG size;
    FILETIME lastWriteTime;
};

struct DirectoryJob
{
    std::wstring source;
    std::wstring target;
    DWORD attributes;
    FILETIME creationTime;
    FILETIME accessTime;
    FILETIME lastWriteTime;
};

struct CopyPlan
{
    std::vector<FileJob> files;
    std::vector<DirectoryJob> directories;
    std::set<std::wstring> sourcePaths;
    std::set<std::wstring> targetPaths;
    ULONGLONG totalBytes;
    bool moveAcrossVolumes;

    CopyPlan(void)
        : totalBytes(0)
        , moveAcrossVolumes(false)
    {
    }
};

// Rollback must never infer ownership from a path alone.  Another process can
// create or replace a path after preflight but before cleanup.  Keep the
// volume/file id obtained from the object created by this operation and mark
// deletion through that same verified handle.
struct CreatedTargetEvidence
{
    std::wstring path;
    DWORD volumeSerialNumber;
    DWORD fileIndexHigh;
    DWORD fileIndexLow;
    bool directory;
    bool valid;

    CreatedTargetEvidence(void)
        : volumeSerialNumber(0), fileIndexHigh(0), fileIndexLow(0),
          directory(false), valid(false)
    {
    }
};

enum StorageKind
{
    StorageUnknown,
    StorageSolidState,
    StorageRotational,
};

struct SharedCopyState
{
    const std::vector<FileJob> *jobs;
    std::vector<CreatedTargetEvidence> *createdFiles;
    std::atomic<size_t> nextJob;
    std::atomic<size_t> finishedJobs;
    std::atomic<size_t> activeWorkers;
    std::atomic<size_t> currentJob;
    std::atomic<size_t> firstFailureJob;
    std::atomic<ULONGLONG> transferredBytes;
    std::atomic<LONG> firstFailure;
    std::atomic<bool> cancelRequested;

    SharedCopyState(void)
        : jobs(NULL)
        , createdFiles(NULL)
        , nextJob(0)
        , finishedJobs(0)
        , activeWorkers(0)
        , currentJob(0)
        , firstFailureJob(static_cast<size_t>(-1))
        , transferredBytes(0)
        , firstFailure(S_OK)
        , cancelRequested(false)
    {
    }
};

struct DeletePlan
{
    std::vector<std::wstring> files;
    std::vector<std::wstring> directories;
};

struct SharedDeleteState
{
    const std::vector<std::wstring> *files;
    std::atomic<size_t> nextFile;
    std::atomic<size_t> completed;
    std::atomic<size_t> activeWorkers;
    std::atomic<DWORD> firstFailure;
    std::atomic<bool> cancelRequested;

    SharedDeleteState(void)
        : files(NULL), nextFile(0), completed(0), activeWorkers(0),
          firstFailure(ERROR_SUCCESS), cancelRequested(false)
    {
    }
};

struct FileCopyContext
{
    SharedCopyState *shared;
    CreatedTargetEvidence *createdEvidence;
    const std::wstring *targetPath;
    ULONGLONG lastTransferred;

    FileCopyContext(void)
        : shared(NULL)
        , createdEvidence(NULL)
        , targetPath(NULL)
        , lastTransferred(0)
    {
    }
};

std::wstring normalizePath(const std::wstring &aPath)
{
    std::wstring sPath(aPath);
    while (sPath.length() > 3 && (sPath.back() == L'\\' || sPath.back() == L'/'))
        sPath.pop_back();

    std::transform(sPath.begin(), sPath.end(), sPath.begin(),
                   [](wchar_t aCharacter) { return static_cast<wchar_t>(::towlower(aCharacter)); });
    return sPath;
}

std::wstring joinPath(const std::wstring &aDirectory, const std::wstring &aLeaf)
{
    std::wstring sPath(aDirectory);
    if (!sPath.empty() && sPath.back() != L'\\')
        sPath.push_back(L'\\');
    sPath += aLeaf;
    return sPath;
}

std::wstring getLeafName(const std::wstring &aPath)
{
    std::wstring sPath(aPath);
    while (sPath.length() > 3 && (sPath.back() == L'\\' || sPath.back() == L'/'))
        sPath.pop_back();

    const size_t sSlash = sPath.find_last_of(L"\\/");
    return (sSlash == std::wstring::npos) ? sPath : sPath.substr(sSlash + 1);
}

bool isUnsupportedAttributes(DWORD aAttributes)
{
    const DWORD sUnsupported = FILE_ATTRIBUTE_REPARSE_POINT |
                               FILE_ATTRIBUTE_SPARSE_FILE |
                               FILE_ATTRIBUTE_ENCRYPTED |
                               FILE_ATTRIBUTE_OFFLINE |
                               FILE_ATTRIBUTE_READONLY;
    if ((aAttributes & sUnsupported) != 0)
        return true;

    // Cloud Files placeholders can trigger an unbounded network recall.  The
    // Windows Shell owns hydration/conflict UI, so keep those objects out of
    // the direct CopyFile2 path.
#ifdef FILE_ATTRIBUTE_RECALL_ON_OPEN
    if ((aAttributes & FILE_ATTRIBUTE_RECALL_ON_OPEN) != 0)
        return true;
#endif
#ifdef FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS
    if ((aAttributes & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS) != 0)
        return true;
#endif
    return false;
}

bool isLocalFileSystemPath(const std::wstring &aPath)
{
    if (aPath.empty() || ::PathIsUNCW(aPath.c_str()))
        return false;

    wchar_t sVolumePath[MAX_PATH + 1] = {0};
    if (!::GetVolumePathNameW(aPath.c_str(), sVolumePath, MAX_PATH))
        return false;

    const UINT sDriveType = ::GetDriveTypeW(sVolumePath);
    return sDriveType == DRIVE_FIXED || sDriveType == DRIVE_REMOVABLE ||
           sDriveType == DRIVE_RAMDISK;
}

bool getVolumeName(const std::wstring &aPath, std::wstring &aVolumeName)
{
    wchar_t sVolumePath[MAX_PATH + 1] = {0};
    wchar_t sVolumeName[MAX_PATH + 1] = {0};
    if (!::GetVolumePathNameW(aPath.c_str(), sVolumePath, MAX_PATH))
        return false;
    if (!::GetVolumeNameForVolumeMountPointW(sVolumePath, sVolumeName, MAX_PATH))
        return false;
    aVolumeName.assign(sVolumeName);
    return true;
}

StorageKind queryStorageKind(const std::wstring &aPath,
                             std::wstring *aVolumeName)
{
    wchar_t sVolumePath[MAX_PATH + 1] = {0};
    wchar_t sVolumeName[MAX_PATH + 1] = {0};
    if (!::GetVolumePathNameW(aPath.c_str(), sVolumePath, MAX_PATH) ||
        !::GetVolumeNameForVolumeMountPointW(sVolumePath, sVolumeName, MAX_PATH))
        return StorageUnknown;

    std::wstring sName(sVolumeName);
    if (aVolumeName != NULL)
        *aVolumeName = normalizePath(sName);
    while (!sName.empty() && sName.back() == L'\\')
        sName.pop_back();

    HANDLE sVolume = ::CreateFileW(sName.c_str(), 0,
                                    FILE_SHARE_READ | FILE_SHARE_WRITE |
                                    FILE_SHARE_DELETE,
                                    NULL, OPEN_EXISTING, 0, NULL);
    if (sVolume == INVALID_HANDLE_VALUE)
        return StorageUnknown;

    STORAGE_PROPERTY_QUERY sQuery = {0};
    sQuery.PropertyId = StorageDeviceSeekPenaltyProperty;
    sQuery.QueryType = PropertyStandardQuery;
    DEVICE_SEEK_PENALTY_DESCRIPTOR sPenalty = {0};
    DWORD sReturned = 0;
    const BOOL sResult = ::DeviceIoControl(sVolume,
                                           IOCTL_STORAGE_QUERY_PROPERTY,
                                           &sQuery, sizeof(sQuery),
                                           &sPenalty, sizeof(sPenalty),
                                           &sReturned, NULL);
    ::CloseHandle(sVolume);
    if (!sResult || sReturned < sizeof(sPenalty))
        return StorageUnknown;
    return sPenalty.IncursSeekPenalty ? StorageRotational : StorageSolidState;
}

bool insertUniquePath(std::set<std::wstring> &aPaths, const std::wstring &aPath)
{
    return aPaths.insert(normalizePath(aPath)).second;
}

bool isPathInside(const std::wstring &aPath, const std::wstring &aParent)
{
    const std::wstring sPath = normalizePath(aPath);
    std::wstring sParent = normalizePath(aParent);
    if (sParent.empty() || sPath.length() < sParent.length())
        return false;
    if (sPath.compare(0, sParent.length(), sParent) != 0)
        return false;
    return sPath.length() == sParent.length() ||
           sPath[sParent.length()] == L'\\';
}

HANDLE openTargetForIdentity(const std::wstring &aPath,
                             bool aDirectory,
                             DWORD aDesiredAccess)
{
    DWORD sFlags = FILE_FLAG_OPEN_REPARSE_POINT;
    if (aDirectory)
        sFlags |= FILE_FLAG_BACKUP_SEMANTICS;
    return ::CreateFileW(aPath.c_str(), aDesiredAccess,
                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                         NULL, OPEN_EXISTING, sFlags, NULL);
}

bool readTargetEvidence(HANDLE aHandle,
                        const std::wstring &aPath,
                        CreatedTargetEvidence &aEvidence)
{
    BY_HANDLE_FILE_INFORMATION sInformation = {0};
    if (aHandle == INVALID_HANDLE_VALUE ||
        !::GetFileInformationByHandle(aHandle, &sInformation))
        return false;

    // A zero file index is not a usable ownership proof.  In that case the
    // target is intentionally left in place and shell fallback is forbidden.
    if (sInformation.nFileIndexHigh == 0 && sInformation.nFileIndexLow == 0)
        return false;

    aEvidence.path = aPath;
    aEvidence.volumeSerialNumber = sInformation.dwVolumeSerialNumber;
    aEvidence.fileIndexHigh = sInformation.nFileIndexHigh;
    aEvidence.fileIndexLow = sInformation.nFileIndexLow;
    aEvidence.directory =
        (sInformation.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    aEvidence.valid = true;
    return true;
}

bool captureTargetEvidence(const std::wstring &aPath,
                           bool aDirectory,
                           CreatedTargetEvidence &aEvidence)
{
    HANDLE sHandle = openTargetForIdentity(aPath, aDirectory,
                                           FILE_READ_ATTRIBUTES);
    if (sHandle == INVALID_HANDLE_VALUE)
        return false;
    CreatedTargetEvidence sEvidence;
    const bool sSucceeded = readTargetEvidence(sHandle, aPath, sEvidence) &&
                            sEvidence.directory == aDirectory;
    ::CloseHandle(sHandle);
    if (sSucceeded)
        aEvidence = sEvidence;
    return sSucceeded;
}

bool handleMatchesEvidence(HANDLE aHandle,
                           const CreatedTargetEvidence &aEvidence)
{
    CreatedTargetEvidence sCurrent;
    return aEvidence.valid &&
           readTargetEvidence(aHandle, aEvidence.path, sCurrent) &&
           sCurrent.directory == aEvidence.directory &&
           sCurrent.volumeSerialNumber == aEvidence.volumeSerialNumber &&
           sCurrent.fileIndexHigh == aEvidence.fileIndexHigh &&
           sCurrent.fileIndexLow == aEvidence.fileIndexLow;
}

bool deleteOwnedTarget(const CreatedTargetEvidence &aEvidence)
{
    if (!aEvidence.valid)
        return false;

    HANDLE sHandle = openTargetForIdentity(aEvidence.path,
                                           aEvidence.directory,
                                           DELETE | FILE_READ_ATTRIBUTES);
    if (sHandle == INVALID_HANDLE_VALUE)
    {
        const DWORD sError = ::GetLastError();
        return sError == ERROR_FILE_NOT_FOUND || sError == ERROR_PATH_NOT_FOUND;
    }
    if (!handleMatchesEvidence(sHandle, aEvidence))
    {
        ::CloseHandle(sHandle);
        return false;
    }

    FILE_DISPOSITION_INFO sDisposition = {0};
    sDisposition.DeleteFile = TRUE;
    const bool sSucceeded =
        ::SetFileInformationByHandle(sHandle, FileDispositionInfo,
                                     &sDisposition, sizeof(sDisposition)) != FALSE;
    ::CloseHandle(sHandle);
    return sSucceeded;
}

bool isWindowsProtectedPath(const std::wstring &aPath)
{
    wchar_t sWindows[MAX_PATH + 1] = {0};
    if (::GetWindowsDirectoryW(sWindows, MAX_PATH) > 0 &&
        isPathInside(aPath, sWindows))
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

bool canOpenForDelete(const std::wstring &aPath, bool aDirectory)
{
    HANDLE sHandle = ::CreateFileW(aPath.c_str(), DELETE | FILE_READ_ATTRIBUTES,
                                    FILE_SHARE_READ | FILE_SHARE_WRITE |
                                    FILE_SHARE_DELETE,
                                    NULL, OPEN_EXISTING,
                                    aDirectory ? FILE_FLAG_BACKUP_SEMANTICS : 0,
                                    NULL);
    if (sHandle == INVALID_HANDLE_VALUE)
        return false;
    ::CloseHandle(sHandle);
    return true;
}

bool enumerateDeletePath(const std::wstring &aPath,
                         DeletePlan &aPlan,
                         DWORD &aError)
{
    if (!isLocalFileSystemPath(aPath) || ::PathIsRootW(aPath.c_str()) ||
        isWindowsProtectedPath(aPath))
        return false;

    WIN32_FIND_DATAW sData = {0};
    HANDLE sFind = ::FindFirstFileW(aPath.c_str(), &sData);
    if (sFind == INVALID_HANDLE_VALUE)
        return false;
    ::FindClose(sFind);
    if (isUnsupportedAttributes(sData.dwFileAttributes) ||
        (sData.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM) != 0)
        return false;

    const bool sDirectory =
        (sData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    if (!canOpenForDelete(aPath, sDirectory))
        return false;

    if (!sDirectory)
    {
        aPlan.files.push_back(aPath);
        return true;
    }

    const std::wstring sPattern = joinPath(aPath, L"*");
    sFind = ::FindFirstFileExW(sPattern.c_str(), FindExInfoBasic,
                               &sData, FindExSearchNameMatch,
                               NULL, FIND_FIRST_EX_LARGE_FETCH);
    if (sFind == INVALID_HANDLE_VALUE)
    {
        aError = ::GetLastError();
        if (aError != ERROR_FILE_NOT_FOUND)
            return false;
        aError = ERROR_SUCCESS;
    }
    else
    {
        bool sOk = true;
        do
        {
            if (wcscmp(sData.cFileName, L".") == 0 ||
                wcscmp(sData.cFileName, L"..") == 0)
                continue;
            if (!enumerateDeletePath(joinPath(aPath, sData.cFileName),
                                     aPlan, aError))
            {
                sOk = false;
                break;
            }
        } while (::FindNextFileW(sFind, &sData));
        if (sOk && ::GetLastError() != ERROR_NO_MORE_FILES)
        {
            aError = ::GetLastError();
            sOk = false;
        }
        ::FindClose(sFind);
        if (!sOk)
            return false;
    }
    aPlan.directories.push_back(aPath);
    return true;
}

bool buildDeletePlan(const SHFILEOPSTRUCT *aOperation,
                     DeletePlan &aPlan,
                     DWORD &aError)
{
    if (aOperation == NULL || aOperation->wFunc != FO_DELETE ||
        aOperation->pFrom == NULL ||
        (aOperation->fFlags & FOF_ALLOWUNDO) != 0)
        return false;

    const wchar_t *sSource = aOperation->pFrom;
    while (*sSource != L'\0')
    {
        if (wcschr(sSource, L'*') != NULL || wcschr(sSource, L'?') != NULL ||
            !enumerateDeletePath(sSource, aPlan, aError))
            return false;
        sSource += wcslen(sSource) + 1;
    }
    return !aPlan.files.empty() || !aPlan.directories.empty();
}

void deleteWorker(SharedDeleteState *aShared)
{
    for (;;)
    {
        if (aShared->cancelRequested.load() ||
            aShared->firstFailure.load() != ERROR_SUCCESS)
            break;
        const size_t sIndex = aShared->nextFile.fetch_add(1);
        if (sIndex >= aShared->files->size())
            break;
        if (!::DeleteFileW((*aShared->files)[sIndex].c_str()))
        {
            DWORD sExpected = ERROR_SUCCESS;
            aShared->firstFailure.compare_exchange_strong(sExpected,
                                                           ::GetLastError());
            aShared->cancelRequested.store(true);
            break;
        }
        aShared->completed.fetch_add(1);
    }
    aShared->activeWorkers.fetch_sub(1);
}

bool addFile(CopyPlan &aPlan,
             const std::wstring &aSource,
             const std::wstring &aTarget,
             const WIN32_FIND_DATAW &aFindData,
             DWORD &aError)
{
    if (isUnsupportedAttributes(aFindData.dwFileAttributes))
        return false;

    if (::GetFileAttributesW(aTarget.c_str()) != INVALID_FILE_ATTRIBUTES)
        return false;
    if (::GetLastError() != ERROR_FILE_NOT_FOUND && ::GetLastError() != ERROR_PATH_NOT_FOUND)
    {
        aError = ::GetLastError();
        return false;
    }

    if (!insertUniquePath(aPlan.sourcePaths, aSource) ||
        !insertUniquePath(aPlan.targetPaths, aTarget))
        return false;

    FileJob sJob;
    sJob.source = aSource;
    sJob.target = aTarget;
    sJob.size = (static_cast<ULONGLONG>(aFindData.nFileSizeHigh) << 32) |
                aFindData.nFileSizeLow;
    sJob.lastWriteTime = aFindData.ftLastWriteTime;
    aPlan.totalBytes += sJob.size;
    aPlan.files.push_back(sJob);
    return true;
}

bool enumerateDirectory(CopyPlan &aPlan,
                        const std::wstring &aSource,
                        const std::wstring &aTarget,
                        const WIN32_FIND_DATAW &aDirectoryData,
                        DWORD &aError)
{
    if (isUnsupportedAttributes(aDirectoryData.dwFileAttributes))
        return false;

    if (::GetFileAttributesW(aTarget.c_str()) != INVALID_FILE_ATTRIBUTES)
        return false;
    if (::GetLastError() != ERROR_FILE_NOT_FOUND && ::GetLastError() != ERROR_PATH_NOT_FOUND)
    {
        aError = ::GetLastError();
        return false;
    }

    if (!insertUniquePath(aPlan.sourcePaths, aSource) ||
        !insertUniquePath(aPlan.targetPaths, aTarget))
        return false;

    DirectoryJob sDirectory;
    sDirectory.source = aSource;
    sDirectory.target = aTarget;
    sDirectory.attributes = aDirectoryData.dwFileAttributes;
    sDirectory.creationTime = aDirectoryData.ftCreationTime;
    sDirectory.accessTime = aDirectoryData.ftLastAccessTime;
    sDirectory.lastWriteTime = aDirectoryData.ftLastWriteTime;
    aPlan.directories.push_back(sDirectory);

    const std::wstring sPattern = joinPath(aSource, L"*");
    WIN32_FIND_DATAW sFindData = {0};
    HANDLE sFind = ::FindFirstFileExW(sPattern.c_str(), FindExInfoBasic,
                                      &sFindData, FindExSearchNameMatch,
                                      NULL, FIND_FIRST_EX_LARGE_FETCH);
    if (sFind == INVALID_HANDLE_VALUE)
    {
        aError = ::GetLastError();
        if (aError == ERROR_FILE_NOT_FOUND)
        {
            aError = ERROR_SUCCESS;
            return true;
        }
        return false;
    }

    bool sSucceeded = true;
    do
    {
        if (wcscmp(sFindData.cFileName, L".") == 0 ||
            wcscmp(sFindData.cFileName, L"..") == 0)
            continue;

        const std::wstring sChildSource = joinPath(aSource, sFindData.cFileName);
        const std::wstring sChildTarget = joinPath(aTarget, sFindData.cFileName);
        if ((sFindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        {
            if (!enumerateDirectory(aPlan, sChildSource, sChildTarget, sFindData, aError))
            {
                sSucceeded = false;
                break;
            }
        }
        else if (!addFile(aPlan, sChildSource, sChildTarget, sFindData, aError))
        {
            sSucceeded = false;
            break;
        }
    } while (::FindNextFileW(sFind, &sFindData));

    if (sSucceeded)
    {
        const DWORD sFindError = ::GetLastError();
        if (sFindError != ERROR_NO_MORE_FILES)
        {
            aError = sFindError;
            sSucceeded = false;
        }
    }
    ::FindClose(sFind);
    return sSucceeded;
}

bool buildPlan(const SHFILEOPSTRUCT *aOperation, CopyPlan &aPlan, DWORD &aError)
{
    if (aOperation == NULL || aOperation->pFrom == NULL || aOperation->pTo == NULL)
        return false;
    if (aOperation->wFunc != FO_COPY && aOperation->wFunc != FO_MOVE)
        return false;
    if ((aOperation->fFlags & (FOF_RENAMEONCOLLISION | FOF_MULTIDESTFILES)) != 0)
        return false;

    const std::wstring sTargetDirectory(aOperation->pTo);
    const DWORD sTargetAttributes = ::GetFileAttributesW(sTargetDirectory.c_str());
    if (sTargetAttributes == INVALID_FILE_ATTRIBUTES ||
        (sTargetAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0 ||
        !isLocalFileSystemPath(sTargetDirectory))
        return false;

    std::wstring sTargetVolume;
    if (!getVolumeName(sTargetDirectory, sTargetVolume))
        return false;

    bool sAllCrossVolume = true;
    const wchar_t *sSource = aOperation->pFrom;
    while (*sSource != L'\0')
    {
        const std::wstring sSourcePath(sSource);
        if (sSourcePath.find_first_of(L"*?") != std::wstring::npos ||
            !isLocalFileSystemPath(sSourcePath))
            return false;

        std::wstring sSourceVolume;
        if (!getVolumeName(sSourcePath, sSourceVolume))
            return false;
        if (_wcsicmp(sSourceVolume.c_str(), sTargetVolume.c_str()) == 0)
            sAllCrossVolume = false;

        WIN32_FIND_DATAW sFindData = {0};
        HANDLE sFind = ::FindFirstFileW(sSourcePath.c_str(), &sFindData);
        if (sFind == INVALID_HANDLE_VALUE)
            return false;
        ::FindClose(sFind);

        const std::wstring sLeaf = getLeafName(sSourcePath);
        if (sLeaf.empty())
            return false;
        const std::wstring sTargetPath = joinPath(sTargetDirectory, sLeaf);

        bool sAdded = false;
        if ((sFindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            sAdded = enumerateDirectory(aPlan, sSourcePath, sTargetPath, sFindData, aError);
        else
            sAdded = addFile(aPlan, sSourcePath, sTargetPath, sFindData, aError);
        if (!sAdded)
            return false;

        sSource += wcslen(sSource) + 1;
    }

    if (aPlan.files.empty() && aPlan.directories.empty())
        return false;

    if (aOperation->wFunc == FO_MOVE)
    {
        // Same-volume moves are metadata renames and the Shell already executes
        // them optimally.  The adaptive copy/delete path is only useful when
        // every selected source is on another volume.
        if (!sAllCrossVolume)
            return false;
        aPlan.moveAcrossVolumes = true;
    }

    return true;
}

bool createDirectories(const CopyPlan &aPlan,
                       std::vector<CreatedTargetEvidence> &aCreatedTargets,
                       DWORD &aError)
{
    for (size_t i = 0; i < aPlan.directories.size(); ++i)
    {
        const DirectoryJob &sDirectory = aPlan.directories[i];
        if (!::CreateDirectoryW(sDirectory.target.c_str(), NULL))
        {
            aError = ::GetLastError();
            return false;
        }
        CreatedTargetEvidence sEvidence;
        if (!captureTargetEvidence(sDirectory.target, true, sEvidence))
        {
            aError = ::GetLastError();
            if (aError == ERROR_SUCCESS)
                aError = ERROR_INVALID_DATA;
            return false;
        }
        aCreatedTargets.push_back(sEvidence);
    }
    return true;
}

COPYFILE2_MESSAGE_ACTION CALLBACK copyProgressRoutine(
    const COPYFILE2_MESSAGE *aMessage,
    PVOID aContext)
{
    FileCopyContext *sContext = static_cast<FileCopyContext *>(aContext);
    if (sContext == NULL || sContext->shared == NULL)
        return COPYFILE2_PROGRESS_CANCEL;
    if (sContext->shared->cancelRequested.load())
        return COPYFILE2_PROGRESS_CANCEL;

    ULONGLONG sTransferred = sContext->lastTransferred;
    if (aMessage != NULL)
    {
        HANDLE sDestination = INVALID_HANDLE_VALUE;
        if (aMessage->Type == COPYFILE2_CALLBACK_CHUNK_STARTED)
            sDestination = aMessage->Info.ChunkStarted.hDestinationFile;
        else if (aMessage->Type == COPYFILE2_CALLBACK_CHUNK_FINISHED)
            sDestination = aMessage->Info.ChunkFinished.hDestinationFile;
        else if (aMessage->Type == COPYFILE2_CALLBACK_STREAM_STARTED)
            sDestination = aMessage->Info.StreamStarted.hDestinationFile;
        else if (aMessage->Type == COPYFILE2_CALLBACK_STREAM_FINISHED)
            sDestination = aMessage->Info.StreamFinished.hDestinationFile;

        // This is CopyFile2's live destination handle, not a later path
        // reopen.  Its volume/file id is therefore direct proof that this
        // exact object was created by the current copy attempt.
        if (sDestination != NULL && sDestination != INVALID_HANDLE_VALUE &&
            sContext->createdEvidence != NULL &&
            !sContext->createdEvidence->valid &&
            sContext->targetPath != NULL)
        {
            CreatedTargetEvidence sEvidence;
            if (readTargetEvidence(sDestination, *sContext->targetPath,
                                   sEvidence) && !sEvidence.directory)
                *sContext->createdEvidence = sEvidence;
        }

        if (aMessage->Type == COPYFILE2_CALLBACK_CHUNK_FINISHED)
            sTransferred = aMessage->Info.ChunkFinished.uliTotalBytesTransferred.QuadPart;
        else if (aMessage->Type == COPYFILE2_CALLBACK_STREAM_FINISHED)
            sTransferred = aMessage->Info.StreamFinished.uliTotalBytesTransferred.QuadPart;
    }

    if (sTransferred > sContext->lastTransferred)
    {
        sContext->shared->transferredBytes.fetch_add(
            sTransferred - sContext->lastTransferred);
        sContext->lastTransferred = sTransferred;
    }
    return COPYFILE2_PROGRESS_CONTINUE;
}

void copyWorker(SharedCopyState *aShared)
{
    for (;;)
    {
        if (aShared->cancelRequested.load() || FAILED(aShared->firstFailure.load()))
            break;

        const size_t sIndex = aShared->nextJob.fetch_add(1);
        if (sIndex >= aShared->jobs->size())
            break;

        aShared->currentJob.store(sIndex);
        const FileJob &sJob = (*aShared->jobs)[sIndex];

        FileCopyContext sContext;
        sContext.shared = aShared;
        if (aShared->createdFiles != NULL &&
            sIndex < aShared->createdFiles->size())
        {
            sContext.createdEvidence = &(*aShared->createdFiles)[sIndex];
            sContext.targetPath = &sJob.target;
        }
        COPYFILE2_EXTENDED_PARAMETERS sParameters = {0};
        sParameters.dwSize = sizeof(sParameters);
        sParameters.dwCopyFlags = COPY_FILE_FAIL_IF_EXISTS;
        sParameters.pProgressRoutine = copyProgressRoutine;
        sParameters.pvCallbackContext = &sContext;

        const HRESULT sResult = ::CopyFile2(sJob.source.c_str(),
                                            sJob.target.c_str(),
                                            &sParameters);
        if (SUCCEEDED(sResult))
        {
            if (aShared->createdFiles != NULL &&
                sIndex < aShared->createdFiles->size())
            {
                if (!(*aShared->createdFiles)[sIndex].valid)
                {
                    const HRESULT sEvidenceFailure =
                        HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                    LONG sExpected = S_OK;
                    if (aShared->firstFailure.compare_exchange_strong(
                            sExpected, sEvidenceFailure))
                        aShared->firstFailureJob.store(sIndex);
                    // The copy exists but its object identity is not known.
                    // Stop before another failure could trigger an unsafe
                    // path-only rollback or shell retry.
                    aShared->cancelRequested.store(true);
                    continue;
                }
            }
            if (sContext.lastTransferred < sJob.size)
                aShared->transferredBytes.fetch_add(sJob.size - sContext.lastTransferred);
            aShared->finishedJobs.fetch_add(1);
        }
        else if (sResult == HRESULT_FROM_WIN32(ERROR_REQUEST_ABORTED) ||
                 sResult == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            aShared->cancelRequested.store(true);
        }
        else
        {
            LONG sExpected = S_OK;
            if (aShared->firstFailure.compare_exchange_strong(sExpected, sResult))
                aShared->firstFailureJob.store(sIndex);
            aShared->cancelRequested.store(true);
        }
    }
    aShared->activeWorkers.fetch_sub(1);
}

unsigned chooseWorkerCount(const CopyPlan &aPlan)
{
    if (aPlan.files.empty())
        return 1;

    ULONGLONG sLargest = 0;
    for (size_t i = 0; i < aPlan.files.size(); ++i)
        sLargest = (std::max)(sLargest, aPlan.files[i].size);

    const ULONGLONG sAverage = aPlan.totalBytes / aPlan.files.size();
    unsigned sWorkers = 1;
    if (aPlan.files.size() >= kManyFileThreshold &&
        sLargest <= 64ULL * 1024ULL * 1024ULL &&
        sAverage <= kSmallFileLimit)
        sWorkers = kMaximumWorkers;
    else if (aPlan.files.size() >= kMediumFileThreshold &&
             sAverage <= 2ULL * kSmallFileLimit)
        sWorkers = 2;

    // Use the documented seek-penalty property instead of guessing from a
    // drive letter.  Unknown and single rotational volumes are deliberately
    // capped: this keeps antivirus/minifilter and cloud-provider contention
    // from turning parallel copy into seek amplification.
    StorageKind sAggregateKind = StorageSolidState;
    bool sUnknown = false;
    bool sRotational = false;
    bool sSameVolume = true;
    std::wstring sTargetVolume;
    if (!aPlan.files.empty())
    {
        const StorageKind sTargetKind =
            queryStorageKind(aPlan.files[0].target, &sTargetVolume);
        sUnknown = (sTargetKind == StorageUnknown);
        sRotational = (sTargetKind == StorageRotational);

        std::set<std::wstring> sQueriedVolumes;
        for (size_t i = 0; i < aPlan.files.size(); ++i)
        {
            std::wstring sSourceVolume;
            if (!getVolumeName(aPlan.files[i].source, sSourceVolume))
            {
                sUnknown = true;
                continue;
            }
            sSourceVolume = normalizePath(sSourceVolume);
            if (!sQueriedVolumes.insert(sSourceVolume).second)
                continue;
            const StorageKind sSourceKind =
                queryStorageKind(aPlan.files[i].source, NULL);
            sUnknown = sUnknown || sSourceKind == StorageUnknown;
            sRotational = sRotational || sSourceKind == StorageRotational;
            sSameVolume = sSameVolume && !sTargetVolume.empty() &&
                          sSourceVolume == sTargetVolume;
        }
    }
    sAggregateKind = sUnknown ? StorageUnknown :
                     (sRotational ? StorageRotational : StorageSolidState);

    if (sAggregateKind == StorageUnknown ||
        (sAggregateKind == StorageRotational && sSameVolume))
        sWorkers = (std::min)(sWorkers, 2U);

    const unsigned sHardware = std::thread::hardware_concurrency();
    if (sHardware > 0)
        sWorkers = (std::min)(sWorkers, sHardware);
    return (std::max)(1U, sWorkers);
}

bool rollbackTargets(const CopyPlan &aPlan,
                      const std::vector<CreatedTargetEvidence> &aCreatedTargets)
{
    // Path absence during preflight is not ownership: another process can
    // create or replace that path before rollback.  Delete only objects whose
    // volume/file id was captured from this operation, through the same
    // verified handle.  Files are removed before child-before-parent dirs.
    bool sSucceeded = true;
    std::set<std::wstring> sHandled;
    for (size_t i = 0; i < aCreatedTargets.size(); ++i)
    {
        const CreatedTargetEvidence &sEvidence = aCreatedTargets[i];
        if (!sEvidence.valid || sEvidence.directory)
            continue;
        const std::wstring sKey = normalizePath(sEvidence.path);
        if (!sHandled.insert(sKey).second)
            continue;
        if (!deleteOwnedTarget(sEvidence))
            sSucceeded = false;
    }
    for (std::vector<CreatedTargetEvidence>::const_reverse_iterator sIt =
             aCreatedTargets.rbegin();
         sIt != aCreatedTargets.rend(); ++sIt)
    {
        if (!sIt->valid || !sIt->directory)
            continue;
        const std::wstring sKey = normalizePath(sIt->path);
        if (!sHandled.insert(sKey).second)
            continue;
        if (!deleteOwnedTarget(*sIt))
            sSucceeded = false;
    }

    // A fallback may enumerate the same original selection again.  It is
    // safe only when every target owned by the adaptive attempt is absent.
    for (size_t i = 0; i < aPlan.files.size(); ++i)
    {
        if (::GetFileAttributesW(aPlan.files[i].target.c_str()) != INVALID_FILE_ATTRIBUTES)
            sSucceeded = false;
        else
        {
            const DWORD sError = ::GetLastError();
            if (sError != ERROR_FILE_NOT_FOUND && sError != ERROR_PATH_NOT_FOUND)
                sSucceeded = false;
        }
    }
    for (size_t i = 0; i < aPlan.directories.size(); ++i)
    {
        if (::GetFileAttributesW(aPlan.directories[i].target.c_str()) !=
            INVALID_FILE_ATTRIBUTES)
            sSucceeded = false;
        else
        {
            const DWORD sError = ::GetLastError();
            if (sError != ERROR_FILE_NOT_FOUND && sError != ERROR_PATH_NOT_FOUND)
                sSucceeded = false;
        }
    }

    return sSucceeded;
}

bool rollbackCopyTargets(
    const CopyPlan &aPlan,
    const std::vector<CreatedTargetEvidence> &aCreatedDirectories,
    const std::vector<CreatedTargetEvidence> &aCreatedFiles)
{
    std::vector<CreatedTargetEvidence> sCreatedTargets(aCreatedDirectories);
    sCreatedTargets.reserve(aCreatedDirectories.size() + aCreatedFiles.size());
    for (size_t i = 0; i < aCreatedFiles.size(); ++i)
    {
        if (aCreatedFiles[i].valid)
            sCreatedTargets.push_back(aCreatedFiles[i]);
    }
    return rollbackTargets(aPlan, sCreatedTargets);
}

bool canRetryWithModernShell(DWORD aError)
{
    switch (aError)
    {
    case ERROR_FILE_NOT_FOUND:
    case ERROR_PATH_NOT_FOUND:
    case ERROR_FILE_INVALID:
    case ERROR_SHARING_VIOLATION:
    case ERROR_LOCK_VIOLATION:
    case ERROR_NOT_SUPPORTED:
    case ERROR_INVALID_FUNCTION:
    case ERROR_NOT_ENOUGH_MEMORY:
    case ERROR_ACCESS_DENIED:
    case ERROR_PRIVILEGE_NOT_HELD:
    case ERROR_CANNOT_MAKE:
    case ERROR_WRITE_PROTECT:
        return true;
    }
    return false;
}

void restoreDirectoryMetadata(const CopyPlan &aPlan)
{
    for (std::vector<DirectoryJob>::const_reverse_iterator sIt = aPlan.directories.rbegin();
         sIt != aPlan.directories.rend(); ++sIt)
    {
        HANDLE sDirectory = ::CreateFileW(sIt->target.c_str(), FILE_WRITE_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          NULL, OPEN_EXISTING,
                                          FILE_FLAG_BACKUP_SEMANTICS, NULL);
        if (sDirectory != INVALID_HANDLE_VALUE)
        {
            ::SetFileTime(sDirectory, &sIt->creationTime, &sIt->accessTime,
                          &sIt->lastWriteTime);
            ::CloseHandle(sDirectory);
        }
        const DWORD sAttributes = sIt->attributes & ~FILE_ATTRIBUTE_DIRECTORY;
        ::SetFileAttributesW(sIt->target.c_str(), sAttributes == 0 ? FILE_ATTRIBUTE_NORMAL : sAttributes);
    }
}

bool sourceSnapshotUnchanged(const CopyPlan &aPlan, DWORD &aError)
{
    for (size_t i = 0; i < aPlan.directories.size(); ++i)
    {
        const DWORD sTargetAttributes =
            ::GetFileAttributesW(aPlan.directories[i].target.c_str());
        if (sTargetAttributes == INVALID_FILE_ATTRIBUTES ||
            (sTargetAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            aError = (sTargetAttributes == INVALID_FILE_ATTRIBUTES) ?
                     ::GetLastError() : ERROR_DIRECTORY;
            return false;
        }
    }

    for (size_t i = 0; i < aPlan.files.size(); ++i)
    {
        WIN32_FILE_ATTRIBUTE_DATA sData = {0};
        if (!::GetFileAttributesExW(aPlan.files[i].source.c_str(), GetFileExInfoStandard, &sData))
        {
            aError = ::GetLastError();
            return false;
        }
        const ULONGLONG sSize = (static_cast<ULONGLONG>(sData.nFileSizeHigh) << 32) |
                                sData.nFileSizeLow;
        if (sSize != aPlan.files[i].size ||
            ::CompareFileTime(&sData.ftLastWriteTime, &aPlan.files[i].lastWriteTime) != 0)
        {
            aError = ERROR_FILE_INVALID;
            return false;
        }

        WIN32_FILE_ATTRIBUTE_DATA sTargetData = {0};
        if (!::GetFileAttributesExW(aPlan.files[i].target.c_str(),
                                    GetFileExInfoStandard, &sTargetData))
        {
            aError = ::GetLastError();
            return false;
        }
        const ULONGLONG sTargetSize =
            (static_cast<ULONGLONG>(sTargetData.nFileSizeHigh) << 32) |
            sTargetData.nFileSizeLow;
        if (sTargetSize != aPlan.files[i].size ||
            ::CompareFileTime(&sTargetData.ftLastWriteTime,
                              &aPlan.files[i].lastWriteTime) != 0)
        {
            aError = ERROR_FILE_INVALID;
            return false;
        }
    }

    // Re-enumerate every source directory and reject a move when an entry was
    // added after the preflight snapshot.  Copy operations do not need this.
    for (size_t i = 0; i < aPlan.directories.size(); ++i)
    {
        WIN32_FIND_DATAW sFindData = {0};
        const std::wstring sPattern = joinPath(aPlan.directories[i].source, L"*");
        HANDLE sFind = ::FindFirstFileExW(sPattern.c_str(), FindExInfoBasic,
                                          &sFindData, FindExSearchNameMatch,
                                          NULL, FIND_FIRST_EX_LARGE_FETCH);
        if (sFind == INVALID_HANDLE_VALUE)
        {
            aError = ::GetLastError();
            if (aError == ERROR_FILE_NOT_FOUND)
                continue;
            return false;
        }

        bool sUnchanged = true;
        do
        {
            if (wcscmp(sFindData.cFileName, L".") == 0 ||
                wcscmp(sFindData.cFileName, L"..") == 0)
                continue;
            const std::wstring sChild = joinPath(aPlan.directories[i].source,
                                                 sFindData.cFileName);
            if (aPlan.sourcePaths.find(normalizePath(sChild)) == aPlan.sourcePaths.end())
            {
                sUnchanged = false;
                aError = ERROR_FILE_INVALID;
                break;
            }
        } while (::FindNextFileW(sFind, &sFindData));
        ::FindClose(sFind);
        if (!sUnchanged)
            return false;
    }
    return true;
}

enum RobocopySelection
{
    RobocopyNotRecommended,
    RobocopyDeclined,
    RobocopyAccepted,
    RobocopyCancelled,
};

enum RobocopyResult
{
    RobocopySucceeded,
    RobocopyFailed,
    RobocopyUserCancelled,
    RobocopyTerminationUncertain,
};

std::wstring quoteCommandArgument(const std::wstring &aArgument)
{
    std::wstring sQuoted(L"\"");
    size_t sBackslashes = 0;
    for (size_t i = 0; i < aArgument.length(); ++i)
    {
        const wchar_t sCharacter = aArgument[i];
        if (sCharacter == L'\\')
        {
            ++sBackslashes;
            continue;
        }
        if (sCharacter == L'\"')
        {
            sQuoted.append(sBackslashes * 2 + 1, L'\\');
            sQuoted.push_back(L'\"');
            sBackslashes = 0;
            continue;
        }
        sQuoted.append(sBackslashes, L'\\');
        sBackslashes = 0;
        sQuoted.push_back(sCharacter);
    }
    sQuoted.append(sBackslashes * 2, L'\\');
    sQuoted.push_back(L'\"');
    return sQuoted;
}

bool getRobocopyPath(std::wstring &aPath)
{
    wchar_t sSystemDirectory[MAX_PATH + 1] = {0};
    const UINT sLength = ::GetSystemDirectoryW(sSystemDirectory, MAX_PATH);
    if (sLength == 0 || sLength >= MAX_PATH)
        return false;
    aPath = joinPath(sSystemDirectory, L"robocopy.exe");
    const DWORD sAttributes = ::GetFileAttributesW(aPath.c_str());
    return sAttributes != INVALID_FILE_ATTRIBUTES &&
           (sAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

bool allTopLevelSourcesAreDirectories(const SHFILEOPSTRUCT *aOperation,
                                      size_t &aTopLevelCount)
{
    aTopLevelCount = 0;
    if (aOperation == NULL || aOperation->pFrom == NULL)
        return false;
    const wchar_t *sSource = aOperation->pFrom;
    while (*sSource != L'\0')
    {
        const DWORD sAttributes = ::GetFileAttributesW(sSource);
        if (sAttributes == INVALID_FILE_ATTRIBUTES ||
            (sAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
            return false;
        ++aTopLevelCount;
        sSource += wcslen(sSource) + 1;
    }
    return aTopLevelCount > 0;
}

bool evidenceStillMatches(const CreatedTargetEvidence &aEvidence)
{
    if (!aEvidence.valid)
        return false;
    HANDLE sHandle = openTargetForIdentity(aEvidence.path,
                                           aEvidence.directory,
                                           FILE_READ_ATTRIBUTES);
    if (sHandle == INVALID_HANDLE_VALUE)
        return false;
    const bool sMatches = handleMatchesEvidence(sHandle, aEvidence);
    ::CloseHandle(sHandle);
    return sMatches;
}

bool prepareRobocopyRoots(const SHFILEOPSTRUCT *aOperation,
                          std::vector<CreatedTargetEvidence> &aOwnedRoots,
                          DWORD &aError)
{
    if (aOperation == NULL || aOperation->pFrom == NULL ||
        aOperation->pTo == NULL)
        return false;

    const wchar_t *sSource = aOperation->pFrom;
    while (*sSource != L'\0')
    {
        const std::wstring sTarget =
            joinPath(aOperation->pTo, getLeafName(sSource));
        if (!::CreateDirectoryW(sTarget.c_str(), NULL))
        {
            aError = ::GetLastError();
            return false;
        }
        CreatedTargetEvidence sEvidence;
        if (!captureTargetEvidence(sTarget, true, sEvidence))
        {
            aError = ::GetLastError();
            if (aError == ERROR_SUCCESS)
                aError = ERROR_INVALID_DATA;
            return false;
        }
        aOwnedRoots.push_back(sEvidence);
        sSource += wcslen(sSource) + 1;
    }
    return !aOwnedRoots.empty();
}

void collectRobocopyTargetEvidence(
    const CopyPlan &aPlan,
    const std::vector<CreatedTargetEvidence> &aOwnedRoots,
    std::vector<CreatedTargetEvidence> &aCreatedTargets)
{
    (void)aPlan;
    aCreatedTargets.clear();
    for (size_t i = 0; i < aOwnedRoots.size(); ++i)
    {
        if (evidenceStillMatches(aOwnedRoots[i]))
            aCreatedTargets.push_back(aOwnedRoots[i]);
    }
    // Robocopy does not expose destination handles.  Capturing child IDs by
    // reopening their paths after process exit cannot prove who created them.
    // Therefore rollback may remove only our pre-created, identity-matched
    // roots, and only if they are still empty.  Any partial child output (or
    // external file) makes root deletion fail, preserving the entire boundary
    // and forbidding automatic shell fallback.
}

unsigned chooseRobocopyThreads(const CopyPlan &aPlan)
{
    if (aPlan.files.empty())
        return 2;
    std::wstring sTargetVolume;
    const StorageKind sTargetKind =
        queryStorageKind(aPlan.files[0].target, &sTargetVolume);
    const StorageKind sSourceKind =
        queryStorageKind(aPlan.files[0].source, NULL);
    std::wstring sSourceVolume;
    const bool sSameVolume = getVolumeName(aPlan.files[0].source, sSourceVolume) &&
                             !sTargetVolume.empty() &&
                             normalizePath(sSourceVolume) == sTargetVolume;
    if (sSourceKind == StorageRotational || sTargetKind == StorageRotational)
        return sSameVolume ? 2 : 4;
    if (sSourceKind == StorageSolidState && sTargetKind == StorageSolidState)
        return 8;
    return 4;
}

RobocopySelection selectRobocopy(const SHFILEOPSTRUCT *aOperation,
                                 const CopyPlan &aPlan,
                                 unsigned &aThreads,
                                 bool &aUnbuffered)
{
    if (aOperation == NULL || aOperation->wFunc != FO_COPY ||
        aOperation->pTo == NULL ||
        (aOperation->fFlags & (FOF_RENAMEONCOLLISION |
                              FOF_MULTIDESTFILES |
                              FOF_ALLOWUNDO)) != 0)
        return RobocopyNotRecommended;

    size_t sTopLevelCount = 0;
    if (!allTopLevelSourcesAreDirectories(aOperation, sTopLevelCount))
        return RobocopyNotRecommended;

    const bool sLargeTree =
        aPlan.files.size() >= kRobocopyManyFileThreshold ||
        aPlan.directories.size() >= kRobocopyManyDirectoryThreshold ||
        aPlan.totalBytes >= kRobocopyLargeTreeThreshold;
    if (!sLargeTree)
        return RobocopyNotRecommended;

    std::wstring sRobocopyPath;
    if (!getRobocopyPath(sRobocopyPath))
        return RobocopyNotRecommended;

    aThreads = chooseRobocopyThreads(aPlan);
    ULONGLONG sLargestFile = 0;
    for (size_t i = 0; i < aPlan.files.size(); ++i)
        sLargestFile = (std::max)(sLargestFile, aPlan.files[i].size);
    const ULONGLONG sAverage = aPlan.files.empty() ? 0 :
                               aPlan.totalBytes / aPlan.files.size();
    aUnbuffered = sLargestFile >= 256ULL * 1024ULL * 1024ULL &&
                  sAverage >= 32ULL * 1024ULL * 1024ULL;

    wchar_t sMessage[1024] = {0};
    _snwprintf_s(
        sMessage, _countof(sMessage), _TRUNCATE,
        L"선택한 작업은 대량 폴더 복제로 판정되었습니다.\n\n"
        L"폴더: %Iu개(최상위 %Iu개)\n파일: %Iu개\n전체 크기: %.2f GiB\n"
        L"권장 Robocopy 설정: /MT:%u /R:2 /W:1%s\n\n"
        L"[예] Robocopy 사용\n[아니요] FxFile 적응형 엔진 사용\n"
        L"[취소] 복사 작업 취소\n\n"
        L"Robocopy는 목적지에 같은 이름이 없는 로컬 폴더 복제에만 사용되며, "
        L"완료 후 파일 수·크기·수정 시각을 FxFile이 다시 검증합니다.",
        aPlan.directories.size(), sTopLevelCount, aPlan.files.size(),
        static_cast<double>(aPlan.totalBytes) /
            (1024.0 * 1024.0 * 1024.0),
        aThreads, aUnbuffered ? L" /J" : L"");
    const int sChoice = ::MessageBoxW(
        aOperation->hwnd, sMessage, L"FxFile - 대량 폴더 복사 엔진 선택",
        MB_YESNOCANCEL | MB_DEFBUTTON1 | MB_ICONINFORMATION | MB_TASKMODAL);
    if (sChoice == IDYES)
        return RobocopyAccepted;
    if (sChoice == IDCANCEL)
        return RobocopyCancelled;
    return RobocopyDeclined;
}

RobocopyResult executeRobocopy(const SHFILEOPSTRUCT *aOperation,
                               const CopyPlan &aPlan,
                               unsigned aThreads,
                               bool aUnbuffered,
                               std::vector<CreatedTargetEvidence> &aOwnedRoots,
                               DWORD &aError)
{
    std::wstring sRobocopyPath;
    if (!getRobocopyPath(sRobocopyPath))
    {
        aError = ERROR_FILE_NOT_FOUND;
        return RobocopyFailed;
    }

    if (!prepareRobocopyRoots(aOperation, aOwnedRoots, aError))
        return RobocopyFailed;

    IProgressDialog *sProgress = NULL;
    HRESULT sProgressResult = ::CoCreateInstance(CLSID_ProgressDialog, NULL,
                                                  CLSCTX_INPROC_SERVER,
                                                  IID_PPV_ARGS(&sProgress));
    if (FAILED(sProgressResult) || sProgress == NULL)
    {
        aError = ERROR_NOT_SUPPORTED;
        return RobocopyFailed;
    }
    sProgress->SetTitle(L"FxFile - Robocopy 대량 폴더 복사");
    sProgress->SetLine(1, L"Robocopy로 대량 폴더를 복사하고 있습니다.", FALSE, NULL);
    sProgress->StartProgressDialog(aOperation->hwnd, NULL,
                                   PROGDLG_NORMAL | PROGDLG_AUTOTIME, NULL);

    SECURITY_ATTRIBUTES sSecurity = {0};
    sSecurity.nLength = sizeof(sSecurity);
    sSecurity.bInheritHandle = TRUE;
    HANDLE sNull = ::CreateFileW(L"NUL", GENERIC_READ | GENERIC_WRITE,
                                 FILE_SHARE_READ | FILE_SHARE_WRITE,
                                 &sSecurity, OPEN_EXISTING, 0, NULL);
    HANDLE sJob = ::CreateJobObjectW(NULL, NULL);
    if (sJob != NULL)
    {
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION sLimits = {0};
        sLimits.BasicLimitInformation.LimitFlags =
            JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
        if (!::SetInformationJobObject(sJob, JobObjectExtendedLimitInformation,
                                       &sLimits, sizeof(sLimits)))
        {
            ::CloseHandle(sJob);
            sJob = NULL;
        }
    }

    bool sCancelled = false;
    bool sTerminationUncertain = false;
    bool sSucceeded = sNull != INVALID_HANDLE_VALUE && sJob != NULL;
    if (!sSucceeded)
        aError = ::GetLastError() == ERROR_SUCCESS ?
                 ERROR_NOT_ENOUGH_MEMORY : ::GetLastError();
    size_t sCompletedRoots = 0;
    const wchar_t *sSource = aOperation->pFrom;
    while (sSucceeded && *sSource != L'\0')
    {
        const std::wstring sSourcePath(sSource);
        const std::wstring sTargetPath =
            joinPath(aOperation->pTo, getLeafName(sSourcePath));
        sProgress->SetLine(2, sSourcePath.c_str(), TRUE, NULL);
        sProgress->SetProgress64(sCompletedRoots,
                                 (std::max)(static_cast<size_t>(1),
                                            aPlan.directories.size()));

        wchar_t sOptions[256] = {0};
        _snwprintf_s(sOptions, _countof(sOptions), _TRUNCATE,
                     L" /E /COPY:DAT /DCOPY:DAT /R:2 /W:1 /MT:%u%s"
                     L" /XJ /NP /NFL /NDL /NJH /NJS",
                     aThreads, aUnbuffered ? L" /J" : L"");
        std::wstring sCommand = quoteCommandArgument(sRobocopyPath) + L" " +
                                quoteCommandArgument(sSourcePath) + L" " +
                                quoteCommandArgument(sTargetPath) + sOptions;
        std::vector<wchar_t> sMutable(sCommand.begin(), sCommand.end());
        sMutable.push_back(L'\0');

        STARTUPINFOW sStartup = {0};
        sStartup.cb = sizeof(sStartup);
        sStartup.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
        sStartup.wShowWindow = SW_HIDE;
        sStartup.hStdInput = NULL;
        sStartup.hStdOutput = sNull;
        sStartup.hStdError = sNull;
        PROCESS_INFORMATION sProcess = {0};
        if (!::CreateProcessW(sRobocopyPath.c_str(), &sMutable[0], NULL, NULL,
                              TRUE, CREATE_NO_WINDOW | CREATE_SUSPENDED,
                              NULL, NULL, &sStartup, &sProcess))
        {
            aError = ::GetLastError();
            sSucceeded = false;
            break;
        }
        if (!::AssignProcessToJobObject(sJob, sProcess.hProcess))
        {
            aError = ::GetLastError();
            ::TerminateProcess(sProcess.hProcess, ERROR_CANCELLED);
            const DWORD sTerminationWait =
                ::WaitForSingleObject(sProcess.hProcess, 5000);
            if (sTerminationWait != WAIT_OBJECT_0)
            {
                sTerminationUncertain = true;
                if (sTerminationWait == WAIT_TIMEOUT)
                    aError = ERROR_TIMEOUT;
            }
            ::CloseHandle(sProcess.hThread);
            ::CloseHandle(sProcess.hProcess);
            sSucceeded = false;
            break;
        }
        if (::ResumeThread(sProcess.hThread) == static_cast<DWORD>(-1))
        {
            aError = ::GetLastError();
            ::TerminateJobObject(sJob, aError);
            const DWORD sTerminationWait =
                ::WaitForSingleObject(sProcess.hProcess, 5000);
            if (sTerminationWait != WAIT_OBJECT_0)
            {
                sTerminationUncertain = true;
                if (sTerminationWait == WAIT_TIMEOUT)
                    aError = ERROR_TIMEOUT;
            }
            ::CloseHandle(sProcess.hThread);
            ::CloseHandle(sProcess.hProcess);
            sSucceeded = false;
            break;
        }
        ::CloseHandle(sProcess.hThread);

        bool sProcessTerminated = false;
        for (;;)
        {
            const DWORD sWait = ::WaitForSingleObject(sProcess.hProcess, 100);
            if (sWait == WAIT_OBJECT_0)
            {
                sProcessTerminated = true;
                break;
            }
            if (sWait == WAIT_FAILED)
            {
                aError = ::GetLastError();
                ::TerminateJobObject(sJob, aError);
                const DWORD sTerminationWait =
                    ::WaitForSingleObject(sProcess.hProcess, 5000);
                sProcessTerminated = sTerminationWait == WAIT_OBJECT_0;
                if (!sProcessTerminated)
                {
                    sTerminationUncertain = true;
                    if (sTerminationWait == WAIT_TIMEOUT)
                        aError = ERROR_TIMEOUT;
                }
                sSucceeded = false;
                break;
            }
            if (sProgress->HasUserCancelled())
            {
                sCancelled = true;
                ::TerminateJobObject(sJob, ERROR_CANCELLED);
                const DWORD sTerminationWait =
                    ::WaitForSingleObject(sProcess.hProcess, 5000);
                sProcessTerminated = sTerminationWait == WAIT_OBJECT_0;
                if (!sProcessTerminated)
                {
                    sTerminationUncertain = true;
                    aError = sTerminationWait == WAIT_TIMEOUT ?
                             ERROR_TIMEOUT : ::GetLastError();
                }
                break;
            }
        }

        DWORD sExitCode = 16;
        if (sProcessTerminated &&
            !::GetExitCodeProcess(sProcess.hProcess, &sExitCode))
        {
            aError = ::GetLastError();
            sSucceeded = false;
        }
        ::CloseHandle(sProcess.hProcess);
        if (sTerminationUncertain || sCancelled)
            break;
        if (!sSucceeded || sExitCode >= 8)
        {
            if (aError == ERROR_SUCCESS)
                aError = ERROR_GEN_FAILURE;
            sSucceeded = false;
            break;
        }
        ++sCompletedRoots;
        sSource += wcslen(sSource) + 1;
    }

    if (sJob != NULL)
        ::CloseHandle(sJob);
    if (sNull != INVALID_HANDLE_VALUE)
        ::CloseHandle(sNull);
    sProgress->StopProgressDialog();
    sProgress->Release();

    // A timeout means Robocopy may still own handles and may still mutate the
    // destination.  Never race it with rollback or a second shell operation.
    if (sTerminationUncertain)
        return RobocopyTerminationUncertain;
    if (sCancelled)
        return RobocopyUserCancelled;
    if (!sSucceeded)
        return RobocopyFailed;
    if (!sourceSnapshotUnchanged(aPlan, aError))
        return RobocopyFailed;
    return RobocopySucceeded;
}

bool removeSourcesAfterMove(const CopyPlan &aPlan, DWORD &aError)
{
    if (!sourceSnapshotUnchanged(aPlan, aError))
        return false;

    for (size_t i = 0; i < aPlan.files.size(); ++i)
    {
        if (!::DeleteFileW(aPlan.files[i].source.c_str()))
        {
            aError = ::GetLastError();
            return false;
        }
    }
    for (std::vector<DirectoryJob>::const_reverse_iterator sIt = aPlan.directories.rbegin();
         sIt != aPlan.directories.rend(); ++sIt)
    {
        if (!::RemoveDirectoryW(sIt->source.c_str()))
        {
            aError = ::GetLastError();
            return false;
        }
    }
    return true;
}

void showOperationError(const SHFILEOPSTRUCT *aOperation,
                        DWORD aError,
                        const wchar_t *aFailedPath = NULL,
                        bool aRollbackSucceeded = true)
{
    if (aOperation == NULL || (aOperation->fFlags & FOF_NOERRORUI) != 0)
        return;

    wchar_t *sSystemMessage = NULL;
    ::FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                     FORMAT_MESSAGE_FROM_SYSTEM |
                     FORMAT_MESSAGE_IGNORE_INSERTS,
                     NULL, aError, 0,
                     reinterpret_cast<wchar_t *>(&sSystemMessage), 0, NULL);

    std::wstring sMessage = L"고성능 파일 작업을 완료하지 못했습니다.";
    if (sSystemMessage != NULL)
    {
        sMessage += L"\n\n";
        sMessage += sSystemMessage;
        ::LocalFree(sSystemMessage);
    }
    if (aFailedPath != NULL && aFailedPath[0] != L'\0')
    {
        sMessage += L"\n문제 경로:\n";
        sMessage += aFailedPath;
    }
    if (!aRollbackSucceeded)
    {
        sMessage += L"\n\n일부 대상 파일을 되돌리지 못했습니다. "
                    L"대상 폴더를 확인한 뒤 다시 시도하십시오.";
    }
    ::MessageBoxW(aOperation->hwnd, sMessage.c_str(), L"FxFile 파일 작업",
                  MB_OK | MB_ICONERROR | MB_TASKMODAL);
}

AdaptiveFileOperation::Result executePermanentDelete(
    SHFILEOPSTRUCT *aOperation,
    const DeletePlan &aPlan,
    DWORD &aError)
{
    const size_t sTotal = aPlan.files.size() + aPlan.directories.size();
    if ((aOperation->fFlags & FOF_NOCONFIRMATION) == 0)
    {
        wchar_t sQuestion[256] = {0};
        _snwprintf_s(sQuestion, _countof(sQuestion), _TRUNCATE,
                     L"선택한 파일 및 폴더 %Iu개를 영구 삭제합니다.\n"
                     L"이 작업은 휴지통에서 복구할 수 없습니다. 계속하시겠습니까?",
                     sTotal);
        if (::MessageBoxW(aOperation->hwnd, sQuestion,
                          L"FxFile 영구 삭제",
                          MB_YESNO | MB_DEFBUTTON2 | MB_ICONWARNING |
                          MB_TASKMODAL) != IDYES)
            return AdaptiveFileOperation::ResultCancelled;
    }

    IProgressDialog *sProgress = NULL;
    HRESULT sResult = ::CoCreateInstance(CLSID_ProgressDialog, NULL,
                                         CLSCTX_INPROC_SERVER,
                                         IID_PPV_ARGS(&sProgress));
    if (FAILED(sResult) || sProgress == NULL)
        return AdaptiveFileOperation::ResultNotApplicable;

    sProgress->SetTitle(L"FxFile - 영구 삭제");
    sProgress->SetLine(1, L"검증된 일반 항목을 안전하게 영구 삭제하고 있습니다.",
                       FALSE, NULL);
    sProgress->StartProgressDialog(aOperation->hwnd, NULL,
                                   PROGDLG_NORMAL | PROGDLG_AUTOTIME, NULL);

    unsigned sWorkers = 1;
    if (aPlan.files.size() >= 32)
        sWorkers = 4;
    else if (aPlan.files.size() >= 8)
        sWorkers = 2;
    if (!aPlan.files.empty())
    {
        const StorageKind sKind = queryStorageKind(aPlan.files[0], NULL);
        if (sKind != StorageSolidState)
            sWorkers = (std::min)(sWorkers, 2U);
    }
    const unsigned sHardware = std::thread::hardware_concurrency();
    if (sHardware > 0)
        sWorkers = (std::min)(sWorkers, sHardware);
    sWorkers = (std::max)(1U, sWorkers);

    SharedDeleteState sShared;
    sShared.files = &aPlan.files;
    std::vector<std::thread> sThreads;
    try
    {
        for (unsigned i = 0; i < sWorkers; ++i)
        {
            sShared.activeWorkers.fetch_add(1);
            try
            {
                sThreads.push_back(std::thread(deleteWorker, &sShared));
            }
            catch (...)
            {
                sShared.activeWorkers.fetch_sub(1);
                throw;
            }
        }
    }
    catch (...)
    {
        sShared.cancelRequested.store(true);
        for (size_t i = 0; i < sThreads.size(); ++i)
            sThreads[i].join();
        sProgress->StopProgressDialog();
        sProgress->Release();
        aError = ERROR_NOT_ENOUGH_MEMORY;
        showOperationError(aOperation, aError);
        return AdaptiveFileOperation::ResultFailed;
    }

    while (sShared.activeWorkers.load() != 0)
    {
        if (sProgress->HasUserCancelled())
            sShared.cancelRequested.store(true);
        sProgress->SetProgress64(sShared.completed.load(),
                                 (std::max)(static_cast<size_t>(1), sTotal));
        ::Sleep(25);
    }
    for (size_t i = 0; i < sThreads.size(); ++i)
        sThreads[i].join();

    bool sCancelled = sShared.cancelRequested.load() &&
                      sShared.firstFailure.load() == ERROR_SUCCESS;
    if (!sCancelled && sShared.firstFailure.load() == ERROR_SUCCESS)
    {
        // enumerateDeletePath records directories in child-before-parent
        // order, so removing in forward order is deterministic and safe.
        for (size_t i = 0; i < aPlan.directories.size(); ++i)
        {
            if (sProgress->HasUserCancelled())
            {
                sCancelled = true;
                break;
            }
            if (!::RemoveDirectoryW(aPlan.directories[i].c_str()))
            {
                aError = ::GetLastError();
                break;
            }
            sShared.completed.fetch_add(1);
            sProgress->SetProgress64(sShared.completed.load(), sTotal);
        }
    }
    else if (sShared.firstFailure.load() != ERROR_SUCCESS)
    {
        aError = sShared.firstFailure.load();
    }

    sProgress->StopProgressDialog();
    sProgress->Release();

    if (sCancelled || aError != ERROR_SUCCESS)
    {
        wchar_t sSummary[384] = {0};
        _snwprintf_s(sSummary, _countof(sSummary), _TRUNCATE,
                     L"영구 삭제가 완료되지 않았습니다.\n\n"
                     L"완료: %Iu개\n남음 또는 실패: %Iu개\n"
                     L"삭제된 항목은 복구할 수 없으며, 남은 항목은 그대로 보존했습니다.",
                     sShared.completed.load(),
                     sTotal - (std::min)(sTotal, sShared.completed.load()));
        if ((aOperation->fFlags & FOF_NOERRORUI) == 0)
            ::MessageBoxW(aOperation->hwnd, sSummary, L"FxFile 영구 삭제",
                          MB_OK | (sCancelled ? MB_ICONINFORMATION : MB_ICONERROR) |
                          MB_TASKMODAL);
        return sCancelled ? AdaptiveFileOperation::ResultCancelled :
                            AdaptiveFileOperation::ResultFailed;
    }
    return AdaptiveFileOperation::ResultSucceeded;
}
} // namespace anonymous

AdaptiveFileOperation::Result AdaptiveFileOperation::tryExecute(
    SHFILEOPSTRUCT *aFileOperation,
    DWORD *aError)
{
    DWORD sError = ERROR_SUCCESS;
    if (aError != NULL)
        *aError = ERROR_SUCCESS;

    if (aFileOperation != NULL && aFileOperation->wFunc == FO_DELETE)
    {
        DeletePlan sDeletePlan;
        if (!buildDeletePlan(aFileOperation, sDeletePlan, sError))
            return ResultNotApplicable;
        const Result sDeleteResult =
            executePermanentDelete(aFileOperation, sDeletePlan, sError);
        if (aError != NULL)
            *aError = sError;
        return sDeleteResult;
    }

    CopyPlan sPlan;
    if (!buildPlan(aFileOperation, sPlan, sError))
        return ResultNotApplicable;

    unsigned sRobocopyThreads = 1;
    bool sRobocopyUnbuffered = false;
    const RobocopySelection sRobocopySelection =
        selectRobocopy(aFileOperation, sPlan, sRobocopyThreads,
                       sRobocopyUnbuffered);
    if (sRobocopySelection == RobocopyCancelled)
        return ResultCancelled;
    if (sRobocopySelection == RobocopyAccepted)
    {
        std::vector<CreatedTargetEvidence> sRobocopyOwnedRoots;
        const RobocopyResult sRobocopyResult =
            executeRobocopy(aFileOperation, sPlan, sRobocopyThreads,
                            sRobocopyUnbuffered, sRobocopyOwnedRoots, sError);
        if (sRobocopyResult == RobocopySucceeded)
            return ResultSucceeded;

        if (sRobocopyResult == RobocopyTerminationUncertain)
        {
            // The child may still be writing.  Keep the partial destination
            // intact and do not start a competing shell operation.
            if (sError == ERROR_SUCCESS)
                sError = ERROR_TIMEOUT;
            showOperationError(aFileOperation, sError, NULL, false);
            if (aError != NULL)
                *aError = sError;
            return ResultFailed;
        }

        std::vector<CreatedTargetEvidence> sRobocopyCreatedTargets;
        collectRobocopyTargetEvidence(sPlan, sRobocopyOwnedRoots,
                                      sRobocopyCreatedTargets);
        const bool sRollbackSucceeded =
            rollbackTargets(sPlan, sRobocopyCreatedTargets);
        if (sRobocopyResult == RobocopyUserCancelled)
        {
            if (sRollbackSucceeded)
                return ResultCancelled;
            showOperationError(aFileOperation, ERROR_CANCELLED, NULL, false);
            return ResultFailed;
        }
        if (sRollbackSucceeded)
        {
            if ((aFileOperation->fFlags & FOF_NOERRORUI) == 0)
            {
                ::MessageBoxW(aFileOperation->hwnd,
                    L"Robocopy 작업이 완료되지 않아 생성된 대상을 안전하게 되돌렸습니다.\n\n"
                    L"FxFile의 Windows 셸 엔진으로 다시 시도합니다.",
                    L"FxFile - Robocopy 복구 및 전환",
                    MB_OK | MB_ICONWARNING | MB_TASKMODAL);
            }
            if (aError != NULL)
                *aError = sError;
            return ResultNotApplicable;
        }
        showOperationError(aFileOperation, sError, NULL, false);
        if (aError != NULL)
            *aError = sError;
        return ResultFailed;
    }

    IProgressDialog *sProgress = NULL;
    HRESULT sResult = ::CoCreateInstance(CLSID_ProgressDialog, NULL,
                                         CLSCTX_INPROC_SERVER,
                                         IID_PPV_ARGS(&sProgress));
    if (FAILED(sResult) || sProgress == NULL)
        return ResultNotApplicable;

    sProgress->SetTitle(sPlan.moveAcrossVolumes ? L"FxFile - 파일 이동" :
                                                 L"FxFile - 파일 복사");
    sProgress->SetLine(1, sPlan.moveAcrossVolumes ?
                       L"고성능 엔진으로 복사한 뒤 원본을 안전하게 확인합니다." :
                       L"고성능 적응형 엔진으로 복사하고 있습니다.", FALSE, NULL);
    sProgress->StartProgressDialog(aFileOperation->hwnd, NULL,
                                   PROGDLG_NORMAL | PROGDLG_AUTOTIME,
                                   NULL);

    std::vector<CreatedTargetEvidence> sCreatedDirectories;
    if (!createDirectories(sPlan, sCreatedDirectories, sError))
    {
        sProgress->StopProgressDialog();
        sProgress->Release();
        rollbackTargets(sPlan, sCreatedDirectories);
        showOperationError(aFileOperation, sError);
        if (aError != NULL)
            *aError = sError;
        return ResultFailed;
    }

    SharedCopyState sShared;
    sShared.jobs = &sPlan.files;
    std::vector<CreatedTargetEvidence> sCreatedFiles(sPlan.files.size());
    sShared.createdFiles = &sCreatedFiles;
    const unsigned sWorkerCount = chooseWorkerCount(sPlan);

    std::vector<std::thread> sWorkers;
    sWorkers.reserve(sWorkerCount);
    try
    {
        for (unsigned i = 0; i < sWorkerCount; ++i)
        {
            sShared.activeWorkers.fetch_add(1);
            try
            {
                sWorkers.push_back(std::thread(copyWorker, &sShared));
            }
            catch (...)
            {
                sShared.activeWorkers.fetch_sub(1);
                throw;
            }
        }
    }
    catch (...)
    {
        sShared.cancelRequested.store(true);
        for (size_t i = 0; i < sWorkers.size(); ++i)
            sWorkers[i].join();
        rollbackCopyTargets(sPlan, sCreatedDirectories, sCreatedFiles);
        sProgress->StopProgressDialog();
        sProgress->Release();
        sError = ERROR_NOT_ENOUGH_MEMORY;
        showOperationError(aFileOperation, sError);
        if (aError != NULL)
            *aError = sError;
        return ResultFailed;
    }

    size_t sLastDisplayedJob = static_cast<size_t>(-1);
    while (sShared.activeWorkers.load() != 0)
    {
        if (sProgress->HasUserCancelled())
            sShared.cancelRequested.store(true);

        if (!sPlan.files.empty())
        {
            const size_t sCurrent = (std::min)(sShared.currentJob.load(),
                                               sPlan.files.size() - 1);
            if (sCurrent != sLastDisplayedJob)
            {
                sProgress->SetLine(2, sPlan.files[sCurrent].source.c_str(),
                                   TRUE, NULL);
                sLastDisplayedJob = sCurrent;
            }
        }

        if (sPlan.totalBytes > 0)
            sProgress->SetProgress64((std::min)(sShared.transferredBytes.load(),
                                                sPlan.totalBytes), sPlan.totalBytes);
        else
            sProgress->SetProgress64(sShared.finishedJobs.load(),
                                     (std::max)(static_cast<size_t>(1), sPlan.files.size()));
        ::Sleep(50);
    }

    for (size_t i = 0; i < sWorkers.size(); ++i)
        sWorkers[i].join();

    const HRESULT sCopyFailure = sShared.firstFailure.load();
    const bool sCancelled = sShared.cancelRequested.load() && SUCCEEDED(sCopyFailure);
    if (FAILED(sCopyFailure) || sCancelled)
    {
        const bool sRollbackSucceeded =
            rollbackCopyTargets(sPlan, sCreatedDirectories, sCreatedFiles);
        sProgress->StopProgressDialog();
        sProgress->Release();
        if (sCancelled)
        {
            if (sRollbackSucceeded)
                return ResultCancelled;
            sError = ERROR_CANCELLED;
            showOperationError(aFileOperation, sError, NULL, false);
            if (aError != NULL)
                *aError = sError;
            return ResultFailed;
        }

        sError = HRESULT_FACILITY(sCopyFailure) == FACILITY_WIN32 ?
                 HRESULT_CODE(sCopyFailure) : ERROR_GEN_FAILURE;
        if (sRollbackSucceeded && canRetryWithModernShell(sError))
        {
            if (aError != NULL)
                *aError = sError;
            return ResultNotApplicable;
        }

        const size_t sFailureJob = sShared.firstFailureJob.load();
        const wchar_t *sFailedPath = sFailureJob < sPlan.files.size() ?
                                     sPlan.files[sFailureJob].source.c_str() : NULL;
        showOperationError(aFileOperation, sError, sFailedPath,
                           sRollbackSucceeded);
        if (aError != NULL)
            *aError = sError;
        return ResultFailed;
    }

    restoreDirectoryMetadata(sPlan);
    if (!sourceSnapshotUnchanged(sPlan, sError))
    {
        // A synchronizer or another application changed the source while it
        // was being copied.  Never report that generation as verified and,
        // for a move, never delete the source.
        const bool sRollbackSucceeded =
            rollbackCopyTargets(sPlan, sCreatedDirectories, sCreatedFiles);
        sProgress->StopProgressDialog();
        sProgress->Release();
        if (sRollbackSucceeded && canRetryWithModernShell(sError))
        {
            if (aError != NULL)
                *aError = sError;
            return ResultNotApplicable;
        }
        showOperationError(aFileOperation, sError);
        if (aError != NULL)
            *aError = sError;
        return ResultFailed;
    }
    if (sPlan.moveAcrossVolumes)
    {
        sProgress->SetLine(1, L"복사 완료. 원본 변경 여부를 확인하고 정리합니다.",
                           FALSE, NULL);
        if (!removeSourcesAfterMove(sPlan, sError))
        {
            // The destination is already complete.  Never roll it back after
            // source deletion has started: keeping an extra copy is safer than
            // risking data loss.
            sProgress->StopProgressDialog();
            sProgress->Release();
            showOperationError(aFileOperation, sError);
            if (aError != NULL)
                *aError = sError;
            return ResultFailed;
        }
    }

    sProgress->SetProgress64(sPlan.totalBytes > 0 ? sPlan.totalBytes :
                             sPlan.files.size(),
                             sPlan.totalBytes > 0 ? sPlan.totalBytes :
                             (std::max)(static_cast<size_t>(1), sPlan.files.size()));
    sProgress->StopProgressDialog();
    sProgress->Release();
    return ResultSucceeded;
}
} // namespace fxfile
