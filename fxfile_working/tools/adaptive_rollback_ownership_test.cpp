#define FXFILE_ADAPTIVE_STANDALONE
#include "../src/fxfile/adaptive_file_operation.cpp"

#include <stdio.h>

namespace
{
bool writeTextFile(const std::wstring &aPath, const char *aText)
{
    HANDLE sFile = ::CreateFileW(aPath.c_str(), GENERIC_WRITE,
                                 FILE_SHARE_READ | FILE_SHARE_WRITE |
                                     FILE_SHARE_DELETE,
                                 NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
    if (sFile == INVALID_HANDLE_VALUE)
        return false;
    DWORD sWritten = 0;
    const DWORD sLength = static_cast<DWORD>(strlen(aText));
    const bool sSucceeded =
        ::WriteFile(sFile, aText, sLength, &sWritten, NULL) != FALSE &&
        sWritten == sLength;
    ::CloseHandle(sFile);
    return sSucceeded;
}

bool fileContains(const std::wstring &aPath, const char *aExpected)
{
    HANDLE sFile = ::CreateFileW(aPath.c_str(), GENERIC_READ,
                                 FILE_SHARE_READ | FILE_SHARE_WRITE |
                                     FILE_SHARE_DELETE,
                                 NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (sFile == INVALID_HANDLE_VALUE)
        return false;
    char sBuffer[64] = {0};
    DWORD sRead = 0;
    const bool sSucceeded =
        ::ReadFile(sFile, sBuffer, sizeof(sBuffer) - 1, &sRead, NULL) != FALSE;
    ::CloseHandle(sFile);
    return sSucceeded && strcmp(sBuffer, aExpected) == 0;
}

fxfile::FileJob makeFileJob(const std::wstring &aTarget)
{
    fxfile::FileJob sJob;
    sJob.target = aTarget;
    sJob.size = 0;
    ::ZeroMemory(&sJob.lastWriteTime, sizeof(sJob.lastWriteTime));
    return sJob;
}

fxfile::DirectoryJob makeDirectoryJob(const std::wstring &aTarget)
{
    fxfile::DirectoryJob sJob;
    sJob.target = aTarget;
    sJob.attributes = FILE_ATTRIBUTE_DIRECTORY;
    ::ZeroMemory(&sJob.creationTime, sizeof(sJob.creationTime));
    ::ZeroMemory(&sJob.accessTime, sizeof(sJob.accessTime));
    ::ZeroMemory(&sJob.lastWriteTime, sizeof(sJob.lastWriteTime));
    return sJob;
}

int fail(const wchar_t *aMessage)
{
    fwprintf(stderr, L"FAIL: %ls (error=%lu)\n", aMessage, ::GetLastError());
    return 1;
}
}

int wmain(int aArgumentCount, wchar_t **aArguments)
{
    if (aArgumentCount != 2)
    {
        fwprintf(stderr, L"usage: adaptive_rollback_ownership_test <empty-test-root>\n");
        return 2;
    }

    const std::wstring sRoot(aArguments[1]);
    if (!::CreateDirectoryW(sRoot.c_str(), NULL))
        return fail(L"test root must not already exist");

    // A target whose identity was captured by this operation is removable.
    const std::wstring sOwned = fxfile::joinPath(sRoot, L"owned.bin");
    if (!writeTextFile(sOwned, "owned"))
        return fail(L"create owned target");
    fxfile::CreatedTargetEvidence sOwnedEvidence;
    if (!fxfile::captureTargetEvidence(sOwned, false, sOwnedEvidence))
        return fail(L"capture owned target identity");
    fxfile::CopyPlan sOwnedPlan;
    sOwnedPlan.files.push_back(makeFileJob(sOwned));
    std::vector<fxfile::CreatedTargetEvidence> sOwnedTargets(1, sOwnedEvidence);
    if (!fxfile::rollbackTargets(sOwnedPlan, sOwnedTargets) ||
        ::GetFileAttributesW(sOwned.c_str()) != INVALID_FILE_ATTRIBUTES)
        return fail(L"owned target rollback");

    // Replace the path with a different, concurrently-created file object.
    // The stale evidence must not delete that external object.
    const std::wstring sReplaced = fxfile::joinPath(sRoot, L"replaced.bin");
    const std::wstring sExternal = fxfile::joinPath(sRoot, L"external.tmp");
    if (!writeTextFile(sReplaced, "old-owned"))
        return fail(L"create replaceable target");
    fxfile::CreatedTargetEvidence sStaleEvidence;
    if (!fxfile::captureTargetEvidence(sReplaced, false, sStaleEvidence))
        return fail(L"capture stale target identity");
    if (!writeTextFile(sExternal, "external-owner") ||
        !::MoveFileExW(sExternal.c_str(), sReplaced.c_str(),
                       MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
        return fail(L"replace target with external object");
    fxfile::CopyPlan sReplacementPlan;
    sReplacementPlan.files.push_back(makeFileJob(sReplaced));
    std::vector<fxfile::CreatedTargetEvidence> sStaleTargets(1, sStaleEvidence);
    if (fxfile::rollbackTargets(sReplacementPlan, sStaleTargets) ||
        !fileContains(sReplaced, "external-owner"))
        return fail(L"external replacement was not preserved");

    // Even inside an operation-created directory, an untracked child must
    // prevent directory deletion and therefore prevent shell fallback.
    const std::wstring sOwnedDirectory =
        fxfile::joinPath(sRoot, L"owned_directory");
    if (!::CreateDirectoryW(sOwnedDirectory.c_str(), NULL))
        return fail(L"create owned directory");
    fxfile::CreatedTargetEvidence sDirectoryEvidence;
    if (!fxfile::captureTargetEvidence(sOwnedDirectory, true,
                                       sDirectoryEvidence))
        return fail(L"capture owned directory identity");
    const std::wstring sForeignChild =
        fxfile::joinPath(sOwnedDirectory, L"foreign.keep");
    if (!writeTextFile(sForeignChild, "foreign-child"))
        return fail(L"create foreign child");
    fxfile::CopyPlan sDirectoryPlan;
    sDirectoryPlan.directories.push_back(makeDirectoryJob(sOwnedDirectory));
    std::vector<fxfile::CreatedTargetEvidence> sDirectoryTargets(
        1, sDirectoryEvidence);
    if (fxfile::rollbackTargets(sDirectoryPlan, sDirectoryTargets) ||
        !fileContains(sForeignChild, "foreign-child"))
        return fail(L"foreign child was not preserved");

    ::DeleteFileW(sReplaced.c_str());
    ::DeleteFileW(sForeignChild.c_str());
    ::RemoveDirectoryW(sOwnedDirectory.c_str());
    if (!::RemoveDirectoryW(sRoot.c_str()))
        return fail(L"fixture cleanup");

    wprintf(L"PASS owned deletion; PASS replacement preservation; "
            L"PASS foreign-child preservation\n");
    return 0;
}
