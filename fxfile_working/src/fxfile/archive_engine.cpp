//
// Copyright (c) 2026 FxFile Project. All rights reserved.
// Use of this source code is governed by a GPLv3 license.
//
// Real 7-Zip Engine Integration via 7z.exe CLI
//

#include "stdafx.h"
#include "archive_engine.h"
#include <algorithm>
#include <sstream>
#include <shlwapi.h>
#include <shlobj.h>
#include <locale>
#include <codecvt>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace archive
{

// Static member
std::wstring ArchiveEngine::s7zExePath;

// ============================================================
// Path normalization
// ============================================================
static std::wstring normalizePath(const std::wstring &aPath)
{
    std::wstring sResult = aPath;
    std::replace(sResult.begin(), sResult.end(), L'\\', L'/');
    while (!sResult.empty() && sResult.front() == L'/')
        sResult.erase(sResult.begin());
    while (!sResult.empty() && sResult.back() == L'/')
        sResult.pop_back();
    return sResult;
}

// ============================================================
// Robust Console Output Decoder (UTF-8 with CP949 / ACP fallback)
// ============================================================
static std::wstring decodeConsoleOutput(const std::string &aBytes)
{
    if (aBytes.empty())
        return L"";

    // 1. Try UTF-8 first
    int sLen = ::MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, aBytes.c_str(), (int)aBytes.size(), NULL, 0);
    if (sLen > 0)
    {
        std::wstring sResult(sLen, 0);
        ::MultiByteToWideChar(CP_UTF8, 0, aBytes.c_str(), (int)aBytes.size(), &sResult[0], sLen);
        return sResult;
    }

    // 2. Fallback to System ANSI CodePage (CP949 on Korean Windows)
    sLen = ::MultiByteToWideChar(CP_ACP, 0, aBytes.c_str(), (int)aBytes.size(), NULL, 0);
    if (sLen > 0)
    {
        std::wstring sResult(sLen, 0);
        ::MultiByteToWideChar(CP_ACP, 0, aBytes.c_str(), (int)aBytes.size(), &sResult[0], sLen);
        return sResult;
    }

    // 3. Fallback to OEM CodePage
    sLen = ::MultiByteToWideChar(CP_OEMCP, 0, aBytes.c_str(), (int)aBytes.size(), NULL, 0);
    if (sLen > 0)
    {
        std::wstring sResult(sLen, 0);
        ::MultiByteToWideChar(CP_OEMCP, 0, aBytes.c_str(), (int)aBytes.size(), &sResult[0], sLen);
        return sResult;
    }

    return L"";
}

// ============================================================
// 7z.exe path auto-detection
// ============================================================
std::wstring ArchiveEngine::find7zExePath()
{
    // Return cached path if already found and valid
    if (!s7zExePath.empty())
    {
        DWORD sAttr = ::GetFileAttributesW(s7zExePath.c_str());
        if (sAttr != INVALID_FILE_ATTRIBUTES)
            return s7zExePath;
        s7zExePath.clear();
    }

    // 0. Check application directory first (Stand-alone bundled 7-Zip)
    wchar_t sExePath[MAX_PATH] = {0};
    if (::GetModuleFileNameW(NULL, sExePath, MAX_PATH) > 0)
    {
        ::PathRemoveFileSpecW(sExePath);
        std::wstring sAppDir = sExePath;

        const wchar_t *sAppCandidates[] = {
            L"\\7zip\\7z.exe",
            L"\\7zip\\7za.exe",
            L"\\7z.exe",
            L"\\7za.exe",
        };

        for (int i = 0; i < _countof(sAppCandidates); ++i)
        {
            std::wstring sCandidate = sAppDir + sAppCandidates[i];
            DWORD sAttr = ::GetFileAttributesW(sCandidate.c_str());
            if (sAttr != INVALID_FILE_ATTRIBUTES)
            {
                s7zExePath = sCandidate;
                return s7zExePath;
            }
        }
    }

    // 1. Check common installation paths
    const wchar_t *sCommonPaths[] = {
        L"C:\\Program Files\\7-Zip\\7z.exe",
        L"C:\\Program Files (x86)\\7-Zip\\7z.exe",
    };
    for (int i = 0; i < _countof(sCommonPaths); ++i)
    {
        DWORD sAttr = ::GetFileAttributesW(sCommonPaths[i]);
        if (sAttr != INVALID_FILE_ATTRIBUTES)
        {
            s7zExePath = sCommonPaths[i];
            return s7zExePath;
        }
    }

    // 2. Check Windows Registry
    HKEY sKey = NULL;
    const wchar_t *sRegPaths[] = {
        L"SOFTWARE\\7-Zip",
        L"SOFTWARE\\WOW6432Node\\7-Zip",
    };
    for (int i = 0; i < _countof(sRegPaths); ++i)
    {
        if (::RegOpenKeyExW(HKEY_LOCAL_MACHINE, sRegPaths[i], 0, KEY_READ, &sKey) == ERROR_SUCCESS)
        {
            wchar_t sPathBuf[MAX_PATH] = {0};
            DWORD sBufSize = sizeof(sPathBuf);
            // Try Path64 first, then Path
            if (::RegQueryValueExW(sKey, L"Path64", NULL, NULL, (LPBYTE)sPathBuf, &sBufSize) == ERROR_SUCCESS ||
                ::RegQueryValueExW(sKey, L"Path", NULL, NULL, (LPBYTE)sPathBuf, &sBufSize) == ERROR_SUCCESS)
            {
                std::wstring sCandidate = sPathBuf;
                if (!sCandidate.empty() && sCandidate.back() != L'\\')
                    sCandidate += L'\\';
                sCandidate += L"7z.exe";
                DWORD sAttr = ::GetFileAttributesW(sCandidate.c_str());
                if (sAttr != INVALID_FILE_ATTRIBUTES)
                {
                    s7zExePath = sCandidate;
                    ::RegCloseKey(sKey);
                    return s7zExePath;
                }
            }
            ::RegCloseKey(sKey);
        }
    }

    // 3. Check PATH environment variable via SearchPath
    wchar_t sFoundPath[MAX_PATH] = {0};
    if (::SearchPathW(NULL, L"7z.exe", NULL, MAX_PATH, sFoundPath, NULL) > 0)
    {
        s7zExePath = sFoundPath;
        return s7zExePath;
    }

    return L"";
}

bool ArchiveEngine::is7zAvailable()
{
    return !find7zExePath().empty();
}

// ============================================================
// Format helpers
// ============================================================
std::wstring ArchiveEngine::getFormatExtension(ArchiveFormat aFormat)
{
    switch (aFormat)
    {
    case Format7z:   return L".7z";
    case FormatZip:  return L".zip";
    case FormatTar:  return L".tar";
    case FormatGzip: return L".gz";
    case FormatBzip2: return L".bz2";
    case FormatXz:   return L".xz";
    case FormatRar:  return L".rar";
    case FormatCab:  return L".cab";
    case FormatIso:  return L".iso";
    default:         return L".7z";
    }
}

std::wstring ArchiveEngine::getFormatSwitch(ArchiveFormat aFormat)
{
    switch (aFormat)
    {
    case Format7z:   return L"-t7z";
    case FormatZip:  return L"-tzip";
    case FormatTar:  return L"-ttar";
    case FormatGzip: return L"-tgzip";
    case FormatBzip2: return L"-tbzip2";
    case FormatXz:   return L"-txz";
    default:         return L"-t7z";
    }
}

// ============================================================
// Process execution: pumped message loop (non-blocking UI)
// ============================================================
bool ArchiveEngine::executePumped(const std::wstring &aCommandLine, const std::wstring &aWorkingDir)
{
    STARTUPINFOW sSi = {0};
    PROCESS_INFORMATION sPi = {0};
    sSi.cb = sizeof(STARTUPINFOW);
    sSi.dwFlags = STARTF_USESHOWWINDOW;
    sSi.wShowWindow = SW_HIDE;

    std::vector<wchar_t> sCmdBuf(aCommandLine.begin(), aCommandLine.end());
    sCmdBuf.push_back(0);

    LPCWSTR sCwd = aWorkingDir.empty() ? NULL : aWorkingDir.c_str();

    BOOL sSuccess = ::CreateProcessW(NULL, sCmdBuf.data(), NULL, NULL, FALSE,
                                     CREATE_NO_WINDOW, NULL, sCwd, &sSi, &sPi);
    if (!sSuccess)
        return false;

    // Pumped wait: wait until process finishes
    for (;;)
    {
        DWORD sWaitResult = ::WaitForSingleObject(sPi.hProcess, 50);
        if (sWaitResult == WAIT_OBJECT_0)
            break;
    }

    DWORD sExitCode = 0;
    ::GetExitCodeProcess(sPi.hProcess, &sExitCode);
    ::CloseHandle(sPi.hProcess);
    ::CloseHandle(sPi.hThread);

    return (sExitCode == 0);
}

// ============================================================
// Process execution: with stdout pipe for progress parsing
// ============================================================
bool ArchiveEngine::executeWithProgress(const std::wstring &aCommandLine,
                                        const std::wstring &aWorkingDir,
                                        ArchiveProgressCallback aCallback)
{
    if (!aCallback)
        return executePumped(aCommandLine, aWorkingDir);

    // Create pipe for stdout
    SECURITY_ATTRIBUTES sSa = {0};
    sSa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sSa.bInheritHandle = TRUE;
    HANDLE sReadPipe = NULL, sWritePipe = NULL;
    if (!::CreatePipe(&sReadPipe, &sWritePipe, &sSa, 0))
        return executePumped(aCommandLine, aWorkingDir);

    ::SetHandleInformation(sReadPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW sSi = {0};
    PROCESS_INFORMATION sPi = {0};
    sSi.cb = sizeof(STARTUPINFOW);
    sSi.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    sSi.wShowWindow = SW_HIDE;
    sSi.hStdOutput = sWritePipe;
    sSi.hStdError = sWritePipe;

    std::vector<wchar_t> sCmdBuf(aCommandLine.begin(), aCommandLine.end());
    sCmdBuf.push_back(0);

    LPCWSTR sCwd = aWorkingDir.empty() ? NULL : aWorkingDir.c_str();

    BOOL sSuccess = ::CreateProcessW(NULL, sCmdBuf.data(), NULL, NULL, TRUE,
                                     CREATE_NO_WINDOW, NULL, sCwd, &sSi, &sPi);
    ::CloseHandle(sWritePipe);

    if (!sSuccess)
    {
        ::CloseHandle(sReadPipe);
        return false;
    }

    // Read stdout in responsive loop
    char sBuffer[4096];
    std::string sAccum;
    bool sCancelled = false;

    for (;;)
    {
        // Read available stdout data
        DWORD sBytesAvail = 0;
        while (::PeekNamedPipe(sReadPipe, NULL, 0, NULL, &sBytesAvail, NULL) && sBytesAvail > 0)
        {
            DWORD sRead = 0;
            DWORD sToRead = (sBytesAvail < sizeof(sBuffer)) ? sBytesAvail : sizeof(sBuffer);
            if (::ReadFile(sReadPipe, sBuffer, sToRead, &sRead, NULL) && sRead > 0)
            {
                sAccum.append(sBuffer, sRead);

                // Parse progress: look for "XX%" pattern in output
                size_t sPos = sAccum.rfind('%');
                if (sPos != std::string::npos && sPos >= 1)
                {
                    // Find the start of the number before %
                    size_t sNumStart = sPos;
                    while (sNumStart > 0 && sAccum[sNumStart - 1] >= '0' && sAccum[sNumStart - 1] <= '9')
                        --sNumStart;
                    if (sNumStart < sPos)
                    {
                        int sPercent = atoi(sAccum.c_str() + sNumStart);
                        if (sPercent >= 0 && sPercent <= 100)
                        {
                            // Extract current filename if present after " - "
                            std::wstring sFileName;
                            size_t sDash = sAccum.find(" - ", sPos);
                            if (sDash != std::string::npos)
                            {
                                size_t sEnd = sAccum.find_first_of("\r\n", sDash + 3);
                                if (sEnd == std::string::npos) sEnd = sAccum.length();
                                std::string sRawName = sAccum.substr(sDash + 3, sEnd - sDash - 3);
                                
                                // Clean up control characters
                                while (!sRawName.empty() && ((unsigned char)sRawName.front() < 32 || sRawName.front() == ' '))
                                    sRawName.erase(sRawName.begin());
                                while (!sRawName.empty() && ((unsigned char)sRawName.back() < 32 || sRawName.back() == ' '))
                                    sRawName.pop_back();

                                sFileName = decodeConsoleOutput(sRawName);
                            }

                            if (!aCallback(sPercent, sFileName))
                            {
                                sCancelled = true;
                                ::TerminateProcess(sPi.hProcess, 1);
                                break;
                            }
                        }
                    }
                }

                // Keep only recent buffer data
                if (sAccum.size() > 8192)
                    sAccum = sAccum.substr(sAccum.size() - 4096);
            }
        }

        if (sCancelled)
            break;

        DWORD sWaitResult = ::WaitForSingleObject(sPi.hProcess, 30);
        if (sWaitResult == WAIT_OBJECT_0)
        {
            break;
        }
    }

    DWORD sExitCode = 0;
    ::GetExitCodeProcess(sPi.hProcess, &sExitCode);
    ::CloseHandle(sPi.hProcess);
    ::CloseHandle(sPi.hThread);
    ::CloseHandle(sReadPipe);

    if (sCancelled)
        return false;

    // Report 100% completion
    if (aCallback)
        aCallback(100, L"");

    return (sExitCode == 0);
}

// ============================================================
// Execute and capture stdout (for list command)
// ============================================================
bool ArchiveEngine::executeAndCapture(const std::wstring &aCommandLine,
                                      const std::wstring &aWorkingDir,
                                      std::wstring &aOutput)
{
    aOutput.clear();

    SECURITY_ATTRIBUTES sSa = {0};
    sSa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sSa.bInheritHandle = TRUE;
    HANDLE sReadPipe = NULL, sWritePipe = NULL;
    if (!::CreatePipe(&sReadPipe, &sWritePipe, &sSa, 0))
        return false;

    ::SetHandleInformation(sReadPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW sSi = {0};
    PROCESS_INFORMATION sPi = {0};
    sSi.cb = sizeof(STARTUPINFOW);
    sSi.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    sSi.wShowWindow = SW_HIDE;
    sSi.hStdOutput = sWritePipe;
    sSi.hStdError = sWritePipe;

    std::vector<wchar_t> sCmdBuf(aCommandLine.begin(), aCommandLine.end());
    sCmdBuf.push_back(0);

    LPCWSTR sCwd = aWorkingDir.empty() ? NULL : aWorkingDir.c_str();

    BOOL sSuccess = ::CreateProcessW(NULL, sCmdBuf.data(), NULL, NULL, TRUE,
                                     CREATE_NO_WINDOW, NULL, sCwd, &sSi, &sPi);
    ::CloseHandle(sWritePipe);

    if (!sSuccess)
    {
        ::CloseHandle(sReadPipe);
        return false;
    }

    std::string sRawOutput;
    char sBuffer[4096];
    DWORD sRead = 0;
    while (::ReadFile(sReadPipe, sBuffer, sizeof(sBuffer), &sRead, NULL) && sRead > 0)
    {
        sRawOutput.append(sBuffer, sRead);
    }

    ::WaitForSingleObject(sPi.hProcess, INFINITE);
    DWORD sExitCode = 0;
    ::GetExitCodeProcess(sPi.hProcess, &sExitCode);
    ::CloseHandle(sPi.hProcess);
    ::CloseHandle(sPi.hThread);
    ::CloseHandle(sReadPipe);

    if (sExitCode == 0 && !sRawOutput.empty())
    {
        aOutput = decodeConsoleOutput(sRawOutput);
    }

    return (sExitCode == 0);
}

// ============================================================
// Constructor / Destructor
// ============================================================
ArchiveEngine::ArchiveEngine()
    : mIsOpen(false)
    , mFormat(FormatUnknown)
{
}

ArchiveEngine::~ArchiveEngine()
{
    closeArchive();
}

// ============================================================
// Format Detection
// ============================================================
bool ArchiveEngine::isArchiveFile(const std::wstring &aFilePath)
{
    return (detectFormat(aFilePath) != FormatUnknown);
}

ArchiveFormat ArchiveEngine::detectFormat(const std::wstring &aFilePath)
{
    LPCWSTR sExt = ::PathFindExtensionW(aFilePath.c_str());
    if (!sExt || sExt[0] == 0)
        return FormatUnknown;

    if (_wcsicmp(sExt, L".7z") == 0)   return Format7z;
    if (_wcsicmp(sExt, L".zip") == 0)  return FormatZip;
    if (_wcsicmp(sExt, L".tar") == 0)  return FormatTar;
    if (_wcsicmp(sExt, L".gz") == 0)   return FormatGzip;
    if (_wcsicmp(sExt, L".tgz") == 0)  return FormatGzip;
    if (_wcsicmp(sExt, L".bz2") == 0)  return FormatBzip2;
    if (_wcsicmp(sExt, L".xz") == 0)   return FormatXz;
    if (_wcsicmp(sExt, L".rar") == 0)  return FormatRar;
    if (_wcsicmp(sExt, L".iso") == 0)  return FormatIso;
    if (_wcsicmp(sExt, L".cab") == 0)  return FormatCab;

    return FormatUnknown;
}

// ============================================================
// Open Archive (reads structure via 7z.exe l -slt)
// ============================================================
bool ArchiveEngine::openArchive(const std::wstring &aArchivePath)
{
    closeArchive();

    if (aArchivePath.empty())
        return false;

    DWORD sAttr = ::GetFileAttributesW(aArchivePath.c_str());
    if (sAttr == INVALID_FILE_ATTRIBUTES || (sAttr & FILE_ATTRIBUTE_DIRECTORY))
        return false;

    mFormat = detectFormat(aArchivePath);
    if (mFormat == FormatUnknown)
        return false;

    mArchivePath = aArchivePath;

    std::wstring s7z = find7zExePath();
    if (s7z.empty())
    {
        // If 7z is not available, we cannot inspect internal archive tree
        return false;
    }

    // Command: 7z.exe l -slt -sccUTF-8 "aArchivePath"
    std::wstring sCmd = L"\"" + s7z + L"\" l -slt -sccUTF-8 \"" + aArchivePath + L"\"";
    std::wstring sOutput;
    if (!executeAndCapture(sCmd, L"", sOutput))
        return false;

    // Build directory tree from 7z output
    mRootNode = std::make_shared<ArchiveDirectoryNode>(L"");
    mAllItems.clear();

    std::wistringstream sStream(sOutput);
    std::wstring sLine;
    ArchiveItem sCurrentItem;
    bool sInItem = false;

    while (std::getline(sStream, sLine))
    {
        while (!sLine.empty() && (sLine.back() == L'\r' || sLine.back() == L'\n'))
            sLine.pop_back();

        if (sLine.rfind(L"Path = ", 0) == 0)
        {
            if (sInItem && !sCurrentItem.mFullPath.empty() && sCurrentItem.mFullPath != aArchivePath)
            {
                mAllItems.push_back(sCurrentItem);
            }
            sCurrentItem = ArchiveItem();
            sCurrentItem.mFullPath = sLine.substr(7);
            sInItem = true;
        }
        else if (sInItem)
        {
            if (sLine.rfind(L"Size = ", 0) == 0)
                sCurrentItem.mUncompressedSize = _wcstoui64(sLine.substr(7).c_str(), NULL, 10);
            else if (sLine.rfind(L"Packed Size = ", 0) == 0)
                sCurrentItem.mCompressedSize = _wcstoui64(sLine.substr(14).c_str(), NULL, 10);
            else if (sLine.rfind(L"Folder = ", 0) == 0)
                sCurrentItem.mIsDirectory = (sLine.substr(9) == L"+");
        }
    }

    if (sInItem && !sCurrentItem.mFullPath.empty() && sCurrentItem.mFullPath != aArchivePath)
    {
        mAllItems.push_back(sCurrentItem);
    }

    // Populate tree nodes
    for (size_t i = 0; i < mAllItems.size(); ++i)
    {
        ArchiveItem &sItem = mAllItems[i];
        std::wstring sNormPath = normalizePath(sItem.mFullPath);
        sItem.mFullPath = sNormPath;

        size_t sSlash = sNormPath.find_last_of(L'/');
        std::wstring sItemDir;
        if (sSlash == std::wstring::npos)
        {
            sItem.mFileName = sNormPath;
            sItemDir = L"";
        }
        else
        {
            sItem.mFileName = sNormPath.substr(sSlash + 1);
            sItemDir = sNormPath.substr(0, sSlash);
        }

        // Add to tree hierarchy
        std::shared_ptr<ArchiveDirectoryNode> sCurr = mRootNode;
        if (!sItemDir.empty())
        {
            std::wistringstream sDirStream(sItemDir);
            std::wstring sSegment;
            while (std::getline(sDirStream, sSegment, L'/'))
            {
                if (sSegment.empty()) continue;
                auto it = sCurr->mSubDirectories.find(sSegment);
                if (it == sCurr->mSubDirectories.end())
                {
                    auto sNewNode = std::make_shared<ArchiveDirectoryNode>(sSegment);
                    sCurr->mSubDirectories[sSegment] = sNewNode;
                    sCurr = sNewNode;
                }
                else
                {
                    sCurr = it->second;
                }
            }
        }

        if (sItem.mIsDirectory)
        {
            if (sCurr->mSubDirectories.find(sItem.mFileName) == sCurr->mSubDirectories.end())
            {
                sCurr->mSubDirectories[sItem.mFileName] = std::make_shared<ArchiveDirectoryNode>(sItem.mFileName);
            }
        }
        else
        {
            sCurr->mFiles.push_back(sItem);
        }
    }

    mIsOpen = true;
    return true;
}

void ArchiveEngine::closeArchive()
{
    mIsOpen = false;
    mArchivePath.clear();
    mFormat = FormatUnknown;
    mAllItems.clear();
    mRootNode.reset();
}

// ============================================================
// Query Directory Listing
// ============================================================
std::shared_ptr<ArchiveDirectoryNode> ArchiveEngine::getNodeByPath(const std::wstring &aVirtualPath) const
{
    if (!mIsOpen || !mRootNode)
        return nullptr;

    std::wstring sNorm = normalizePath(aVirtualPath);
    if (sNorm.empty())
        return mRootNode;

    std::wistringstream sStream(sNorm);
    std::wstring sSeg;
    std::shared_ptr<ArchiveDirectoryNode> sCurr = mRootNode;

    while (std::getline(sStream, sSeg, L'/'))
    {
        if (sSeg.empty()) continue;
        auto it = sCurr->mSubDirectories.find(sSeg);
        if (it == sCurr->mSubDirectories.end())
            return nullptr;
        sCurr = it->second;
    }

    return sCurr;
}

bool ArchiveEngine::getDirectoryListing(const std::wstring &aVirtualPath, std::vector<ArchiveItem> &aOutItems) const
{
    aOutItems.clear();
    auto sNode = getNodeByPath(aVirtualPath);
    if (!sNode)
        return false;

    // 1. Add directories first
    for (auto const &pair : sNode->mSubDirectories)
    {
        ArchiveItem sDirItem;
        sDirItem.mFileName = pair.first;
        sDirItem.mFullPath = aVirtualPath.empty() ? pair.first : (normalizePath(aVirtualPath) + L"/" + pair.first);
        sDirItem.mIsDirectory = true;
        sDirItem.mUncompressedSize = 0;
        aOutItems.push_back(sDirItem);
    }

    // 2. Add files
    for (size_t i = 0; i < sNode->mFiles.size(); ++i)
    {
        aOutItems.push_back(sNode->mFiles[i]);
    }

    return true;
}

// ============================================================
// Extract operations (real 7z.exe)
// ============================================================
bool ArchiveEngine::extractFile(const std::wstring &aVirtualFilePath, const std::wstring &aDestFilePath)
{
    if (!mIsOpen)
        return false;

    std::wstring s7z = find7zExePath();
    if (s7z.empty())
        return false;

    std::wstring sDestDir = aDestFilePath;
    size_t sLastSlash = sDestDir.find_last_of(L"\\/");
    if (sLastSlash != std::wstring::npos)
    {
        std::wstring sParent = sDestDir.substr(0, sLastSlash);
        ::SHCreateDirectoryExW(NULL, sParent.c_str(), NULL);
    }

    std::wstring sCmd = L"\"" + s7z + L"\" e -sccUTF-8 \"" + mArchivePath + L"\" -o\"" +
                        sDestDir.substr(0, sLastSlash) + L"\" \"" + aVirtualFilePath + L"\" -y";
    return executePumped(sCmd, L"");
}

bool ArchiveEngine::extractDirectory(const std::wstring &aVirtualDirPath, const std::wstring &aDestDirPath)
{
    if (!mIsOpen)
        return false;

    std::wstring s7z = find7zExePath();
    if (s7z.empty())
        return false;

    ::SHCreateDirectoryExW(NULL, aDestDirPath.c_str(), NULL);

    std::wstring sCmd = L"\"" + s7z + L"\" x -sccUTF-8 \"" + mArchivePath + L"\" -o\"" +
                        aDestDirPath + L"\" \"" + aVirtualDirPath + L"\\*\" -y";
    return executePumped(sCmd, L"");
}

bool ArchiveEngine::extractAll(const std::wstring &aDestDirPath, ArchiveProgressCallback aCallback)
{
    if (!mIsOpen || mArchivePath.empty())
        return false;

    std::wstring s7z = find7zExePath();
    if (s7z.empty())
    {
        // Fallback to PowerShell if no 7z
        ::SHCreateDirectoryExW(NULL, aDestDirPath.c_str(), NULL);
        std::wstring sPsCmd = L"powershell.exe -NoProfile -NonInteractive -Command \"Expand-Archive -LiteralPath '" +
                              mArchivePath + L"' -DestinationPath '" + aDestDirPath + L"' -Force\"";
        bool sOk = executePumped(sPsCmd, aDestDirPath);
        if (sOk)
            ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, aDestDirPath.c_str(), NULL);
        return sOk;
    }

    ::SHCreateDirectoryExW(NULL, aDestDirPath.c_str(), NULL);

    // 7z.exe x -sccUTF-8 -bsp1 -y "archive" -o"dest"
    std::wstring sCmd = L"\"" + s7z + L"\" x -sccUTF-8 -bsp1 -y \"" + mArchivePath + L"\" -o\"" +
                        aDestDirPath + L"\"";
    bool sOk = executeWithProgress(sCmd, L"", aCallback);
    if (sOk)
        ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, aDestDirPath.c_str(), NULL);
    return sOk;
}

// ============================================================
// Archive modification
// ============================================================
bool ArchiveEngine::addFiles(const std::vector<std::wstring> &aSrcPhysicalFiles, const std::wstring &aTargetVirtualDir)
{
    if (!mIsOpen)
        return false;

    std::wstring s7z = find7zExePath();
    if (s7z.empty())
        return false;

    std::wstring sCmd = L"\"" + s7z + L"\" a -sccUTF-8 \"" + mArchivePath + L"\"";
    for (size_t i = 0; i < aSrcPhysicalFiles.size(); ++i)
    {
        sCmd += L" \"" + aSrcPhysicalFiles[i] + L"\"";
    }
    sCmd += L" -y";

    return executePumped(sCmd, L"");
}

bool ArchiveEngine::deleteFiles(const std::vector<std::wstring> &aVirtualFilePaths)
{
    if (!mIsOpen)
        return false;

    std::wstring s7z = find7zExePath();
    if (s7z.empty())
        return false;

    std::wstring sCmd = L"\"" + s7z + L"\" d -sccUTF-8 \"" + mArchivePath + L"\"";
    for (size_t i = 0; i < aVirtualFilePaths.size(); ++i)
    {
        sCmd += L" \"" + aVirtualFilePaths[i] + L"\"";
    }
    sCmd += L" -y";

    return executePumped(sCmd, L"");
}

// ============================================================
// Create Archive (real 7z.exe with high speed multithreading)
// ============================================================
bool ArchiveEngine::createArchive(const std::wstring &aArchivePath,
                                  const std::vector<std::wstring> &aSrcFiles,
                                  ArchiveFormat aFormat,
                                  int aCompressionLevel,
                                  ArchiveProgressCallback aCallback)
{
    if (aArchivePath.empty() || aSrcFiles.empty())
        return false;

    // Determine parent directory of sources
    std::wstring sFirst = aSrcFiles[0];
    size_t sSlash = sFirst.find_last_of(L"\\/");
    std::wstring sParentDir = (sSlash != std::wstring::npos) ? sFirst.substr(0, sSlash) : L"";

    std::wstring s7z = find7zExePath();

    if (!s7z.empty())
    {
        // High-Speed Multi-Threaded Real 7z.exe compression
        // 7z.exe a -t7z -mx=5 -mmt=on -sccUTF-8 -bsp1 -y "archive.7z" "file1" "file2"
        std::wstring sFormatSwitch = getFormatSwitch(aFormat);
        wchar_t sLevelBuf[32];
        int sLevel = (aCompressionLevel >= 1 && aCompressionLevel <= 9) ? aCompressionLevel : 5;
        _snwprintf_s(sLevelBuf, _countof(sLevelBuf), _TRUNCATE, L"-mx=%d -mmt=on", sLevel);

        std::wstring sCmd = L"\"" + s7z + L"\" a " + sFormatSwitch + L" " + sLevelBuf +
                            L" -sccUTF-8 -bsp1 -y \"" + aArchivePath + L"\"";
        for (size_t i = 0; i < aSrcFiles.size(); ++i)
        {
            sCmd += L" \"" + aSrcFiles[i] + L"\"";
        }

        bool sOk = executeWithProgress(sCmd, sParentDir, aCallback);
        if (sOk)
        {
            DWORD sAttr = ::GetFileAttributesW(aArchivePath.c_str());
            if (sAttr != INVALID_FILE_ATTRIBUTES)
            {
                ::SHChangeNotify(SHCNE_CREATE, SHCNF_PATH, aArchivePath.c_str(), NULL);
                if (!sParentDir.empty())
                    ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, sParentDir.c_str(), NULL);
                return true;
            }
        }
        return false;
    }

    // Fallback: tar.exe for .zip only
    std::wstring sActualPath = aArchivePath;
    LPCWSTR sExt = ::PathFindExtensionW(sActualPath.c_str());
    if (sExt && _wcsicmp(sExt, L".7z") == 0)
    {
        sActualPath = sActualPath.substr(0, sActualPath.length() - 3) + L".zip";
    }

    std::wstring sTarCmd = L"tar.exe -a -c -f \"" + sActualPath + L"\"";
    for (size_t i = 0; i < aSrcFiles.size(); ++i)
    {
        LPCWSTR sFilePart = ::PathFindFileNameW(aSrcFiles[i].c_str());
        sTarCmd += L" \"" + std::wstring(sFilePart) + L"\"";
    }

    if (executePumped(sTarCmd, sParentDir))
    {
        DWORD sAttr = ::GetFileAttributesW(sActualPath.c_str());
        if (sAttr != INVALID_FILE_ATTRIBUTES)
        {
            ::SHChangeNotify(SHCNE_CREATE, SHCNF_PATH, sActualPath.c_str(), NULL);
            if (!sParentDir.empty())
                ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, sParentDir.c_str(), NULL);
            return true;
        }
    }

    // Final fallback: PowerShell
    std::wstring sPsCmd = L"powershell.exe -NoProfile -NonInteractive -Command \"$files = @(";
    for (size_t i = 0; i < aSrcFiles.size(); ++i)
    {
        if (i > 0) sPsCmd += L",";
        sPsCmd += L"'" + aSrcFiles[i] + L"'";
    }
    sPsCmd += L"); Compress-Archive -LiteralPath $files -DestinationPath '" + sActualPath + L"' -Force\"";

    bool sOk = executePumped(sPsCmd, sParentDir);
    if (sOk)
    {
        ::SHChangeNotify(SHCNE_CREATE, SHCNF_PATH, sActualPath.c_str(), NULL);
        if (!sParentDir.empty())
            ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH, sParentDir.c_str(), NULL);
    }
    return sOk;
}

bool ArchiveEngine::extractItems(const std::wstring &aArchivePath,
                                 const std::vector<std::wstring> &aItemNames,
                                 const std::wstring &aDestDirPath,
                                 ArchiveProgressCallback aCallback)
{
    std::wstring s7z = find7zExePath();
    if (s7z.empty() || aItemNames.empty())
        return false;

    // Ensure destination directory exists
    ::SHCreateDirectoryExW(NULL, aDestDirPath.c_str(), NULL);

    std::wostringstream sCmd;
    if (aCallback != nullptr)
    {
        sCmd << L"\"" << s7z << L"\" x \"" << aArchivePath << L"\" -o\"" << aDestDirPath << L"\" -mmt=on -bsp1 -y -aoa -r";
    }
    else
    {
        sCmd << L"\"" << s7z << L"\" x \"" << aArchivePath << L"\" -o\"" << aDestDirPath << L"\" -mmt=on -bso0 -bse0 -bsp0 -y -aoa -r";
    }

    for (size_t i = 0; i < aItemNames.size(); ++i)
    {
        sCmd << L" \"" << aItemNames[i] << L"\"";
    }

    bool sOk = false;
    if (aCallback != nullptr)
    {
        sOk = executeWithProgress(sCmd.str(), L"", aCallback);
    }
    else
    {
        std::wstring sOutput;
        sOk = executeAndCapture(sCmd.str(), L"", sOutput);
    }

    if (sOk)
    {
        ::SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATH | SHCNF_FLUSH, aDestDirPath.c_str(), NULL);
    }
    return sOk;
}

bool ArchiveEngine::addItemsToArchive(const std::wstring &aArchivePath,
                                      const std::wstring &aVirtualSubDir,
                                      const std::vector<std::wstring> &aSrcPhysicalPaths,
                                      ArchiveProgressCallback aCallback)
{
    std::wstring s7z = find7zExePath();
    if (s7z.empty() || aSrcPhysicalPaths.empty())
        return false;

    std::wostringstream sCmd;
    if (aCallback != nullptr)
    {
        sCmd << L"\"" << s7z << L"\" a \"" << aArchivePath << L"\" -mmt=on -bsp1 -y -ssw";
    }
    else
    {
        sCmd << L"\"" << s7z << L"\" a \"" << aArchivePath << L"\" -mmt=on -bso0 -bse0 -bsp0 -y -ssw";
    }

    for (size_t i = 0; i < aSrcPhysicalPaths.size(); ++i)
    {
        sCmd << L" \"" << aSrcPhysicalPaths[i] << L"\"";
    }

    if (aCallback != nullptr)
    {
        return executeWithProgress(sCmd.str(), L"", aCallback);
    }

    std::wstring sOutput;
    return executeAndCapture(sCmd.str(), L"", sOutput);
}

} // namespace archive
} // namespace fxfile