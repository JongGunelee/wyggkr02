//
// Copyright (c) 2026 FxFile Project. All rights reserved.
// Use of this source code is governed by a GPLv3 license.
//
// Real 7-Zip Engine Integration via 7z.exe CLI
//

#ifndef __FXFILE_ARCHIVE_ENGINE_H__
#define __FXFILE_ARCHIVE_ENGINE_H__ 1
#pragma once

#include <string>
#include <vector>
#include <map>
#include <memory>
#include <cstdint>
#include <functional>
#include <windows.h>

namespace fxfile
{
namespace archive
{

enum ArchiveFormat
{
    FormatUnknown = 0,
    Format7z,
    FormatZip,
    FormatTar,
    FormatGzip,
    FormatBzip2,
    FormatXz,
    FormatRar,
    FormatCab,
    FormatIso
};

struct ArchiveItem
{
    std::wstring mFullPath;       // Virtual full path inside archive: e.g. "sub/deep/file.txt"
    std::wstring mFileName;       // "file.txt"
    bool         mIsDirectory;    // true if folder
    uint64_t     mUncompressedSize;
    uint64_t     mCompressedSize;
    FILETIME     mLastModified;
    uint32_t     mAttributes;
    uint32_t     mIndex;          // Index in archive

    ArchiveItem()
        : mIsDirectory(false)
        , mUncompressedSize(0)
        , mCompressedSize(0)
        , mAttributes(0)
        , mIndex(0)
    {
        memset(&mLastModified, 0, sizeof(FILETIME));
    }
};

struct ArchiveDirectoryNode
{
    std::wstring mVirtualDirPath;  // e.g. "sub/deep"
    std::wstring mDirName;         // "deep"
    std::vector<ArchiveItem> mFiles;
    std::map<std::wstring, std::shared_ptr<ArchiveDirectoryNode>> mSubDirectories;

    ArchiveDirectoryNode(const std::wstring &aPath = L"", const std::wstring &aName = L"")
        : mVirtualDirPath(aPath)
        , mDirName(aName)
    {
    }

    void clear()
    {
        mFiles.clear();
        mSubDirectories.clear();
    }
};

// Progress callback: receives percent (0-100) and current file name.
// Return false from callback to request cancellation.
typedef std::function<bool(int aPercent, const std::wstring &aCurrentFile)> ArchiveProgressCallback;

class ArchiveEngine
{
public:
    ArchiveEngine();
    ~ArchiveEngine();

    // 7z.exe path detection and validation
    static std::wstring find7zExePath();
    static bool is7zAvailable();

    static bool isArchiveFile(const std::wstring &aFilePath);
    static ArchiveFormat detectFormat(const std::wstring &aFilePath);
    static std::wstring getFormatExtension(ArchiveFormat aFormat);
    static std::wstring getFormatSwitch(ArchiveFormat aFormat);

    bool openArchive(const std::wstring &aArchivePath);
    void closeArchive();

    bool isOpen() const { return mIsOpen; }
    const std::wstring &getArchivePath() const { return mArchivePath; }

    // Navigation & Query
    std::shared_ptr<ArchiveDirectoryNode> getRootNode() const { return mRootNode; }
    std::shared_ptr<ArchiveDirectoryNode> getNodeByPath(const std::wstring &aVirtualPath) const;
    bool getDirectoryListing(const std::wstring &aVirtualPath, std::vector<ArchiveItem> &aOutItems) const;

    // Operations with progress
    bool extractFile(const std::wstring &aVirtualFilePath, const std::wstring &aDestFilePath);
    bool extractDirectory(const std::wstring &aVirtualDirPath, const std::wstring &aDestDirPath);
    bool extractAll(const std::wstring &aDestDirPath, ArchiveProgressCallback aCallback = nullptr);

    bool addFiles(const std::vector<std::wstring> &aSrcPhysicalFiles, const std::wstring &aTargetVirtualDir);
    bool deleteFiles(const std::vector<std::wstring> &aVirtualFilePaths);

    // Static helpers for selective extract and direct add
    static bool extractItems(const std::wstring &aArchivePath,
                             const std::vector<std::wstring> &aItemNames,
                             const std::wstring &aDestDirPath,
                             ArchiveProgressCallback aCallback = nullptr);

    static bool addItemsToArchive(const std::wstring &aArchivePath,
                                  const std::wstring &aVirtualSubDir,
                                  const std::vector<std::wstring> &aSrcPhysicalPaths,
                                  ArchiveProgressCallback aCallback = nullptr);

    // Create New Archive with progress
    static bool createArchive(const std::wstring &aArchivePath,
                              const std::vector<std::wstring> &aSrcFiles,
                              ArchiveFormat aFormat = Format7z,
                              int aCompressionLevel = 9,
                              ArchiveProgressCallback aCallback = nullptr);

private:
    void parseArchiveStructure();
    void addPathToTree(const ArchiveItem &aItem);

    // Internal: execute 7z.exe with stdout pipe for progress parsing
    static bool executeWithProgress(const std::wstring &aCommandLine,
                                    const std::wstring &aWorkingDir,
                                    ArchiveProgressCallback aCallback);

    // Internal: execute process with pumped message loop (no progress)
    static bool executePumped(const std::wstring &aCommandLine, const std::wstring &aWorkingDir);

    // Internal: run command and capture stdout
    static bool executeAndCapture(const std::wstring &aCommandLine,
                                  const std::wstring &aWorkingDir,
                                  std::wstring &aOutput);

    std::wstring mArchivePath;
    ArchiveFormat mFormat;
    bool mIsOpen;
    std::vector<ArchiveItem> mAllItems;
    std::shared_ptr<ArchiveDirectoryNode> mRootNode;

    static std::wstring s7zExePath; // Cached 7z.exe path
};

} // namespace archive
} // namespace fxfile

#endif // __FXFILE_ARCHIVE_ENGINE_H__