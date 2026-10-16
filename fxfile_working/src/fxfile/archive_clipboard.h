//
// Copyright (c) 2026 FxFile Project. All rights reserved.
// Use of this source code is governed by a GPLv3 license.
//

#ifndef __FXFILE_ARCHIVE_CLIPBOARD_H__
#define __FXFILE_ARCHIVE_CLIPBOARD_H__ 1
#pragma once

#include <string>
#include <vector>

namespace fxfile
{
class ArchiveClipboard
{
public:
    static ArchiveClipboard &instance()
    {
        static ArchiveClipboard sInstance;
        return sInstance;
    }

    void setSelection(const std::wstring &aArchivePath,
                      const std::wstring &aSubDir,
                      const std::vector<std::wstring> &aItemNames)
    {
        mHasData = true;
        mArchivePath = aArchivePath;
        mSubDir = aSubDir;
        mItemNames = aItemNames;
    }

    void clear()
    {
        mHasData = false;
        mArchivePath.clear();
        mSubDir.clear();
        mItemNames.clear();
    }

    bool hasData() const { return mHasData && !mItemNames.empty(); }
    const std::wstring &getArchivePath() const { return mArchivePath; }
    const std::wstring &getSubDir() const { return mSubDir; }
    const std::vector<std::wstring> &getItemNames() const { return mItemNames; }

private:
    ArchiveClipboard() : mHasData(false) {}
    bool mHasData;
    std::wstring mArchivePath;
    std::wstring mSubDir;
    std::vector<std::wstring> mItemNames;
};
} // namespace fxfile

#endif // __FXFILE_ARCHIVE_CLIPBOARD_H__
