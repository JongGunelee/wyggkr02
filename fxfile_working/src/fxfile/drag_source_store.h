//
// Copyright (c) 2026 FxFile Project. All rights reserved.
// Use of this source code is governed by a GPLv3 license.
//

#ifndef __FXFILE_DRAG_SOURCE_STORE_H__
#define __FXFILE_DRAG_SOURCE_STORE_H__ 1
#pragma once

#include <string>
#include <vector>

namespace fxfile
{

class DragSourceStore
{
public:
    static DragSourceStore &instance()
    {
        static DragSourceStore sInstance;
        return sInstance;
    }

    void setFiles(const std::vector<std::wstring> &aFiles)
    {
        mFiles = aFiles;
        mActive = !mFiles.empty();
    }

    void clear()
    {
        mFiles.clear();
        mActive = false;
    }

    bool hasFiles() const { return mActive && !mFiles.empty(); }
    const std::vector<std::wstring> &getFiles() const { return mFiles; }

private:
    DragSourceStore() : mActive(false) {}
    bool mActive;
    std::vector<std::wstring> mFiles;
};

} // namespace fxfile

#endif // __FXFILE_DRAG_SOURCE_STORE_H__
