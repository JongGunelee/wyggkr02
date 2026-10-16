//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_RECENT_FILE_LIST_H__
#define __FXFILE_RECENT_FILE_LIST_H__ 1
#pragma once

#include "pattern.h"

namespace fxfile
{
namespace base
{
class ConfFileEx;
} // namespace base
} // namespace fxfile

namespace fxfile
{
class RecentFileList : public fxfile::base::Singleton<RecentFileList>
{
    friend class fxfile::base::Singleton<RecentFileList>;

public:
    RecentFileList(void);
    virtual ~RecentFileList(void);

public:
    void load(fxfile::base::ConfFileEx &aConfFile);
    xpr_bool_t load(const xpr_tchar_t *aFilePath);
    void scheduleLoad(const xpr_tchar_t *aFilePath);
    // Resolve the lazy source while it still exists. Configuration-directory
    // migration moves fxfile-main.conf, so deferring this read until the next
    // save would otherwise look at the old path and erase the recent list.
    void prepareForConfigDirMove(void);
    void save(fxfile::base::ConfFileEx &aConfFile) const;

public:
    void               addFile(const xpr_tchar_t *aPath);
    xpr_size_t         getFileCount(void) const;
    const xpr_tchar_t *getFile(xpr_size_t aIndex) const;
    void               clear(void);

protected:
    void ensureLoaded(void);

    typedef std::deque<xpr::string> FileDeque;
    FileDeque mFileDeque;
    xpr_bool_t mLoadPending;
    xpr::string mPendingPath;
};
} // namespace fxfile

#endif // __FXFILE_RECENT_FILE_LIST_H__
