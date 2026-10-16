//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_ITEM_DATA_H__
#define __FXFILE_ITEM_DATA_H__ 1
#pragma once

namespace fxfile
{
enum
{
    IDT_ARCHIVE_ITEM = 200
};

typedef struct LVITEMDATA
{
    xpr::string    mName;
    xpr_uint_t     mSignature;
    DWORD          mItemType;
    LPSHELLFOLDER  mShellFolder;
    LPSHELLFOLDER2 mShellFolder2;
    LPITEMIDLIST   mPidl;
    xpr_ulong_t    mShellAttributes;
    DWORD          mFileAttributes;
    xpr_uint_t     mThumbImageId;
    xpr_sint_t     mCachedIconIndex;
    xpr_bool_t     mIconResolved;
    xpr_bool_t     mIconRequestIssued;
    xpr_uint_t     mCachedOverlayState;
    xpr_bool_t     mOverlayResolved;
    xpr_bool_t     mOverlayRequestIssued;

    // Archive browsing metadata
    uint64_t       mArchiveFileSize;
    bool           mArchiveIsDir;

    // Task 064b: Cached filesystem path resolved once via GetName(SHGDN_FORPARSING).
    // Avoids repeated synchronous IShellFolder::GetDisplayName calls on every
    // LVN_GETDISPINFO (LVIF_IMAGE) paint callback which can stall the UI thread
    // when a shell extension intercepts IShellFolder for that item.
    xpr_tchar_t    mCachedPath[XPR_MAX_PATH + 1];
    xpr_bool_t     mPathResolved;

    LVITEMDATA()
        : mSignature(0)
        , mItemType(0)
        , mShellFolder(XPR_NULL)
        , mShellFolder2(XPR_NULL)
        , mPidl(XPR_NULL)
        , mShellAttributes(0)
        , mFileAttributes(0)
        , mThumbImageId(0)
        , mCachedIconIndex(-1)
        , mIconResolved(XPR_FALSE)
        , mIconRequestIssued(XPR_FALSE)
        , mCachedOverlayState(0)
        , mOverlayResolved(XPR_FALSE)
        , mOverlayRequestIssued(XPR_FALSE)
        , mArchiveFileSize(0)
        , mArchiveIsDir(false)
        , mPathResolved(XPR_FALSE)
    {
        mCachedPath[0] = XPR_STRING_LITERAL('\0');
    }
} LVITEMDATA, *LPLVITEMDATA;

typedef struct TVITEMDATA
{
    LPSHELLFOLDER mShellFolder;
    LPITEMIDLIST  mPidl;
    LPITEMIDLIST  mFullPidl;
    xpr_ulong_t   mShellAttributes;
    DWORD         mFileAttributes;
    //WatchId     mWatchId;
    //AdvWatchId  mAdvWatchId;
    xpr_bool_t    mExpandPartial;
} TVITEMDATA, *LPTVITEMDATA;

typedef struct ABITEMDATA
{
    LPSHELLFOLDER mShellFolder;
    LPITEMIDLIST  mPidl;
    LPITEMIDLIST  mFullPidl;
    xpr_ulong_t   mShellAttributes;
    xpr_uint_t    mLevel;
    xpr_uint_t    mType;
} ABITEMDATA, *LPABITEMDATA;
} // namespace fxfile

#endif // __FXFILE_ITEM_DATA_H__
