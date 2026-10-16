//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_DIRECTORY_ENUMERATION_WORKER_H__
#define __FXFILE_DIRECTORY_ENUMERATION_WORKER_H__ 1
#pragma once

#include <vector>

namespace fxfile
{
class DirectoryEnumerationWorker
{
public:
    struct Item
    {
        LPITEMIDLIST mPidl;
        xpr_ulong_t mShellAttributes;
        DWORD mFileAttributes;
        xpr::string mName;
        xpr_bool_t mHasKnownMetadata;

        Item(void);
    };

    struct Batch
    {
        xpr_uint_t mGeneration;
        xpr_uint_t mOwnerToken;
        xpr_bool_t mComplete;
        xpr_bool_t mSucceeded;
        xpr_uint64_t mEnumerationMilliseconds;
        HANDLE mFlowSemaphore;
        std::vector<Item> mItems;

        Batch(void);
    };

public:
    static xpr_bool_t start(HWND aOwnerHwnd,
                            xpr_uint_t aMessage,
                            LPCITEMIDLIST aFolderFullPidl,
                            xpr_sint_t aListType,
                            xpr_sint_t aAttributes,
                            xpr_uint_t aGeneration,
                            xpr_uint_t aOwnerToken,
                            HANDLE aCancelEvent);
    // Posted messages carry an opaque pointer value.  The UI must claim it
    // before dereferencing; shutdown can then reclaim unclaimed batches even
    // when Windows discards messages for a destroyed pane.
    static Batch *claimBatch(Batch *aBatch);
    static void cleanupOwnerBatches(xpr_uint_t aOwnerToken);
    static void destroyBatch(Batch *aBatch);
};
} // namespace fxfile

#endif // __FXFILE_DIRECTORY_ENUMERATION_WORKER_H__
