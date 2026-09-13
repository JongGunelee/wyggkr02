//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_ADAPTIVE_FILE_OPERATION_H__
#define __FXFILE_ADAPTIVE_FILE_OPERATION_H__ 1
#pragma once

namespace fxfile
{
class AdaptiveFileOperation
{
public:
    enum Result
    {
        ResultNotApplicable,
        ResultSucceeded,
        ResultCancelled,
        ResultFailed,
    };

    enum Engine
    {
        EngineNone,
        EngineCopyFile2,
        EngineRobocopy,
        EngineDirectPermanentDelete,
    };

    enum SelectionReason
    {
        ReasonNotApplicable,
        ReasonLocalCollisionFreeCopy,
        ReasonCrossVolumeVerifiedMove,
        ReasonUserApprovedLargeFolderRobocopy,
        ReasonLocalUnprotectedPermanentDelete,
    };

    struct ExecutionInfo
    {
        Engine mEngine;
        SelectionReason mReason;
        xpr_size_t mFileCount;
        xpr_size_t mDirectoryCount;
        xpr_uint64_t mTotalBytes;
        xpr_uint_t mWorkerCount;
        xpr_bool_t mUnbuffered;
        xpr_uint64_t mPlanningMilliseconds;

        ExecutionInfo(void);
    };

public:
    // Executes only the conservative high-performance subset.  The caller
    // must fall back to the Shell engine when ResultNotApplicable is returned.
    static Result tryExecute(SHFILEOPSTRUCT *aFileOperation, DWORD *aError,
                             ExecutionInfo *aExecutionInfo = NULL);
};
} // namespace fxfile

#endif // __FXFILE_ADAPTIVE_FILE_OPERATION_H__
