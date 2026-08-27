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

public:
    // Executes only the conservative high-performance subset.  The caller
    // must fall back to the Shell engine when ResultNotApplicable is returned.
    static Result tryExecute(SHFILEOPSTRUCT *aFileOperation, DWORD *aError);
};
} // namespace fxfile

#endif // __FXFILE_ADAPTIVE_FILE_OPERATION_H__
