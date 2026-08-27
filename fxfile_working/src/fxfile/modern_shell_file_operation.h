//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_MODERN_SHELL_FILE_OPERATION_H__
#define __FXFILE_MODERN_SHELL_FILE_OPERATION_H__ 1
#pragma once

namespace fxfile
{
class ModernShellFileOperation
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
    // Executes path-based copy/move/delete operations with IFileOperation.
    // Collision-renaming and legacy multi-destination semantics remain with
    // SHFileOperation because callers depend on its name-mapping handle.
    static Result tryExecute(SHFILEOPSTRUCT *aFileOperation, HRESULT *aError);
};
} // namespace fxfile

#endif // __FXFILE_MODERN_SHELL_FILE_OPERATION_H__
