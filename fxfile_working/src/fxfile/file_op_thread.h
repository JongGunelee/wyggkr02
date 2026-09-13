//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#ifndef __FXFILE_FILE_OP_THREAD_H__
#define __FXFILE_FILE_OP_THREAD_H__ 1
#pragma once

#include <vector>
#include "adaptive_file_operation.h"
#include "modern_shell_file_operation.h"

namespace fxfile
{
typedef struct HandleToMappings
{
    xpr_uint_t      mNumberOfMappings;  // number of mappings in array
    LPSHNAMEMAPPING mSHNameMapping;    // pointer to array of mappings
} HandleToMappings;

typedef struct PasteSelItems
{
    const xpr_tchar_t *mTarget;
    const xpr_tchar_t *mSource;
    HandleToMappings  *mHandleToMappings;
} PasteSelItems;

class FileOpThread : public CWnd
{
public:
    FileOpThread(void);

public:
    static void setCompleteFlash(xpr_bool_t aCompleteFlash);

public:
    void start(
        SHFILEOPSTRUCT *aShFileOpStruct,
        xpr_bool_t      aSelectItem = XPR_FALSE,
        HWND            aNotifyHwnd = XPR_NULL,
        xpr_uint_t      aMsg = 0);
    void start(
        xpr_bool_t      aCopy,
        xpr_tchar_t    *aSource,
        xpr_sint_t      aSourceCount,
        xpr_tchar_t    *aTarget,
        xpr_bool_t      aDuplicate = XPR_FALSE,
        xpr_bool_t      aSelectedItem = XPR_FALSE,
        HWND            aNotifyHwnd = XPR_NULL,
        xpr_uint_t      aMsg = 0);

    void setUndo(xpr_bool_t aUndo);

public:
    static xpr_sint_t getRefCount(void) { return mRefCount; }
    xpr_sint_t getRefIndex(void) { return mRefIndex; }

protected:
    static unsigned __stdcall FileOpProc(LPVOID lpParam);
    void OnFileOp(void);
    void captureOperationSources(void);
    void captureOperationResult(void);
    void reconcileOperationResult(void);

    struct SourceSnapshot
    {
        xpr::string mPath;
        xpr::string mExpectedTargetPath;
        xpr_bool_t  mDirectory;
        xpr_bool_t  mExistedBefore;
    };

    // Captured by the file-operation worker.  UI handlers must never probe
    // source/target paths here: disconnected shares, cloud placeholders and
    // anti-virus filters may block GetFileAttributes for seconds.
    struct ResultSnapshot
    {
        enum Outcome
        {
            OutcomeSucceeded,
            OutcomeFailed,
            OutcomeCancelled,
            OutcomeUnprocessed,
        };

        xpr::string mSourcePath;
        xpr::string mTargetPath;
        xpr_bool_t  mDirectory;
        xpr_bool_t  mSourceGone;
        xpr_bool_t  mTargetExists;
        Outcome     mOutcome;
        DWORD       mError;
    };

protected:
    SHFILEOPSTRUCT *mShFileOpStruct;
    xpr_bool_t      mSelectItem;
    HWND            mNotifyHwnd;
    xpr_uint_t      mMsg;

    HANDLE          mThread;
    unsigned        mThreadId;
    HANDLE          mStopEvent;
    xpr_bool_t      mUndo;
    xpr_bool_t      mOperationSucceeded;
    DWORD           mOperationError;
    xpr_uint64_t    mOperationId;
    xpr_uint64_t    mAcceptedTick;
    xpr_uint64_t    mWorkerStartedTick;
    xpr_uint64_t    mPreparationFinishedTick;
    xpr_uint64_t    mExecutionFinishedTick;
    xpr_uint64_t    mVerificationFinishedTick;
    xpr::string     mEngineName;
    xpr::string     mEngineReason;
    AdaptiveFileOperation::ExecutionInfo mAdaptiveExecutionInfo;
    ModernShellFileOperation::ExecutionInfo mModernExecutionInfo;
    std::vector<SourceSnapshot> mSourceSnapshots;
    std::vector<ResultSnapshot> mResultSnapshots;

    static xpr_bool_t mCompleteFlash;
    static volatile LONG mNextOperationId;

protected:
    static xpr_sint_t mRefCount;
    xpr_sint_t mRefIndex;

public:
    virtual xpr_bool_t DestroyWindow(void);

protected:
    virtual void PostNcDestroy(void);

protected:
    virtual ~FileOpThread(void);

protected:
    DECLARE_MESSAGE_MAP()
    afx_msg void OnTimer(UINT_PTR nIDEvent);
    afx_msg void OnClose(void);
    afx_msg LRESULT OnPostEnd(WPARAM, LPARAM);
};
} // namespace fxfile

#endif // __FXFILE_FILE_OP_THREAD_H__
