//
// Copyright (c) 2001-2012 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "file_op_undo.h"

#include "file_op_thread.h"
#include "bb_info_file.h"           // for BitBucket Info File
#include "operation_path_pairs.h"

#include <mutex>
#include <cstring>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace
{
// FileOpThread permits more than one file operation at a time.  The undo
// journal, however, is a single process-wide stack (mPath/mIndex/mMode).
// Serialize every public journal transaction so two worker/UI threads cannot
// allocate the same undoNNNN.dat file or replace mPath while it is in use.
std::recursive_mutex gFileOpUndoMutex;
int gPendingUndoIndex = -1;
}

xpr_tchar_t *FileOpUndo::mPath  = XPR_NULL;
xpr_sint_t   FileOpUndo::mIndex = -1;
xpr_sint_t   FileOpUndo::mMode  = -1;

#define UNDO_DIR_NAME_FORMAT  XPR_STRING_LITERAL("\\undo%04d")
#define UNDO_FILE_NAME_FORMAT XPR_STRING_LITERAL("%s\\undo%04d.dat")

FileOpUndo::FileOpUndo(void)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);
    validateUndoDir();
}

FileOpUndo::~FileOpUndo(void)
{
}

void FileOpUndo::validateUndoDir(void)
{
    if (XPR_IS_NOT_NULL(mPath))
    {
        if (IsExistFile(mPath) == XPR_TRUE)
            return;

        XPR_SAFE_DELETE_ARRAY(mPath);
    }

    // ex) C:\\Documents and Settings\\flychk\\Local Settings\\Temp\\fxfile\\undo

    mPath = new xpr_tchar_t[XPR_MAX_PATH + 1];

    GetTempPath(XPR_MAX_PATH, mPath);
    if (mPath[_tcslen(mPath)-1] == XPR_STRING_LITERAL('\\'))
        mPath[_tcslen(mPath)-1] = XPR_STRING_LITERAL('\0');

    _tcscat(mPath, XPR_STRING_LITERAL("\\fxfile"));
    if (IsExistFile(mPath) == XPR_FALSE)
        ::CreateDirectory(mPath, XPR_NULL);

    _tcscat(mPath, XPR_STRING_LITERAL("\\undo"));
    if (IsExistFile(mPath) == XPR_FALSE)
        ::CreateDirectory(mPath, XPR_NULL);

    xpr_sint_t i;
    xpr_tchar_t *sFileName = mPath + _tcslen(mPath);
    for (i = 0; ; ++i)
    {
        _stprintf(sFileName, UNDO_DIR_NAME_FORMAT, i);
        if (IsExistFile(mPath) == XPR_FALSE)
        {
            ::CreateDirectory(mPath, XPR_NULL);
            break;
        }
    }

    mIndex = -1;
    mMode = -1;
}

void FileOpUndo::deleteUndoDir(void)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);

    if (XPR_IS_NOT_NULL(mPath))
    {
        mPath[_tcslen(mPath) + 1] = '\0';
        SHSilentDeleteFile(mPath);

        XPR_SAFE_DELETE_ARRAY(mPath);
    }
}

//----------------------------------------------------------------------
// Undo File Structure
//
// +---------------------+----------+
// | Undo Mode           |        4 |
// +---------------------+----------+
// | Source Count        |        4 |
// +---------------------+----------+
// | Source Total Length |        4 |
// +---------------------+----------+
// | Source              |  Str Len |
// +---------------------+----------+
// | ...                 |  Str Len |
// +---------------------+----------+
// | Target Count        |        4 |
// +---------------------+----------+
// | Target Length       |        4 |
// +---------------------+----------+
// | Target              |  Str Len |
// +---------------------+----------+

// copy -> delete
// move -> move
// recycled bin -> restore

xpr_bool_t FileOpUndo::beginAddFile(xpr::FileIo &aFileIo)
{
    validateUndoDir();

    mIndex++;

    xpr_tchar_t sPath[XPR_MAX_PATH + 1] = {0};
    _stprintf(sPath, UNDO_FILE_NAME_FORMAT, mPath, mIndex);

    xpr_rcode_t sRcode;
    xpr_sint_t sOpenMode = xpr::FileIo::OpenModeCreate | xpr::FileIo::OpenModeTruncate | xpr::FileIo::OpenModeWriteOnly;

    sRcode = aFileIo.open(sPath, sOpenMode);
    if (XPR_RCODE_IS_ERROR(sRcode))
    {
        --mIndex;
        return XPR_FALSE;
    }

    return XPR_TRUE;
}

void FileOpUndo::endAddFile(xpr::FileIo &aFileIo)
{
    aFileIo.close();
}

xpr_bool_t FileOpUndo::beginUndoFile(xpr::FileIo &aFileIo, xpr_bool_t aUpdate)
{
    validateUndoDir();

    xpr_tchar_t sPath[XPR_MAX_PATH + 1] = {0};
    _stprintf(sPath, UNDO_FILE_NAME_FORMAT, mPath, mIndex);

    xpr_rcode_t sRcode;

    sRcode = aFileIo.open(sPath, xpr::FileIo::OpenModeReadOnly);
    if (XPR_RCODE_IS_ERROR(sRcode))
        return XPR_FALSE;

    return XPR_TRUE;
}

void FileOpUndo::endUndoFile(xpr::FileIo &aFileIo, xpr_bool_t aUpdate)
{
    aFileIo.close();

    if (XPR_IS_TRUE(aUpdate))
    {
        xpr_tchar_t sPath[XPR_MAX_PATH + 1];
        _stprintf(sPath, UNDO_FILE_NAME_FORMAT, mPath, mIndex);

        sPath[_tcslen(sPath) + 1] = '\0';
        SHSilentDeleteFile(sPath);

        mIndex--;
        if (mIndex < -1)
            mIndex = -1;
    }
}

//
// undo
//   - copy, move, rename: post-processing
//   - delete: pre-processing
//
void FileOpUndo::addOperation(SHFILEOPSTRUCT *aShFileOpStruct)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);

    xpr_sint_t sMode = -1;
    switch (aShFileOpStruct->wFunc)
    {
    case FO_COPY: sMode = MODE_COPY; break;
    case FO_MOVE: sMode = MODE_MOVE; break;
    case FO_DELETE:
        if (aShFileOpStruct->fFlags & FOF_ALLOWUNDO)
            sMode = MODE_TRASH;
        break;
    default: break;
    }

    if (sMode == -1)
        return;

    std::swap(mMode, sMode);

    const xpr_tchar_t *sSource = aShFileOpStruct->pFrom;
    const xpr_tchar_t *sTarget = aShFileOpStruct->pTo;
    HandleToMappings *sHandleToMappings = (HandleToMappings *)aShFileOpStruct->hNameMappings;

    xpr::string sPath;
    const xpr_tchar_t *sParsing = sSource;
    xpr_sint_t nLen = XPR_IS_NOT_NULL(sTarget) ? (xpr_sint_t)_tcslen(sTarget) : 4;

    // ex) C:\\Documents and Settings\\Adminstrator\\Application Data\\fxfile\\undo0000\\undo0000.dat

    xpr_rcode_t sRcode;
    xpr_ssize_t sWritten;
    xpr::FileIo sFileIo;

    if (beginAddFile(sFileIo) == XPR_FALSE)
    {
        std::swap(mMode, sMode);
        return;
    }

    // Undo Mode
    sRcode = sFileIo.write(&mMode, sizeof(xpr_sint_t), &sWritten);

    // Temp Source Path Total Count, Length
    sRcode = sFileIo.write(&mMode, sizeof(xpr_sint_t), &sWritten);
    sRcode = sFileIo.write(&mMode, sizeof(xpr_sint_t), &sWritten);

    HANDLE sFindFile = XPR_NULL;
    WIN32_FIND_DATA sWin32FindData = {0};

    std::deque<xpr_tchar_t *> sSourceDeque;
    std::deque<FILETIME> sTimeDeque;

    xpr_sint_t i;
    xpr_tchar_t *sSplit;
    while (*(sParsing) != '\0')
    {
        sPath.clear();

        if (mMode == MODE_COPY)
        {
            sSplit = (xpr_tchar_t *)_tcsrchr(sParsing, '\\');
            if (XPR_IS_NOT_NULL(sSplit))
            {
                sPath  = sTarget;
                sPath += (nLen <= 3) ? (sSplit+1) : (sSplit);

                if (XPR_IS_NOT_NULL(sHandleToMappings))
                {
                    for (i = 0; i < (xpr_sint_t)sHandleToMappings->mNumberOfMappings; ++i)
                    {
                        if (_tcsicmp(sPath.c_str(), sHandleToMappings->mSHNameMapping[i].pszOldPath) == 0)
                        {
                            sPath = sHandleToMappings->mSHNameMapping[i].pszNewPath;
                            break;
                        }
                    }
                }
            }

            if (sPath.empty() == XPR_FALSE)
            {
                sSourceDeque.push_back(new xpr_tchar_t[sPath.length() + 1]);
                _tcscpy(*(sSourceDeque.end()-1), sPath.c_str());
            }
        }
        else if (mMode == MODE_MOVE)
        {
            sSourceDeque.push_back(new xpr_tchar_t[_tcslen(sParsing) + 1]);
            _tcscpy(*(sSourceDeque.end()-1), sParsing);
        }
        else if (mMode == MODE_TRASH)
        {
            sFindFile = ::FindFirstFile(sParsing, &sWin32FindData);
            if (sFindFile != INVALID_HANDLE_VALUE)
            {
                ::FindClose(sFindFile);

                // Creation Time
                sTimeDeque.push_back(sWin32FindData.ftCreationTime);

                // Long Path
                sSourceDeque.push_back(new xpr_tchar_t[_tcslen(sParsing) + 1]);
                _tcscpy(*(sSourceDeque.end()-1), sParsing);
            }
        }

        sParsing += _tcslen(sParsing) + 1;
    }

    xpr_sint_t sSourceLen = 0;
    xpr_sint_t sSourceCount = (xpr_sint_t)sSourceDeque.size();
    for (i = 0; i < sSourceCount; ++i)
    {
        sRcode = sFileIo.write(sSourceDeque[i], (_tcslen(sSourceDeque[i]) + 1)*sizeof(xpr_tchar_t), &sWritten);
        sSourceLen += (xpr_sint_t)_tcslen(sSourceDeque[i]) + 1;

        delete[] sSourceDeque[i];
    }

    if (mMode == MODE_TRASH)
    {
        for (i = 0; i < sSourceCount; ++i)
            sRcode = sFileIo.write(&sTimeDeque[i], sizeof(FILETIME), &sWritten);
    }

    sSourceDeque.clear();
    sTimeDeque.clear();

    sRcode = sFileIo.seekFromBegin(sizeof(xpr_sint_t));

    // Source Path Count, Total Length
    sRcode = sFileIo.write(&sSourceCount, sizeof(xpr_sint_t), &sWritten);
    sRcode = sFileIo.write(&sSourceLen,   sizeof(xpr_sint_t), &sWritten);

    if (XPR_IS_NOT_NULL(sTarget))
    {
        sRcode = sFileIo.seekToEnd();

        // Temp Target Path Total Count, Length
        xpr_sint64_t sPos = sFileIo.tell();
        sRcode = sFileIo.write(&mMode, sizeof(xpr_sint_t), &sWritten);
        sRcode = sFileIo.write(&mMode, sizeof(xpr_sint_t), &sWritten);

        // Targer Path
        xpr_sint_t sTargetLen = 0;
        xpr_sint_t sTargetCount = 0;
        sParsing = sTarget;
        while (*(sParsing) != '\0')
        {
            sRcode = sFileIo.write(sParsing, (_tcslen(sParsing)+1) * sizeof(xpr_tchar_t), &sWritten);
            sTargetLen += (xpr_sint_t)_tcslen(sParsing) + 1;

            sParsing += _tcslen(sParsing) + 1;
            sTargetCount++;
        }

        sRcode = sFileIo.seekFromBegin(sPos);

        // Target Path Count
        sRcode = sFileIo.write(&sTargetCount, sizeof(xpr_sint_t), &sWritten);

        // Target Path Total Length
        sRcode = sFileIo.write(&sTargetLen,   sizeof(xpr_sint_t), &sWritten);
    }

    endAddFile(sFileIo);
}

void FileOpUndo::addRename(const xpr::string &aSource, const xpr::string &aTarget)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);
    addRename(aSource.c_str(), aTarget.c_str());
}

void FileOpUndo::addRename(const xpr_tchar_t *aSource, const xpr_tchar_t *aTarget)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);

    xpr::FileIo sFileIo;
    if (beginAddFile(sFileIo) == XPR_FALSE)
        return;

    xpr_rcode_t sRcode;
    xpr_ssize_t sWritten;

    mMode = MODE_RENAME;

    // Undo Mode
    sRcode = sFileIo.write(&mMode, sizeof(xpr_sint_t), &sWritten);

    // Source Path Count
    xpr_sint_t sSourceCount = 1;
    sRcode = sFileIo.write(&sSourceCount, sizeof(xpr_sint_t), &sWritten);

    // Source Path Length
    xpr_sint_t sSourceLen = (xpr_sint_t)_tcslen(aSource) + 1;
    sRcode = sFileIo.write(&sSourceLen, sizeof(xpr_sint_t), &sWritten);

    // Source Path
    sRcode = sFileIo.write(aSource, sSourceLen * sizeof(xpr_tchar_t), &sWritten);

    // Target Path Count
    xpr_sint_t sTargetCount = 1;
    sRcode = sFileIo.write(&sTargetCount, sizeof(xpr_sint_t), &sWritten);

    // Target Path Length
    xpr_sint_t sTargetLen = (xpr_sint_t)_tcslen(aTarget) + 1;
    sRcode = sFileIo.write(&sTargetLen, sizeof(xpr_sint_t), &sWritten);

    // Target Path
    sRcode = sFileIo.write(aTarget, sTargetLen * sizeof(xpr_tchar_t), &sWritten);

    endAddFile(sFileIo);
}

xpr_bool_t FileOpUndo::isUndo(void)
{
    // Menu updates must never wait for a worker's filesystem/journal transaction.
    std::unique_lock<std::recursive_mutex> sGuard(gFileOpUndoMutex, std::try_to_lock);
    return sGuard.owns_lock() && XPR_IS_NOT_NULL(mPath) && (mIndex > -1) &&
        gPendingUndoIndex < 0 && FileOpThread::getRefCount() == 0;
}

xpr_sint_t FileOpUndo::getMode(void)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);
    return mMode;
}

xpr_sint_t FileOpUndo::getCount(void)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);
    return mIndex + 1;
}

void FileOpUndo::start(void)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);
    if (gPendingUndoIndex >= 0 || FileOpThread::getRefCount() != 0) return;

    xpr::FileIo sFileIo;
    if (beginUndoFile(sFileIo) == XPR_FALSE)
        return;

    xpr_rcode_t sRcode;
    xpr_ssize_t sRead;

    sRcode = sFileIo.read(&mMode, sizeof(xpr_sint_t), &sRead);

    SHFILEOPSTRUCT *sShFileOpStruct = XPR_NULL;
    std::vector<FILETIME> sRecycleTimes;
    switch (mMode)
    {
    case MODE_COPY:   sShFileOpStruct = startCopy(sFileIo);       break;
    case MODE_MOVE:   sShFileOpStruct = startMove(sFileIo);       break;
    case MODE_RENAME: sShFileOpStruct = startRename(sFileIo);     break;
    case MODE_TRASH:  sShFileOpStruct = startRecycleBin(sFileIo, sRecycleTimes); break;
    }

    if (XPR_IS_NOT_NULL(sShFileOpStruct))
    {
        FileOpThread *sFileOpThread = new FileOpThread;
        sFileOpThread->setUndo(XPR_FALSE);
        sFileOpThread->setRecycleUndoTimes(sRecycleTimes);
        gPendingUndoIndex = mIndex;
        sFileOpThread->setUndoJournalIndex(mIndex);

        sFileOpThread->start(sShFileOpStruct);
    }

    // A missing/unfinished recycle entry must not silently consume the journal.
    endUndoFile(sFileIo, XPR_FALSE);

    // Update Mode
    if (beginUndoFile(sFileIo, XPR_FALSE) == XPR_TRUE)
        sRcode = sFileIo.read(&mMode, sizeof(xpr_sint_t), &sRead);
    else
        mMode = -1;

    endUndoFile(sFileIo, XPR_FALSE);
}

void FileOpUndo::finishUndo(xpr_sint_t aIndex, xpr_bool_t aSucceeded)
{
    std::lock_guard<std::recursive_mutex> sGuard(gFileOpUndoMutex);
    if (gPendingUndoIndex != aIndex) return;
    if (XPR_IS_TRUE(aSucceeded) && mIndex == aIndex)
    {
        FileOpUndo sUndo;
        xpr::FileIo sFile;
        if (sUndo.beginUndoFile(sFile) == XPR_TRUE) sUndo.endUndoFile(sFile);
        xpr_ssize_t sRead = 0;
        if (sUndo.beginUndoFile(sFile, XPR_FALSE) == XPR_TRUE)
        {
            sFile.read(&mMode, sizeof(mMode), &sRead);
            sUndo.endUndoFile(sFile, XPR_FALSE);
        }
        else mMode = -1;
    }
    // Cancelled/failed restoration retains the journal for retry.
    gPendingUndoIndex = -1;
}

SHFILEOPSTRUCT *FileOpUndo::startCopy(xpr::FileIo &aFileIo)
{
    xpr_rcode_t sRcode;
    xpr_ssize_t sRead;

    xpr_sint_t sSourceCount = 0;
    sRcode = aFileIo.read(&sSourceCount, sizeof(xpr_sint_t), &sRead);

    xpr_sint_t sSourceLen = 0;
    sRcode = aFileIo.read(&sSourceLen, sizeof(xpr_sint_t), &sRead);

    if (sSourceCount <= 0 || sSourceLen <= 0)
        return XPR_NULL;

    xpr_tchar_t *sSource2 = new xpr_tchar_t[sSourceLen + XPR_MAX_PATH + 1];
    sSource2[0] = '\0';

    sRcode = aFileIo.read(sSource2, sSourceLen * sizeof(xpr_tchar_t), &sRead);
    sSource2[sSourceLen  ] = '\0';
    sSource2[sSourceLen+1] = '\0';

    SHFILEOPSTRUCT *sShFileOpStruct = XPR_NULL;
    if (sSource2[0] != '\0')
    {
        sShFileOpStruct = new SHFILEOPSTRUCT;
        memset(sShFileOpStruct, 0, sizeof(SHFILEOPSTRUCT));
        sShFileOpStruct->hwnd   = AfxGetMainWnd()->GetSafeHwnd();
        sShFileOpStruct->wFunc  = FO_DELETE;
        sShFileOpStruct->fFlags = 0;
        sShFileOpStruct->pFrom  = sSource2;
        sShFileOpStruct->pTo    = XPR_NULL;
    }
    else
    {
        XPR_SAFE_DELETE_ARRAY(sSource2);
    }

    return sShFileOpStruct;
}

SHFILEOPSTRUCT *FileOpUndo::startMove(xpr::FileIo &aFileIo)
{
    xpr_rcode_t sRcode;
    xpr_ssize_t sRead;

    xpr_sint_t sSourceCount = 0;
    sRcode = aFileIo.read(&sSourceCount, sizeof(xpr_sint_t), &sRead);

    xpr_sint_t sSourceLen = 0;
    sRcode = aFileIo.read(&sSourceLen, sizeof(xpr_sint_t), &sRead);

    if (sSourceCount <= 0 || sSourceLen <= 0)
        return XPR_NULL;

    xpr_tchar_t *sSourceTemp = new xpr_tchar_t[sSourceLen + XPR_MAX_PATH + 1];
    sRcode = aFileIo.read(sSourceTemp, sSourceLen * sizeof(xpr_tchar_t), &sRead);
    sSourceTemp[sSourceLen  ] = '\0';
    sSourceTemp[sSourceLen+1] = '\0';

    xpr_sint_t sTargetCount = 0;
    sRcode = aFileIo.read(&sTargetCount, sizeof(xpr_sint_t), &sRead);
    if (sSourceCount != sTargetCount)
        sTargetCount = 1;

    xpr_sint_t sTargetLen = 0;
    sRcode = aFileIo.read(&sTargetLen, sizeof(xpr_sint_t), &sRead);

    xpr_tchar_t *sTargetTemp = new xpr_tchar_t[sTargetLen + XPR_MAX_PATH + 1];
    sRcode = aFileIo.read(sTargetTemp, sTargetLen * sizeof(xpr_tchar_t), &sRead);

    xpr_tchar_t *sSource2 = new xpr_tchar_t[(XPR_MAX_PATH + 1) * sSourceCount];
    xpr_tchar_t *sTarget2 = new xpr_tchar_t[(XPR_MAX_PATH + 1) * sSourceCount];
    sSource2[0] = '\0';
    sTarget2[0] = '\0';

    xpr_tchar_t *sSplit = XPR_NULL;
    xpr_tchar_t *sSourceEnum = sSource2;
    xpr_tchar_t *sTargetEnum = sTarget2;
    xpr_tchar_t *sSourceTempEnum = sSourceTemp;
    xpr_tchar_t *sTargetTempEnum = sTargetTemp;

    if (sTargetCount == 1 && XPR_IS_NOT_NULL(sTargetTemp))
    {
        if (sTargetTemp[_tcslen(sTargetTemp) - 1] == '\\')
            sTargetTemp[_tcslen(sTargetTemp) - 1] = '\0';
    }

    xpr_sint_t i;
    xpr_sint_t sSourceTempLen = 0;
    xpr_tchar_t *sSourceDir = XPR_NULL;
    xpr_tchar_t *sTargetDir = XPR_NULL;
    xpr_tchar_t *sFileName  = XPR_NULL;
    for (i = 0; i < sSourceCount; ++i)
    {
        sSourceTempLen = (xpr_sint_t)_tcslen(sSourceTempEnum);
        sSplit = _tcsrchr(sSourceTempEnum, '\\');
        if (XPR_IS_NOT_NULL(sSplit))
        {
            *sSplit = '\0';
            sSplit++;
        }
        else
        {
            continue;
        }

        sSourceDir = sTargetTemp;
        sTargetDir = sSourceTemp;
        sFileName  = sSplit;

        // consist of source
        if (sTargetCount == 1)
            _tcscpy(sSourceEnum, sSourceDir);
        else
        {
            _tcscpy(sSourceEnum, sTargetTempEnum);
            if (sSourceEnum[_tcslen(sSourceEnum) - 1] == '\\')
                sSourceEnum[_tcslen(sSourceEnum) - 1] = '\0';
        }

        _tcscat(sSourceEnum, XPR_STRING_LITERAL("\\"));
        _tcscat(sSourceEnum, sFileName);

        // consist of target
        _tcscpy(sTargetEnum, sTargetDir);
        _tcscat(sTargetEnum, XPR_STRING_LITERAL("\\"));
        _tcscat(sTargetEnum, sFileName);

        sSourceEnum += _tcslen(sSourceEnum) + 1;
        sTargetEnum += _tcslen(sTargetEnum) + 1;

        sSourceTempEnum += sSourceTempLen + 1;
        if (sTargetCount > 1)
            sTargetTempEnum += _tcslen(sTargetTempEnum) + 1;
    }

    // last double null
    sSourceEnum[0] = '\0';
    sTargetEnum[0] = '\0';

    // free temp buffer
    XPR_SAFE_DELETE_ARRAY(sSourceTemp);
    XPR_SAFE_DELETE_ARRAY(sTargetTemp);

    SHFILEOPSTRUCT *sShFileOpStruct = XPR_NULL;
    if (sSource2[0] != '\0' && sTarget2[0] != '\0')
    {
        sShFileOpStruct = new SHFILEOPSTRUCT;
        memset(sShFileOpStruct, 0, sizeof(SHFILEOPSTRUCT));
        sShFileOpStruct->hwnd   = AfxGetMainWnd()->GetSafeHwnd();
        sShFileOpStruct->wFunc  = FO_MOVE;
        sShFileOpStruct->fFlags = FOF_MULTIDESTFILES;
        sShFileOpStruct->pFrom  = sSource2;
        sShFileOpStruct->pTo    = sTarget2;
    }
    else
    {
        XPR_SAFE_DELETE_ARRAY(sSource2);
        XPR_SAFE_DELETE_ARRAY(sTarget2);
    }

    return sShFileOpStruct;
}

SHFILEOPSTRUCT *FileOpUndo::startRename(xpr::FileIo &aFileIo)
{
    xpr_rcode_t sRcode;
    xpr_ssize_t sRead;

    xpr_sint_t sSourceCount = 0;
    sRcode = aFileIo.read(&sSourceCount, sizeof(xpr_sint_t), &sRead);

    xpr_sint_t sSourceLen = 0;
    sRcode = aFileIo.read(&sSourceLen, sizeof(xpr_sint_t), &sRead);

    if (sSourceCount <= 0 || sSourceLen <= 0)
        return XPR_NULL;

    xpr_tchar_t *sSource2 = new xpr_tchar_t[sSourceLen + XPR_MAX_PATH + 1];
    sSource2[0] = '\0';

    sRcode = aFileIo.read(sSource2, sSourceLen * sizeof(xpr_tchar_t), &sRead);

    xpr_sint_t sTargetCount = 0;
    sRcode = aFileIo.read(&sTargetCount, sizeof(xpr_sint_t), &sRead);

    xpr_sint_t sTargetLen = 0;
    sRcode = aFileIo.read(&sTargetLen, sizeof(xpr_sint_t), &sRead);

    xpr_tchar_t *sTarget2 = new xpr_tchar_t[sTargetLen + XPR_MAX_PATH + 1];
    sTarget2[0] = '\0';

    sRcode = aFileIo.read(sTarget2, sTargetLen * sizeof(xpr_tchar_t), &sRead);

    sSource2[_tcslen(sSource2) + 1] = '\0';
    sTarget2[_tcslen(sTarget2) + 1] = '\0';

    SHFILEOPSTRUCT *sShFileOpStruct = XPR_NULL;
    if (sSource2[0] != '\0' && sTarget2[0] != '\0')
    {
        sShFileOpStruct = new SHFILEOPSTRUCT;
        memset(sShFileOpStruct, 0, sizeof(SHFILEOPSTRUCT));
        sShFileOpStruct->hwnd   = AfxGetMainWnd()->GetSafeHwnd();
        sShFileOpStruct->wFunc  = FO_RENAME;
        sShFileOpStruct->fFlags = 0;
        sShFileOpStruct->pFrom  = sTarget2;
        sShFileOpStruct->pTo    = sSource2;
    }
    else
    {
        XPR_SAFE_DELETE_ARRAY(sSource2);
        XPR_SAFE_DELETE_ARRAY(sTarget2);
    }

    return sShFileOpStruct;
}

SHFILEOPSTRUCT *FileOpUndo::startRecycleBin(xpr::FileIo &aFileIo, std::vector<FILETIME> &aTimes)
{
    xpr_rcode_t sRcode;
    xpr_ssize_t sRead;

    xpr_sint_t sSourceCount = 0;
    sRcode = aFileIo.read(&sSourceCount, sizeof(xpr_sint_t), &sRead);

    xpr_sint_t sSourceLen = 0;
    sRcode = aFileIo.read(&sSourceLen, sizeof(xpr_sint_t), &sRead);

    if (sSourceCount <= 0 || sSourceCount > 1000000 || sSourceLen <= 0 || sSourceLen > 16 * 1024 * 1024)
        return XPR_NULL;

    xpr_tchar_t *sSource = new xpr_tchar_t[sSourceLen + XPR_MAX_PATH + 1];
    sRcode = aFileIo.read(sSource, sSourceLen * sizeof(xpr_tchar_t), &sRead);
    sSource[sSourceLen  ] = '\0';
    sSource[sSourceLen+1] = '\0';

    if (XPR_RCODE_IS_ERROR(sRcode) || sRead != sSourceLen * sizeof(xpr_tchar_t))
    { delete[] sSource; return XPR_NULL; }
    aTimes.resize(sSourceCount);
    sRcode = aFileIo.read(&aTimes[0], sizeof(FILETIME) * sSourceCount, &sRead);
    if (XPR_RCODE_IS_ERROR(sRcode) || sRead != sizeof(FILETIME) * sSourceCount)
    { delete[] sSource; aTimes.clear(); return XPR_NULL; }
    xpr_sint_t sEntries = 0;
    for (xpr_sint_t pos = 0; pos < sSourceLen; ++pos)
        if (!sSource[pos]) ++sEntries;
    if (sEntries != sSourceCount || !sSource[0] || sSource[sSourceLen - 1])
    { delete[] sSource; aTimes.clear(); return XPR_NULL; }
    SHFILEOPSTRUCT *sOperation = new SHFILEOPSTRUCT();
    sOperation->hwnd = AfxGetMainWnd()->GetSafeHwnd();
    sOperation->wFunc = FO_MOVE;
    sOperation->fFlags = FOF_MULTIDESTFILES;
    sOperation->pFrom = sSource;
    return sOperation;
}

xpr_bool_t FileOpUndo::resolveRecycleUndo(SHFILEOPSTRUCT *aOperation, const std::vector<FILETIME> &aTimes)
{
    if (!aOperation || !aOperation->pFrom || aTimes.empty()) return XPR_FALSE;
    BBInfoFile sBin;
    xpr_tchar_t sDrive = 0;
    OperationPathPairs<xpr_tchar_t> sPaths;
    const xpr_tchar_t *sOriginal = aOperation->pFrom;
    for (size_t i = 0; i < aTimes.size() && *sOriginal; ++i)
    {
        const xpr_tchar_t *sNext = sOriginal + _tcslen(sOriginal) + 1;
        if (::GetFileAttributes(sOriginal) == INVALID_FILE_ATTRIBUTES)
        {
            DWORD sError = ::GetLastError();
            if (sError == ERROR_FILE_NOT_FOUND || sError == ERROR_PATH_NOT_FOUND)
            {
                if (sDrive != sOriginal[0])
                {
                    if (!sBin.open(sOriginal[0])) { sOriginal = sNext; continue; }
                    sDrive = sOriginal[0];
                }
                xpr_tchar_t sRecycled[XPR_MAX_PATH * 2 + 1] = {0};
                if (sBin.findOriginalPath(sOriginal, aTimes[i], sRecycled))
                {
                    sPaths.add(sRecycled, sOriginal);
                }
            }
        }
        sOriginal = sNext;
    }
    if (sPaths.sources.empty()) return XPR_FALSE;
    xpr_tchar_t *sSources = new xpr_tchar_t[sPaths.sources.size() + 1];
    xpr_tchar_t *sTargets = new xpr_tchar_t[sPaths.targets.size() + 1];
    std::memcpy(sSources, sPaths.sources.c_str(), (sPaths.sources.size() + 1) * sizeof(xpr_tchar_t));
    std::memcpy(sTargets, sPaths.targets.c_str(), (sPaths.targets.size() + 1) * sizeof(xpr_tchar_t));
    delete[] aOperation->pFrom;
    aOperation->pFrom = sSources;
    aOperation->pTo = sTargets;
    return XPR_TRUE;
}
} // namespace fxfile
