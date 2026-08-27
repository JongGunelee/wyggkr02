//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "conf_dir.h"
#include "conf_file_ex.h"
#include "env_path.h"
#include "path.h"
#include <xpr_file_sys.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace fxfile;
using namespace fxfile::base;

namespace fxfile
{
namespace
{
#define CFG_PATH_FILENAME             XPR_STRING_LITERAL("fxfile.ini")
#define CFG_PATH_FILENAME_OLD         XPR_STRING_LITERAL(".fxfile")

const xpr_tchar_t kLocalFxFilePath   [] = XPR_STRING_LITERAL("%fxfile%\\")CFG_PATH_FILENAME;
const xpr_tchar_t kLocalFxFilePathOld[] = XPR_STRING_LITERAL("%fxfile%\\")CFG_PATH_FILENAME_OLD;
const xpr_tchar_t kFxFilePath        [] = XPR_STRING_LITERAL("%AppData%\\fxfile\\")CFG_PATH_FILENAME_OLD;

const xpr_tchar_t kProgramConfDir    [] = XPR_STRING_LITERAL("%fxfile%\\fxfile");
const xpr_tchar_t kDefConfDir        [] = XPR_STRING_LITERAL("%AppData%\\fxfile\\conf");

const xpr_tchar_t kFxFileSection [] = XPR_STRING_LITERAL(".fxfile");
const xpr_tchar_t kConfHomeKey   [] = XPR_STRING_LITERAL("conf_home");
} // namespace anonymous

ConfDir::ConfDir(void)
    : mReadOnly(XPR_FALSE)
{
}

ConfDir::~ConfDir(void)
{
    clear();
}

void ConfDir::clear(void)
{
}

const xpr_tchar_t *ConfDir::getConfDir(void) const
{
    return mConfDir.c_str();
}

const xpr_tchar_t *ConfDir::getOldConfDir(void) const
{
    return mOldConfDir.c_str();
}

void ConfDir::setConfDir(const xpr_tchar_t *aConfDir, xpr_bool_t aReadOnly)
{
    XPR_ASSERT(aConfDir != XPR_NULL);

    mConfDir  = aConfDir;
    mReadOnly = aReadOnly;
}

xpr_bool_t ConfDir::getDir(const xpr_tchar_t *aConfDir, xpr_tchar_t *aDir, xpr_size_t aMaxLen) const
{
    if (XPR_IS_NULL(aDir) || aMaxLen <= 0)
        return XPR_FALSE;

    xpr::string sDir;
    GetEnvRealPath(aConfDir, sDir);

    if (sDir.empty() == XPR_TRUE)
        return XPR_FALSE;

    if (sDir.length() > aMaxLen)
        return XPR_FALSE;

    _tcscpy(aDir, sDir.c_str());

    return XPR_TRUE;
}

xpr_bool_t ConfDir::getDir(xpr_tchar_t *aDir, xpr_size_t aMaxLen) const
{
    const xpr_tchar_t *sConfDir;

    sConfDir = getConfDir();
    XPR_ASSERT(sConfDir != XPR_NULL);

    return getDir(sConfDir, aDir, aMaxLen);
}

xpr_bool_t ConfDir::getPath(xpr_sint_t aType, const xpr_tchar_t *aDir, xpr_tchar_t *aPath, xpr_size_t aMaxLen) const
{
    const xpr_tchar_t *sFileName = ConfDir::getFileName(aType);
    if (XPR_IS_NULL(sFileName))
        return XPR_FALSE;

    xpr::string sPath(aDir);
    sPath += XPR_STRING_LITERAL('\\');
    sPath += sFileName;

    if (sPath.length() > aMaxLen)
        return XPR_FALSE;

    _tcscpy(aPath, sPath.c_str());

    return XPR_TRUE;
}

xpr_bool_t ConfDir::getPath(xpr_sint_t aType, xpr_tchar_t *aPath, xpr_size_t aMaxLen) const
{
    xpr_tchar_t sDir[XPR_MAX_PATH + 1] = {0};
    getDir(sDir, XPR_MAX_PATH + 1);

    return getPath(aType, sDir, aPath, aMaxLen);
}

xpr_bool_t ConfDir::getLoadPath(xpr_sint_t aType, xpr_tchar_t *aPath, xpr_size_t aMaxLen) const
{
    return getPath(aType, aPath, aMaxLen);
}

xpr_bool_t ConfDir::getSavePath(xpr_sint_t aType, xpr_tchar_t *aPath, xpr_size_t aMaxLen) const
{
    if (getPath(aType, aPath, aMaxLen) == XPR_FALSE)
        return XPR_FALSE;

    if (CreateDirectoryLevel(aPath, 0, XPR_FALSE) == XPR_FALSE)
        return XPR_FALSE;

    return XPR_TRUE;
}

xpr_bool_t ConfDir::getLoadDir(xpr_tchar_t *aDir, xpr_size_t aMaxLen) const
{
    return getDir(aDir, aMaxLen);
}

xpr_bool_t ConfDir::getSaveDir(xpr_tchar_t *aDir, xpr_size_t aMaxLen) const
{
    if (getDir(aDir, aMaxLen) == XPR_FALSE)
        return XPR_FALSE;

    if (IsExistFile(aDir) == XPR_TRUE)
        return XPR_TRUE;

    if (CreateDirectoryLevel(aDir) == XPR_FALSE)
        return XPR_FALSE;

    return XPR_TRUE;
}

void ConfDir::setBackup(void)
{
    mOldConfDir = mConfDir;
}

xpr_bool_t ConfDir::checkChangedConfDir(void)
{
    if (_tcsicmp(mOldConfDir.c_str(), mConfDir.c_str()) == 0)
    {
        return XPR_FALSE;
    }

    xpr_tchar_t sOldDir[XPR_MAX_PATH + 1] = {0};
    xpr_tchar_t sNewDir[XPR_MAX_PATH + 1] = {0};

    if (getDir(mOldConfDir.c_str(), sOldDir, XPR_MAX_PATH) == XPR_TRUE &&
        getDir(mConfDir.c_str(), sNewDir, XPR_MAX_PATH) == XPR_TRUE)
    {
        if (_tcsicmp(sOldDir, sNewDir) == 0)
        {
            return XPR_FALSE;
        }
    }

    return XPR_TRUE;
}

xpr_bool_t ConfDir::moveToNewConfDir(void)
{
    xpr_sint_t  sType;
    xpr_tchar_t sOldConfDir[XPR_MAX_PATH + 1] = {0};
    xpr_tchar_t sNewConfDir[XPR_MAX_PATH + 1] = {0};
    xpr_tchar_t sOldPath[XPR_MAX_PATH + 1] = {0};
    xpr_tchar_t sNewPath[XPR_MAX_PATH + 1] = {0};

    getDir(mOldConfDir.c_str(), sOldConfDir, XPR_MAX_PATH);

    if (getSaveDir(sNewConfDir, XPR_MAX_PATH) == XPR_FALSE)
    {
        return XPR_FALSE;
    }

    for (sType = TypeBegin; sType < TypeEnd; ++sType)
    {
        sOldPath[0] = XPR_STRING_LITERAL('\0');
        sNewPath[0] = XPR_STRING_LITERAL('\0');

        getPath(sType, sOldConfDir, sOldPath, XPR_MAX_PATH);
        getSavePath(sType, sNewPath, XPR_MAX_PATH);

        if (xpr::FileSys::exist(sOldPath))
        {
            if (_tcsicmp(sOldPath, sNewPath) != 0)
            {
                if (xpr::FileSys::exist(sNewPath))
                {
                    xpr::FileSys::remove(sNewPath);
                }

                xpr::FileSys::rename(sOldPath, sNewPath);
            }
        }
    }

    return XPR_TRUE;
}

const xpr_tchar_t *ConfDir::getFileName(xpr_sint_t aType)
{
    switch (aType)
    {
    case TypeMain:           return XPR_STRING_LITERAL("fxfile-main.conf");
    case TypeConfig:         return XPR_STRING_LITERAL("fxfile.conf");
    case TypeBookmark:       return XPR_STRING_LITERAL("fxfile-bookmark.conf");
    case TypeFileScrap:      return XPR_STRING_LITERAL("fxfile-file_scrap.conf");
    case TypeSearchDir:      return XPR_STRING_LITERAL("fxfile-search_dir.conf");
    case TypeFolderLayout:   return XPR_STRING_LITERAL("fxfile-folder_layout.conf");
    case TypeDlgState:       return XPR_STRING_LITERAL("fxfile-dlg_state.conf");
    case TypeAccel:          return XPR_STRING_LITERAL("fxfile-accel.dat");
    case TypeCoolBar:        return XPR_STRING_LITERAL("fxfile-coolbar.dat");
    case TypeToolBar:        return XPR_STRING_LITERAL("fxfile-toolbar.dat");
    case TypeThumbnailData:  return XPR_STRING_LITERAL("fxfile-thumbnail.dat");
    case TypeThumbnailIndex: return XPR_STRING_LITERAL("fxfile-thumbnail.idx");
    case TypeLauncher:       return XPR_STRING_LITERAL("fxfile-launcher.conf");
    case TypeUpchecker:      return XPR_STRING_LITERAL("fxfile-upchecker.conf");
    case TypeViewSet:        return XPR_STRING_LITERAL("fxfile-view_set.conf");
    case TypeUpdater:        return XPR_STRING_LITERAL("fxfile-updater.conf");
    }

    return XPR_NULL;
}

const xpr_tchar_t *ConfDir::getDefConfDir(void)
{
    return kDefConfDir;
}

const xpr_tchar_t *ConfDir::getProgramConfDir(void)
{
    return kProgramConfDir;
}

xpr_bool_t ConfDir::load(void)
{
    clear();

    xpr::string sLocalPath;
    xpr::string sLocalPathOld;
    xpr::string sAppDataPath;
    xpr::string sProgramConfDir;
    xpr::string sProgramConfigPath;
    xpr::string sProgramMainPath;
    
    GetEnvRealPath(kLocalFxFilePath, sLocalPath);
    GetEnvRealPath(kLocalFxFilePathOld, sLocalPathOld);
    GetEnvRealPath(kFxFilePath, sAppDataPath);
    GetEnvRealPath(kProgramConfDir, sProgramConfDir);

    sProgramConfigPath = sProgramConfDir;
    sProgramConfigPath += XPR_STRING_LITERAL("\\fxfile.conf");

    sProgramMainPath = sProgramConfDir;
    sProgramMainPath += XPR_STRING_LITERAL("\\fxfile-main.conf");

    const DWORD sProgramConfigAttrs = ::GetFileAttributes(sProgramConfigPath.c_str());
    const DWORD sProgramMainAttrs = ::GetFileAttributes(sProgramMainPath.c_str());

    const xpr_bool_t sHasProgramConfig =
        (sProgramConfigAttrs != INVALID_FILE_ATTRIBUTES &&
         !XPR_TEST_BITS(sProgramConfigAttrs, FILE_ATTRIBUTE_DIRECTORY));
    const xpr_bool_t sHasProgramMain =
        (sProgramMainAttrs != INVALID_FILE_ATTRIBUTES &&
         !XPR_TEST_BITS(sProgramMainAttrs, FILE_ATTRIBUTE_DIRECTORY));

    xpr_bool_t sResult = XPR_FALSE;
    xpr::string sFinalPath;

    if (xpr::FileSys::exist(sLocalPath.c_str()))
    {
        sFinalPath = sLocalPath;
    }
    else if (xpr::FileSys::exist(sLocalPathOld.c_str()))
    {
        sFinalPath = sLocalPathOld;
    }
    else if (XPR_IS_TRUE(sHasProgramConfig) && XPR_IS_TRUE(sHasProgramMain))
    {
        // A copied portable package already contains its saved environment.
        // Prefer it to a machine-wide AppData pointer without requiring an INI.
        setConfDir(kProgramConfDir, XPR_TRUE);
        sResult = XPR_TRUE;
    }
    else if (xpr::FileSys::exist(sAppDataPath.c_str()))
    {
        sFinalPath = sAppDataPath;
    }

    if (!sFinalPath.empty())
    {
        fxfile::base::ConfFileEx sConfFile(sFinalPath.c_str());
        if (sConfFile.load() == XPR_TRUE)
        {
            const xpr_tchar_t *sValue;
            ConfFile::Section *sSection;

            sSection = sConfFile.findSection(kFxFileSection);
            if (XPR_IS_NOT_NULL(sSection))
            {
                sValue = sConfFile.getValueS(sSection, kConfHomeKey, XPR_NULL);
                if (XPR_IS_NOT_NULL(sValue))
                {
                    setConfDir(sValue);
                    sResult = XPR_TRUE;
                }
            }
        }
    }

    if (XPR_IS_FALSE(sResult))
    {
        setConfDir(kProgramConfDir);
        sResult = XPR_TRUE;
    }

    return sResult;
}

xpr_bool_t ConfDir::save(void) const
{
    if (XPR_IS_TRUE(mReadOnly))
    {
        return XPR_TRUE;
    }

    xpr::string sPath;
    xpr::string sLocalPath;
    xpr::string sLocalPathOld;
    GetEnvRealPath(kLocalFxFilePath, sLocalPath);
    GetEnvRealPath(kLocalFxFilePathOld, sLocalPathOld);

    if (xpr::FileSys::exist(sLocalPath.c_str()) == XPR_FALSE &&
        xpr::FileSys::exist(sLocalPathOld.c_str()) == XPR_FALSE &&
        _tcsicmp(mConfDir.c_str(), kProgramConfDir) == 0)
    {
        // The program configuration directory is self-discovering. Do not
        // create fxfile.ini merely because the configuration dialog was saved.
        return XPR_TRUE;
    }
    
    if (xpr::FileSys::exist(sLocalPath.c_str()) || mConfDir.find(XPR_STRING_LITERAL("%fxfile%")) != xpr::string::npos)
    {
        sPath = sLocalPath;
    }
    else
    {
        GetEnvRealPath(kFxFilePath, sPath);
    }

    fxfile::base::ConfFileEx sConfFile(sPath.c_str());
    sConfFile.setComment(XPR_STRING_LITERAL("fxfile configuration file"));

    ConfFile::Section *sSection;

    sSection = sConfFile.addSection(kFxFileSection);
    XPR_ASSERT(sSection != XPR_NULL);

    sConfFile.setValueS(sSection, kConfHomeKey, mConfDir.c_str());

    DWORD sFileAttributes = ::GetFileAttributes(sPath.c_str());
    if (sFileAttributes != INVALID_FILE_ATTRIBUTES)
    {
        sFileAttributes &= ~FILE_ATTRIBUTE_HIDDEN;
        ::SetFileAttributes(sPath.c_str(), sFileAttributes);
    }

    sConfFile.save(xpr::CharSetUtf16);

    if (sPath.find(XPR_STRING_LITERAL(".fxfile")) != xpr::string::npos)
    {
        sFileAttributes = ::GetFileAttributes(sPath.c_str());
        if (sFileAttributes != INVALID_FILE_ATTRIBUTES)
        {
            sFileAttributes |= FILE_ATTRIBUTE_HIDDEN;
            ::SetFileAttributes(sPath.c_str(), sFileAttributes);
        }
    }

    return XPR_TRUE;
}
} // namespace fxfile
