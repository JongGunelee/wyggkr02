//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "recent_file_list.h"

#include "conf_file_ex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace fxfile;
using namespace fxfile::base;

namespace fxfile
{
namespace
{
const xpr_tchar_t kRecentFileListSection[] = XPR_STRING_LITERAL("recent_file_list");
const xpr_tchar_t kFileKey              [] = XPR_STRING_LITERAL("recent_file%d");
} // namespace anonymous

RecentFileList::RecentFileList(void)
    : mLoadPending(XPR_FALSE)
{
}

RecentFileList::~RecentFileList(void)
{
    clear();
}

void RecentFileList::addFile(const xpr_tchar_t *aPath)
{
    if (XPR_IS_NULL(aPath))
        return;

    ensureLoaded();
    mFileDeque.push_front(aPath);
}

xpr_size_t RecentFileList::getFileCount(void) const
{
    const_cast<RecentFileList *>(this)->ensureLoaded();
    return mFileDeque.size();
}

const xpr_tchar_t *RecentFileList::getFile(xpr_size_t aIndex) const
{
    const_cast<RecentFileList *>(this)->ensureLoaded();
    if (!FXFILE_STL_IS_INDEXABLE(aIndex, mFileDeque))
        return XPR_NULL;

    return mFileDeque[aIndex].c_str();
}

void RecentFileList::clear(void)
{
    mFileDeque.clear();
    mLoadPending = XPR_FALSE;
    mPendingPath.clear();
}

void RecentFileList::load(fxfile::base::ConfFileEx &aConfFile)
{
    clear();

    xpr_sint_t         i;
    xpr_tchar_t        sKey[XPR_MAX_PATH + 1];
    const xpr_tchar_t *sValue;
    ConfFile::Section *sSection;

    sSection = aConfFile.findSection(kRecentFileListSection);
    if (XPR_IS_NOT_NULL(sSection))
    {
        for (i = 0; ; ++i)
        {
            _stprintf(sKey, kFileKey, i + 1);

            sValue = aConfFile.getValueS(sSection, sKey, XPR_NULL);
            if (XPR_IS_NULL(sValue))
                break;

            mFileDeque.push_back(sValue);
        }
    }
}

xpr_bool_t RecentFileList::load(const xpr_tchar_t *aFilePath)
{
    if (XPR_IS_NULL(aFilePath))
        return XPR_FALSE;

    HANDLE sFile = ::CreateFileW((const wchar_t *)aFilePath,
                                 GENERIC_READ,
                                 FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                 XPR_NULL,
                                 OPEN_EXISTING,
                                 FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
                                 XPR_NULL);
    if (sFile == INVALID_HANDLE_VALUE)
        return XPR_FALSE;

    LARGE_INTEGER sFileSize = {0};
    if (::GetFileSizeEx(sFile, &sFileSize) == FALSE ||
        sFileSize.QuadPart < 2 || sFileSize.QuadPart > 0x7fffffff)
    {
        ::CloseHandle(sFile);
        return XPR_FALSE;
    }

    std::vector<xpr_byte_t> sBytes((xpr_size_t)sFileSize.QuadPart + sizeof(wchar_t), 0);
    DWORD sReadBytes = 0;
    xpr_bool_t sRead = (::ReadFile(sFile,
                                  &sBytes[0],
                                  (DWORD)sFileSize.QuadPart,
                                  &sReadBytes,
                                  XPR_NULL) != FALSE &&
                        sReadBytes == (DWORD)sFileSize.QuadPart);
    ::CloseHandle(sFile);
    if (XPR_IS_FALSE(sRead) || sBytes[0] != 0xff || sBytes[1] != 0xfe)
        return XPR_FALSE;

    mFileDeque.clear();
    mLoadPending = XPR_FALSE;
    mPendingPath.clear();

    const wchar_t *sCursor = (const wchar_t *)(&sBytes[2]);
    const wchar_t *sEnd = (const wchar_t *)(&sBytes[0] + sReadBytes);
    xpr_bool_t sInRecentSection = XPR_FALSE;
    const wchar_t sRecentSection[] = L"[recent_file_list]";
    const wchar_t sRecentPrefix[] = L"recent_file";

    while (sCursor < sEnd && *sCursor != 0)
    {
        const wchar_t *sLineEnd = sCursor;
        while (sLineEnd < sEnd && *sLineEnd != L'\r' && *sLineEnd != L'\n')
            ++sLineEnd;

        const wchar_t *sText = sCursor;
        while (sText < sLineEnd && iswspace(*sText))
            ++sText;

        if (sText < sLineEnd && *sText == L'[')
        {
            size_t sLength = (size_t)(sLineEnd - sText);
            sInRecentSection = (sLength == _countof(sRecentSection) - 1 &&
                                _wcsnicmp(sText, sRecentSection, sLength) == 0)
                             ? XPR_TRUE : XPR_FALSE;
        }
        else if (XPR_IS_TRUE(sInRecentSection) &&
                 (size_t)(sLineEnd - sText) > _countof(sRecentPrefix) - 1 &&
                 _wcsnicmp(sText, sRecentPrefix, _countof(sRecentPrefix) - 1) == 0)
        {
            const wchar_t *sValue = sText;
            while (sValue < sLineEnd && *sValue != L'=')
                ++sValue;
            if (sValue < sLineEnd)
            {
                ++sValue;
                while (sValue < sLineEnd && iswspace(*sValue))
                    ++sValue;

                const wchar_t *sValueEnd = sLineEnd;
                while (sValueEnd > sValue && iswspace(*(sValueEnd - 1)))
                    --sValueEnd;

                xpr::string sPath;
                sPath.assign(sValue, (xpr_size_t)(sValueEnd - sValue));
                mFileDeque.push_back(sPath);
            }
        }

        sCursor = sLineEnd;
        while (sCursor < sEnd && (*sCursor == L'\r' || *sCursor == L'\n'))
            ++sCursor;
    }

    return XPR_TRUE;
}

void RecentFileList::scheduleLoad(const xpr_tchar_t *aFilePath)
{
    mFileDeque.clear();
    mPendingPath = XPR_IS_NOT_NULL(aFilePath) ? aFilePath : XPR_STRING_LITERAL("");
    mLoadPending = mPendingPath.empty() ? XPR_FALSE : XPR_TRUE;
}

void RecentFileList::prepareForConfigDirMove(void)
{
    ensureLoaded();
}

void RecentFileList::ensureLoaded(void)
{
    if (XPR_IS_FALSE(mLoadPending))
        return;

    xpr::string sPath = mPendingPath;
    mLoadPending = XPR_FALSE;
    mPendingPath.clear();

    if (load(sPath.c_str()) == XPR_FALSE)
    {
        // Compatibility fallback for legacy non-UTF-16 configurations.
        fxfile::base::ConfFileEx sConfFile(sPath);
        if (sConfFile.load() == XPR_TRUE)
            load(sConfFile);
    }
}

void RecentFileList::save(fxfile::base::ConfFileEx &aConfFile) const
{
    const_cast<RecentFileList *>(this)->ensureLoaded();

    xpr_sint_t         i;
    xpr_tchar_t        sKey[XPR_MAX_PATH + 1];
    ConfFile::Section *sSection;
    FileDeque::const_iterator sIterator;

    sSection = aConfFile.addSection(kRecentFileListSection);
    XPR_ASSERT(sSection != XPR_NULL);

    sIterator = mFileDeque.begin();
    for (i = 0; sIterator != mFileDeque.end(); ++sIterator, ++i)
    {
        const xpr::string &sPath = *sIterator;

        _stprintf(sKey, kFileKey, i + 1);

        aConfFile.setValueS(sSection, sKey, sPath);
    }
}
} // namespace fxfile
