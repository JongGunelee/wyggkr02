//
// Copyright (c) 2012-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "language_table.h"

#include "xpr_xml.h"

namespace fxfile
{
namespace base
{
LanguageTable::LanguageTable(void)
    : mLoadedLanaguage(XPR_NULL)
    , mStringTable(XPR_NULL)
    , mFormatStringTable(XPR_NULL)
{
    mDir[0] = 0;
}

LanguageTable::~LanguageTable(void)
{
    clear();
}

void LanguageTable::setDir(const xpr_tchar_t *aDir)
{
    if (XPR_IS_NOT_NULL(aDir))
        _tcscpy(mDir, aDir);
}

xpr_bool_t LanguageTable::scan(const xpr_tchar_t *aPreloadLanguage)
{
    clear();

    // [x64 Fix] Use native wchar_t and W-APIs to handle Korean paths reliably
    wchar_t sDir[XPR_MAX_PATH + 1] = {0};
    wcsncpy_s(sDir, (const wchar_t*)mDir, XPR_MAX_PATH);

    // Ensure trailing backslash for the directory
    size_t sDirLen = wcslen(sDir);
    if (sDirLen > 0 && sDir[sDirLen - 1] != L'\\')
        wcscat_s(sDir, XPR_MAX_PATH, L"\\");

    wchar_t sPattern[XPR_MAX_PATH + 1] = {0};
    wcscpy_s(sPattern, sDir);
    wcscat_s(sPattern, L"*.*");

    HANDLE sFile;
    WIN32_FIND_DATAW sWin32FindData = {0};
    LanguagePack *sLanguagePack;

    // Direct call to Wide-version API
    sFile = ::FindFirstFileW(sPattern, &sWin32FindData);
    if (sFile == INVALID_HANDLE_VALUE)
    {
        wchar_t sErrDbg[1024];
        swprintf_s(sErrDbg, L"Language scan failed!\nSearching in: %ls\nError: %u", sPattern, GetLastError());
        ::MessageBoxW(NULL, sErrDbg, L"fxfile Unicode Debug", MB_OK | MB_ICONSTOP);
        return XPR_FALSE;
    }

    do
    {
        // Skip directory navigation entries
        if (wcscmp(sWin32FindData.cFileName, L".") == 0 || wcscmp(sWin32FindData.cFileName, L"..") == 0)
            continue;

        if ((sWin32FindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            continue;

        const wchar_t *sExtension = wcsrchr(sWin32FindData.cFileName, L'.');
        if (sExtension == XPR_NULL || _wcsicmp(sExtension, L".xml") != 0)
            continue;

        // Construct exact file path using native wchar_t
        wchar_t sFullFilePath[XPR_MAX_PATH + 1] = {0};
        swprintf_s(sFullFilePath, L"%ls%ls", sDir, sWin32FindData.cFileName);

        sLanguagePack = new LanguagePack;
        if (XPR_IS_NULL(sLanguagePack))
            continue;

        xpr_bool_t sPreload = XPR_FALSE;
        if (XPR_IS_NOT_NULL(aPreloadLanguage))
        {
            wchar_t sFileLanguage[LanguagePack::kMaxLanguageLength + 1] = {0};
            size_t sBaseLength = (size_t)(sExtension - sWin32FindData.cFileName);
            if (sBaseLength <= LanguagePack::kMaxLanguageLength)
            {
                wcsncpy_s(sFileLanguage,
                          LanguagePack::kMaxLanguageLength + 1,
                          sWin32FindData.cFileName,
                          sBaseLength);
                sPreload = (_wcsicmp(sFileLanguage, (const wchar_t *)aPreloadLanguage) == 0) ? XPR_TRUE : XPR_FALSE;
            }
        }

        StringTable *sPreloadedStringTable = XPR_NULL;
        if (XPR_IS_TRUE(sPreload))
            sPreloadedStringTable = new StringTable;

        // Parse the selected language and its string table in one pass.  The
        // former scan parsed the same XML once for metadata and again for the
        // strings, doubling startup cost.
        xpr_bool_t sLoaded = XPR_IS_TRUE(sPreload)
                           ? sLanguagePack->load((const xpr_tchar_t *)sFullFilePath, sPreloadedStringTable)
                           : sLanguagePack->load((const xpr_tchar_t *)sFullFilePath);
        if (XPR_IS_FALSE(sLoaded))
        {
            XPR_SAFE_DELETE(sPreloadedStringTable);
            XPR_SAFE_DELETE(sLanguagePack);
            continue;
        }

        mLanguageMap[sLanguagePack->getLanguageDesc()->mLanguage] = sLanguagePack;
        mLanguageDeque.push_back(sLanguagePack);

        if (XPR_IS_TRUE(sPreload))
        {
            mLoadedLanaguage = sLanguagePack;
            mStringTable = sPreloadedStringTable;
            mFormatStringTable = new FormatStringTable(*mStringTable);
        }
    }
    while (::FindNextFileW(sFile, &sWin32FindData));

    ::FindClose(sFile);

    return XPR_TRUE;
}

xpr_bool_t LanguageTable::loadLanguage(const xpr_tchar_t *aLanguage)
{
    if (XPR_IS_NULL(aLanguage))
        return XPR_FALSE;

    if (XPR_IS_NOT_NULL(mLoadedLanaguage) && mLoadedLanaguage->equalLanguage(aLanguage) == XPR_TRUE)
        return XPR_TRUE;

    LanguageMap::iterator sIterator = mLanguageMap.find(aLanguage);
    if (sIterator == mLanguageMap.end())
    {
        // not exist and just one languague pack
        if (mLanguageMap.size() == 1)
        {
            sIterator = mLanguageMap.begin();
        }
        else
        {
            return XPR_FALSE;
        }
    }

    mLoadedLanaguage = sIterator->second;

    XPR_SAFE_DELETE(mStringTable);
    XPR_SAFE_DELETE(mFormatStringTable);

    mStringTable = new StringTable;
    if (XPR_IS_NULL(mStringTable))
    {
        mLoadedLanaguage = XPR_NULL;

        XPR_SAFE_DELETE(mStringTable);

        return XPR_FALSE;
    }

    if (mLoadedLanaguage->loadStringTable(mStringTable) == XPR_FALSE)
    {
        mLoadedLanaguage = XPR_NULL;

        XPR_SAFE_DELETE(mStringTable);

        return XPR_FALSE;
    }

    mFormatStringTable = new FormatStringTable(*mStringTable);
    if (XPR_IS_NULL(mFormatStringTable))
    {
        mLoadedLanaguage = XPR_NULL;

        XPR_SAFE_DELETE(mStringTable);

        return XPR_FALSE;
    }

    return XPR_TRUE;
}

void LanguageTable::unloadLanguage(void)
{
    mLoadedLanaguage = XPR_NULL;

    XPR_SAFE_DELETE(mStringTable);
    XPR_SAFE_DELETE(mFormatStringTable);
}

xpr_size_t LanguageTable::getLanguageCount(void) const
{
    return mLanguageDeque.size();
}

const LanguagePack::Desc *LanguageTable::getLanguageDesc(xpr_size_t aIndex) const
{
    if (!FXFILE_STL_IS_INDEXABLE(aIndex, mLanguageDeque))
        return XPR_NULL;

    return mLanguageDeque[aIndex]->getLanguageDesc();
}

const LanguagePack::Desc *LanguageTable::getLanguageDesc(const xpr_tchar_t *aLanguage) const
{
    if (XPR_IS_NULL(aLanguage))
        return XPR_NULL;

    LanguageMap::const_iterator sIterator = mLanguageMap.find(aLanguage);
    if (sIterator == mLanguageMap.end())
        return XPR_NULL;

    const LanguagePack *sLanguagePack = sIterator->second;
    const LanguagePack::Desc *sLanguagePackDesc = sLanguagePack->getLanguageDesc();

    return sLanguagePackDesc;
}

const LanguagePack::Desc *LanguageTable::getLanguageDesc(void) const
{
    if (XPR_IS_NULL(mLoadedLanaguage))
        return XPR_NULL;

    return mLoadedLanaguage->getLanguageDesc();
}

StringTable *LanguageTable::getStringTable(void) const
{
    return mStringTable;
}

FormatStringTable *LanguageTable::getFormatStringTable(void) const
{
    return mFormatStringTable;
}

void LanguageTable::clear(void)
{
    LanguagePack *sLanguagePack;
    LanguageDeque::iterator sIterator;

    sIterator = mLanguageDeque.begin();
    for (; sIterator != mLanguageDeque.end(); ++sIterator)
    {
        sLanguagePack = *sIterator;
        XPR_SAFE_DELETE(sLanguagePack);
    }

    mLanguageMap.clear();
    mLanguageDeque.clear();

    mLoadedLanaguage = XPR_NULL;
    XPR_SAFE_DELETE(mStringTable);
    XPR_SAFE_DELETE(mFormatStringTable);
}
} // namespace base
} // namespace fxfile
