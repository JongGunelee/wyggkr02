//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "thumb_img_list.h"

#include "conf_dir.h"
#include <xpr_file_sys.h>
#include <set>
#include <vector>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace
{
const DWORD     kThumbnailDataMagic   = 0x44435846; // "FXCD"
const DWORD     kThumbnailIndexMagic  = 0x49435846; // "FXCI"
const DWORD     kThumbnailCacheVersion = 1;
// Cache loading currently occurs on the UI thread. Conservative limits keep
// corrupt or unexpectedly huge caches from exhausting x86 address space or
// producing a long "Not Responding" interval.
#if defined(_WIN64)
const ULONGLONG kMaxThumbnailDataSize = 64ULL * 1024ULL * 1024ULL;
#else
const ULONGLONG kMaxThumbnailDataSize = 32ULL * 1024ULL * 1024ULL;
#endif
const ULONGLONG kMaxThumbnailIndexSize = 4ULL * 1024ULL * 1024ULL;
const DWORD     kMaxThumbnailRecords   = 4096;
const xpr_uint_t kInvalidThumbImageId  = 0xffffffff;

const xpr_tchar_t kThumbnailDataFileName [] = XPR_STRING_LITERAL("fxfile-thumbnail.dat");
const xpr_tchar_t kThumbnailIndexFileName[] = XPR_STRING_LITERAL("fxfile-thumbnail.idx");

xpr_bool_t fileExists(const xpr_tchar_t *aPath)
{
    const DWORD sAttrs = ::GetFileAttributes(aPath);
    return (sAttrs != INVALID_FILE_ATTRIBUTES &&
            !XPR_TEST_BITS(sAttrs, FILE_ATTRIBUTE_DIRECTORY)) ? XPR_TRUE : XPR_FALSE;
}

xpr_bool_t readExact(HANDLE aFile, void *aBuffer, DWORD aBytes)
{
    DWORD sRead = 0;
    return (::ReadFile(aFile, aBuffer, aBytes, &sRead, XPR_NULL) != FALSE &&
            sRead == aBytes) ? XPR_TRUE : XPR_FALSE;
}

xpr_bool_t writeExact(HANDLE aFile, const void *aBuffer, DWORD aBytes)
{
    DWORD sWritten = 0;
    return (::WriteFile(aFile, aBuffer, aBytes, &sWritten, XPR_NULL) != FALSE &&
            sWritten == aBytes) ? XPR_TRUE : XPR_FALSE;
}

xpr_bool_t getFileSize(const xpr_tchar_t *aPath, ULONGLONG &aSize)
{
    aSize = 0;

    HANDLE sFile = ::CreateFile(aPath, GENERIC_READ,
                                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                XPR_NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, XPR_NULL);
    if (sFile == INVALID_HANDLE_VALUE)
        return XPR_FALSE;

    LARGE_INTEGER sSize = {0};
    const xpr_bool_t sResult = (::GetFileSizeEx(sFile, &sSize) != FALSE &&
                                sSize.QuadPart >= 0) ? XPR_TRUE : XPR_FALSE;
    ::CloseHandle(sFile);

    if (XPR_IS_TRUE(sResult))
        aSize = (ULONGLONG)sSize.QuadPart;

    return sResult;
}

ULONGLONG createCacheGeneration(void)
{
    FILETIME sFileTime = {0};
    ::GetSystemTimeAsFileTime(&sFileTime);

    ULONGLONG sGeneration = ((ULONGLONG)sFileTime.dwHighDateTime << 32) |
                             (ULONGLONG)sFileTime.dwLowDateTime;
    sGeneration ^= ((ULONGLONG)::GetCurrentProcessId() << 32);
    sGeneration ^= (ULONGLONG)::GetTickCount();
    return sGeneration;
}

xpr_bool_t moveReplace(const xpr_tchar_t *aSource, const xpr_tchar_t *aTarget)
{
    return (::MoveFileEx(aSource, aTarget,
                         MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) != FALSE)
         ? XPR_TRUE : XPR_FALSE;
}

void restoreBackup(const xpr_tchar_t *aBackup, const xpr_tchar_t *aTarget,
                   xpr_bool_t aHadOriginal)
{
    ::DeleteFile(aTarget);
    if (XPR_IS_TRUE(aHadOriginal))
        moveReplace(aBackup, aTarget);
}

xpr_bool_t commitCachePair(const xpr_tchar_t *aDataPath,
                           const xpr_tchar_t *aIndexPath,
                           const xpr_tchar_t *aTempDataPath,
                           const xpr_tchar_t *aTempIndexPath,
                           const xpr_tchar_t *aBackupDataPath,
                           const xpr_tchar_t *aBackupIndexPath)
{
    ::DeleteFile(aBackupDataPath);
    ::DeleteFile(aBackupIndexPath);

    const xpr_bool_t sHadData  = fileExists(aDataPath);
    const xpr_bool_t sHadIndex = fileExists(aIndexPath);

    if (XPR_IS_TRUE(sHadData) &&
        XPR_IS_FALSE(moveReplace(aDataPath, aBackupDataPath)))
    {
        return XPR_FALSE;
    }

    if (XPR_IS_TRUE(sHadIndex) &&
        XPR_IS_FALSE(moveReplace(aIndexPath, aBackupIndexPath)))
    {
        restoreBackup(aBackupDataPath, aDataPath, sHadData);
        return XPR_FALSE;
    }

    if (XPR_IS_FALSE(moveReplace(aTempDataPath, aDataPath)))
    {
        restoreBackup(aBackupDataPath,  aDataPath,  sHadData);
        restoreBackup(aBackupIndexPath, aIndexPath, sHadIndex);
        return XPR_FALSE;
    }

    if (XPR_IS_FALSE(moveReplace(aTempIndexPath, aIndexPath)))
    {
        restoreBackup(aBackupDataPath,  aDataPath,  sHadData);
        restoreBackup(aBackupIndexPath, aIndexPath, sHadIndex);
        return XPR_FALSE;
    }

    ::DeleteFile(aBackupDataPath);
    ::DeleteFile(aBackupIndexPath);
    return XPR_TRUE;
}
} // namespace anonymous

ThumbImgList::ThumbImgList(void)
{
}

ThumbImgList::~ThumbImgList(void)
{
    removeAll();
}

xpr_bool_t ThumbImgList::create(xpr_sint_t cx, xpr_sint_t cy, xpr_uint_t aFlags)
{
    return mImageList.Create(cx, cy, aFlags, 100, 50);
}

xpr_bool_t ThumbImgList::resize(xpr_sint_t cx, xpr_sint_t cy)
{
    if (XPR_IS_NULL(mImageList.m_hImageList))
        return create(cx, cy, ILC_COLOR24);

    removeAll();
    return (::ImageList_SetIconSize(mImageList.m_hImageList, cx, cy) != FALSE)
         ? XPR_TRUE : XPR_FALSE;
}

void ThumbImgList::setCacheDir(const xpr_tchar_t *aCacheDir)
{
    mCacheDir = XPR_IS_NOT_NULL(aCacheDir)
              ? aCacheDir
              : XPR_STRING_LITERAL("");
}

const xpr_tchar_t *ThumbImgList::getCacheDir(void) const
{
    return mCacheDir.c_str();
}

xpr_bool_t ThumbImgList::getCachePaths(xpr_bool_t aForSave,
                                       xpr::string &aDataPath,
                                       xpr::string &aIndexPath) const
{
    aDataPath.clear();
    aIndexPath.clear();

    if (mCacheDir.empty())
    {
        xpr_tchar_t sDataPath [XPR_MAX_PATH + 1] = {0};
        xpr_tchar_t sIndexPath[XPR_MAX_PATH + 1] = {0};

        xpr_bool_t sDataResult;
        xpr_bool_t sIndexResult;
        if (XPR_IS_TRUE(aForSave))
        {
            sDataResult = ConfDir::instance().getSavePath(
                ConfDir::TypeThumbnailData, sDataPath, XPR_MAX_PATH);
            sIndexResult = ConfDir::instance().getSavePath(
                ConfDir::TypeThumbnailIndex, sIndexPath, XPR_MAX_PATH);
        }
        else
        {
            sDataResult = ConfDir::instance().getLoadPath(
                ConfDir::TypeThumbnailData, sDataPath, XPR_MAX_PATH);
            sIndexResult = ConfDir::instance().getLoadPath(
                ConfDir::TypeThumbnailIndex, sIndexPath, XPR_MAX_PATH);
        }

        if (XPR_IS_FALSE(sDataResult) || XPR_IS_FALSE(sIndexResult))
            return XPR_FALSE;

        aDataPath  = sDataPath;
        aIndexPath = sIndexPath;
        return XPR_TRUE;
    }

    if (mCacheDir.length() > XPR_MAX_PATH)
        return XPR_FALSE;

    if (XPR_IS_TRUE(aForSave))
    {
        const DWORD sAttrs = ::GetFileAttributes(mCacheDir.c_str());
        if (sAttrs == INVALID_FILE_ATTRIBUTES)
        {
            if (XPR_RCODE_IS_ERROR(xpr::FileSys::mkdir_recursive(mCacheDir)))
                return XPR_FALSE;
        }
        else if (!XPR_TEST_BITS(sAttrs, FILE_ATTRIBUTE_DIRECTORY))
        {
            return XPR_FALSE;
        }
    }

    aDataPath = mCacheDir;
    if (!aDataPath.empty() &&
        aDataPath[aDataPath.length() - 1] != XPR_STRING_LITERAL('\\') &&
        aDataPath[aDataPath.length() - 1] != XPR_STRING_LITERAL('/'))
    {
        aDataPath += XPR_STRING_LITERAL('\\');
    }
    aIndexPath = aDataPath;
    aDataPath  += kThumbnailDataFileName;
    aIndexPath += kThumbnailIndexFileName;

    if (aDataPath.length() > XPR_MAX_PATH || aIndexPath.length() > XPR_MAX_PATH)
        return XPR_FALSE;

    return XPR_TRUE;
}

xpr_sint_t ThumbImgList::add(ThumbImage *aThumbImage, xpr_uint_t aThumbImageId)
{
    if (XPR_IS_NULL(aThumbImage))
        return -1;

    if (XPR_IS_NULL(aThumbImage->mImage) || aThumbImage->mPath.empty() == XPR_TRUE || aThumbImage->mWidth <= 0 || aThumbImage->mHeight <= 0 || aThumbImage->mDepth <= 0)
        return -1;

    // Keep GDI bitmap memory bounded during long thumbnail browsing sessions.
    // The same limit is already enforced for persistent cache loading/saving.
    if (mThumbDeque.size() >= kMaxThumbnailRecords)
        return -1;

    xpr_sint_t sImageIndex = mImageList.Add(aThumbImage->mImage, CLR_NONE);
    if (sImageIndex < 0)
        return -1;

    ThumbElement sThumbElement;
    sThumbElement.mPath             = aThumbImage->mPath;
    sThumbElement.mImageIndex       = sImageIndex;
    sThumbElement.mThumbImageId     = aThumbImageId;
    sThumbElement.mWidth            = aThumbImage->mWidth;
    sThumbElement.mHeight           = aThumbImage->mHeight;
    sThumbElement.mDepth            = aThumbImage->mDepth;
    sThumbElement.mModifiedFileTime = aThumbImage->mModifiedFileTime;

    mThumbDeque.push_back(sThumbElement);

    return sImageIndex;
}

xpr_sint_t ThumbImgList::getImageCount(void) const
{
    return mImageList.GetImageCount();
}

ThumbImgList::ThumbElement *ThumbImgList::getImage(xpr_sint_t aImageIndex)
{
    if (!FXFILE_STL_IS_INDEXABLE(aImageIndex, mThumbDeque))
        return XPR_NULL;

    return &mThumbDeque[aImageIndex];
}

ThumbImgList::ThumbElement *ThumbImgList::getImage(const xpr_tchar_t *aPath)
{
    if (XPR_IS_NULL(aPath))
        return XPR_NULL;

    ThumbDeque::iterator sIterator;

    sIterator = mThumbDeque.begin();
    for (; sIterator != mThumbDeque.end(); ++sIterator)
    {
        ThumbElement &sThumbElement = *sIterator;

        if (_tcsicmp(sThumbElement.mPath.c_str(), aPath) == 0)
            return &sThumbElement;
    }

    return XPR_NULL;
}

// return value
// (-1) Not Exist
// ( 0) Not Modified
// ( 1) Modified
xpr_sint_t ThumbImgList::compareTime(xpr_sint_t nImage, LPFILETIME aModifiedFileTime)
{
    if (XPR_IS_NULL(aModifiedFileTime))
        return -1;

    xpr_sint_t sResult = -1;

    ThumbDeque::iterator sIterator;

    sIterator = mThumbDeque.begin();
    for (; sIterator != mThumbDeque.end(); ++sIterator)
    {
        ThumbElement &sThumbElement = *sIterator;

        if (sThumbElement.mImageIndex == nImage)
        {
            sResult = CompareFileTime(&sThumbElement.mModifiedFileTime, aModifiedFileTime) ? 1 : 0;
            break;
        }
    }

    return sResult;
}

CImageList *ThumbImgList::getImageList(void)
{
    return &mImageList;
}

xpr_bool_t ThumbImgList::draw(CDC *aDC, xpr_sint_t aImageIndex, POINT aPoint, xpr_uint_t aStyle)
{
    return mImageList.Draw(aDC, aImageIndex, aPoint, aStyle);
}

// if image remove in image list, then mImageIndex shoud change.
xpr_bool_t ThumbImgList::remove(xpr_sint_t aImageIndex)
{
    ThumbDeque::iterator sIterator;

    sIterator = mThumbDeque.begin();
    while (sIterator != mThumbDeque.end())
    {
        ThumbElement &sThumbElement = *sIterator;

        if (sThumbElement.mImageIndex == aImageIndex)
        {
            sIterator = mThumbDeque.erase(sIterator);
            break;
        }

        sIterator++;
    }

    sIterator = mThumbDeque.begin();
    for (; sIterator != mThumbDeque.end(); ++sIterator)
    {
        ThumbElement &sThumbElement = *sIterator;

        if (sThumbElement.mImageIndex > aImageIndex)
            sThumbElement.mImageIndex--;
    }

    return mImageList.Remove(aImageIndex);
}

void ThumbImgList::removeAll(void)
{
    if (XPR_IS_NOT_NULL(mImageList.m_hImageList))
    {
        xpr_sint_t i, sCount;

        sCount = mImageList.GetImageCount();
        if (sCount > 0)
        {
            for (i = sCount - 1; i >= 0; --i)
                mImageList.Remove(i);
        }
    }

    mThumbDeque.clear();
}

xpr_bool_t ThumbImgList::save(void)
{
    const xpr_sint_t sImageCount = mImageList.GetImageCount();
    if (sImageCount <= 0)
        return deleteCacheFiles();

    if ((xpr_size_t)sImageCount != mThumbDeque.size() ||
        mThumbDeque.size() > kMaxThumbnailRecords)
    {
        return XPR_FALSE;
    }

    // Reject an oversized generation before CImageList::Write creates a huge
    // temporary file. The image-list payload is at least a 32-bit bitmap per
    // cell; the conservative estimate deliberately ignores compression.
    xpr_sint_t sIconWidth = 0;
    xpr_sint_t sIconHeight = 0;
    if (::ImageList_GetIconSize(mImageList.m_hImageList,
                                &sIconWidth, &sIconHeight) == FALSE ||
        sIconWidth <= 0 || sIconHeight <= 0)
    {
        return XPR_FALSE;
    }

    const ULONGLONG sEstimatedPayload =
        (ULONGLONG)sImageCount * (ULONGLONG)sIconWidth *
        (ULONGLONG)sIconHeight * 4ULL;
    if (sEstimatedPayload > kMaxThumbnailDataSize)
        return XPR_FALSE;

    xpr::string sDataPath;
    xpr::string sIndexPath;
    if (getCachePaths(XPR_TRUE, sDataPath, sIndexPath) == XPR_FALSE)
        return XPR_FALSE;

    // Re-check free space at the actual save boundary. The setting may have
    // been accepted hours earlier and the volume can become nearly full in
    // the meantime. Preserve a 512 MiB safety reserve and room for both temp
    // payloads/transactional overhead.
    xpr_size_t sSeparator = sDataPath.find_last_of(XPR_STRING_LITERAL("\\/"));
    if (sSeparator == xpr::string::npos)
        return XPR_FALSE;

    xpr::string sCacheDirectory = sDataPath.substr(0, sSeparator + 1);
    ULARGE_INTEGER sAvailable = {0};
    const ULONGLONG kCacheSafetyReserve = 512ULL * 1024ULL * 1024ULL;
    const ULONGLONG sRequiredFreeSpace =
        kCacheSafetyReserve + (sEstimatedPayload * 2ULL) +
        kMaxThumbnailIndexSize;
    if (::GetDiskFreeSpaceEx(sCacheDirectory.c_str(), &sAvailable,
                             XPR_NULL, XPR_NULL) == FALSE ||
        sAvailable.QuadPart < sRequiredFreeSpace)
    {
        return XPR_FALSE;
    }

    xpr_tchar_t sSuffix[96] = {0};
    _stprintf_s(sSuffix, XPR_COUNT_OF(sSuffix),
                XPR_STRING_LITERAL(".%lu.%lu"),
                ::GetCurrentProcessId(), ::GetTickCount());

    xpr::string sTempDataPath   = sDataPath  + XPR_STRING_LITERAL(".tmp") + sSuffix;
    xpr::string sTempIndexPath  = sIndexPath + XPR_STRING_LITERAL(".tmp") + sSuffix;
    xpr::string sBackupDataPath = sDataPath  + XPR_STRING_LITERAL(".bak") + sSuffix;
    xpr::string sBackupIndexPath= sIndexPath + XPR_STRING_LITERAL(".bak") + sSuffix;

    if (sTempDataPath.length()    > XPR_MAX_PATH ||
        sTempIndexPath.length()   > XPR_MAX_PATH ||
        sBackupDataPath.length()  > XPR_MAX_PATH ||
        sBackupIndexPath.length() > XPR_MAX_PATH)
    {
        return XPR_FALSE;
    }

    ::DeleteFile(sTempDataPath.c_str());
    ::DeleteFile(sTempIndexPath.c_str());

    const ULONGLONG sGeneration = createCacheGeneration();
    xpr_bool_t sDataSaved = XPR_FALSE;

    try
    {
        CFile sFile(sTempDataPath.c_str(),
                    CFile::modeCreate | CFile::modeWrite |
                    CFile::shareExclusive | CFile::osSequentialScan);
        sFile.Write(&kThumbnailDataMagic,    sizeof(kThumbnailDataMagic));
        sFile.Write(&kThumbnailCacheVersion, sizeof(kThumbnailCacheVersion));
        sFile.Write(&sGeneration,            sizeof(sGeneration));

        CArchive sArchive(&sFile, CArchive::store);
        sDataSaved = mImageList.Write(&sArchive) ? XPR_TRUE : XPR_FALSE;
        if (XPR_IS_TRUE(sDataSaved))
            sArchive.Close();
        else
            sArchive.Abort();

        sFile.Flush();
        sFile.Close();
    }
    catch (CException *sException)
    {
        sException->Delete();
        sDataSaved = XPR_FALSE;
    }
    catch (...)
    {
        sDataSaved = XPR_FALSE;
    }

    ULONGLONG sDataSize = 0;
    if (XPR_IS_FALSE(sDataSaved) ||
        XPR_IS_FALSE(getFileSize(sTempDataPath.c_str(), sDataSize)) ||
        sDataSize == 0 || sDataSize > kMaxThumbnailDataSize)
    {
        ::DeleteFile(sTempDataPath.c_str());
        return XPR_FALSE;
    }

    HANDLE sIndexFile = ::CreateFile(sTempIndexPath.c_str(), GENERIC_WRITE, 0,
                                     XPR_NULL, CREATE_ALWAYS,
                                     FILE_ATTRIBUTE_TEMPORARY |
                                     FILE_FLAG_SEQUENTIAL_SCAN |
                                     FILE_FLAG_WRITE_THROUGH,
                                     XPR_NULL);
    if (sIndexFile == INVALID_HANDLE_VALUE)
    {
        ::DeleteFile(sTempDataPath.c_str());
        return XPR_FALSE;
    }

    const DWORD sRecordCount = (DWORD)mThumbDeque.size();
    xpr_bool_t sIndexSaved =
        writeExact(sIndexFile, &kThumbnailIndexMagic,   sizeof(kThumbnailIndexMagic)) &&
        writeExact(sIndexFile, &kThumbnailCacheVersion, sizeof(kThumbnailCacheVersion)) &&
        writeExact(sIndexFile, &sGeneration,            sizeof(sGeneration)) &&
        writeExact(sIndexFile, &sRecordCount,           sizeof(sRecordCount));

    ThumbDeque::const_iterator sIterator = mThumbDeque.begin();
    for (; XPR_IS_TRUE(sIndexSaved) && sIterator != mThumbDeque.end(); ++sIterator)
    {
        const ThumbElement &sThumbElement = *sIterator;
        if (sThumbElement.mPath.empty() ||
            sThumbElement.mPath.length() > XPR_MAX_PATH ||
            sThumbElement.mImageIndex < 0 ||
            sThumbElement.mImageIndex >= sImageCount ||
            sThumbElement.mWidth <= 0 || sThumbElement.mHeight <= 0 ||
            sThumbElement.mDepth <= 0)
        {
            sIndexSaved = XPR_FALSE;
            break;
        }

        xpr_wchar_t sPathW[XPR_MAX_PATH + 1] = {0};
        xpr_size_t sInputBytes  = sThumbElement.mPath.length() * sizeof(xpr_tchar_t);
        xpr_size_t sOutputBytes = XPR_MAX_PATH * sizeof(xpr_wchar_t);
        XPR_TCS_TO_UTF16(sThumbElement.mPath.c_str(), &sInputBytes,
                         sPathW, &sOutputBytes);
        if (sOutputBytes > XPR_MAX_PATH * sizeof(xpr_wchar_t))
        {
            sIndexSaved = XPR_FALSE;
            break;
        }

        sPathW[sOutputBytes / sizeof(xpr_wchar_t)] = 0;
        const xpr_sint_t sPathBytes = (xpr_sint_t)(sOutputBytes + sizeof(xpr_wchar_t));

        sIndexSaved =
            writeExact(sIndexFile, &sPathBytes,                         sizeof(sPathBytes)) &&
            writeExact(sIndexFile, sPathW,                              (DWORD)sPathBytes) &&
            writeExact(sIndexFile, &sThumbElement.mThumbImageId,        sizeof(xpr_uint_t)) &&
            writeExact(sIndexFile, &sThumbElement.mImageIndex,          sizeof(xpr_sint_t)) &&
            writeExact(sIndexFile, &sThumbElement.mWidth,               sizeof(xpr_sint_t)) &&
            writeExact(sIndexFile, &sThumbElement.mHeight,              sizeof(xpr_sint_t)) &&
            writeExact(sIndexFile, &sThumbElement.mDepth,               sizeof(xpr_sint_t)) &&
            writeExact(sIndexFile, &sThumbElement.mModifiedFileTime,    sizeof(FILETIME));
    }

    if (XPR_IS_TRUE(sIndexSaved))
        sIndexSaved = (::FlushFileBuffers(sIndexFile) != FALSE) ? XPR_TRUE : XPR_FALSE;
    ::CloseHandle(sIndexFile);

    ULONGLONG sIndexSize = 0;
    if (XPR_IS_FALSE(sIndexSaved) ||
        XPR_IS_FALSE(getFileSize(sTempIndexPath.c_str(), sIndexSize)) ||
        sIndexSize == 0 || sIndexSize > kMaxThumbnailIndexSize)
    {
        ::DeleteFile(sTempDataPath.c_str());
        ::DeleteFile(sTempIndexPath.c_str());
        return XPR_FALSE;
    }

    const xpr_bool_t sCommitted = commitCachePair(
        sDataPath.c_str(), sIndexPath.c_str(),
        sTempDataPath.c_str(), sTempIndexPath.c_str(),
        sBackupDataPath.c_str(), sBackupIndexPath.c_str());

    if (XPR_IS_FALSE(sCommitted))
    {
        ::DeleteFile(sTempDataPath.c_str());
        ::DeleteFile(sTempIndexPath.c_str());
    }

    return sCommitted;
}

xpr_bool_t ThumbImgList::load(void)
{
    xpr::string sDataPath;
    xpr::string sIndexPath;
    if (getCachePaths(XPR_FALSE, sDataPath, sIndexPath) == XPR_FALSE ||
        XPR_IS_FALSE(fileExists(sDataPath.c_str())) ||
        XPR_IS_FALSE(fileExists(sIndexPath.c_str())))
    {
        return XPR_FALSE;
    }

    ULONGLONG sDataSize  = 0;
    ULONGLONG sIndexSize = 0;
    if (XPR_IS_FALSE(getFileSize(sDataPath.c_str(), sDataSize)) ||
        XPR_IS_FALSE(getFileSize(sIndexPath.c_str(), sIndexSize)) ||
        sDataSize == 0 || sDataSize > kMaxThumbnailDataSize ||
        sIndexSize == 0 || sIndexSize > kMaxThumbnailIndexSize)
    {
        return XPR_FALSE;
    }

    HANDLE sIndexFile = ::CreateFile(sIndexPath.c_str(), GENERIC_READ,
                                     FILE_SHARE_READ | FILE_SHARE_DELETE,
                                     XPR_NULL, OPEN_EXISTING,
                                     FILE_FLAG_SEQUENTIAL_SCAN, XPR_NULL);
    if (sIndexFile == INVALID_HANDLE_VALUE)
        return XPR_FALSE;

    LARGE_INTEGER sZero = {0};
    DWORD sFirst = 0;
    xpr_bool_t sIndexValid = readExact(sIndexFile, &sFirst, sizeof(sFirst));
    xpr_bool_t sNewFormat = XPR_FALSE;
    ULONGLONG sGeneration = 0;
    DWORD sDeclaredCount = 0;

    if (XPR_IS_TRUE(sIndexValid) && sFirst == kThumbnailIndexMagic)
    {
        DWORD sVersion = 0;
        sNewFormat = XPR_TRUE;
        sIndexValid =
            readExact(sIndexFile, &sVersion,       sizeof(sVersion)) &&
            readExact(sIndexFile, &sGeneration,    sizeof(sGeneration)) &&
            readExact(sIndexFile, &sDeclaredCount, sizeof(sDeclaredCount)) &&
            sVersion == kThumbnailCacheVersion &&
            sDeclaredCount <= kMaxThumbnailRecords;
    }
    else if (XPR_IS_TRUE(sIndexValid))
    {
        sIndexValid = (::SetFilePointerEx(sIndexFile, sZero, XPR_NULL, FILE_BEGIN) != FALSE)
                    ? XPR_TRUE : XPR_FALSE;
    }

    ThumbDeque sLoadedThumbs;
    while (XPR_IS_TRUE(sIndexValid))
    {
        LARGE_INTEGER sPosition = {0};
        if (::SetFilePointerEx(sIndexFile, sZero, &sPosition, FILE_CURRENT) == FALSE)
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        if ((ULONGLONG)sPosition.QuadPart == sIndexSize)
            break;
        if ((ULONGLONG)sPosition.QuadPart > sIndexSize ||
            sLoadedThumbs.size() >= kMaxThumbnailRecords ||
            (XPR_IS_TRUE(sNewFormat) && sLoadedThumbs.size() >= sDeclaredCount))
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        xpr_sint_t sPathBytes = 0;
        if (XPR_IS_FALSE(readExact(sIndexFile, &sPathBytes, sizeof(sPathBytes))) ||
            sPathBytes < (xpr_sint_t)sizeof(xpr_wchar_t) ||
            sPathBytes > (xpr_sint_t)((XPR_MAX_PATH + 1) * sizeof(xpr_wchar_t)) ||
            (sPathBytes % sizeof(xpr_wchar_t)) != 0)
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        const xpr_size_t sPathChars = sPathBytes / sizeof(xpr_wchar_t);
        std::vector<xpr_wchar_t> sPathW(sPathChars, 0);
        ThumbElement sThumbElement;
        memset(&sThumbElement.mModifiedFileTime, 0, sizeof(FILETIME));

        if (XPR_IS_FALSE(readExact(sIndexFile, &sPathW[0], (DWORD)sPathBytes)) ||
            sPathW[sPathChars - 1] != 0 ||
            (XPR_IS_TRUE(sNewFormat) &&
             XPR_IS_FALSE(readExact(sIndexFile, &sThumbElement.mThumbImageId, sizeof(xpr_uint_t)))) ||
            XPR_IS_FALSE(readExact(sIndexFile, &sThumbElement.mImageIndex, sizeof(xpr_sint_t))) ||
            XPR_IS_FALSE(readExact(sIndexFile, &sThumbElement.mWidth,      sizeof(xpr_sint_t))) ||
            XPR_IS_FALSE(readExact(sIndexFile, &sThumbElement.mHeight,     sizeof(xpr_sint_t))) ||
            XPR_IS_FALSE(readExact(sIndexFile, &sThumbElement.mDepth,      sizeof(xpr_sint_t))) ||
            XPR_IS_FALSE(readExact(sIndexFile, &sThumbElement.mModifiedFileTime, sizeof(FILETIME))) ||
            sThumbElement.mImageIndex < 0 ||
            sThumbElement.mWidth <= 0 || sThumbElement.mWidth > 32768 ||
            sThumbElement.mHeight <= 0 || sThumbElement.mHeight > 32768 ||
            sThumbElement.mDepth <= 0 || sThumbElement.mDepth > 128)
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        xpr_size_t sActualChars = 0;
        while (sActualChars < sPathChars && sPathW[sActualChars] != 0)
            ++sActualChars;
        if (sActualChars == 0 || sActualChars != sPathChars - 1)
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        xpr_tchar_t sPath[XPR_MAX_PATH + 1] = {0};
        xpr_size_t sInputBytes  = sActualChars * sizeof(xpr_wchar_t);
        xpr_size_t sOutputBytes = XPR_MAX_PATH * sizeof(xpr_tchar_t);
        XPR_UTF16_TO_TCS(&sPathW[0], &sInputBytes, sPath, &sOutputBytes);
        if (sOutputBytes > XPR_MAX_PATH * sizeof(xpr_tchar_t))
        {
            sIndexValid = XPR_FALSE;
            break;
        }
        sPath[sOutputBytes / sizeof(xpr_tchar_t)] = 0;
        if (sPath[0] == 0)
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        if (sThumbElement.mImageIndex != (xpr_sint_t)sLoadedThumbs.size())
        {
            sIndexValid = XPR_FALSE;
            break;
        }

        sThumbElement.mPath = sPath;
        if (XPR_IS_FALSE(sNewFormat))
            sThumbElement.mThumbImageId = (xpr_uint_t)sThumbElement.mImageIndex;
        sLoadedThumbs.push_back(sThumbElement);
    }

    ::CloseHandle(sIndexFile);

    if (XPR_IS_FALSE(sIndexValid) || sLoadedThumbs.empty() ||
        (XPR_IS_TRUE(sNewFormat) && sLoadedThumbs.size() != sDeclaredCount))
    {
        return XPR_FALSE;
    }

    CImageList sLoadedImages;
    xpr_bool_t sDataValid = XPR_FALSE;
    try
    {
        CFile sFile(sDataPath.c_str(),
                    CFile::modeRead | CFile::shareDenyWrite | CFile::osSequentialScan);

        if (XPR_IS_TRUE(sNewFormat))
        {
            DWORD sMagic = 0;
            DWORD sVersion = 0;
            ULONGLONG sDataGeneration = 0;
            sDataValid =
                sFile.Read(&sMagic,          sizeof(sMagic)) == sizeof(sMagic) &&
                sFile.Read(&sVersion,        sizeof(sVersion)) == sizeof(sVersion) &&
                sFile.Read(&sDataGeneration, sizeof(sDataGeneration)) == sizeof(sDataGeneration) &&
                sMagic == kThumbnailDataMagic &&
                sVersion == kThumbnailCacheVersion &&
                sDataGeneration == sGeneration;
        }
        else
        {
            sFile.SeekToBegin();
            sDataValid = XPR_TRUE;
        }

        if (XPR_IS_TRUE(sDataValid))
        {
            CArchive sArchive(&sFile, CArchive::load);
            sDataValid = sLoadedImages.Read(&sArchive) ? XPR_TRUE : XPR_FALSE;
            if (XPR_IS_TRUE(sDataValid))
                sArchive.Close();
            else
                sArchive.Abort();
        }

        sFile.Close();
    }
    catch (CException *sException)
    {
        sException->Delete();
        sDataValid = XPR_FALSE;
    }
    catch (...)
    {
        sDataValid = XPR_FALSE;
    }

    const xpr_sint_t sImageCount = sLoadedImages.GetImageCount();
    if (XPR_IS_FALSE(sDataValid) || sImageCount <= 0 ||
        (xpr_size_t)sImageCount != sLoadedThumbs.size())
    {
        return XPR_FALSE;
    }

    std::vector<xpr_bool_t> sUsedImageIndexes(sImageCount, XPR_FALSE);
    std::set<xpr_uint_t> sUsedThumbImageIds;
    ThumbDeque::const_iterator sIterator = sLoadedThumbs.begin();
    for (; sIterator != sLoadedThumbs.end(); ++sIterator)
    {
        const xpr_sint_t sImageIndex = sIterator->mImageIndex;
        if (sImageIndex < 0 || sImageIndex >= sImageCount ||
            XPR_IS_TRUE(sUsedImageIndexes[sImageIndex]) ||
            sIterator->mThumbImageId == kInvalidThumbImageId ||
            !sUsedThumbImageIds.insert(sIterator->mThumbImageId).second)
        {
            return XPR_FALSE;
        }
        sUsedImageIndexes[sImageIndex] = XPR_TRUE;
    }

    HIMAGELIST sImageListHandle = sLoadedImages.Detach();
    if (XPR_IS_NULL(sImageListHandle))
        return XPR_FALSE;

    mImageList.DeleteImageList();
    if (mImageList.Attach(sImageListHandle) == FALSE)
    {
        ::ImageList_Destroy(sImageListHandle);
        return XPR_FALSE;
    }

    mThumbDeque.swap(sLoadedThumbs);
    return XPR_TRUE;
}

xpr_bool_t ThumbImgList::deleteCacheFiles(void)
{
    xpr::string sDataPath;
    xpr::string sIndexPath;
    if (XPR_IS_FALSE(getCachePaths(XPR_TRUE, sDataPath, sIndexPath)))
        return XPR_FALSE;

    xpr_bool_t sDataDeleted = XPR_TRUE;
    xpr_bool_t sIndexDeleted = XPR_TRUE;

    if (XPR_IS_TRUE(fileExists(sDataPath.c_str())))
        sDataDeleted = (::DeleteFile(sDataPath.c_str()) != FALSE) ? XPR_TRUE : XPR_FALSE;
    if (XPR_IS_TRUE(fileExists(sIndexPath.c_str())))
        sIndexDeleted = (::DeleteFile(sIndexPath.c_str()) != FALSE) ? XPR_TRUE : XPR_FALSE;

    return (XPR_IS_TRUE(sDataDeleted) && XPR_IS_TRUE(sIndexDeleted))
         ? XPR_TRUE : XPR_FALSE;
}

void ThumbImgList::verify(void)
{
    xpr_sint_t i;
    ThumbDeque::iterator sIterator;
    ThumbDeque::iterator sIterator2;

    i = 0;
    sIterator = mThumbDeque.begin();
    while (sIterator != mThumbDeque.end())
    {
        ThumbElement &sThumbElement = *sIterator;

        if (IsExistFile(sThumbElement.mPath.c_str()) == XPR_FALSE)
        {
            sIterator = mThumbDeque.erase(sIterator);

            sIterator2 = mThumbDeque.begin();
            for (; sIterator2 != mThumbDeque.end(); ++sIterator2)
            {
                ThumbElement &sThumbElement2 = *sIterator2;

                if (sThumbElement2.mImageIndex > i)
                    sThumbElement2.mImageIndex--;
            }

            mImageList.Remove(i);

            continue;
        }

        sIterator++;
        i++;
    }
}

xpr_bool_t ThumbImgList::clear(void)
{
    mThumbDeque.clear();

    return mImageList.DeleteImageList();
}
} // namespace fxfile
