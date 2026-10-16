//
// SHA-256 based file integrity monitoring dialog.
//

#include "stdafx.h"
#include "file_integrity_dlg.h"

#include "resource.h"

#include <bcrypt.h>
#pragma comment(lib, "bcrypt.lib")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cmd
{
namespace
{
const UINT_PTR kMonitorTimerId = 1;
const UINT     kMonitorInterval = 3000;

xpr_bool_t calculateSha256(
    const xpr_tchar_t *aPath,
    xpr::string       &aHash,
    ULONGLONG         &aSize,
    FILETIME          &aWriteTime)
{
    aHash.clear();
    aSize = 0;
    ::ZeroMemory(&aWriteTime, sizeof(aWriteTime));

    HANDLE sFile = ::CreateFile(
        aPath,
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        XPR_NULL,
        OPEN_EXISTING,
        FILE_FLAG_SEQUENTIAL_SCAN,
        XPR_NULL);

    if (sFile == INVALID_HANDLE_VALUE)
        return XPR_FALSE;

    LARGE_INTEGER sFileSize = {0};
    if (::GetFileSizeEx(sFile, &sFileSize) == XPR_FALSE ||
        ::GetFileTime(sFile, XPR_NULL, XPR_NULL, &aWriteTime) == XPR_FALSE)
    {
        ::CloseHandle(sFile);
        return XPR_FALSE;
    }
    aSize = (ULONGLONG)sFileSize.QuadPart;

    BCRYPT_ALG_HANDLE  sAlgorithm = XPR_NULL;
    BCRYPT_HASH_HANDLE sHash      = XPR_NULL;
    PUCHAR sHashObject = XPR_NULL;
    PUCHAR sHashBytes  = XPR_NULL;
    PUCHAR sBuffer     = XPR_NULL;
    DWORD sObjectSize = 0;
    DWORD sHashSize   = 0;
    DWORD sResultSize = 0;
    xpr_bool_t sSuccess = XPR_FALSE;

    if (BCryptOpenAlgorithmProvider(&sAlgorithm, BCRYPT_SHA256_ALGORITHM, XPR_NULL, 0) < 0)
        goto Cleanup;
    if (BCryptGetProperty(sAlgorithm, BCRYPT_OBJECT_LENGTH, (PUCHAR)&sObjectSize, sizeof(sObjectSize), &sResultSize, 0) < 0)
        goto Cleanup;
    if (BCryptGetProperty(sAlgorithm, BCRYPT_HASH_LENGTH, (PUCHAR)&sHashSize, sizeof(sHashSize), &sResultSize, 0) < 0)
        goto Cleanup;

    sHashObject = new UCHAR[sObjectSize];
    sHashBytes  = new UCHAR[sHashSize];
    sBuffer     = new UCHAR[1024 * 1024];
    if (XPR_IS_NULL(sHashObject) || XPR_IS_NULL(sHashBytes) || XPR_IS_NULL(sBuffer))
        goto Cleanup;

    if (BCryptCreateHash(sAlgorithm, &sHash, sHashObject, sObjectSize, XPR_NULL, 0, 0) < 0)
        goto Cleanup;

    for (;;)
    {
        DWORD sRead = 0;
        if (::ReadFile(sFile, sBuffer, 1024 * 1024, &sRead, XPR_NULL) == XPR_FALSE)
            goto Cleanup;
        if (sRead == 0)
            break;
        if (BCryptHashData(sHash, sBuffer, sRead, 0) < 0)
            goto Cleanup;
    }

    if (BCryptFinishHash(sHash, sHashBytes, sHashSize, 0) < 0)
        goto Cleanup;

    {
        static const xpr_tchar_t kHex[] = XPR_STRING_LITERAL("0123456789ABCDEF");
        xpr::string sText;
        sText.resize(sHashSize * 2);
        for (DWORD i = 0; i < sHashSize; ++i)
        {
            sText[i * 2]     = kHex[(sHashBytes[i] >> 4) & 0x0f];
            sText[i * 2 + 1] = kHex[sHashBytes[i] & 0x0f];
        }
        aHash.swap(sText);
    }

    sSuccess = XPR_TRUE;

Cleanup:
    if (XPR_IS_NOT_NULL(sHash))
        BCryptDestroyHash(sHash);
    if (XPR_IS_NOT_NULL(sAlgorithm))
        BCryptCloseAlgorithmProvider(sAlgorithm, 0);
    XPR_SAFE_DELETE_ARRAY(sHashObject);
    XPR_SAFE_DELETE_ARRAY(sHashBytes);
    XPR_SAFE_DELETE_ARRAY(sBuffer);
    ::CloseHandle(sFile);
    return sSuccess;
}
} // anonymous namespace

FileIntegrityDlg::Item::Item(void)
    : mBaselineSize(0), mCurrentSize(0), mBaselineValid(XPR_FALSE), mCurrentValid(XPR_FALSE)
{
    ::ZeroMemory(&mBaselineWriteTime, sizeof(mBaselineWriteTime));
    ::ZeroMemory(&mCurrentWriteTime, sizeof(mCurrentWriteTime));
}

FileIntegrityDlg::FileIntegrityDlg(void)
    : super(IDD_FILE_INTEGRITY_MONITOR, XPR_NULL), mBusy(XPR_FALSE)
{
}

FileIntegrityDlg::~FileIntegrityDlg(void)
{
}

void FileIntegrityDlg::addPath(const xpr_tchar_t *aPath)
{
    if (XPR_IS_NULL(aPath) || *aPath == XPR_STRING_LITERAL('\0'))
        return;

    Item sItem;
    sItem.mPath = aPath;
    mItems.push_back(sItem);
}

BEGIN_MESSAGE_MAP(FileIntegrityDlg, super)
    ON_BN_CLICKED(IDC_FIM_REFRESH, OnRefresh)
    ON_BN_CLICKED(IDC_FIM_RESET_BASELINE, OnResetBaseline)
    ON_BN_CLICKED(IDC_FIM_COPY_REPORT, OnCopyReport)
    ON_BN_CLICKED(IDC_FIM_AUTO_MONITOR, OnAutoMonitor)
    ON_WM_TIMER()
END_MESSAGE_MAP()

xpr_bool_t FileIntegrityDlg::OnInitDialog(void)
{
    super::OnInitDialog();

    SetWindowText(gApp.loadString(XPR_STRING_LITERAL("popup.fim.title")));
    SetDlgItemText(IDC_FIM_REFRESH,        gApp.loadString(XPR_STRING_LITERAL("popup.fim.button.refresh")));
    SetDlgItemText(IDC_FIM_RESET_BASELINE, gApp.loadString(XPR_STRING_LITERAL("popup.fim.button.reset_baseline")));
    SetDlgItemText(IDC_FIM_COPY_REPORT,    gApp.loadString(XPR_STRING_LITERAL("popup.fim.button.copy_report")));
    SetDlgItemText(IDC_FIM_AUTO_MONITOR,   gApp.loadString(XPR_STRING_LITERAL("popup.fim.check.auto")));
    SetDlgItemText(IDCANCEL,               gApp.loadString(XPR_STRING_LITERAL("popup.fim.button.close")));

    mListCtrl.SubclassDlgItem(IDC_FIM_LIST, this);
    mProgressCtrl.SubclassDlgItem(IDC_FIM_PROGRESS, this);
    mListCtrl.SetExtendedStyle(mListCtrl.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);
    mListCtrl.InsertColumn(0, gApp.loadString(XPR_STRING_LITERAL("popup.fim.column.file")),       LVCFMT_LEFT, 210);
    mListCtrl.InsertColumn(1, gApp.loadString(XPR_STRING_LITERAL("popup.fim.column.size")),       LVCFMT_RIGHT, 85);
    mListCtrl.InsertColumn(2, gApp.loadString(XPR_STRING_LITERAL("popup.fim.column.modified")),   LVCFMT_LEFT, 120);
    mListCtrl.InsertColumn(3, XPR_STRING_LITERAL("SHA-256"),                                     LVCFMT_LEFT, 330);
    mListCtrl.InsertColumn(4, gApp.loadString(XPR_STRING_LITERAL("popup.fim.column.integrity")),  LVCFMT_LEFT, 105);
    mListCtrl.InsertColumn(5, gApp.loadString(XPR_STRING_LITERAL("popup.fim.column.difference")), LVCFMT_LEFT, 130);

    mProgressCtrl.SetRange32(0, (xpr_sint_t)mItems.size());
    scan(XPR_TRUE);
    return XPR_TRUE;
}

void FileIntegrityDlg::OnCancel(void)
{
    KillTimer(kMonitorTimerId);
    super::OnCancel();
}

void FileIntegrityDlg::OnRefresh(void)
{
    scan(XPR_FALSE);
}

void FileIntegrityDlg::OnResetBaseline(void)
{
    scan(XPR_TRUE);
}

void FileIntegrityDlg::OnAutoMonitor(void)
{
    if (((CButton *)GetDlgItem(IDC_FIM_AUTO_MONITOR))->GetCheck() == BST_CHECKED)
        SetTimer(kMonitorTimerId, kMonitorInterval, XPR_NULL);
    else
        KillTimer(kMonitorTimerId);
}

void FileIntegrityDlg::OnTimer(UINT_PTR aEventId)
{
    if (aEventId == kMonitorTimerId && XPR_IS_FALSE(mBusy))
        scan(XPR_FALSE);
    else
        super::OnTimer(aEventId);
}

void FileIntegrityDlg::scan(xpr_bool_t aResetBaseline)
{
    if (XPR_IS_TRUE(mBusy))
        return;

    mBusy = XPR_TRUE;
    CWaitCursor sWaitCursor;
    mProgressCtrl.SetPos(0);

    for (xpr_size_t i = 0; i < mItems.size(); ++i)
    {
        Item &sItem = mItems[i];
        sItem.mCurrentValid = calculateSha256(sItem.mPath.c_str(), sItem.mCurrentHash, sItem.mCurrentSize, sItem.mCurrentWriteTime);

        if (XPR_IS_TRUE(aResetBaseline))
        {
            sItem.mBaselineValid     = sItem.mCurrentValid;
            sItem.mBaselineHash      = sItem.mCurrentHash;
            sItem.mBaselineSize      = sItem.mCurrentSize;
            sItem.mBaselineWriteTime = sItem.mCurrentWriteTime;
        }

        updateRow((xpr_sint_t)i);
        mProgressCtrl.SetPos((xpr_sint_t)i + 1);
    }

    updateStatus();
    mBusy = XPR_FALSE;
}

xpr::string FileIntegrityDlg::formatFileTime(const FILETIME &aFileTime) const
{
    FILETIME sLocalFileTime = {0};
    SYSTEMTIME sSystemTime = {0};
    if (::FileTimeToLocalFileTime(&aFileTime, &sLocalFileTime) == XPR_FALSE ||
        ::FileTimeToSystemTime(&sLocalFileTime, &sSystemTime) == XPR_FALSE)
        return XPR_STRING_LITERAL("-");

    xpr_tchar_t sText[64] = {0};
    _sntprintf_s(sText, XPR_COUNT_OF(sText), _TRUNCATE,
        XPR_STRING_LITERAL("%04d-%02d-%02d %02d:%02d:%02d"),
        sSystemTime.wYear, sSystemTime.wMonth, sSystemTime.wDay,
        sSystemTime.wHour, sSystemTime.wMinute, sSystemTime.wSecond);
    return sText;
}

xpr::string FileIntegrityDlg::getIntegrityStatus(const Item &aItem) const
{
    if (XPR_IS_FALSE(aItem.mCurrentValid))
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.status.unreadable"));
    if (XPR_IS_FALSE(aItem.mBaselineValid))
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.status.no_baseline"));
    if (aItem.mCurrentSize != aItem.mBaselineSize)
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.status.size_changed"));
    if (aItem.mCurrentHash != aItem.mBaselineHash)
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.status.content_changed"));
    return gApp.loadString(XPR_STRING_LITERAL("popup.fim.status.unchanged"));
}

xpr::string FileIntegrityDlg::getDifference(xpr_sint_t aIndex) const
{
    if (mItems.empty())
        return XPR_STRING_LITERAL("");

    const Item &sItem = mItems[aIndex];
    const Item &sReference = mItems[0];
    if (XPR_IS_FALSE(sItem.mCurrentValid))
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.diff.unavailable"));
    if (aIndex == 0)
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.diff.reference"));
    if (XPR_IS_FALSE(sReference.mCurrentValid))
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.diff.no_reference"));
    if (sItem.mCurrentSize != sReference.mCurrentSize)
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.diff.size"));
    if (sItem.mCurrentHash == sReference.mCurrentHash)
        return gApp.loadString(XPR_STRING_LITERAL("popup.fim.diff.identical"));
    return gApp.loadString(XPR_STRING_LITERAL("popup.fim.diff.content"));
}

void FileIntegrityDlg::updateRow(xpr_sint_t aIndex)
{
    Item &sItem = mItems[aIndex];
    if (aIndex >= mListCtrl.GetItemCount())
        mListCtrl.InsertItem(aIndex, sItem.mPath.c_str());
    else
        mListCtrl.SetItemText(aIndex, 0, sItem.mPath.c_str());

    xpr_tchar_t sSize[64] = {0};
    if (XPR_IS_TRUE(sItem.mCurrentValid))
        _sntprintf_s(sSize, XPR_COUNT_OF(sSize), _TRUNCATE, XPR_STRING_LITERAL("%I64u"), sItem.mCurrentSize);
    else
        _tcscpy_s(sSize, XPR_STRING_LITERAL("-"));

    mListCtrl.SetItemText(aIndex, 1, sSize);
    mListCtrl.SetItemText(aIndex, 2, XPR_IS_TRUE(sItem.mCurrentValid) ? formatFileTime(sItem.mCurrentWriteTime).c_str() : XPR_STRING_LITERAL("-"));
    mListCtrl.SetItemText(aIndex, 3, XPR_IS_TRUE(sItem.mCurrentValid) ? sItem.mCurrentHash.c_str() : XPR_STRING_LITERAL("-"));
    mListCtrl.SetItemText(aIndex, 4, getIntegrityStatus(sItem).c_str());
    mListCtrl.SetItemText(aIndex, 5, getDifference(aIndex).c_str());
}

void FileIntegrityDlg::updateStatus(void)
{
    xpr_sint_t sUnchanged = 0;
    xpr_sint_t sChanged = 0;
    xpr_sint_t sFailed = 0;
    for (ItemVector::const_iterator sIterator = mItems.begin(); sIterator != mItems.end(); ++sIterator)
    {
        const Item &sItem = *sIterator;
        if (XPR_IS_FALSE(sItem.mCurrentValid) || XPR_IS_FALSE(sItem.mBaselineValid))
            ++sFailed;
        else if (sItem.mCurrentSize == sItem.mBaselineSize && sItem.mCurrentHash == sItem.mBaselineHash)
            ++sUnchanged;
        else
            ++sChanged;
    }

    xpr_tchar_t sText[256] = {0};
    _sntprintf_s(sText, XPR_COUNT_OF(sText), _TRUNCATE,
        gApp.loadString(XPR_STRING_LITERAL("popup.fim.summary")),
        (xpr_sint_t)mItems.size(), sUnchanged, sChanged, sFailed);
    SetDlgItemText(IDC_FIM_STATUS, sText);
}

void FileIntegrityDlg::OnCopyReport(void)
{
    xpr::string sReport = gApp.loadString(XPR_STRING_LITERAL("popup.fim.report.header"));
    sReport += XPR_STRING_LITERAL("\r\n");
    for (xpr_size_t i = 0; i < mItems.size(); ++i)
    {
        const Item &sItem = mItems[i];
        sReport += sItem.mPath;
        sReport += XPR_STRING_LITERAL("\t");
        sReport += sItem.mCurrentHash.empty() ? XPR_STRING_LITERAL("-") : sItem.mCurrentHash;
        sReport += XPR_STRING_LITERAL("\t");
        sReport += getIntegrityStatus(sItem);
        sReport += XPR_STRING_LITERAL("\t");
        sReport += getDifference((xpr_sint_t)i);
        sReport += XPR_STRING_LITERAL("\r\n");
    }

    if (OpenClipboard() == XPR_FALSE)
        return;
    EmptyClipboard();
    const SIZE_T sBytes = (sReport.length() + 1) * sizeof(xpr_tchar_t);
    HGLOBAL sMemory = GlobalAlloc(GMEM_MOVEABLE, sBytes);
    if (XPR_IS_NOT_NULL(sMemory))
    {
        void *sData = GlobalLock(sMemory);
        if (XPR_IS_NOT_NULL(sData))
        {
            memcpy(sData, sReport.c_str(), sBytes);
            GlobalUnlock(sMemory);
#if defined(UNICODE) || defined(_UNICODE)
            SetClipboardData(CF_UNICODETEXT, sMemory);
#else
            SetClipboardData(CF_TEXT, sMemory);
#endif
            sMemory = XPR_NULL;
        }
    }
    if (XPR_IS_NOT_NULL(sMemory))
        GlobalFree(sMemory);
    CloseClipboard();
}
} // namespace cmd
} // namespace fxfile
