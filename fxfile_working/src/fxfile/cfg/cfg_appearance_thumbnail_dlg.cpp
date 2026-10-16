//
// Copyright (c) 2001-2013 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.

#include "stdafx.h"
#include "cfg_appearance_thumbnail_dlg.h"

#include "../thumbnail.h"
#include "../option.h"
#include "../resource.h"
#include "../explorer_ctrl.h"
#include "../shell.h"

#include "gfl/libgfl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cfg
{
CfgAppearanceThumbnailDlg::CfgAppearanceThumbnailDlg(void)
    : super(IDD_CFG_APPEARANCE_THUMBNAIL, XPR_NULL)
{
}

void CfgAppearanceThumbnailDlg::DoDataExchange(CDataExchange* pDX)
{
    super::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_CFG_THUMBNAIL_PRIORITY,        mPriorityComboBox);
    DDX_Control(pDX, IDC_CFG_THUMBNAIL_SUPPORTED_IMAGE, mImageFormatComboBox);
}

BEGIN_MESSAGE_MAP(CfgAppearanceThumbnailDlg, super)
    ON_BN_CLICKED(IDC_CFG_THUMBNAIL_CACHE_CLEAN, OnCacheClean)
    ON_BN_CLICKED(IDC_CFG_THUMBNAIL_CACHE_INIT,  OnCacheInit)
    ON_BN_CLICKED(IDC_CFG_THUMBNAIL_CACHE_PATH_BROWSE, OnCachePathBrowse)
END_MESSAGE_MAP()

xpr_bool_t CfgAppearanceThumbnailDlg::OnInitDialog(void) 
{
    super::OnInitDialog();

    // disable apply button event
    addIgnoreApply(IDC_CFG_THUMBNAIL_SUPPORTED_IMAGE);
    addIgnoreApply(IDC_CFG_THUMBNAIL_CACHE_CLEAN);
    addIgnoreApply(IDC_CFG_THUMBNAIL_CACHE_INIT);
    addIgnoreApply(IDC_CFG_THUMBNAIL_CACHE_PATH_BROWSE);

    ((CEdit *)GetDlgItem(IDC_CFG_THUMBNAIL_CACHE_PATH))->LimitText(XPR_MAX_PATH);

    CComboBox *sComboBox;
    sComboBox = (CComboBox *)GetDlgItem(IDC_CFG_THUMBNAIL_PRIORITY);
    sComboBox->AddString(gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.priority.low")));
    sComboBox->AddString(gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.priority.normal")));
    sComboBox->AddString(gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.priority.high")));

    CSpinButtonCtrl *sSpinCtrl;
    sSpinCtrl = (CSpinButtonCtrl *)GetDlgItem(IDC_CFG_THUMBNAIL_WIDTH_SPIN);
    sSpinCtrl->SetRange(MIN_THUMB_SIZE, MAX_THUMB_SIZE);
    sSpinCtrl = (CSpinButtonCtrl *)GetDlgItem(IDC_CFG_THUMBNAIL_HEIGHT_SPIN);
    sSpinCtrl->SetRange(MIN_THUMB_SIZE, MAX_THUMB_SIZE);

    xpr_sint_t i, sCount;
    xpr_tchar_t sText[0xff] = {0};
    xpr_tchar_t sDesc[0xff] = {0};
    xpr_tchar_t sExt[0xff] = {0};
    GFL_FORMAT_INFORMATION sGflFormatInformation;
    xpr_size_t sInputBytes;
    xpr_size_t sOutputBytes;

    sCount = gflGetNumberOfFormat();
    for (i = 0; i < sCount; ++i)
    {
        memset(&sGflFormatInformation, 0, sizeof(GFL_FORMAT_INFORMATION));
        if (gflGetFormatInformationByIndex(i, &sGflFormatInformation) != GFL_NO_ERROR)
            continue;

        sInputBytes = strlen(sGflFormatInformation.Description) * sizeof(xpr_char_t);
        sOutputBytes = 0xfe * sizeof(xpr_tchar_t);
        XPR_MBS_TO_TCS(sGflFormatInformation.Description, &sInputBytes, sDesc, &sOutputBytes);
        sDesc[sOutputBytes / sizeof(xpr_tchar_t)] = 0;

        sInputBytes = strlen(sGflFormatInformation.Extension[0]) * sizeof(xpr_char_t);
        sOutputBytes = 0xfe * sizeof(xpr_tchar_t);
        XPR_MBS_TO_TCS(sGflFormatInformation.Extension[0], &sInputBytes, sExt, &sOutputBytes);
        sExt[sOutputBytes / sizeof(xpr_tchar_t)] = 0;

        _stprintf(sText, XPR_STRING_LITERAL("%s (*.%s)"), sDesc, sExt);
        mImageFormatComboBox.AddString(sText);
    }

    _stprintf(sText, gApp.loadFormatString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.limit_size"), XPR_STRING_LITERAL("%d,%d")), MIN_THUMB_SIZE, MAX_THUMB_SIZE);
    SetDlgItemText(IDC_CFG_THUMBNAIL_LIMIT_SIZE, sText);

    _stprintf(sText, gApp.loadFormatString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.supported_image_count"), XPR_STRING_LITERAL("%d")), sCount);
    SetDlgItemText(IDC_CFG_THUMBNAIL_LABEL_SUPPORTED_IMAGE_COUNT, sText);

    SetDlgItemText(IDC_CFG_THUMBNAIL_LABEL_WIDTH,           gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.width")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_LABEL_HEIGHT,          gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.height")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_LABEL_PRIORITY,        gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.priority")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_LABEL_CAUTION,         gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.caution")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_LABEL_SUPPORTED_IMAGE, gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.supported_image")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_LOAD_BY_EXT,           gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.check.load_by_extension")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE,                 gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.check.cache")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_CLEAN,           gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.button.cache_clean")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_INIT,            gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.button.cache_initalization")));
    SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_PATH_LABEL,      gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.label.cache_path")));

    return XPR_TRUE;
}

void CfgAppearanceThumbnailDlg::onInit(const Option::Config &aConfig)
{
    SetDlgItemInt(IDC_CFG_THUMBNAIL_WIDTH,  aConfig.mThumbnailWidth);
    SetDlgItemInt(IDC_CFG_THUMBNAIL_HEIGHT, aConfig.mThumbnailHeight);
    ((CButton *)GetDlgItem(IDC_CFG_THUMBNAIL_CACHE))->SetCheck(aConfig.mThumbnailSaveCache);
    ((CButton *)GetDlgItem(IDC_CFG_THUMBNAIL_LOAD_BY_EXT))->SetCheck(aConfig.mThumbnailLoadByExt);
    mPriorityComboBox.SetCurSel(aConfig.mThumbnailPriority);
    SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_PATH, aConfig.mThumbnailCachePath);
}

void CfgAppearanceThumbnailDlg::onApply(Option::Config &aConfig)
{
    aConfig.mThumbnailWidth     = GetDlgItemInt(IDC_CFG_THUMBNAIL_WIDTH);
    aConfig.mThumbnailHeight    = GetDlgItemInt(IDC_CFG_THUMBNAIL_HEIGHT);
    aConfig.mThumbnailSaveCache = ((CButton *)GetDlgItem(IDC_CFG_THUMBNAIL_CACHE))->GetCheck();
    aConfig.mThumbnailPriority  = mPriorityComboBox.GetCurSel();
    aConfig.mThumbnailLoadByExt = ((CButton *)GetDlgItem(IDC_CFG_THUMBNAIL_LOAD_BY_EXT))->GetCheck();

    xpr_tchar_t sCachePath[XPR_MAX_PATH + 1] = {0};
    GetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_PATH, sCachePath, XPR_MAX_PATH + 1);
    if (XPR_IS_TRUE(validateCachePath(sCachePath)))
    {
        _tcsncpy_s(aConfig.mThumbnailCachePath, XPR_MAX_PATH + 1,
                   sCachePath, _TRUNCATE);
    }
    else
    {
        SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_PATH, aConfig.mThumbnailCachePath);
        MessageBox(gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.msg.invalid_cache_path")),
                   XPR_NULL, MB_OK | MB_ICONSTOP);
    }

    if (aConfig.mThumbnailWidth  < MIN_THUMB_SIZE) aConfig.mThumbnailWidth  = MIN_THUMB_SIZE;
    if (aConfig.mThumbnailWidth  > MAX_THUMB_SIZE) aConfig.mThumbnailWidth  = MAX_THUMB_SIZE;
    if (aConfig.mThumbnailHeight < MIN_THUMB_SIZE) aConfig.mThumbnailHeight = MIN_THUMB_SIZE;
    if (aConfig.mThumbnailHeight > MAX_THUMB_SIZE) aConfig.mThumbnailHeight = MAX_THUMB_SIZE;
}

namespace
{
xpr_sint_t CALLBACK ThumbnailCacheBrowseCallback(HWND aHwnd, xpr_uint_t aMsg,
                                                  LPARAM aLParam, LPARAM aData)
{
    if (aMsg == BFFM_INITIALIZED && aData != 0)
        ::SendMessage(aHwnd, BFFM_SETSELECTION, XPR_FALSE, aData);

    return 0;
}
}

void CfgAppearanceThumbnailDlg::OnCachePathBrowse(void)
{
    xpr_tchar_t sCurrentPath[XPR_MAX_PATH + 1] = {0};
    GetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_PATH, sCurrentPath, XPR_MAX_PATH + 1);

    LPITEMIDLIST sOldPidl = XPR_NULL;
    if (sCurrentPath[0] != 0)
        sOldPidl = Path2Pidl(sCurrentPath);

    BROWSEINFO sBrowseInfo = {0};
    sBrowseInfo.hwndOwner = GetSafeHwnd();
    sBrowseInfo.ulFlags   = BIF_RETURNONLYFSDIRS;
    sBrowseInfo.lpszTitle = gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.folder_browse.title"));
    sBrowseInfo.lpfn      = (BFFCALLBACK)ThumbnailCacheBrowseCallback;
    sBrowseInfo.lParam    = (LPARAM)sOldPidl;

    LPITEMIDLIST sNewPidl = ::SHBrowseForFolder(&sBrowseInfo);
    if (XPR_IS_NOT_NULL(sNewPidl))
    {
        xpr_tchar_t sNewPath[XPR_MAX_PATH + 1] = {0};
        if (XPR_IS_TRUE(GetName(sNewPidl, SHGDN_FORPARSING, sNewPath)))
        {
            SetDlgItemText(IDC_CFG_THUMBNAIL_CACHE_PATH, sNewPath);
            setModified();
        }
    }

    COM_FREE(sOldPidl);
    COM_FREE(sNewPidl);
}

xpr_bool_t CfgAppearanceThumbnailDlg::validateCachePath(const xpr_tchar_t *aPath) const
{
    if (XPR_IS_NULL(aPath) || aPath[0] == 0)
        return XPR_TRUE;

    const xpr_size_t sLength = _tcslen(aPath);
    // Leave room for cache file names and transactional temp suffixes.
    if (sLength < 4 || sLength > XPR_MAX_PATH - 64 ||
        !((aPath[0] >= XPR_STRING_LITERAL('A') && aPath[0] <= XPR_STRING_LITERAL('Z')) ||
          (aPath[0] >= XPR_STRING_LITERAL('a') && aPath[0] <= XPR_STRING_LITERAL('z'))) ||
        aPath[1] != XPR_STRING_LITERAL(':') ||
        (aPath[2] != XPR_STRING_LITERAL('\\') && aPath[2] != XPR_STRING_LITERAL('/')))
    {
        return XPR_FALSE;
    }

    xpr_tchar_t sRoot[] = { aPath[0], XPR_STRING_LITERAL(':'), XPR_STRING_LITERAL('\\'), 0 };
    if (::GetDriveType(sRoot) != DRIVE_FIXED)
        return XPR_FALSE;

    DWORD sAttributes = ::GetFileAttributes(aPath);
    if (sAttributes == INVALID_FILE_ATTRIBUTES)
    {
        if (XPR_IS_FALSE(CreateDirectoryLevel(aPath, 0, XPR_TRUE)))
            return XPR_FALSE;
        sAttributes = ::GetFileAttributes(aPath);
    }

    if (sAttributes == INVALID_FILE_ATTRIBUTES ||
        !XPR_TEST_BITS(sAttributes, FILE_ATTRIBUTE_DIRECTORY) ||
        XPR_TEST_BITS(sAttributes, FILE_ATTRIBUTE_READONLY) ||
        XPR_TEST_BITS(sAttributes, FILE_ATTRIBUTE_REPARSE_POINT) ||
        XPR_TEST_BITS(sAttributes, FILE_ATTRIBUTE_OFFLINE))
    {
        return XPR_FALSE;
    }

    ULARGE_INTEGER sAvailable = {0};
    if (::GetDiskFreeSpaceEx(aPath, &sAvailable, XPR_NULL, XPR_NULL) == FALSE ||
        sAvailable.QuadPart < 1024ULL * 1024ULL * 1024ULL)
    {
        return XPR_FALSE;
    }

    xpr_tchar_t sProbePath[XPR_MAX_PATH + 1] = {0};
    _stprintf_s(sProbePath, XPR_COUNT_OF(sProbePath),
                XPR_STRING_LITERAL("%s%s.fxfile-cache-probe-%lu-%lu.tmp"),
                aPath,
                (aPath[sLength - 1] == XPR_STRING_LITERAL('\\') ||
                 aPath[sLength - 1] == XPR_STRING_LITERAL('/')) ? XPR_STRING_LITERAL("") : XPR_STRING_LITERAL("\\"),
                ::GetCurrentProcessId(), ::GetTickCount());

    HANDLE sProbe = ::CreateFile(sProbePath, GENERIC_WRITE, 0, XPR_NULL,
                                 CREATE_NEW, FILE_ATTRIBUTE_TEMPORARY |
                                 FILE_FLAG_DELETE_ON_CLOSE, XPR_NULL);
    if (sProbe == INVALID_HANDLE_VALUE)
        return XPR_FALSE;

    const xpr_char_t sMarker[] = "FXFILE_CACHE_PATH_PROBE";
    DWORD sWritten = 0;
    const xpr_bool_t sWriteSucceeded =
        (::WriteFile(sProbe, sMarker, sizeof(sMarker), &sWritten, XPR_NULL) != FALSE &&
         sWritten == sizeof(sMarker) && ::FlushFileBuffers(sProbe) != FALSE)
        ? XPR_TRUE : XPR_FALSE;
    ::CloseHandle(sProbe);
    ::DeleteFile(sProbePath);

    return sWriteSucceeded;
}

void CfgAppearanceThumbnailDlg::OnCacheClean(void) 
{
    CWaitCursor sWaitCursor;

    Thumbnail::instance().verifyThumb();

    const xpr_tchar_t *sMsg = gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.msg.completed_clean"));
    MessageBox(sMsg, XPR_NULL, MB_OK | MB_ICONINFORMATION);
}

void CfgAppearanceThumbnailDlg::OnCacheInit(void) 
{
    const xpr_tchar_t *sMsg;

    sMsg = gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.msg.warning_initialize"));
    xpr_sint_t sMsgId = MessageBox(sMsg, XPR_NULL, MB_YESNO | MB_ICONQUESTION);
    if (sMsgId != IDYES)
        return;

    CWaitCursor aWaitCursor;

    Thumbnail::instance().initThumb();

    sMsg = gApp.loadString(XPR_STRING_LITERAL("popup.cfg.body.appearnace.thumbnail.msg.completed_initialize"));
    MessageBox(sMsg, XPR_NULL, MB_OK | MB_ICONINFORMATION);
}
} // namespace cfg
} // namespace fxfile
