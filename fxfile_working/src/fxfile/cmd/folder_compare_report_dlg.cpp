//
// Copyright (c) 2012-2026 Leon Lee author. All rights reserved.
//
//   homepage: http://www.flychk.com
//   e-mail:   mailto:flychk@flychk.com
//
// Use of this source code is governed by a GPLv3 license that can be
// found in the LICENSE file.
//

#include "stdafx.h"
#include "folder_compare_report_dlg.h"
#include "main_frame.h"
#include "folder_sync_dlg.h"
#include "explorer_ctrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

namespace fxfile
{
namespace cmd
{
BEGIN_MESSAGE_MAP(FolderCompareReportDlg, CResizingDialog)
    ON_CBN_SELCHANGE(IDC_COMPARE_REPORT_FILTER,      OnSelchangeFilter)
    ON_BN_CLICKED(IDC_COMPARE_REPORT_BTN_COPY,       OnBtnCopyReport)
    ON_BN_CLICKED(IDC_COMPARE_REPORT_BTN_SAVE,       OnBtnSaveReport)
    ON_BN_CLICKED(IDC_COMPARE_REPORT_BTN_SELECT,     OnBtnSelectInWindow)
    ON_BN_CLICKED(IDC_COMPARE_REPORT_BTN_SYNC,       OnBtnSyncTool)
END_MESSAGE_MAP()

FolderCompareReportDlg::FolderCompareReportDlg(MainFrame *aMainFrame, CWnd *pParent)
    : super(FolderCompareReportDlg::IDD, pParent)
    , mMainFrame(aMainFrame)
    , mSyncDirs(XPR_NULL)
    , mCountEqual(0)
    , mCountDiff(0)
    , mCountLeftOnly(0)
    , mCountRightOnly(0)
    , mCountTotal(0)
    , mSizeEqual(0)
    , mSizeDiff(0)
    , mSizeLeftOnly(0)
    , mSizeRightOnly(0)
{
}

FolderCompareReportDlg::~FolderCompareReportDlg(void)
{
}

void FolderCompareReportDlg::setResult(SyncDirs *aSyncDirs, const xpr_tchar_t *aDir1, const xpr_tchar_t *aDir2)
{
    mSyncDirs = aSyncDirs;
    if (aDir1) mDir1 = aDir1;
    if (aDir2) mDir2 = aDir2;
}

void FolderCompareReportDlg::DoDataExchange(CDataExchange* pDX)
{
    super::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMPARE_REPORT_FILTER, mComboFilter);
    DDX_Control(pDX, IDC_COMPARE_REPORT_LIST,   mListCtrl);
}

CString FolderCompareReportDlg::formatFileSize(xpr_sint64_t aSize) const
{
    if (aSize < 0) return _T("-");
    if (aSize < 1024)
    {
        CString s;
        s.Format(_T("%lld B"), aSize);
        return s;
    }
    double kb = (double)aSize / 1024.0;
    if (kb < 1024.0)
    {
        CString s;
        s.Format(_T("%.1f KB"), kb);
        return s;
    }
    double mb = kb / 1024.0;
    if (mb < 1024.0)
    {
        CString s;
        s.Format(_T("%.2f MB"), mb);
        return s;
    }
    double gb = mb / 1024.0;
    CString s;
    s.Format(_T("%.2f GB"), gb);
    return s;
}

CString FolderCompareReportDlg::formatFileTime(const FILETIME &aFt) const
{
    if (aFt.dwLowDateTime == 0 && aFt.dwHighDateTime == 0)
        return _T("-");

    FILETIME localFt;
    FileTimeToLocalFileTime(&aFt, &localFt);
    SYSTEMTIME st;
    FileTimeToSystemTime(&localFt, &st);

    CString s;
    s.Format(_T("%04d-%02d-%02d %02d:%02d:%02d"),
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    return s;
}

void FolderCompareReportDlg::populateItems(void)
{
    mAllItems.clear();
    mCountEqual = 0;
    mCountDiff = 0;
    mCountLeftOnly = 0;
    mCountRightOnly = 0;
    mCountTotal = 0;

    mSizeEqual = 0;
    mSizeDiff = 0;
    mSizeLeftOnly = 0;
    mSizeRightOnly = 0;

    if (XPR_IS_NULL(mSyncDirs))
        return;

    size_t count = mSyncDirs->getCount();
    mCountTotal = count;

    for (size_t i = 0; i < count; ++i)
    {
        SyncItem *item = mSyncDirs->getSyncItem((xpr_sint_t)i);
        if (XPR_IS_NULL(item))
            continue;

        ItemViewData v;
        v.pItem = item;
        v.subPath = item->getSubPath();

        bool isLeft  = item->isExistLeft() == XPR_TRUE;
        bool isRight = item->isExistRight() == XPR_TRUE;

        if (isLeft)
        {
            v.size1Text = formatFileSize(item->mFileSize[0]);
            v.time1Text = formatFileTime(item->mModifiedFileTime[0]);
        }
        else
        {
            v.size1Text = _T("-");
            v.time1Text = _T("-");
        }

        if (isRight)
        {
            v.size2Text = formatFileSize(item->mFileSize[1]);
            v.time2Text = formatFileTime(item->mModifiedFileTime[1]);
        }
        else
        {
            v.size2Text = _T("-");
            v.time2Text = _T("-");
        }

        if (isLeft && isRight)
        {
            if (XPR_TEST_BITS(item->mDiff, CompareDiffEqualed))
            {
                v.statusText = _T("일치 (==)");
                v.diffReason = _T("내용·크기 동일");
                v.statusType = 0; // Equal
                mCountEqual++;
                mSizeEqual += item->mFileSize[0];
            }
            else
            {
                v.statusText = _T("불일치 (=/=)");
                v.statusType = 1; // Diff
                mCountDiff++;
                mSizeDiff += item->mFileSize[0];

                CString r;
                if (XPR_TEST_BITS(item->mDiff, CompareDiffSize))
                    r += _T("크기 상이; ");
                if (XPR_TEST_BITS(item->mDiff, CompareDiffTime))
                    r += _T("수정일시 상이; ");
                if (XPR_TEST_BITS(item->mDiff, CompareDiffContentsBytes) ||
                    XPR_TEST_BITS(item->mDiff, CompareDiffContentsHashMD5) ||
                    XPR_TEST_BITS(item->mDiff, CompareDiffContentsHashSFV))
                    r += _T("내용 상이 (데이터 불일치); ");
                if (XPR_TEST_BITS(item->mDiff, CompareDiffAttributes))
                    r += _T("파일 속성 상이; ");

                if (r.IsEmpty()) r = _T("세부 속성 상이");
                v.diffReason = r;
            }
        }
        else if (isLeft && !isRight)
        {
            v.statusText = _T("기준만 (<-)");
            v.diffReason = _T("대상(우측) 폴더에 없음 (누락)");
            v.statusType = 2; // LeftOnly
            mCountLeftOnly++;
            mSizeLeftOnly += item->mFileSize[0];
        }
        else if (!isLeft && isRight)
        {
            v.statusText = _T("대상만 (->)");
            v.diffReason = _T("기준(좌측) 폴더에 없음 (누락)");
            v.statusType = 3; // RightOnly
            mCountRightOnly++;
            mSizeRightOnly += item->mFileSize[1];
        }

        mAllItems.push_back(v);
    }
}

void FolderCompareReportDlg::applyFilter(int aFilterMode)
{
    mListCtrl.DeleteAllItems();
    mFilteredIndices.clear();

    for (size_t i = 0; i < mAllItems.size(); ++i)
    {
        const ItemViewData &v = mAllItems[i];
        bool include = false;

        switch (aFilterMode)
        {
        case 0: // 차이점만 보기 (기본)
            include = (v.statusType != 0);
            break;
        case 1: // 전체 보기
            include = true;
            break;
        case 2: // 불일치(=/=)만
            include = (v.statusType == 1);
            break;
        case 3: // 좌측만(<-)
            include = (v.statusType == 2);
            break;
        case 4: // 우측만(->)
            include = (v.statusType == 3);
            break;
        case 5: // 일치(==)만
            include = (v.statusType == 0);
            break;
        default:
            include = true;
            break;
        }

        if (include)
        {
            mFilteredIndices.push_back(i);
            int itemIdx = mListCtrl.InsertItem((int)mFilteredIndices.size() - 1, v.statusText);
            mListCtrl.SetItemText(itemIdx, 1, v.subPath);
            mListCtrl.SetItemText(itemIdx, 2, v.size1Text);
            mListCtrl.SetItemText(itemIdx, 3, v.time1Text);
            mListCtrl.SetItemText(itemIdx, 4, v.size2Text);
            mListCtrl.SetItemText(itemIdx, 5, v.time2Text);
            mListCtrl.SetItemText(itemIdx, 6, v.diffReason);
        }
    }
}

xpr_bool_t FolderCompareReportDlg::OnInitDialog(void)
{
    super::OnInitDialog();

    // 리사이징 등록
    AddControl(IDC_COMPARE_REPORT_PATH_INFO,    sizeResize, sizeNone);
    AddControl(IDC_COMPARE_REPORT_SUMMARY_TEXT, sizeResize, sizeNone);
    AddControl(IDC_COMPARE_REPORT_LIST,         sizeResize, sizeResize);
    AddControl(IDC_COMPARE_REPORT_BTN_COPY,     sizeNone,   sizeRepos);
    AddControl(IDC_COMPARE_REPORT_BTN_SAVE,     sizeNone,   sizeRepos);
    AddControl(IDC_COMPARE_REPORT_BTN_SELECT,   sizeNone,   sizeRepos);
    AddControl(IDC_COMPARE_REPORT_BTN_SYNC,     sizeNone,   sizeRepos);
    AddControl(IDCANCEL,                        sizeRepos,  sizeRepos);

    DWORD style = mListCtrl.GetExtendedStyle();
    style |= LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;
    mListCtrl.SetExtendedStyle(style);

    mListCtrl.InsertColumn(0, _T("비교 상태"),           LVCFMT_CENTER, 85);
    mListCtrl.InsertColumn(1, _T("하위 경로 및 파일명"),   LVCFMT_LEFT, 200);
    mListCtrl.InsertColumn(2, _T("기준 크기 (좌)"),       LVCFMT_RIGHT,  85);
    mListCtrl.InsertColumn(3, _T("기준 수정일시 (좌)"),   LVCFMT_LEFT, 130);
    mListCtrl.InsertColumn(4, _T("대상 크기 (우)"),       LVCFMT_RIGHT,  85);
    mListCtrl.InsertColumn(5, _T("대상 수정일시 (우)"),   LVCFMT_LEFT, 130);
    mListCtrl.InsertColumn(6, _T("비교 상세 사유"),       LVCFMT_LEFT, 150);

    // 경로 배너 설정
    CString pathBanner;
    pathBanner.Format(_T("[기준 폴더 (좌측)]: %s\n[대상 폴더 (우측)]: %s"), mDir1.c_str(), mDir2.c_str());
    SetDlgItemText(IDC_COMPARE_REPORT_PATH_INFO, pathBanner);

    // 데이터 분석 및 요약 집계
    populateItems();

    // 대시보드 통계 요약 텍스트 구성
    CString summaryText;
    summaryText.Format(
        _T("■ 총 비교 대상: %zu개  |  [완전 일치]: %zu개 (%s)\n")
        _T("■ 불일치(크기·시간·내용 상이): %zu개  |  [기준(좌측) 전용]: %zu개 (%s)  |  [대상(우측) 전용]: %zu개 (%s)"),
        mCountTotal, mCountEqual, (LPCTSTR)formatFileSize(mSizeEqual),
        mCountDiff,
        mCountLeftOnly, (LPCTSTR)formatFileSize(mSizeLeftOnly),
        mCountRightOnly, (LPCTSTR)formatFileSize(mSizeRightOnly));
    SetDlgItemText(IDC_COMPARE_REPORT_SUMMARY_TEXT, summaryText);

    // 콤보박스 필터 초기화
    mComboFilter.AddString(_T("차이점만 보기 (기본)"));
    mComboFilter.AddString(_T("전체 비교 목록 보기"));
    mComboFilter.AddString(_T("불일치 항목만 보기 (=/=)"));
    mComboFilter.AddString(_T("기준(좌측)에만 존재 (<-)"));
    mComboFilter.AddString(_T("대상(우측)에만 존재 (->)"));
    mComboFilter.AddString(_T("완전 일치 항목만 보기 (==)"));
    mComboFilter.SetCurSel(0);

    applyFilter(0);

    return XPR_TRUE;
}

void FolderCompareReportDlg::OnSelchangeFilter(void)
{
    int sel = mComboFilter.GetCurSel();
    if (sel >= 0)
    {
        applyFilter(sel);
    }
}

CString FolderCompareReportDlg::generateMarkdownReport(void) const
{
    CString md;
    md.Append(_T("# FxFile 폴더 비교 총괄 보고서 (Folder Comparison Summary Report)\r\n\r\n"));
    
    SYSTEMTIME st;
    GetLocalTime(&st);
    CString timeStr;
    timeStr.Format(_T("- **보고서 생성 일시**: %04d-%02d-%02d %02d:%02d:%02d\r\n"),
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    md.Append(timeStr);
    
    CString paths;
    paths.Format(_T("- **기준 폴더 (Left)**: `%s`\r\n- **대상 폴더 (Right)**: `%s`\r\n\r\n"),
        mDir1.c_str(), mDir2.c_str());
    md.Append(paths);

    md.Append(_T("## 1. 비교 요약 통계 대시보드\r\n\r\n"));
    md.Append(_T("| 구분 | 항목 수 | 총 파일 크기 |\r\n"));
    md.Append(_T("|---|---|---|\r\n"));
    
    CString row;
    row.Format(_T("| **총 비교 대상** | **%zu개** | - |\r\n"), mCountTotal);
    md.Append(row);
    row.Format(_T("| 완전 일치 (Equal) | %zu개 | %s |\r\n"), mCountEqual, (LPCTSTR)formatFileSize(mSizeEqual));
    md.Append(row);
    row.Format(_T("| 불일치 (Different) | %zu개 | %s |\r\n"), mCountDiff, (LPCTSTR)formatFileSize(mSizeDiff));
    md.Append(row);
    row.Format(_T("| 기준 폴더 전용 (Left Only) | %zu개 | %s |\r\n"), mCountLeftOnly, (LPCTSTR)formatFileSize(mSizeLeftOnly));
    md.Append(row);
    row.Format(_T("| 대상 폴더 전용 (Right Only) | %zu개 | %s |\r\n\r\n"), mCountRightOnly, (LPCTSTR)formatFileSize(mSizeRightOnly));
    md.Append(row);

    md.Append(_T("## 2. 차이 내역 상세 목록 (Differences Details)\r\n\r\n"));
    md.Append(_T("| 상태 | 상대 경로 및 파일명 | 기준 크기 (좌) | 기준 수정일시 | 대상 크기 (우) | 대상 수정일시 | 비교 상세 사유 |\r\n"));
    md.Append(_T("|---|---|---|---|---|---|---|\r\n"));

    for (size_t i = 0; i < mAllItems.size(); ++i)
    {
        const ItemViewData &v = mAllItems[i];
        if (v.statusType == 0) continue; // 일치 항목 제외하고 차이 항목만 리포트

        CString itemRow;
        itemRow.Format(_T("| %s | `%s` | %s | %s | %s | %s | %s |\r\n"),
            (LPCTSTR)v.statusText,
            (LPCTSTR)v.subPath,
            (LPCTSTR)v.size1Text,
            (LPCTSTR)v.time1Text,
            (LPCTSTR)v.size2Text,
            (LPCTSTR)v.time2Text,
            (LPCTSTR)v.diffReason);
        md.Append(itemRow);
    }

    md.Append(_T("\r\n---\r\n*Report generated by FxFile Modernized Folder Compare System*\r\n"));
    return md;
}

void FolderCompareReportDlg::OnBtnCopyReport(void)
{
    CString report = generateMarkdownReport();

    if (::OpenClipboard(GetSafeHwnd()))
    {
        ::EmptyClipboard();
        size_t bytes = (report.GetLength() + 1) * sizeof(TCHAR);
        HGLOBAL hMem = ::GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (hMem)
        {
            void *pDest = ::GlobalLock(hMem);
            memcpy(pDest, (LPCTSTR)report, bytes);
            ::GlobalUnlock(hMem);
#ifdef _UNICODE
            ::SetClipboardData(CF_UNICODETEXT, hMem);
#else
            ::SetClipboardData(CF_TEXT, hMem);
#endif
        }
        ::CloseClipboard();
        MessageBox(_T("비교 총괄 보고서가 클립보드에 마크다운 형식으로 복사되었습니다.\r\n(메모장, 마크다운 뷰어, 업무 보고서에 바로 붙여넣기 할 수 있습니다.)"),
            _T("보고서 복사 완료"), MB_OK | MB_ICONINFORMATION);
    }
    else
    {
        MessageBox(_T("클립보드를 열 수 없습니다."), _T("오류"), MB_OK | MB_ICONWARNING);
    }
}

void FolderCompareReportDlg::OnBtnSaveReport(void)
{
    CFileDialog dlg(FALSE, _T("txt"), _T("FolderCompareReport.txt"),
        OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
        _T("텍스트/마크다운 파일 (*.txt;*.md)|*.txt;*.md|모든 파일 (*.*)|*.*||"), this);

    if (dlg.DoModal() == IDOK)
    {
        CString path = dlg.GetPathName();
        CString report = generateMarkdownReport();

        FILE *fp = _tfopen(path, _T("wb"));
        if (fp)
        {
            // UTF-8 BOM
            unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
            fwrite(bom, 1, 3, fp);

            int utf8Len = WideCharToMultiByte(CP_UTF8, 0, (LPCTSTR)report, -1, NULL, 0, NULL, NULL);
            if (utf8Len > 0)
            {
                std::vector<char> buf(utf8Len);
                WideCharToMultiByte(CP_UTF8, 0, (LPCTSTR)report, -1, &buf[0], utf8Len, NULL, NULL);
                fwrite(&buf[0], 1, utf8Len - 1, fp);
            }
            fclose(fp);

            CString msg;
            msg.Format(_T("보고서가 다음 파일로 저장되었습니다:\n%s"), (LPCTSTR)path);
            MessageBox(msg, _T("보고서 저장 완료"), MB_OK | MB_ICONINFORMATION);
        }
        else
        {
            MessageBox(_T("파일을 저장할 수 없습니다."), _T("오류"), MB_OK | MB_ICONSTOP);
        }
    }
}

void FolderCompareReportDlg::OnBtnSelectInWindow(void)
{
    if (mMainFrame && mSyncDirs)
    {
        ExplorerCtrl *sExplorerCtrl[2];
        sExplorerCtrl[0] = mMainFrame->getExplorerCtrl();
        sExplorerCtrl[1] = mMainFrame->getExplorerCtrl(-2);

        if (XPR_IS_NOT_NULL(sExplorerCtrl[0])) sExplorerCtrl[0]->unselectAll();
        if (XPR_IS_NOT_NULL(sExplorerCtrl[1])) sExplorerCtrl[1]->unselectAll();

        xpr_bool_t sFirst[2] = { XPR_TRUE, XPR_TRUE };
        xpr::string sDir[2];
        mSyncDirs->getDir(sDir[0], sDir[1]);

        size_t count = mSyncDirs->getCount();
        for (size_t i = 0; i < count; ++i)
        {
            SyncItem *sSyncItem = mSyncDirs->getSyncItem((xpr_sint_t)i);
            if (XPR_IS_NULL(sSyncItem)) continue;

            for (int j = 0; j < 2; ++j)
            {
                if (XPR_IS_NULL(sExplorerCtrl[j])) continue;

                if (!XPR_TEST_BITS(sSyncItem->mExist, (j == 0) ? CompareExistLeft : CompareExistRight))
                    continue;

                if (XPR_TEST_BITS(sSyncItem->mExist, CompareExistEqual))
                {
                    if (!XPR_TEST_BITS(sSyncItem->mDiff, CompareDiffNotEqualed))
                        continue;
                }

                xpr::string sPath = sDir[j];
                sPath += XPR_STRING_LITERAL('\\');
                sPath += sSyncItem->mSubPath;

                xpr_sint_t sFind = sExplorerCtrl[j]->findItemPath(sPath.c_str());
                if (sFind >= 0)
                {
                    xpr_uint_t sMask  = LVIS_SELECTED;
                    xpr_uint_t sState = LVIS_SELECTED;

                    if (sFirst[j])
                    {
                        sMask  |= LVIS_FOCUSED;
                        sState |= LVIS_FOCUSED;
                        sFirst[j] = XPR_FALSE;
                    }

                    sExplorerCtrl[j]->SetItemState(sFind, sState, sMask);
                }
            }
        }

        MessageBox(_T("메인 화면의 탐색기 창에 차이 나는 파일들이 선택 상태로 반영되었습니다."),
            _T("창 선택 반영 완료"), MB_OK | MB_ICONINFORMATION);
    }
}

void FolderCompareReportDlg::OnBtnSyncTool(void)
{
    FolderSyncDlg sDlg;
    sDlg.setDir(mDir1.c_str(), mDir2.c_str());
    sDlg.DoModal();
}

} // namespace cmd
} // namespace fxfile