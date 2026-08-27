// Safe file/folder lock manager dialog.
#include "stdafx.h"
#include "file_lock_manager_dlg.h"

#include "file_operation_lock_store.h"
#include "resource.h"

#include <restartmanager.h>
#include <shlobj.h>

namespace fxfile
{
namespace cmd
{
FileLockManagerDlg::FileLockManagerDlg(
    const std::vector<std::wstring> &aPaths)
    : super(IDD_FILE_LOCK_MANAGER, NULL), mPaths(aPaths)
{
}

BEGIN_MESSAGE_MAP(FileLockManagerDlg, super)
    ON_BN_CLICKED(IDC_FILE_LOCK_REFRESH, OnRefresh)
    ON_BN_CLICKED(IDC_FILE_LOCK_FX_LOCK, OnFxLock)
    ON_BN_CLICKED(IDC_FILE_LOCK_FX_UNLOCK, OnFxUnlock)
    ON_BN_CLICKED(IDC_FILE_LOCK_READONLY, OnReadOnly)
    ON_BN_CLICKED(IDC_FILE_LOCK_WRITABLE, OnWritable)
    ON_BN_CLICKED(IDC_FILE_LOCK_SECURITY, OnSecurity)
END_MESSAGE_MAP()

BOOL FileLockManagerDlg::OnInitDialog(void)
{
    super::OnInitDialog();
    mPathList.SubclassDlgItem(IDC_FILE_LOCK_PATHS, this);
    mProcessList.SubclassDlgItem(IDC_FILE_LOCK_PROCESSES, this);
    mPathList.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
    mProcessList.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
    mPathList.InsertColumn(0, L"대상", LVCFMT_LEFT, 285);
    mPathList.InsertColumn(1, L"상태", LVCFMT_LEFT, 165);
    mProcessList.InsertColumn(0, L"PID", LVCFMT_RIGHT, 60);
    mProcessList.InsertColumn(1, L"잠금 프로그램/서비스", LVCFMT_LEFT, 280);
    for (size_t i = 0; i < mPaths.size(); ++i)
        mPathList.InsertItem(static_cast<int>(i), mPaths[i].c_str());
    if (!mPaths.empty())
        mPathList.SetItemState(0, LVIS_SELECTED | LVIS_FOCUSED,
                              LVIS_SELECTED | LVIS_FOCUSED);
    SetWindowTextW(L"파일·폴더 잠금 관리");
    SetDlgItemTextW(IDC_FILE_LOCK_REFRESH, L"새로 고침");
    SetDlgItemTextW(IDC_FILE_LOCK_FX_LOCK, L"FxFile 잠금");
    SetDlgItemTextW(IDC_FILE_LOCK_FX_UNLOCK, L"FxFile 잠금 해제");
    SetDlgItemTextW(IDC_FILE_LOCK_READONLY, L"읽기 전용");
    SetDlgItemTextW(IDC_FILE_LOCK_WRITABLE, L"쓰기 가능");
    SetDlgItemTextW(IDC_FILE_LOCK_SECURITY, L"Windows 보안");
    SetDlgItemTextW(IDOK, L"닫기");
    SetDlgItemTextW(IDC_FILE_LOCK_NOTICE,
        L"FxFile 작업 잠금은 이동·삭제·이름 변경을 안전하게 차단합니다. "
        L"폴더의 읽기 전용 속성은 Windows에서 보호 기능이 아니므로 적용하지 않습니다. "
        L"WRP 보호 파일은 변경하지 않습니다.");
    initializeToolTips();
    refresh();
    return TRUE;
}

void FileLockManagerDlg::initializeToolTips(void)
{
    if (!mToolTipCtrl.Create(this, TTS_ALWAYSTIP | TTS_NOPREFIX))
        return;

    mToolTipCtrl.SetMaxTipWidth(560);
    mToolTipCtrl.SetDelayTime(TTDT_INITIAL, 350);
    mToolTipCtrl.SetDelayTime(TTDT_RESHOW, 100);
    mToolTipCtrl.SetDelayTime(TTDT_AUTOPOP, 30000);

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_PATHS),
        L"[대상 목록]\r\n"
        L"목적: 잠금·속성·Windows 권한을 확인할 파일과 폴더를 보여 줍니다.\r\n"
        L"방법: 한 항목을 선택한 뒤 '새로 고침'을 누르면 그 항목을 사용 중인 "
        L"프로그램과 서비스를 다시 진단합니다.\r\n"
        L"효과: Windows 보안 버튼은 선택한 한 항목에 적용됩니다. FxFile 잠금/해제와 "
        L"읽기 전용/쓰기 가능 버튼은 이 목록에 들어온 전체 대상에 적용됩니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_REFRESH),
        L"[새로 고침]\r\n"
        L"목적: 파일·폴더의 현재 상태와 사용 중 프로세스를 다시 확인합니다.\r\n"
        L"방법: 대상 목록에서 확인할 항목을 선택한 뒤 누르십시오.\r\n"
        L"효과: 대상 존재 여부, FxFile 잠금, 읽기 전용/쓰기 가능, Windows 보호 여부와 "
        L"Restart Manager 진단 목록을 최신 상태로 갱신합니다. 파일·권한·잠금 상태는 변경하지 않습니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_FX_LOCK),
        L"[FxFile 잠금]\r\n"
        L"목적: 실수로 인한 복사 대상 덮어쓰기, 이동, 삭제, 이름 변경 및 파일 수집함 작업을 "
        L"FxFile 내부에서 차단합니다.\r\n"
        L"방법: 보호할 파일·폴더를 선택해 이 창을 연 뒤 버튼을 누르십시오. 목록의 전체 대상이 "
        L"잠기며, 폴더 잠금은 그 하위 파일과 폴더에도 적용됩니다.\r\n"
        L"효과: 잠금은 현재 설정 폴더의 fxfile-operation-locks.conf에 저장되어 FxFile을 "
        L"다시 실행해도 유지됩니다. NTFS 권한이나 다른 프로그램의 작업까지 막는 기능은 아닙니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_FX_UNLOCK),
        L"[FxFile 잠금 해제]\r\n"
        L"목적: 목록에 있는 대상의 FxFile 내부 작업 차단을 해제합니다.\r\n"
        L"방법: 해제할 대상들로 이 창을 연 뒤 버튼을 누르고 상태 열이 'FxFile 해제'로 "
        L"바뀌었는지 확인하십시오.\r\n"
        L"효과: FxFile에서 다시 복사·이동·삭제·이름 변경할 수 있습니다. 읽기 전용 속성, "
        L"Windows ACL 권한, 다른 프로세스가 잡은 핸들은 변경하거나 해제하지 않습니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_READONLY),
        L"[읽기 전용]\r\n"
        L"목적: 일반 파일의 내용이 프로그램에 의해 쉽게 수정되지 않도록 Windows Read-only "
        L"속성을 설정합니다.\r\n"
        L"방법: 목록에 파일을 넣고 버튼을 누르십시오. 전체 일반 파일에 적용됩니다.\r\n"
        L"효과: 파일 쓰기는 제한되지만 완전한 보안 잠금이나 권한 차단은 아닙니다. 폴더, "
        L"Windows/WRP 보호 항목, 존재하지 않는 대상 및 변경 실패 항목은 안전하게 건너뜁니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_WRITABLE),
        L"[쓰기 가능]\r\n"
        L"목적: 일반 파일의 Read-only 속성을 제거하여 정상 편집과 저장을 허용합니다.\r\n"
        L"방법: 목록에 파일을 넣고 버튼을 누른 뒤 변경/건너뜀 개수와 상태 열을 확인하십시오.\r\n"
        L"효과: 읽기 전용 비트만 제거합니다. NTFS 쓰기 권한을 새로 부여하거나 소유권을 가져오지 "
        L"않으며, 폴더와 Windows/WRP 보호 항목은 변경하지 않습니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_SECURITY),
        L"[Windows 보안]\r\n"
        L"목적: 선택한 한 항목의 NTFS 사용자·그룹·읽기·쓰기·수정 권한을 Windows 공식 보안 "
        L"화면에서 확인하거나 변경합니다.\r\n"
        L"방법: 대상 목록에서 정확한 항목 하나를 선택한 뒤 누르고, Windows가 표시하는 보안 "
        L"탭과 UAC 안내를 검토하십시오.\r\n"
        L"효과/주의: 권한 변경은 시스템 전체에 영향을 줄 수 있습니다. FxFile은 Windows 암호를 "
        L"받거나 저장하지 않고, 소유권 탈취나 보호 파일 변경을 자동화하지 않습니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDC_FILE_LOCK_PROCESSES),
        L"[잠금 사용 프로그램·서비스 - 진단 전용]\r\n"
        L"목적: 선택한 경로를 현재 사용 중인 프로세스와 서비스를 PID와 함께 확인합니다.\r\n"
        L"방법: 대상 목록에서 항목을 고르고 '새로 고침'을 누르십시오.\r\n"
        L"효과: Restart Manager의 진단 정보만 표시합니다. FxFile은 나열된 프로세스를 종료하거나 "
        L"핸들을 강제로 닫지 않으므로 필요한 프로그램은 사용자가 정상 저장·종료해야 합니다.");

    mToolTipCtrl.AddTool(GetDlgItem(IDOK),
        L"[닫기]\r\n"
        L"목적: 파일·폴더 잠금 관리 창을 닫고 FxFile로 돌아갑니다.\r\n"
        L"효과/주의: 이 창의 잠금·잠금 해제·속성 변경은 버튼을 누르는 즉시 적용됩니다. "
        L"닫기는 이미 적용된 변경을 취소하거나 이전 상태로 되돌리지 않습니다.");

    mToolTipCtrl.Activate(TRUE);
}

BOOL FileLockManagerDlg::PreTranslateMessage(MSG *aMsg)
{
    if (::IsWindow(mToolTipCtrl.GetSafeHwnd()))
        mToolTipCtrl.RelayEvent(aMsg);
    return super::PreTranslateMessage(aMsg);
}

bool FileLockManagerDlg::selectedPath(std::wstring &aPath) const
{
    const int sIndex = mPathList.GetNextItem(-1, LVNI_SELECTED);
    if (sIndex < 0 || static_cast<size_t>(sIndex) >= mPaths.size())
        return false;
    aPath = mPaths[sIndex];
    return true;
}

bool FileLockManagerDlg::isProtected(const std::wstring &aPath) const
{
    wchar_t sWindows[MAX_PATH + 1] = {0};
    ::GetWindowsDirectoryW(sWindows, MAX_PATH);
    wchar_t sFull[MAX_PATH + 1] = {0};
    ::GetFullPathNameW(aPath.c_str(), MAX_PATH, sFull, NULL);
    const size_t sWindowsLen = wcslen(sWindows);
    if (_wcsnicmp(sFull, sWindows, sWindowsLen) == 0 &&
        (sFull[sWindowsLen] == L'\\' || sFull[sWindowsLen] == L'\0'))
        return true;

    typedef BOOL (WINAPI *SfcIsFileProtectedProc)(HANDLE, LPCWSTR);
    HMODULE sSfc = ::LoadLibraryW(L"sfc.dll");
    if (sSfc == NULL)
        return false;
    SfcIsFileProtectedProc sFunction =
        reinterpret_cast<SfcIsFileProtectedProc>(
            ::GetProcAddress(sSfc, "SfcIsFileProtected"));
    const bool sResult = sFunction != NULL &&
                         sFunction(NULL, aPath.c_str()) != FALSE;
    ::FreeLibrary(sSfc);
    return sResult;
}

std::wstring FileLockManagerDlg::describePath(const std::wstring &aPath) const
{
    const DWORD sAttributes = ::GetFileAttributesW(aPath.c_str());
    if (sAttributes == INVALID_FILE_ATTRIBUTES)
        return L"대상 없음";
    std::wstring sStatus;
    if (isProtected(aPath))
        sStatus += L"Windows 보호; ";
    if (FileOperationLockStore::instance().isLocked(aPath))
        sStatus += L"FxFile 잠금; ";
    else
        sStatus += L"FxFile 해제; ";
    if ((sAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
        sStatus += L"폴더";
    else if ((sAttributes & FILE_ATTRIBUTE_READONLY) != 0)
        sStatus += L"읽기 전용";
    else
        sStatus += L"쓰기 가능";
    return sStatus;
}

void FileLockManagerDlg::refreshLockingProcesses(
    const std::wstring &aPath)
{
    mProcessList.DeleteAllItems();
    DWORD sSession = 0;
    wchar_t sKey[CCH_RM_SESSION_KEY + 1] = {0};
    if (::RmStartSession(&sSession, 0, sKey) != ERROR_SUCCESS)
        return;
    LPCWSTR sFile = aPath.c_str();
    DWORD sResult = ::RmRegisterResources(sSession, 1, &sFile, 0, NULL,
                                           0, NULL);
    if (sResult == ERROR_SUCCESS)
    {
        UINT sNeeded = 0;
        UINT sCount = 0;
        DWORD sReasons = 0;
        sResult = ::RmGetList(sSession, &sNeeded, &sCount, NULL, &sReasons);
        if (sResult == ERROR_MORE_DATA && sNeeded > 0 && sNeeded < 4096)
        {
            std::vector<RM_PROCESS_INFO> sApps(sNeeded);
            sCount = sNeeded;
            if (::RmGetList(sSession, &sNeeded, &sCount, &sApps[0],
                            &sReasons) == ERROR_SUCCESS)
            {
                for (UINT i = 0; i < sCount; ++i)
                {
                    wchar_t sPid[32] = {0};
                    _snwprintf_s(sPid, _countof(sPid), _TRUNCATE, L"%lu",
                                 sApps[i].Process.dwProcessId);
                    const int sRow = mProcessList.InsertItem(i, sPid);
                    std::wstring sName = sApps[i].strAppName;
                    if (sName.empty())
                        sName = sApps[i].strServiceShortName;
                    mProcessList.SetItemText(sRow, 1, sName.c_str());
                }
            }
        }
    }
    ::RmEndSession(sSession);
}

void FileLockManagerDlg::refresh(void)
{
    for (size_t i = 0; i < mPaths.size(); ++i)
        mPathList.SetItemText(static_cast<int>(i), 1,
                              describePath(mPaths[i]).c_str());
    std::wstring sPath;
    if (selectedPath(sPath))
        refreshLockingProcesses(sPath);
}

void FileLockManagerDlg::OnRefresh(void)
{
    refresh();
}

void FileLockManagerDlg::OnFxLock(void)
{
    if (FileOperationLockStore::instance().setLocked(mPaths, true))
        refresh();
    else
        AfxMessageBox(L"잠금 정보를 원자적으로 저장하지 못했습니다.",
                      MB_OK | MB_ICONERROR);
}

void FileLockManagerDlg::OnFxUnlock(void)
{
    if (FileOperationLockStore::instance().setLocked(mPaths, false))
        refresh();
    else
        AfxMessageBox(L"잠금 해제 정보를 원자적으로 저장하지 못했습니다.",
                      MB_OK | MB_ICONERROR);
}

void FileLockManagerDlg::setReadOnly(bool aReadOnly)
{
    size_t sChanged = 0;
    size_t sSkipped = 0;
    for (size_t i = 0; i < mPaths.size(); ++i)
    {
        DWORD sAttributes = ::GetFileAttributesW(mPaths[i].c_str());
        if (sAttributes == INVALID_FILE_ATTRIBUTES ||
            (sAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0 ||
            isProtected(mPaths[i]))
        {
            ++sSkipped;
            continue;
        }
        DWORD sNew = aReadOnly ? (sAttributes | FILE_ATTRIBUTE_READONLY) :
                                (sAttributes & ~FILE_ATTRIBUTE_READONLY);
        if (::SetFileAttributesW(mPaths[i].c_str(), sNew))
            ++sChanged;
        else
            ++sSkipped;
    }
    wchar_t sMessage[256] = {0};
    _snwprintf_s(sMessage, _countof(sMessage), _TRUNCATE,
                 L"변경: %Iu개\n보호/폴더/실패로 건너뜀: %Iu개",
                 sChanged, sSkipped);
    AfxMessageBox(sMessage, MB_OK | (sSkipped ? MB_ICONINFORMATION : 0));
    refresh();
}

void FileLockManagerDlg::OnReadOnly(void)
{
    setReadOnly(true);
}

void FileLockManagerDlg::OnWritable(void)
{
    setReadOnly(false);
}

void FileLockManagerDlg::OnSecurity(void)
{
    std::wstring sPath;
    if (!selectedPath(sPath))
        return;
    // The OS owns the credential/UAC and ACL UI.  FxFile never receives or
    // stores a Windows password and never takes ownership silently.
    if (!::SHObjectProperties(GetSafeHwnd(), SHOP_FILEPATH,
                              sPath.c_str(), L"security"))
        ::SHObjectProperties(GetSafeHwnd(), SHOP_FILEPATH,
                             sPath.c_str(), NULL);
}
} // namespace cmd
} // namespace fxfile
