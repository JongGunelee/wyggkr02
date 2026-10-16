#include "stdafx.h"
#include "file_operation_confirmation.h"
#include "resource.h"
#include "clip_format.h"
#include "gui/ResizingDialog.h"
#include <shlobj.h>

namespace fxfile {
namespace {
// The modal loop can deliver watcher notifications. The dialog owns strings,
// and callers must retain the exact operation/data object they previewed.
class ConfirmationDialog : public CResizingDialog {
public:
    ConfirmationDialog(HWND owner, const std::wstring& summary,
                       const std::wstring& details, bool notice = false)
        : CResizingDialog(IDD_FILE_OPERATION_CONFIRM, CWnd::FromHandle(owner)),
          mSummary(summary), mDetails(details), mNotice(notice) {}
protected:
    void OnOK() override {
        if (mNotice) OnCancel();
        else CResizingDialog::OnOK();
    }
    BOOL PreTranslateMessage(MSG* message) override {
        // Never let a queued Enter or an MFC default-button translation
        // approve a file request. Explicit Confirm button activation only.
        if ((message->message == WM_KEYDOWN || message->message == WM_SYSKEYDOWN) &&
            message->wParam == VK_RETURN) {
            OnCancel();
            return TRUE;
        }
        if ((message->message == WM_CHAR || message->message == WM_KEYUP) &&
            message->wParam == VK_RETURN) return TRUE;
        return CResizingDialog::PreTranslateMessage(message);
    }
    BOOL OnInitDialog() override {
        SetWindowText(L"FxFile 파일 작업 확인");
        SetDlgItemText(IDC_FILE_CONFIRM_SUMMARY, mSummary.c_str());
        CEdit* edit = static_cast<CEdit*>(GetDlgItem(IDC_FILE_CONFIRM_DETAILS));
        edit->LimitText(0);
        edit->SetWindowText(mDetails.c_str());
        SetDlgItemText(IDOK, L"확인");
        SetDlgItemText(IDCANCEL, mNotice ? L"닫기" : L"취소");
        if (mNotice) GetDlgItem(IDOK)->ShowWindow(SW_HIDE);
        // The 12-point dialog font sets both text and dialog units. Keep the
        // whole window inside the owner's monitor, including high DPI/small
        // screens, and lay out controls using the actual font metrics.
        MONITORINFO monitor = {sizeof(monitor)};
        if (::GetMonitorInfo(::MonitorFromWindow(GetSafeHwnd(), MONITOR_DEFAULTTONEAREST), &monitor)) {
            CRect window; GetWindowRect(window);
            int width = min(window.Width(), monitor.rcWork.right - monitor.rcWork.left - 24);
            int height = min(window.Height(), monitor.rcWork.bottom - monitor.rcWork.top - 24);
            SetWindowPos(NULL, monitor.rcWork.left + (monitor.rcWork.right - monitor.rcWork.left - width) / 2,
                monitor.rcWork.top + (monitor.rcWork.bottom - monitor.rcWork.top - height) / 2,
                width, height, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        CClientDC dc(this);
        CFont* previous = dc.SelectObject(GetFont());
        TEXTMETRIC metrics = {}; dc.GetTextMetrics(&metrics);
        int line = metrics.tmHeight + metrics.tmExternalLeading;
        int margin = max(12, line);
        CRect client; GetClientRect(client);
        CRect summary(0, 0, client.Width() - margin * 2, 0);
        dc.DrawText(mSummary.c_str(), &summary, DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX);
        dc.SelectObject(previous);
        int summaryHeight = max(line * 3, summary.Height() + line / 2);
        int buttonHeight = line * 2;
        int buttonWidth = max(100, metrics.tmAveCharWidth * 14);
        GetDlgItem(IDC_FILE_CONFIRM_SUMMARY)->MoveWindow(margin, margin,
            client.Width() - margin * 2, summaryHeight);
        edit->MoveWindow(margin, margin * 2 + summaryHeight, client.Width() - margin * 2,
            max(line * 3, client.Height() - margin * 4 - summaryHeight - buttonHeight));
        GetDlgItem(IDOK)->MoveWindow(client.Width() - margin * 3 - buttonWidth * 2,
            client.Height() - margin - buttonHeight, buttonWidth, buttonHeight);
        GetDlgItem(IDCANCEL)->MoveWindow(client.Width() - margin - buttonWidth,
            client.Height() - margin - buttonHeight, buttonWidth, buttonHeight);
        CResizingDialog::OnInitDialog();
        SetDefID(IDCANCEL);
        GetDlgItem(IDCANCEL)->SetFocus();
        AddControl(IDC_FILE_CONFIRM_SUMMARY, sizeResize, sizeNone);
        AddControl(IDC_FILE_CONFIRM_DETAILS, sizeResize, sizeResize);
        AddControl(IDOK, sizeRepos, sizeRepos);
        AddControl(IDCANCEL, sizeRepos, sizeRepos);
        return FALSE;
    }
private:
    std::wstring mSummary, mDetails;
    bool mNotice;
};

bool rejectUnknown(HWND owner) {
    ConfirmationDialog dialog(owner, L"작업을 실행하지 않았습니다.",
        L"작업 대상이나 도착지 또는 클립보드가 바뀌어 요청을 안전하게 확인할 수 없습니다.\r\n\r\n"
        L"대상을 다시 선택한 뒤 다시 시도하십시오.\r\n"
        L"이 요청으로 파일을 이동하거나 삭제하지 않았습니다.", true);
    dialog.DoModal();
    return false;
}

bool shellPaths(IShellFolder* folder, LPCITEMIDLIST* pidls, UINT count,
                std::vector<std::wstring>& result) {
    if (!folder) return false;
    if (!count) {
        IPersistFolder2* persist = NULL;
        PIDLIST_ABSOLUTE current = NULL;
        if (SUCCEEDED(folder->QueryInterface(IID_PPV_ARGS(&persist)))) {
            persist->GetCurFolder(&current);
            persist->Release();
        }
        if (!current) return false;
        wchar_t path[XPR_MAX_PATH + 1] = {};
        GetName(current, SHGDN_FORPARSING, path);
        ::CoTaskMemFree(current);
        if (!*path) return false;
        result.push_back(path);
        return true;
    }
    if (!pidls) return false;
    for (UINT i = 0; i < count; ++i) {
        wchar_t path[XPR_MAX_PATH + 1] = {};
        GetName(folder, pidls[i], SHGDN_FORPARSING, path);
        if (!*path) return false;
        result.push_back(path);
    }
    return !result.empty();
}

bool validPidl(const BYTE* data, SIZE_T bytes, UINT offset) {
    SIZE_T pos = offset;
    while (pos <= bytes && bytes - pos >= sizeof(USHORT)) {
        USHORT size = 0;
        memcpy(&size, data + pos, sizeof(size));
        if (!size) return true;
        if (size < sizeof(USHORT) || size > bytes - pos) return false;
        pos += size;
    }
    return false;
}
}

std::vector<std::wstring> FileOperationConfirmation::paths(const wchar_t* multi) {
    std::vector<std::wstring> result;
    if (multi) for (const wchar_t* p = multi; *p; p += wcslen(p) + 1)
        result.push_back(p);
    return result;
}

bool FileOperationConfirmation::confirm(HWND owner, UINT operation,
    const std::vector<std::wstring>& sources,
    const std::vector<std::wstring>& targets, WORD flags) {
    if (sources.empty() || ((operation == FO_COPY || operation == FO_MOVE ||
        operation == Link || operation == ExternalDrop) && targets.empty()))
        return rejectUnknown(owner);
    const wchar_t* action = L"";
    switch (operation) {
    case FO_COPY: action = L"복사 / 붙여넣기"; break;
    case FO_MOVE: action = L"이동 / 잘라낸 항목 붙여넣기"; break;
    case FO_DELETE: action = (flags & FOF_ALLOWUNDO) ? L"휴지통으로 삭제" : L"영구 삭제"; break;
    case Cut: action = L"잘라내기 (클립보드에 이동 대기 등록)"; break;
    case Link: action = L"바로가기 만들기 / 붙여넣기"; break;
    case ExternalDrop: action = L"외부 프로그램으로 드래그 전달"; break;
    default: return rejectUnknown(owner);
    }
    std::wstring summary = action;
    summary += L" — " + std::to_wstring(sources.size()) + L"개 선택 항목\r\n"
        L"전체 대상과 도착지를 확인하십시오. 확인 버튼을 누른 경우에만 진행합니다.";
    std::wstring details = L"작업: "; details += action;
    details += L"\r\n\r\n[대상 전체 경로 — 폴더는 하위 내용 포함]\r\n";
    for (size_t i = 0; i < sources.size(); ++i)
        details += std::to_wstring(i + 1) + L". " + sources[i] + L"\r\n";
    details += L"\r\n[도착지]\r\n";
    for (const auto& target : targets) details += target + L"\r\n";
    if (operation == FO_DELETE) details += (flags & FOF_ALLOWUNDO) ?
        L"휴지통 (Windows에서 휴지통을 지원하지 않는 대상은 별도 경고를 확인하십시오.)\r\n" :
        L"영구 삭제 — 휴지통에 보관되지 않습니다.\r\n";
    if (operation == Cut) details += L"아직 이동하지 않습니다. 붙여넣을 때 도착지를 다시 확인합니다.\r\n";
    if (operation == FO_MOVE) details += L"성공한 항목은 위 출발지에서 사라지고 도착지로 이동합니다.\r\n";
    if (flags & FOF_RENAMEONCOLLISION)
        details += L"같은 폴더 또는 중복 이름: 새 이름의 추가 복사본이 만들어질 수 있습니다.\r\n";
    if (operation == FO_COPY || operation == FO_MOVE)
        details += L"기존 이름 충돌의 덮어쓰기는 별도 확인을 따릅니다.\r\n";
    details += L"\r\n취소 / Esc / 닫기: 이 요청의 파일 작업을 실행하지 않습니다.";
    ConfirmationDialog dialog(owner, summary, details);
    return dialog.DoModal() == IDOK;
}

bool FileOperationConfirmation::confirm(const SHFILEOPSTRUCT* operation) {
    if (!operation) return false;
    // Rename undo is an existing separate operation, outside this consent policy.
    if (operation->wFunc == FO_RENAME) return true;
    std::vector<std::wstring> targets;
    if (operation->pTo && *operation->pTo) {
        if (operation->fFlags & FOF_MULTIDESTFILES) targets = paths(operation->pTo);
        else targets.push_back(operation->pTo);
    }
    return confirm(operation->hwnd, operation->wFunc, paths(operation->pFrom),
                   targets, operation->fFlags);
}

bool FileOperationConfirmation::dataObjectPaths(IDataObject* data,
                                               std::vector<std::wstring>& result) {
    result.clear();
    if (!data) return false;
    FORMATETC format = {CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM medium = {};
    if (SUCCEEDED(data->GetData(&format, &medium))) {
        bool valid = medium.tymed == TYMED_HGLOBAL && medium.hGlobal;
        if (valid) {
            HDROP drop = reinterpret_cast<HDROP>(medium.hGlobal);
            UINT count = ::DragQueryFileW(drop, 0xFFFFFFFF, NULL, 0);
            valid = count > 0 && count <= 1000000;
            for (UINT i = 0; valid && i < count; ++i) {
                UINT size = ::DragQueryFileW(drop, i, NULL, 0);
                if (!size || size > 32767) { valid = false; break; }
                std::vector<wchar_t> path(size + 1, 0);
                if (::DragQueryFileW(drop, i, path.data(), size + 1) != size) { valid = false; break; }
                result.push_back(path.data());
            }
        }
        ::ReleaseStgMedium(&medium);
        if (!valid) result.clear();
        return valid;
    }
    format.cfFormat = ClipFormat::instance().mShellIDList;
    medium = STGMEDIUM();
    if (FAILED(data->GetData(&format, &medium))) return false;
    bool valid = false;
    if (medium.tymed == TYMED_HGLOBAL && medium.hGlobal) {
        SIZE_T bytes = ::GlobalSize(medium.hGlobal);
        const BYTE* raw = static_cast<const BYTE*>(::GlobalLock(medium.hGlobal));
        if (raw && bytes >= sizeof(UINT) * 2) {
            const CIDA* ida = reinterpret_cast<const CIDA*>(raw);
            UINT count = ida->cidl;
            SIZE_T header = (static_cast<SIZE_T>(count) + 2) * sizeof(UINT);
            valid = count > 0 && count <= 1000000 && header <= bytes;
            for (UINT i = 0; valid && i <= count; ++i)
                valid = ida->aoffset[i] >= header && validPidl(raw, bytes, ida->aoffset[i]);
            for (UINT i = 0; valid && i < count; ++i) {
                PIDLIST_ABSOLUTE full = ::ILCombine(
                    reinterpret_cast<LPCITEMIDLIST>(raw + ida->aoffset[0]),
                    reinterpret_cast<LPCITEMIDLIST>(raw + ida->aoffset[i + 1]));
                wchar_t path[XPR_MAX_PATH + 1] = {};
                if (full) { GetName(full, SHGDN_FORPARSING, path); ::ILFree(full); }
                if (!*path) { valid = false; break; }
                result.push_back(path);
            }
        }
        if (raw) ::GlobalUnlock(medium.hGlobal);
    }
    ::ReleaseStgMedium(&medium);
    if (!valid) result.clear();
    return valid;
}

bool FileOperationConfirmation::confirmCut(HWND owner, IDataObject* data) {
    std::vector<std::wstring> sources;
    return dataObjectPaths(data, sources) ? confirm(owner, Cut, sources, {}) : rejectUnknown(owner);
}

bool FileOperationConfirmation::confirmDrop(HWND owner, IDataObject* data,
                                           DWORD effect, const wchar_t* target) {
    std::vector<std::wstring> sources;
    if (!dataObjectPaths(data, sources) || !target || !*target) return rejectUnknown(owner);
    UINT action = effect == DROPEFFECT_MOVE ? FO_MOVE :
        effect == DROPEFFECT_COPY ? FO_COPY : effect == DROPEFFECT_LINK ? Link : ExternalDrop;
    return confirm(owner, action, sources, {target});
}

DWORD FileOperationConfirmation::chooseDropEffect(HWND owner, DWORD allowed) {
    CMenu menu;
    if (!menu.CreatePopupMenu()) return DROPEFFECT_NONE;
    if (allowed & DROPEFFECT_COPY) menu.AppendMenu(MF_STRING, 1, L"복사");
    if (allowed & DROPEFFECT_MOVE) menu.AppendMenu(MF_STRING, 2, L"이동");
    if (allowed & DROPEFFECT_LINK) menu.AppendMenu(MF_STRING, 3, L"바로가기 만들기");
    menu.AppendMenu(MF_SEPARATOR); menu.AppendMenu(MF_STRING, 4, L"취소");
    POINT point; ::GetCursorPos(&point);
    UINT choice = ::TrackPopupMenu(menu.GetSafeHmenu(), TPM_RETURNCMD | TPM_NONOTIFY | TPM_RIGHTBUTTON,
                                  point.x, point.y, 0, owner, NULL);
    return choice == 1 ? DROPEFFECT_COPY : choice == 2 ? DROPEFFECT_MOVE :
           choice == 3 ? DROPEFFECT_LINK : DROPEFFECT_NONE;
}

bool FileOperationConfirmation::confirmShellVerb(HWND owner, const wchar_t* verb,
    IShellFolder* folder, LPCITEMIDLIST* pidls, UINT count) {
    if (!verb) return false;
    bool cut = !_wcsicmp(verb, L"cut"), del = !_wcsicmp(verb, L"delete");
    bool paste = !_wcsicmp(verb, L"paste") || !_wcsicmp(verb, L"pastelink");
    if (!cut && !del && !paste) return true;
    std::vector<std::wstring> selection;
    if (!shellPaths(folder, pidls, count, selection)) return rejectUnknown(owner);
    if (!paste) return confirm(owner, cut ? Cut : FO_DELETE, selection, {},
        del && ::GetAsyncKeyState(VK_SHIFT) >= 0 ? FOF_ALLOWUNDO : 0);
    DWORD sequence = ::GetClipboardSequenceNumber();
    IDataObject* data = NULL;
    if (FAILED(::OleGetClipboard(&data))) return rejectUnknown(owner);
    std::vector<std::wstring> sources;
    bool valid = dataObjectPaths(data, sources);
    DWORD effect = DROPEFFECT_COPY;
    FORMATETC format = {ClipFormat::instance().mDropEffect, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM medium = {};
    if (SUCCEEDED(data->GetData(&format, &medium))) {
        if (medium.tymed == TYMED_HGLOBAL && medium.hGlobal && ::GlobalSize(medium.hGlobal) >= sizeof(DWORD)) {
            void* ptr = ::GlobalLock(medium.hGlobal);
            if (ptr) { memcpy(&effect, ptr, sizeof(effect)); ::GlobalUnlock(medium.hGlobal); }
        }
        ::ReleaseStgMedium(&medium);
    }
    data->Release();
    if (!valid) return rejectUnknown(owner);
    UINT action = !_wcsicmp(verb, L"pastelink") ? Link : effect == DROPEFFECT_MOVE ? FO_MOVE : FO_COPY;
    if (!confirm(owner, action, sources, selection)) return false;
    if (::GetClipboardSequenceNumber() != sequence) return rejectUnknown(owner);
    return true;
}
}
