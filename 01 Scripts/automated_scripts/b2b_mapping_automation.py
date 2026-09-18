# -*- coding: utf-8 -*-
import os
import sys
import re
import gc
import time
import shutil
import datetime
import traceback
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32con
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

EXCEL_CALCULATION_MANUAL = -4135
NO_HEADER_MATCH_OPTION = '매칭 되지 않음'


def is_no_header_match_option(option_text):
    return str(option_text or '').strip().startswith(NO_HEADER_MATCH_OPTION)


def ensure_excel_com_available():
    """Excel만 설치된 PC에서도 pywin32가 없으면 1회 자동 설치를 시도한다."""
    global COM_AVAILABLE, win32com, pythoncom, win32gui, win32con
    if COM_AVAILABLE:
        return True
    try:
        import win32com.client as _win32_client
        import pythoncom as _pythoncom
        import win32gui as _win32gui
        import win32con as _win32con
    except ImportError:
        try:
            import subprocess
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', '--user', 'pywin32'],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            import win32com.client as _win32_client
            import pythoncom as _pythoncom
            import win32gui as _win32gui
            import win32con as _win32con
        except Exception as exc:
            raise RuntimeError(
                'Excel 자동화 구성요소(pywin32)를 사용할 수 없습니다. '
                '인터넷이 가능한 환경에서 다시 실행하거나, 명령 프롬프트에서 '
                '"python -m pip install --user pywin32" 실행 후 재시도하세요.'
            ) from exc

    class _Win32ComModule:
        client = _win32_client

    win32com = _Win32ComModule()
    pythoncom = _pythoncom
    win32gui = _win32gui
    win32con = _win32con
    COM_AVAILABLE = True
    return True


def set_excel_fast_mode(excel):
    """대량 COM 작업 중 화면 갱신/이벤트/자동계산을 잠시 꺼서 속도를 안정화한다."""
    state = {}
    for attr in ('DisplayAlerts', 'ScreenUpdating', 'EnableEvents', 'Calculation'):
        try:
            state[attr] = getattr(excel, attr)
        except Exception:
            pass
    try:
        excel.DisplayAlerts = False
    except Exception:
        pass
    try:
        excel.ScreenUpdating = False
    except Exception:
        pass
    try:
        excel.EnableEvents = False
    except Exception:
        pass
    try:
        excel.Calculation = EXCEL_CALCULATION_MANUAL
    except Exception:
        pass
    return state


def restore_excel_mode(excel, state):
    if not excel or not state:
        return
    for attr, value in reversed(list(state.items())):
        try:
            setattr(excel, attr, value)
        except Exception:
            pass


def _flatten_excel_column(values, expected_count):
    if expected_count <= 0:
        return []
    if expected_count == 1:
        if isinstance(values, tuple):
            first = values[0]
            return [first[0] if isinstance(first, tuple) else first]
        return [values]
    if values is None:
        return [None] * expected_count
    result = []
    for row in values:
        result.append(row[0] if isinstance(row, tuple) else row)
    if len(result) < expected_count:
        result.extend([None] * (expected_count - len(result)))
    return result[:expected_count]


def read_excel_columns(ws, start_row, end_row, columns):
    """셀 단위 COM 호출 대신 열 범위를 한 번에 읽어 대량 처리 속도를 높인다."""
    row_count = max(0, end_row - start_row + 1)
    data = {}
    for col in dict.fromkeys(int(c) for c in columns if c):
        try:
            values = ws.Range(ws.Cells(start_row, col), ws.Cells(end_row, col)).Value
            data[col] = _flatten_excel_column(values, row_count)
        except Exception:
            data[col] = [safe_cell_value(ws, r, col) for r in range(start_row, end_row + 1)]
    return data


def create_timestamped_backup(file_path, suffix='.bak'):
    base_path = os.path.abspath(file_path)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    backup_path = f'{base_path}_{timestamp}{suffix}'
    shutil.copy2(base_path, backup_path)
    return backup_path


def safe_cell_value(ws, r, c):
    try:
        val = ws.Cells(r, c).Value
        if val is None:
            return None
        s_val = str(val).strip()
        if s_val.startswith('#'):
            return None
        return val
    except Exception:
        try:
            val2 = ws.Cells(r, c).Value2
            if val2 is None:
                return None
            s_val2 = str(val2).strip()
            if s_val2.startswith('#'):
                return None
            return val2
        except Exception:
            return None

def extract_room_key(text):
    if not text: return ''
    s = str(text).strip()
    m = re.search(r'(\d+)\s*동\s*(\d+)\s*호?', s)
    if m:
        return f'{int(m.group(1))}-{int(m.group(2))}'
    m2 = re.search(r'(\d+)[\s\-_/]+(\d+)', s)
    if m2:
        return f'{int(m2.group(1))}-{int(m2.group(2))}'
    m3 = re.search(r'([A-Za-z0-9]+)\s*동', s)
    if m3:
        return f'{m3.group(1).upper()}동'
    return ''

def extract_floors(text):
    if not text: return set()
    s = str(text).strip()
    # Match floor range e.g. 1층~4층, 1~4층, 1-4층, 1층-4층
    m_range = re.search(r'(\d+)\s*층?\s*[~\-_]\s*(\d+)\s*층', s)
    if m_range:
        f1, f2 = int(m_range.group(1)), int(m_range.group(2))
        return set(range(min(f1, f2), max(f1, f2) + 1))
    # Match single floor e.g. 5층, 5F
    m_single = re.findall(r'(\d+)\s*층', s)
    if m_single:
        return set(int(x) for x in m_single)
    return set()

def are_floors_compatible(target_floors, source_floors):
    if not target_floors or not source_floors:
        return True
    return not target_floors.isdisjoint(source_floors)


def extract_discipline(text):
    if not text: return ''
    s = str(text)
    disciplines = ['기계', '건축', '전기', '토목', '소방', '통신', '조경', '수장', '설비', '위생', '가스', '소화기', '덕트', '배관']
    for d in disciplines:
        if d in s:
            return d
    return ''

def extract_round_num(text):
    if not text: return ''
    m = re.search(r'(\d+)\s*차', str(text))
    if m:
        return f'{m.group(1)}차'
    return ''

def normalize_text(text):
    if not text: return ''
    t = str(text).replace(' ', '').lower()
    t = t.replace('수장작업', '수장공사').replace('누수수리', '누수보수')
    t = t.replace('작업', '공사').replace('보수', '수리')
    return t

def col_num_to_letter(col):
    """컬럼 번호(1-based)를 엑셀 열 문자(A, B, ..., Z, AA, AB, ...)로 변환"""
    result = ''
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result

def col_letter_to_num(letter):
    """엑셀 열 문자(A, B, ..., AA, AB, ...)를 컬럼 번호(1-based)로 변환"""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - 64)
    return result


class HoverTooltip:
    """Tkinter 위젯에 상세 사용 가이드를 지연 표시하는 경량 툴팁."""

    def __init__(self, widget, text, delay=450, wraplength=620):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.wraplength = wraplength
        self.after_id = None
        self.tip_window = None
        widget.bind('<Enter>', self.schedule, add='+')
        widget.bind('<Leave>', self.hide, add='+')
        widget.bind('<ButtonPress>', self.hide, add='+')
        widget._hover_tooltip = self

    def schedule(self, _event=None):
        self.hide()
        self.after_id = self.widget.after(self.delay, self.show)

    def show(self):
        if self.tip_window or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 18
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            return
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f'+{x}+{y}')
        label = tk.Label(
            self.tip_window,
            text=self.text,
            justify='left',
            background='#FFF8DC',
            foreground='#1A202C',
            relief='solid',
            borderwidth=1,
            padx=10,
            pady=8,
            font=('Segoe UI', 9),
            wraplength=self.wraplength,
        )
        label.pack()

    def hide(self, _event=None):
        if self.after_id:
            try:
                self.widget.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None
        if self.tip_window:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None


class B2BMappingAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title('B2B 데이터 매핑 자동화 스마트 솔루션 v3.0')
        self.root.geometry('980x900')
        self.root.minsize(900, 800)

        self.source_file_path = tk.StringVar()
        self.target_file_path = tk.StringVar()
        self.selected_sheet_name = tk.StringVar()
        self.selected_range_address = tk.StringVar()
        self.source_info = None
        self.target_info = None
        self.target_range_info = None
        self.source_match_col = None
        self.source_match_header = ''
        self.target_match_col = None
        self.target_match_header = ''
        self.range_popup = None
        self.excel_activation_busy = False
        self.interactive_excel_com_initialized = False
        self.interactive_excel_app = None
        self.interactive_target_wb = None

        # 상태 추적 변수 (Smart Action Queue)
        self.active_tasks = {'source': False, 'target': False}
        self.queued_action = None

        # 헤더 매핑 데이터 구조 (자동 기본값으로 초기화)
        self.header_mappings = self.get_default_mappings()

        self.setup_styles()
        self.build_ui()

        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(self.root.attributes, '-topmost', False)
        self.root.focus_force()
        self.root.protocol("WM_DELETE_WINDOW", self.on_app_close)

    def on_app_close(self):
        self.release_interactive_excel_session('앱 종료')
        self.root.destroy()

    def release_interactive_excel_session(self, reason=''):
        """사용자 Excel 닫기 후 EXCEL.EXE가 COM 참조 때문에 남지 않도록 활성화 세션 참조를 해제한다."""
        for attr in ('interactive_target_wb', 'persist_target_wb', 'interactive_excel_app', 'persist_excel_app'):
            try:
                setattr(self, attr, None)
            except Exception:
                pass
        if COM_AVAILABLE and getattr(self, 'interactive_excel_com_initialized', False):
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self.interactive_excel_com_initialized = False
        gc.collect()
        if reason:
            try:
                self.log('INFO', f'Excel 활성화 COM 세션 해제 완료 ({reason})')
            except Exception:
                pass

    def collect_excel_process_visibility(self):
        excel_pids = set()
        visible_pids = set()

        try:
            ensure_excel_com_available()
            import win32process

            def enum_windows_callback(hwnd, _extra):
                try:
                    if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == 'XLMAIN':
                        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
                        if pid:
                            visible_pids.add(int(pid))
                except Exception:
                    pass
                return True

            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception:
            pass

        try:
            import csv
            import subprocess
            creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
            output = subprocess.check_output(
                ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
                text=True,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags,
            )
            for row in csv.reader(output.splitlines()):
                if len(row) >= 2 and row[0].strip().lower() == 'excel.exe':
                    try:
                        excel_pids.add(int(row[1]))
                    except ValueError:
                        pass
        except Exception:
            pass

        excel_pids.update(visible_pids)
        return {
            'all': sorted(excel_pids),
            'visible': sorted(excel_pids & visible_pids),
            'hidden': sorted(excel_pids - visible_pids),
        }

    def force_close_excel_pids(self, pids):
        import subprocess
        import time

        killed = []
        failed = []
        creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
        for pid in sorted(set(int(p) for p in pids if p)):
            try:
                result = subprocess.run(
                    ['taskkill', '/PID', str(pid), '/F', '/T'],
                    capture_output=True,
                    text=True,
                    timeout=8,
                    creationflags=creationflags,
                    check=False,
                )
                if result.returncode == 0:
                    killed.append(pid)
                else:
                    failed.append(pid)
            except Exception:
                failed.append(pid)
        if killed:
            self.release_interactive_excel_session('숨김 Excel 강제 종료')
            time.sleep(0.8)
        return killed, failed

    def confirm_excel_cleanup_before_activation(self, parent=None, show_when_visible_only=True):
        inventory = self.collect_excel_process_visibility()
        all_pids = inventory['all']
        hidden_pids = inventory['hidden']
        visible_pids = inventory['visible']
        self.log('INFO', f'Excel 활성화 사전 점검: 전체 {len(all_pids)}개, 보이는 창 {len(visible_pids)}개, 숨김/백그라운드 {len(hidden_pids)}개')
        if not all_pids:
            return True
        if not hidden_pids and not show_when_visible_only:
            return True

        parent = parent or self.root
        pid_text = ', '.join(str(pid) for pid in hidden_pids[:12])
        if len(hidden_pids) > 12:
            pid_text += ', ...'
        if not pid_text:
            pid_text = '없음'
        msg = (
            '기존 Excel 실행 상태가 감지되었습니다.\n\n'
            f'전체 EXCEL.EXE: {len(inventory["all"])}개\n'
            f'보이는 Excel 창: {len(inventory["visible"])}개\n'
            f'강제 종료 대상(창 없음/숨김): {len(hidden_pids)}개\n'
            f'대상 PID: {pid_text}\n\n'
            '[예] 숨김/백그라운드 Excel이 있으면 강제 닫고, 내부 COM 세션을 초기화한 뒤 대상 파일을 활성화합니다.\n'
            '[아니오] 현재 Excel 상태를 유지하고 그대로 활성화합니다.\n'
            '[취소] Excel 활성화를 중단합니다.\n\n'
            '보이는 Excel 창은 사용자 문서 보호를 위해 강제 종료하지 않습니다.'
        )
        restore_topmost = None
        try:
            restore_topmost = bool(parent.attributes('-topmost'))
            parent.attributes('-topmost', True)
            parent.lift()
        except Exception:
            pass
        try:
            choice = messagebox.askyesnocancel('Excel 실행 상태 정리 확인', msg, parent=parent)
        finally:
            try:
                if restore_topmost is False:
                    parent.attributes('-topmost', False)
            except Exception:
                pass
        if choice is None:
            self.log('INFO', '사용자가 Excel 활성화를 취소했습니다.')
            return False
        if choice is False:
            self.log('INFO', '숨김/백그라운드 Excel을 유지한 상태로 활성화를 진행합니다.')
            return True

        self.release_interactive_excel_session('Excel 활성화 사전 초기화')
        if hidden_pids:
            killed, failed = self.force_close_excel_pids(hidden_pids)
            if killed:
                self.log('SUCCESS', f'숨김/백그라운드 Excel 강제 종료 완료: PID {", ".join(map(str, killed))}')
            if failed:
                self.log('WARN', f'일부 Excel 프로세스 강제 종료 실패 또는 이미 종료됨: PID {", ".join(map(str, failed))}')
        else:
            self.log('INFO', '강제 종료 대상 숨김 Excel은 없어 내부 COM 세션만 초기화했습니다.')
        return True

    def update_dynamic_ui_state(self):
        is_busy = self.active_tasks['source'] or self.active_tasks['target']
        if is_busy:
            self.btn_drag_select.config(text='⏳ 엑셀 구조 분석 중 (예약 가능)')
            self.btn_run.config(text='⏳ 분석 대기 중 (자동 예약)')
        else:
            self.btn_drag_select.config(text='🖥️ 엑셀 창 열기 & 마우스 범위 확정')
            self.btn_run.config(text='▶️ 스마트 매핑 실행')

    def on_parsing_finished(self):
        self.update_dynamic_ui_state()
        if not self.active_tasks['source'] and not self.active_tasks['target']:
            if self.queued_action:
                action = self.queued_action
                self.queued_action = None
                self.log('SUCCESS', '✅ 백그라운드 분석 완료! 예약된 작업을 자동으로 시작합니다.')
                self.root.after(300, action)

    def get_default_mappings(self):
        """자동 기본 7쌍 헤더 매핑 반환 (실제 발주내역.xlsb 엑셀 1행 헤더명 100% 반영)"""
        return [
            {'source_name': '계약금액', 'source_aliases': ['계약금액', '낙찰금액'], 'target_col': 25, 'target_name': '낙찰금액', 'default_source_col': 4, 'mode': '자동'},
            {'source_name': 'PO전송견적금액', 'source_aliases': ['PO전송견적금액', '견적금액', 'PO금액'], 'target_col': 26, 'target_name': 'PO금액', 'default_source_col': 6, 'mode': '자동'},
            {'source_name': 'PO번호', 'source_aliases': ['PO번호'], 'target_col': 27, 'target_name': 'PO번호', 'default_source_col': 7, 'mode': '자동'},
            {'source_name': 'PR번호', 'source_aliases': ['PR번호'], 'target_col': 28, 'target_name': 'PR번호', 'default_source_col': 8, 'mode': '자동'},
            {'source_name': '생성일자', 'source_aliases': ['생성일자', 'PR발생일', 'PR생성일'], 'target_col': 30, 'target_name': 'PR 발생일', 'default_source_col': 1, 'mode': '자동'},
            {'source_name': '공사/입찰/계약번호', 'source_aliases': ['공사/입찰/계약번호', '공사/유지보수번호', '입찰번호', '계약번호'], 'target_col': 31, 'target_name': '공사/ 유지보수 번호', 'default_source_col': 2, 'mode': '자동'},
            {'source_name': '계약일자(체결일)', 'source_aliases': ['계약일자(체결일)', '계약체결일', '체결일', '계약일자'], 'target_col': 34, 'target_name': '계약 체결일', 'default_source_col': 5, 'mode': '자동'},
        ]

    def build_source_header_map(self):
        header_map = {
            1: '생성일자',
            2: '공사/입찰/계약번호',
            3: '공사/입찰명',
            4: '계약금액',
            5: '계약일자(체결일)',
            6: 'PO전송견적금액',
            7: 'PO번호',
            8: 'PR번호',
        }
        if self.source_info and 'header_names' in self.source_info:
            for col_idx, header_name in enumerate(self.source_info['header_names'], 1):
                if header_name:
                    header_map[col_idx] = self.clean_basis_title(header_name)
        return header_map

    def build_target_header_map(self):
        header_map = {
            24: '계약명',
            25: '낙찰금액',
            26: 'PO금액',
            27: 'PO번호',
            28: 'PR번호',
            29: 'PR금액',
            30: 'PR 발생일',
            31: '공사/ 유지보수 번호',
            32: '공사 요청일',
            33: '공사 완료일',
            34: '계약 체결일',
        }
        if self.target_info and 'header_names' in self.target_info:
            for col_idx, header_name in self.target_info['header_names'].items():
                if header_name:
                    header_map[int(col_idx)] = self.clean_basis_title(header_name)
        for mapping in self.header_mappings:
            target_col = int(mapping.get('target_col', 0) or 0)
            if target_col and target_col not in header_map:
                header_map[target_col] = mapping.get('target_name') or mapping.get('source_name') or f'Col {target_col}'
        return header_map

    def is_probable_data_value(self, value):
        text = str(value or '').strip()
        if not text:
            return False
        if hasattr(value, 'strftime') and not isinstance(value, str):
            return True
        compact = re.sub(r'\s+', '', text)
        if re.fullmatch(r'\d{1,2}월\d{1,2}일', compact):
            return True
        if re.fullmatch(r'\d{1,2}월', compact):
            return True
        if re.fullmatch(r'\d{4}[-./]\d{1,2}[-./]\d{1,2}', compact):
            return True
        if re.fullmatch(r'\d{4}년\d{1,2}월\d{1,2}일', compact):
            return True
        if re.fullmatch(r'\d{1,2}[-./]\d{1,2}', compact):
            return True
        if re.fullmatch(r'[+-]?\d{1,3}(,\d{3})+(\.\d+)?', compact):
            return True
        if re.fullmatch(r'[+-]?\d+(\.\d+)?', compact) and len(compact) >= 3:
            return True
        if len(text) >= 12 and any('\uac00' <= ch <= '\ud7a3' for ch in text) and re.search(r'\d+[\s\-]?\d*', text):
            return True
        if re.search(r'\d+\s*(동|호|층|차)', text):
            return True
        if len(text) >= 18 and any(word in text for word in ['보수', '수리', '작업', '공사', '배관', '화장실', '사택', '기숙사']):
            return True
        return False

    def compose_header_label(self, top_value, second_value=''):
        top = str(top_value or '').strip().replace('\n', ' ').replace('\r', ' ')
        second = str(second_value or '').strip().replace('\n', ' ').replace('\r', ' ')
        if top.startswith('#'):
            top = ''
        if second.startswith('#'):
            second = ''
        if top and second and second != top and not self.is_probable_data_value(second):
            return f'{top} / {second}'
        if top:
            return top
        return '' if self.is_probable_data_value(second) else second

    def clean_basis_title(self, title):
        parts = [p.strip() for p in str(title or '').split('/') if p.strip()]
        if not parts:
            return ''
        kept = []
        for idx, part in enumerate(parts):
            if idx > 0 and self.is_probable_data_value(part):
                continue
            kept.append(part)
        return ' / '.join(kept) if kept else ''

    def infer_basis_field_key_from_title(self, title):
        clean_title = self.clean_basis_title(title)
        if not clean_title:
            return None
        raw = str(clean_title).upper()
        normalized = normalize_text(clean_title)
        if any(token in normalized for token in ['마감', '월말', '정산월', '회계월', '회계마감']):
            return None
        if 'PR' in raw or 'pr' in normalized:
            return 'pr_no'
        if 'PO' in raw or 'po' in normalized or 'WBS' in raw:
            return 'po_no'
        if '번호' in clean_title:
            return 'po_no'
        if any(token in clean_title for token in ['발행', '상태', '구분']):
            return 'issue_type'
        if any(token in clean_title for token in ['유형', '계정', '서비스']):
            return 'contract_type'
        if any(token in clean_title for token in ['금액', '총액', '비용', '세액', '단가', '정산액', '낙찰액']):
            return 'amount'
        if any(token in clean_title for token in ['계약명', '공사명', '입찰명', '유지관리명', '관리명', '적요', '세부항목', '작업명', '업무명', '대상명', '시설명', 'WBS명']):
            return 'contract_name'
        return None

    def get_keyword_groups_for_basis_title(self, title):
        field_key = self.infer_basis_field_key_from_title(title)
        if field_key == 'pr_no':
            return [['PR번호'], ['PR'], ['계약번호'], ['공사/입찰/계약번호']]
        if field_key == 'po_no':
            return [['PO번호'], ['PO'], ['계약번호'], ['공사/입찰/계약번호'], ['공사번호'], ['WBS']]
        if field_key == 'issue_type':
            return [['발행구분'], ['발행상태'], ['정산상태'], ['상태'], ['구분']]
        if field_key == 'contract_type':
            return [['계정코드'], ['서비스유형'], ['계약유형'], ['정산유형'], ['유형']]
        if field_key == 'amount':
            return [
                ['낙찰금액'],
                ['계약금액'],
                ['정산금액'],
                ['PO전송견적금액'],
                ['PO금액'],
                ['PR금액'],
                ['청구계획금액'],
                ['기성계획금액'],
                ['총금액'],
                ['금액'],
                ['세액'],
            ]
        if field_key == 'contract_name':
            return [['계약명'], ['공사/입찰명'], ['공사명'], ['입찰명'], ['유지관리명'], ['관리명'], ['적요'], ['세부항목'], ['작업명'], ['업무명'], ['대상명'], ['시설명'], ['WBS명']]
        return []

    def find_header_col_by_keywords(self, header_map, keyword_groups, fallback=None):
        if not keyword_groups:
            return fallback
        normalized_headers = {col: normalize_text(self.clean_basis_title(name)) for col, name in header_map.items() if name}
        for keywords in keyword_groups:
            normalized_keywords = [normalize_text(keyword) for keyword in keywords]
            for col, header_text in normalized_headers.items():
                if any(keyword and (keyword in header_text or header_text in keyword) for keyword in normalized_keywords):
                    return col
        return fallback

    def find_header_col_by_title(self, header_map, title, fallback=None):
        normalized_title = normalize_text(self.clean_basis_title(title))
        if not normalized_title:
            return fallback
        normalized_headers = {col: normalize_text(self.clean_basis_title(name)) for col, name in header_map.items() if name}
        for col, header_text in normalized_headers.items():
            if header_text == normalized_title:
                return col
        for col, header_text in normalized_headers.items():
            if normalized_title in header_text or header_text in normalized_title:
                return col
        return fallback

    def update_linked_b2b_mappings_for_basis(self, header_title, source_col, target_col):
        title_norm = normalize_text(header_title)
        updated = []
        for mapping in self.header_mappings:
            source_names = [mapping.get('source_name', '')] + mapping.get('source_aliases', [])
            target_names = [mapping.get('target_name', '')]
            source_linked = any(title_norm and title_norm == normalize_text(name) for name in source_names)
            target_linked = any(title_norm and title_norm == normalize_text(name) for name in target_names)
            col_linked = int(mapping.get('target_col', 0) or 0) == int(target_col)
            if source_linked or target_linked or col_linked:
                mapping['source_name'] = header_title
                mapping['source_aliases'] = list(dict.fromkeys([header_title] + source_names))
                mapping['target_col'] = int(target_col)
                mapping['target_name'] = header_title
                mapping['default_source_col'] = int(source_col)
                mapping['mode'] = '수동-기준연계'
                updated.append(header_title)
        return updated

    def build_auto_basis_sync_result(self, header_title, range_target_col, source_header_map, target_header_map):
        title = self.clean_basis_title(header_title)
        target_col = int(range_target_col or 24)
        if title:
            target_header_map[target_col] = title
        else:
            title = self.clean_basis_title(target_header_map.get(target_col, f'Col {target_col}'))
        field_key = self.infer_basis_field_key_from_title(title)
        source_col = None
        if field_key:
            source_col = self.find_header_col_by_title(source_header_map, title)
            if not source_col:
                source_col = self.find_header_col_by_keywords(source_header_map, self.get_keyword_groups_for_basis_title(title), None)
        source_header = source_header_map.get(source_col, title) if source_col else NO_HEADER_MATCH_OPTION
        target_header = self.clean_basis_title(target_header_map.get(target_col, title))
        return source_col, source_header, target_col, target_header

    def format_source_header_option(self, col, header_map=None):
        header_map = header_map or self.build_source_header_map()
        col = int(col or 1)
        header_name = self.clean_basis_title(header_map.get(col, ''))
        return f'Col {col:02d} - {header_name}' if header_name else f'Col {col:02d}'

    def format_target_header_option(self, col, header_map=None):
        header_map = header_map or self.build_target_header_map()
        col = int(col or 1)
        letter = col_num_to_letter(col)
        header_name = self.clean_basis_title(header_map.get(col, ''))
        return f'{letter} (Col {col}) - {header_name}' if header_name else f'{letter} (Col {col})'

    def parse_col_num_from_option(self, option_text, fallback=1):
        if not option_text:
            return fallback
        if is_no_header_match_option(option_text):
            return fallback
        m = re.search(r'Col\s*(\d+)', option_text)
        if m:
            return int(m.group(1))
        m_letter = re.match(r'\s*([A-Z]+)\s*', option_text.strip(), re.IGNORECASE)
        if m_letter:
            return col_letter_to_num(m_letter.group(1))
        m_num = re.findall(r'\d+', option_text)
        return int(m_num[0]) if m_num else fallback

    def parse_optional_col_num_from_option(self, option_text):
        if not option_text or is_no_header_match_option(option_text):
            return None
        return self.parse_col_num_from_option(option_text, None)

    def extract_header_name_from_option(self, option_text, header_map, col):
        if not col or is_no_header_match_option(option_text):
            return NO_HEADER_MATCH_OPTION
        if option_text and ' - ' in option_text:
            return self.clean_basis_title(option_text.split(' - ', 1)[1])
        return self.clean_basis_title(header_map.get(col, f'Col {col}'))

    def is_identifier_match_basis(self, *headers):
        joined = ' '.join(str(header or '') for header in headers)
        return any(keyword in joined.upper() for keyword in ['PO', 'PR', 'NO', '번호', '계약번호', '공사번호', '입찰번호'])

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        self.COLOR_BG = '#F4F6F9'
        self.COLOR_PRIMARY = '#1A365D'
        self.COLOR_SECONDARY = '#2B6CB0'
        self.COLOR_ACCENT = '#D69E2E'
        self.COLOR_SUCCESS = '#2F855A'
        self.COLOR_CARD = '#FFFFFF'

        self.root.configure(bg=self.COLOR_BG)

        self.style.configure('TFrame', background=self.COLOR_BG)
        self.style.configure('Card.TFrame', background=self.COLOR_CARD, relief='groove', borderwidth=1)
        self.style.configure('Header.TLabel', background=self.COLOR_PRIMARY, foreground='#FFFFFF', font=('Segoe UI', 16, 'bold'))
        self.style.configure('SubHeader.TLabel', background=self.COLOR_PRIMARY, foreground='#CBD5E0', font=('Segoe UI', 9))
        self.style.configure('Title.TLabel', background=self.COLOR_CARD, foreground=self.COLOR_PRIMARY, font=('Segoe UI', 11, 'bold'))
        self.style.configure('Status.TLabel', background=self.COLOR_CARD, foreground='#4A5568', font=('Segoe UI', 9))
        
        self.style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), background=self.COLOR_SECONDARY, foreground='#FFFFFF')
        self.style.map('Primary.TButton', background=[('active', '#2C5282')])

        self.style.configure('Accent.TButton', font=('Segoe UI', 11, 'bold'), background=self.COLOR_ACCENT, foreground='#FFFFFF')
        self.style.map('Accent.TButton', background=[('active', '#B7791F')])

        self.style.configure('Copy.TButton', font=('Segoe UI', 10, 'bold'), background='#2F855A', foreground='#FFFFFF')
        self.style.map('Copy.TButton', background=[('active', '#276749')])

        self.style.configure('Reset.TButton', font=('Segoe UI', 9, 'bold'), background='#E53E3E', foreground='#FFFFFF')
        self.style.map('Reset.TButton', background=[('active', '#C53030')])

        self.style.configure('Setting.TButton', font=('Segoe UI', 10, 'bold'), background='#805AD5', foreground='#FFFFFF')
        self.style.map('Setting.TButton', background=[('active', '#6B46C1')])

        self.style.configure('Treeview', font=('Consolas', 9), rowheight=28)
        self.style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'), background='#E2E8F0', foreground='#1A365D')

    def build_ui(self):
        header_frame = tk.Frame(self.root, bg=self.COLOR_PRIMARY, height=75)
        header_frame.pack(fill='x', side='top')
        
        lbl_title = ttk.Label(header_frame, text='B2B 데이터 ↔ 발주내역 스마트 매핑 시스템', style='Header.TLabel')
        lbl_title.pack(anchor='w', padx=20, pady=(12, 2))
        
        lbl_sub = ttk.Label(header_frame, text='원천 B2BIEAMS 엑셀 데이터를 마우스 범위 지정 기반으로 발주내역 시트에 정밀 동적 매핑 (범위 정밀 타격 + 수동 매핑 지원)', style='SubHeader.TLabel')
        lbl_sub.pack(anchor='w', padx=20, pady=(0, 10))

        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill='both', expand=True)

        card1 = ttk.Frame(main_container, style='Card.TFrame', padding=15)
        card1.pack(fill='x', pady=(0, 12))

        ttk.Label(card1, text='📌 1단계: 소스 파일 선택 (B2BIEAMS정보_YYYYMMDD.xlsx)', style='Title.TLabel').pack(anchor='w', pady=(0, 8))

        f1_input = ttk.Frame(card1, style='Card.TFrame')
        f1_input.pack(fill='x')

        self.txt_source = ttk.Entry(f1_input, textvariable=self.source_file_path, font=('Consolas', 9))
        self.txt_source.pack(side='left', fill='x', expand=True, padx=(0, 8))

        btn_browse_source = ttk.Button(f1_input, text='📂 소스 파일 찾기...', command=self.on_select_source_file)
        btn_browse_source.pack(side='right')

        self.lbl_source_status = ttk.Label(card1, text='ⓘ 파일 선택 시 헤더 행 개수, 병합 상태, 컬럼 구조를 자동 분석하여 공지합니다.', style='Status.TLabel')
        self.lbl_source_status.pack(anchor='w', pady=(8, 0))

        card2 = ttk.Frame(main_container, style='Card.TFrame', padding=15)
        card2.pack(fill='x', pady=(0, 12))

        ttk.Label(card2, text='🎯 2단계: 타겟 파일 선택 & 매핑 대상 마우스 범위 지정 (재선택/초기화 가능)', style='Title.TLabel').pack(anchor='w', pady=(0, 8))

        f2_input = ttk.Frame(card2, style='Card.TFrame')
        f2_input.pack(fill='x')

        self.txt_target = ttk.Entry(f2_input, textvariable=self.target_file_path, font=('Consolas', 9))
        self.txt_target.pack(side='left', fill='x', expand=True, padx=(0, 8))

        btn_browse_target = ttk.Button(f2_input, text='📂 타겟 파일 찾기...', command=self.on_select_target_file)
        btn_browse_target.pack(side='right')

        f2_interactive = ttk.Frame(card2, style='Card.TFrame')
        f2_interactive.pack(fill='x', pady=(10, 0))

        self.btn_drag_select = ttk.Button(f2_interactive, text='🖥️ 엑셀 창 열기 & 마우스 범위 확정', style='Primary.TButton', command=self.on_interactive_drag_select)
        self.btn_drag_select.pack(side='left', padx=(0, 8))
        HoverTooltip(self.btn_drag_select, (
            "[B2B 매핑 자동화] 대상 기준 수동 변경 워크플로\n\n"
            "1. 소스 파일과 타겟 발주내역 파일을 먼저 선택하고 헤더 분석 완료 로그를 확인합니다.\n"
            "2. 이 버튼을 눌러 발주내역 Excel을 전면으로 연 뒤, 계약명에 한정하지 말고 원하는 기준 열의 행 범위를 드래그합니다.\n"
            "   예: 계약명 X4942:X4963, PO번호 AA4942:AA4963, PR번호 AB4942:AB4963.\n"
            "3. 팝업에서 [소스 파일 기준 헤더]와 [기준 파일 기준 헤더]를 같은 의미의 헤더로 맞춥니다.\n"
            "   예: 소스 PO번호 Col 07 ↔ 기준 PO번호 AA열 Col 27.\n"
            "4. 미리보기에 표시되는 Col 좌표를 확인하고 [범위 확정]을 누릅니다.\n"
            "5. 데이터 쓰기 대상 컬럼도 바꿔야 한다면 [헤더 매핑 설정]에서 소스↔타겟 쓰기 매핑을 별도로 편집합니다.\n\n"
            "주의: Excel에서 드래그한 범위는 '행과 기준 열'을 잡는 단계이고, 실제 기준 의미는 팝업의 두 헤더 콤보박스가 결정합니다."
        ))

        self.btn_reset_range = ttk.Button(f2_interactive, text='🔄 범위 선택 초기화', style='Reset.TButton', command=self.on_reset_range_selection)
        self.btn_reset_range.pack(side='left', padx=(0, 10))

        self.lbl_target_range = ttk.Label(f2_interactive, text='지정된 범위: 미선택 (엑셀에서 직접 마우스로 선택하세요)', style='Status.TLabel')
        self.lbl_target_range.pack(side='left', fill='x', expand=True)

        card3 = ttk.Frame(main_container, style='Card.TFrame', padding=15)
        card3.pack(fill='both', expand=True, pady=(0, 10))

        f3_top = ttk.Frame(card3, style='Card.TFrame')
        f3_top.pack(fill='x', pady=(0, 4))

        ttk.Label(f3_top, text='🚀 3단계: 스마트 데이터 매핑 실행 및 실시간 진행 현황', style='Title.TLabel').pack(side='left')
        
        self.btn_run = ttk.Button(f3_top, text='▶️ 스마트 매핑 실행', style='Accent.TButton', command=self.on_run_mapping)
        self.btn_run.pack(side='right')

        # 헤더 매핑 설정 버튼 + 모드 라벨
        f3_mapping = ttk.Frame(card3, style='Card.TFrame')
        f3_mapping.pack(fill='x', pady=(0, 6))

        self.btn_header_mapping = ttk.Button(f3_mapping, text='⚙️ 헤더 매핑 설정 (자동/수동)', style='Setting.TButton', command=self.show_header_mapping_editor)
        self.btn_header_mapping.pack(side='left', padx=(0, 10))
        HoverTooltip(self.btn_header_mapping, (
            "[B2B 매핑 자동화] 소스↔타겟 헤더 매핑 설정 가이드\n\n"
            "1. 이 설정은 매칭 성공 후 어떤 소스 값을 발주내역의 어떤 열에 쓸지 결정합니다.\n"
            "2. 목록에서 항목을 더블클릭하거나 [선택 항목 편집]을 눌러 소스 헤더와 타겟 헤더를 각각 선택합니다.\n"
            "3. 콤보박스에는 실제 분석된 헤더명과 Col 좌표가 함께 표시됩니다.\n"
            "4. PO번호 기준으로 행 매칭을 했다면, 필요 시 PO번호 쓰기 매핑도 소스 PO번호 ↔ 타겟 PO번호(예: AA Col 27)로 확인합니다.\n"
            "5. [확정 저장] 후 실행하면 현재 매핑 목록이 실행부에 반영됩니다.\n\n"
            "행 매칭 기준 변경은 [엑셀 창 열기 & 마우스 범위 확정] 팝업에서 하고, 값 쓰기 컬럼 변경은 이 화면에서 합니다."
        ))

        self.lbl_mapping_mode = ttk.Label(f3_mapping, text='매핑 모드: 📋 자동 (7쌍 기본 연결)', font=('Segoe UI', 9, 'bold'), foreground='#805AD5', background=self.COLOR_CARD)
        self.lbl_mapping_mode.pack(side='left', fill='x', expand=True)

        self.lbl_mapping_status = ttk.Label(card3, text='ⓘ 준비: 매핑 실행 버튼을 누르면 실시간 작업 상황이 여기에 명확하게 공지됩니다.', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0', background=self.COLOR_CARD)
        self.lbl_mapping_status.pack(anchor='w', pady=(0, 6))

        self.progress = ttk.Progressbar(card3, mode='determinate')
        self.progress.pack(fill='x', pady=(0, 8))

        self.log_area = scrolledtext.ScrolledText(card3, height=10, font=('Consolas', 9), bg='#1E1E1E', fg='#D4D4D4', insertbackground='#FFFFFF')
        self.log_area.pack(fill='both', expand=True)

        self.log_area.tag_config('INFO', foreground='#4EC9B0')
        self.log_area.tag_config('SUCCESS', foreground='#6A9955', font=('Consolas', 9, 'bold'))
        self.log_area.tag_config('WARN', foreground='#CE9178')
        self.log_area.tag_config('ERROR', foreground='#F44747', font=('Consolas', 9, 'bold'))

        self.log('INFO', '시스템이 준비되었습니다. 소스 파일과 타겟 파일을 순서대로 선택하세요.')
        if not COM_AVAILABLE:
            self.log('WARN', 'win32com 모듈이 없습니다. Excel 작업 시작 시 pywin32 자동 설치를 1회 시도합니다.')

    def log(self, tag, message):
        timestamp = datetime.datetime.now().strftime('[%H:%M:%S] ')

        def _append():
            self.log_area.insert(tk.END, timestamp + message + '\n', tag)
            self.log_area.see(tk.END)

        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.root.after(0, _append)

    def update_mapping_status(self, progress_val, status_text):
        def _update():
            self.progress['value'] = progress_val
            self.lbl_mapping_status.config(text=status_text, foreground='#1A365D')
            self.root.update_idletasks()
        self.root.after(0, _update)

    def on_reset_range_selection(self):
        self.target_range_info = None
        self.selected_sheet_name.set('')
        self.selected_range_address.set('')
        self.lbl_target_range.config(
            text='지정된 범위: 미선택 (버튼을 누르고 엑셀 화면에서 마우스로 범위를 지정하세요)',
            foreground='#4A5568',
            font=('Segoe UI', 9)
        )
        self.log('INFO', '기존 마우스 범위 선택이 완전 초기화되었습니다.')
        messagebox.showinfo('범위 초기화 완료', '선택된 매핑 범위가 초기화되었습니다.\n[엑셀 창 열기 & 범위 지정 확정] 버튼을 눌러 새로 드래그하세요.')

    def on_select_source_file(self):
        file_path = filedialog.askopenfilename(
            title='B2B 소스 엑셀 파일 선택',
            filetypes=[('Excel Files', '*.xlsx;*.xlsb;*.xls'), ('All Files', '*.*')]
        )
        if not file_path:
            return

        self.source_file_path.set(file_path)
        self.log('INFO', f'소스 파일 선택됨: {os.path.basename(file_path)}')

        self.active_tasks['source'] = True
        self.update_dynamic_ui_state()

        threading.Thread(target=self.inspect_source_header, args=(file_path,), daemon=True).start()

    def inspect_source_header(self, file_path):
        self.log('INFO', '소스 파일 구조 사전 자가 검증 중...')
        ensure_excel_com_available()
        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(os.path.abspath(file_path), ReadOnly=True)
            ws = wb.Worksheets(1)

            used_range = ws.UsedRange
            total_rows = used_range.Rows.Count
            total_cols = used_range.Columns.Count

            header_rows_count = 1
            merged_cells_info = []
            
            for r in range(1, min(6, total_rows + 1)):
                for c in range(1, min(30, total_cols + 1)):
                    try:
                        cell = ws.Cells(r, c)
                        if getattr(cell, 'MergeCells', False):
                            addr = cell.MergeArea.Address.replace('$', '')
                            if addr not in merged_cells_info:
                                merged_cells_info.append(addr)
                            rows_in_merge = cell.MergeArea.Rows.Count
                            if rows_in_merge > header_rows_count:
                                header_rows_count = rows_in_merge
                    except Exception:
                        pass

            header_names = []
            header_row_idx = header_rows_count
            for c in range(1, total_cols + 1):
                val = safe_cell_value(ws, header_row_idx, c)
                if val and not self.is_probable_data_value(val):
                    if header_row_idx > 1:
                        val_top = safe_cell_value(ws, 1, c)
                        header_names.append(self.compose_header_label(val_top, val) or str(val).strip())
                    else:
                        header_names.append(self.clean_basis_title(val))
                else:
                    val_top = safe_cell_value(ws, 1, c)
                    clean_top = self.clean_basis_title(val_top)
                    header_names.append(clean_top if clean_top else f'Unassigned_{c}')

            required_columns = [
                '계약금액', 'PO전송견적금액', 'PO번호', 'PR번호', 
                '생성일자', '공사/입찰/계약번호', '계약일자(체결일)', '공사/입찰명'
            ]
            
            found_cols = []
            missing_cols = []
            for req in required_columns:
                matched = any(req in col_name or col_name in req for col_name in header_names)
                if matched:
                    found_cols.append(req)
                else:
                    missing_cols.append(req)

            self.source_info = {
                'header_rows': header_rows_count,
                'merged_cells': merged_cells_info,
                'total_rows': total_rows - header_rows_count,
                'found_cols': found_cols,
                'missing_cols': missing_cols,
                'header_names': header_names
            }

            merged_str = ', '.join(merged_cells_info[:5]) + ('...' if len(merged_cells_info) > 5 else '') if merged_cells_info else '없음 (단일 셀 구획)'
            status_msg = f'✅ 사전 검증 완료: 헤더 {header_rows_count}개 행 | 병합 상태: [{merged_str}] | 전체 데이터 레코드: {total_rows - header_rows_count:,}건\n   [필수 컬럼 확인]: {len(found_cols)}/8 항목 식별 완료'
            
            self.root.after(0, lambda: self.lbl_source_status.config(text=status_msg, foreground='#2F855A'))
            self.root.after(0, lambda: self.show_source_analysis_popup(self.source_info))

        except Exception as e:
            err_msg = f'소스 사전 검증 중 오류: {e!s}'
            self.log('ERROR', err_msg)
            self.root.after(0, lambda msg=err_msg: self.lbl_source_status.config(text=f'❌ 소스 파일 사전 검증 실패: {msg}', foreground='#C53030'))
        finally:
            if wb:
                try: wb.Close(False)
                except Exception: pass
            if excel:
                try: excel.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
            self.active_tasks['source'] = False
            self.root.after(0, self.on_parsing_finished)

    def show_source_analysis_popup(self, info):
        popup = tk.Toplevel(self.root)
        popup.title('📋 소스 파일 구조 분석 사전 공지 리포트')
        popup.geometry('620x450')
        popup.grab_set()

        frame = ttk.Frame(popup, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text='📊 B2BIEAMS 소스 데이터 검증 결과', font=('Segoe UI', 12, 'bold'), foreground=self.COLOR_PRIMARY).pack(anchor='w', pady=(0, 10))

        info_text = f'• 헤더 행 갯수: {info["header_rows"]}개 행 (Header Offset)\n• 헤더 병합 셀 구획: {len(info["merged_cells"])}개 영역 탐지됨\n• 순수 데이터 레코드 수: {info["total_rows"]:,} 건\n\n✔ 감지된 매핑 컬럼 ({len(info["found_cols"])}개):\n   {", ".join(info["found_cols"])}\n'
        if info['missing_cols']:
            info_text += f'\n⚠️ 유의: 아래 컬럼명을 직접 찾지 못해 유사 탐색이 수행됩니다:\n   {", ".join(info["missing_cols"])}\n'

        lbl_info = ttk.Label(frame, text=info_text, justify='left', font=('Segoe UI', 9))
        lbl_info.pack(anchor='w', fill='x', pady=(0, 10))

        ttk.Label(frame, text='헤더 전체 컬럼 목록 (미리보기):', font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(5, 2))
        
        list_box = tk.Listbox(frame, height=8, font=('Consolas', 9))
        list_box.pack(fill='both', expand=True, pady=(0, 10))
        for idx, col in enumerate(info['header_names'], 1):
            list_box.insert(tk.END, f'Col {idx:02d}: {col}')

        btn_confirm = ttk.Button(frame, text='확인 및 계속 진행', style='Primary.TButton', command=popup.destroy)
        btn_confirm.pack(anchor='e')

        self.log('SUCCESS', f'소스 구조 검증 완료 (헤더 {info["header_rows"]}행, 데이터 {info["total_rows"]:,}건)')

    def on_select_target_file(self):
        file_path = filedialog.askopenfilename(
            title='발주내역 타겟 엑셀 파일 선택',
            filetypes=[('Excel Files', '*.xlsb;*.xlsx;*.xls'), ('All Files', '*.*')]
        )
        if not file_path:
            return

        self.target_file_path.set(file_path)
        self.log('INFO', f'타겟 파일 선택됨: {os.path.basename(file_path)}')

        self.active_tasks['target'] = True
        self.update_dynamic_ui_state()

        threading.Thread(target=self.inspect_target_header, args=(file_path,), daemon=True).start()

    def inspect_target_header(self, file_path):
        """타겟 발주내역 엑셀 파일의 실제 헤더명을 직접 읽어와 자가 파싱"""
        self.log('INFO', '타겟(발주내역) 파일 실제 엑셀 헤더 구조 정밀 파싱 중...')
        ensure_excel_com_available()
        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(os.path.abspath(file_path), ReadOnly=True)
            ws = None
            try:
                ws = wb.Worksheets('집계')
            except Exception:
                ws = wb.Worksheets(1)

            target_headers = {}
            for c in range(1, 53):
                h_val = None
                for r in range(1, 6):
                    v = safe_cell_value(ws, r, c)
                    if v and not self.is_probable_data_value(v):
                        sv = self.clean_basis_title(v)
                        if len(sv) >= 1 and not sv.startswith('Row') and not sv.startswith('No'):
                            h_val = sv
                            break
                if h_val:
                    target_headers[c] = h_val

            self.target_info = {
                'header_names': target_headers
            }

            # 현재 매핑 항목들의 타겟 헤더명을 실제 파싱된 엑셀 셀 명칭으로 즉시 업데이트
            updated_count = 0
            for m in self.header_mappings:
                c = m['target_col']
                if c in target_headers:
                    m['target_name'] = target_headers[c]
                    updated_count += 1

            self.log('SUCCESS', f'🎯 타겟 엑셀 실제 헤더명 동기화 완료! ({len(target_headers)}개 컬럼 파싱, {updated_count}개 매핑 일치)')

        except Exception as e:
            self.log('WARN', f'타겟 엑셀 헤더 사전 파싱 중 참고: {e}')
        finally:
            if wb:
                try: wb.Close(False)
                except Exception: pass
            if excel:
                try: excel.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
            self.active_tasks['target'] = False
            self.root.after(0, self.on_parsing_finished)

    def bring_excel_to_front_and_get_selection(self, target_path):
        ensure_excel_com_available()
        if not self.interactive_excel_com_initialized:
            pythoncom.CoInitialize()
            self.interactive_excel_com_initialized = True
        sel_addr = ''
        abs_target = os.path.abspath(target_path)
        fn = os.path.basename(abs_target)
        fn_lower = fn.lower()
        activation_started_at = time.perf_counter()
        excel = None
        wb = None

        def is_usable_excel_app(app):
            if app is None:
                return False
            try:
                _ = int(app.Hwnd)
                _ = app.Workbooks.Count
                return True
            except Exception:
                return False

        def find_target_workbook(app):
            try:
                for open_wb in app.Workbooks:
                    try:
                        wb_full_name = str(open_wb.FullName or '').lower()
                        wb_name = str(open_wb.Name or '').lower()
                        if wb_full_name == abs_target.lower() or wb_name == fn_lower:
                            return open_wb
                    except Exception:
                        pass
            except Exception:
                pass
            return None

        def configure_visible_excel(app):
            try:
                app.Visible = True
                app.UserControl = True
                app.DisplayAlerts = True
                app.WindowState = -4137  # xlMaximized
                _ = int(app.Hwnd)
                self.interactive_excel_app = app
                self.persist_excel_app = app
                return True
            except Exception:
                return False

        try:
            excel_candidates = [getattr(self, 'interactive_excel_app', None), getattr(self, 'persist_excel_app', None)]
            try:
                excel_candidates.append(win32com.client.GetActiveObject('Excel.Application'))
            except Exception:
                pass

            for candidate in excel_candidates:
                if not (is_usable_excel_app(candidate) and configure_visible_excel(candidate)):
                    continue
                candidate_wb = find_target_workbook(candidate)
                if candidate_wb:
                    excel = candidate
                    wb = candidate_wb
                    break
                if excel is None:
                    excel = candidate

            if excel:
                wb = wb or find_target_workbook(excel)

            if not wb:
                self.log('INFO', '대상 파일이 현재 Excel에 열려 있지 않아 Windows 기본 열기로 빠르게 활성화합니다.')
                shell_open_started = False
                try:
                    os.startfile(abs_target)
                    shell_open_started = True
                except Exception as exc:
                    self.log('WARN', f'OS 기본 열기로 Excel 파일을 열지 못했습니다. COM 열기로 재시도합니다: {exc}')

                if shell_open_started:
                    deadline = time.perf_counter() + 6.0
                    while time.perf_counter() < deadline and not wb:
                        try:
                            candidate = win32com.client.GetActiveObject('Excel.Application')
                            if is_usable_excel_app(candidate) and configure_visible_excel(candidate):
                                excel = candidate
                                wb = find_target_workbook(excel)
                                if wb:
                                    break
                        except Exception:
                            pass
                        time.sleep(0.15)

            if not wb:
                if excel is None:
                    self.log('WARN', '기존 Excel COM 인스턴스를 찾지 못해 새 Excel 창을 생성해 전면 활성화를 재시도합니다.')
                    for factory in (win32com.client.DispatchEx, win32com.client.Dispatch):
                        try:
                            candidate = factory('Excel.Application')
                            if configure_visible_excel(candidate):
                                excel = candidate
                                break
                        except Exception:
                            continue
                if excel:
                    try:
                        wb = excel.Workbooks.Open(abs_target)
                    except Exception as exc:
                        self.log('WARN', f'타겟 파일을 Excel COM으로 열지 못했습니다: {exc}')

            if excel:
                if wb:
                    self.interactive_target_wb = wb
                    self.persist_target_wb = wb
                    try:
                        wb.Activate()
                    except Exception:
                        pass
                    try:
                        wb.Windows(1).Activate()
                    except Exception:
                        pass
            elif not wb:
                try:
                    os.startfile(abs_target)
                except Exception as exc:
                    self.log('ERROR', f'Excel 실행 실패: {exc}')

            if COM_AVAILABLE and excel:
                try:
                    import win32api
                    import win32process
                    target_hwnd = None
                    for hwnd_getter in (
                        lambda: int(excel.Hwnd),
                        lambda: int(wb.Windows(1).Hwnd) if wb else None,
                    ):
                        try:
                            candidate_hwnd = hwnd_getter()
                            if candidate_hwnd and win32gui.IsWindow(candidate_hwnd):
                                target_hwnd = candidate_hwnd
                                break
                        except Exception:
                            pass

                    if target_hwnd and win32gui.IsWindow(target_hwnd):
                        if win32gui.IsIconic(target_hwnd):
                            win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                        else:
                            win32gui.ShowWindow(target_hwnd, win32con.SW_SHOW)
                        win32gui.ShowWindow(target_hwnd, win32con.SW_MAXIMIZE)
                        win32gui.SetWindowPos(
                            target_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                        )
                        win32gui.SetWindowPos(
                            target_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                        )

                        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                        attached_threads = []
                        app_thread = win32api.GetCurrentThreadId()
                        try:
                            fore_thread = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[0]
                            target_thread = win32process.GetWindowThreadProcessId(target_hwnd)[0]
                            for thread_id in {fore_thread, target_thread}:
                                if thread_id and thread_id != app_thread:
                                    try:
                                        win32process.AttachThreadInput(app_thread, thread_id, True)
                                        attached_threads.append(thread_id)
                                    except Exception:
                                        pass
                            for activate_call in (
                                lambda: win32gui.SetForegroundWindow(target_hwnd),
                                lambda: win32gui.BringWindowToTop(target_hwnd),
                                lambda: win32gui.SetActiveWindow(target_hwnd),
                                lambda: win32gui.SetFocus(target_hwnd),
                            ):
                                try:
                                    activate_call()
                                except Exception:
                                    pass
                        finally:
                            for thread_id in attached_threads:
                                try:
                                    win32process.AttachThreadInput(app_thread, thread_id, False)
                                except Exception:
                                    pass
                            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)

                        try:
                            shell = win32com.client.Dispatch('WScript.Shell')
                            app_titles = []
                            if wb:
                                app_titles.append(wb.Name)
                            app_titles.extend([fn, excel.Caption, 'Microsoft Excel', 'Excel'])
                            for title in app_titles:
                                try:
                                    if title and shell.AppActivate(title):
                                        break
                                except Exception:
                                    pass
                        except Exception:
                            pass
                except Exception:
                    pass

            if excel:
                try:
                    if excel.Selection:
                        sel_addr = str(excel.Selection.Address).replace('$', '')
                except Exception:
                    pass
        except Exception as e:
            self.log('WARN', f'엑셀 활성화 시도 중 경고: {e}')
        self.log('INFO', f'Excel 활성화 처리 완료: {time.perf_counter() - activation_started_at:.2f}초')
        return sel_addr

    def on_interactive_drag_select(self):
        if self.active_tasks['source'] or self.active_tasks['target']:
            self.queued_action = self.on_interactive_drag_select
            self.log('INFO', '⏳ 백그라운드 분석이 진행 중입니다. 완료되는 즉시 [마우스 범위 확정] 창을 자동으로 엽니다.')
            return

        target_path = self.target_file_path.get()
        if not target_path or not os.path.exists(target_path):
            messagebox.showwarning('입력 확인', '먼저 올바른 타겟 엑셀 파일(.xlsb/.xlsx/.xls)을 선택하세요.')
            return

        if self.excel_activation_busy:
            self.log('INFO', 'Excel 전면 활성화가 이미 진행 중입니다. 잠시 후 다시 시도하세요.')
            return

        if self.range_popup is not None and self.range_popup.winfo_exists():
            if not self.confirm_excel_cleanup_before_activation(parent=self.range_popup, show_when_visible_only=False):
                return
            self.range_popup.lift()
            return

        if not self.confirm_excel_cleanup_before_activation(parent=self.root):
            return

        self.log('INFO', 'Target Excel 화면을 전면 최대로 활성화합니다...')
        self.excel_activation_busy = True
        try:
            curr_sel = self.bring_excel_to_front_and_get_selection(target_path)
        finally:
            self.excel_activation_busy = False
        self.show_range_confirmation_dialog(target_path, curr_sel)

    def show_range_confirmation_dialog(self, target_path, curr_sel):
        if self.range_popup is not None and self.range_popup.winfo_exists():
            self.range_popup.lift()
            return

        popup = tk.Toplevel(self.root)
        self.range_popup = popup
        popup.title('🎯 매핑 대상 행 범위 확인 및 확정')
        popup.geometry('780x560')
        popup.attributes('-topmost', True)

        f = ttk.Frame(popup, padding=20)
        f.pack(fill='both', expand=True)

        lbl_title = ttk.Label(f, text='🎯 매핑 대상 행 범위 확정 (수동 기준 헤더 연계)', font=('Segoe UI', 12, 'bold'), foreground=self.COLOR_PRIMARY)
        lbl_title.pack(anchor='w', pady=(0, 8))
        HoverTooltip(lbl_title, (
            "이 창의 목적은 '어떤 행들을 매핑할지'와 '그 행을 어떤 기준 헤더로 찾을지'를 동시에 확정하는 것입니다.\n\n"
            "계약명만 사용할 필요는 없습니다. Excel에서 PO번호, PR번호, 계약번호 등 원하는 기준 열을 드래그한 뒤 아래 두 콤보박스에서 같은 의미의 소스/기준 헤더를 선택하세요."
        ))

        msg_desc = (
            '엑셀 창이 화면 전면에 열렸습니다! 엑셀 화면에서 마우스로 대상 행 영역\n'
            '(예: X4942:X4963 또는 4942:4963)을 드래그하거나 직접 입력한 뒤,\n'
            '소스 파일과 기준 파일의 행 매칭 기준 헤더를 확인 후 확정하세요.'
        )
        ttk.Label(f, text=msg_desc, justify='left', font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 12))

        source_header_map = self.build_source_header_map()
        target_header_map = self.build_target_header_map()
        max_source_col = max(50, max(source_header_map.keys()) if source_header_map else 50)
        max_target_col = max(52, max(target_header_map.keys()) if target_header_map else 52)
        source_options = [NO_HEADER_MATCH_OPTION] + [self.format_source_header_option(i, source_header_map) for i in range(1, max_source_col + 1)]
        target_options = [NO_HEADER_MATCH_OPTION] + [self.format_target_header_option(i, target_header_map) for i in range(1, max_target_col + 1)]

        source_default_col = self.source_match_col or self.find_header_col_by_keywords(
            source_header_map,
            [['공사/입찰명', '공사명', '계약명', '입찰명'], ['PO번호', '계약번호', '공사/입찰/계약번호']],
            None
        )
        target_default_col = self.target_match_col or self.find_header_col_by_keywords(
            target_header_map,
            [['계약명', '공사/입찰명', '공사명'], ['PO번호', '계약번호', '공사번호']],
            24
        )

        initial_val = 'X4942:X4963'
        if curr_sel and ':' in curr_sel:
            initial_val = curr_sel
        elif curr_sel:
            m_r = re.findall(r'\d+', curr_sel)
            if m_r and int(m_r[0]) > 10:
                initial_val = curr_sel

        def extract_range_target_col(range_text):
            m_col = re.search(r'([A-Z]+)\s*\$?\d+', str(range_text or '').upper().replace('$', ''))
            if m_col:
                col_num = col_letter_to_num(m_col.group(1))
                return col_num, col_num_to_letter(col_num)
            return target_default_col, col_num_to_letter(target_default_col)

        f_entry = ttk.Frame(f)
        f_entry.pack(fill='x', pady=(0, 12))

        ttk.Label(f_entry, text='매핑 행 범위:', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=(0, 8))
        var_input = tk.StringVar(value=initial_val)
        ent_range = ttk.Entry(f_entry, textvariable=var_input, font=('Consolas', 11, 'bold'))
        ent_range.pack(side='left', fill='x', expand=True, padx=(0, 8))
        HoverTooltip(ent_range, (
            "Excel에서 드래그한 행 범위입니다.\n\n"
            "입력 예시:\n"
            "- 계약명 기준: X4942:X4963\n"
            "- PO번호 기준: AA4942:AA4963\n"
            "- 행 번호만 수동 입력: 4942:4963\n\n"
            "열 문자가 들어오면 아래 [기준 파일 기준 헤더]에서 선택한 열 문자와 좌표로 최종 범위가 저장됩니다."
        ))

        def bring_and_refresh():
            if self.excel_activation_busy:
                self.log('INFO', 'Excel 전면 활성화가 이미 진행 중입니다. 잠시 후 다시 시도하세요.')
                return
            if not self.confirm_excel_cleanup_before_activation(parent=popup, show_when_visible_only=False):
                return
            self.excel_activation_busy = True
            btn_act_excel.config(state='disabled')
            try:
                new_sel = self.bring_excel_to_front_and_get_selection(target_path)
                if new_sel and new_sel != 'X4' and new_sel != 'A1':
                    var_input.set(new_sel)
                    autofill_title_from_range(force=True)
                    sync_basis_from_title(commit=False)
                    self.log('INFO', f'엑셀 최신 드래그 영역 획득: {new_sel}')
                else:
                    self.log('INFO', f'엑셀 창 전면 활성화 완료 (현재 선택: {new_sel})')
            finally:
                btn_act_excel.config(state='normal')
                self.excel_activation_busy = False

        btn_act_excel = ttk.Button(f_entry, text='🖥️ 엑셀 창 열기 & 영역 읽기', command=bring_and_refresh)
        btn_act_excel.pack(side='right')
        HoverTooltip(btn_act_excel, (
            "활성화된 발주내역 Excel에서 현재 선택한 범위를 다시 읽습니다.\n\n"
            "절차:\n"
            "1. Excel 화면에서 원하는 기준 열의 행 범위를 드래그합니다.\n"
            "2. 이 버튼을 누릅니다.\n"
            "3. 읽힌 주소가 [매핑 행 범위]에 반영되는지 확인합니다.\n"
            "4. 아래 기준 헤더 콤보박스가 드래그한 열과 같은 의미인지 확인합니다."
        ))

        selected_col_for_title, _ = extract_range_target_col(initial_val)
        initial_title = target_header_map.get(selected_col_for_title, target_header_map.get(target_default_col, '계약명'))
        f_title_input = ttk.Frame(f)
        f_title_input.pack(fill='x', pady=(0, 10))
        ttk.Label(f_title_input, text='선택 열 제목:', font=('Segoe UI', 10, 'bold'), width=14, anchor='w').pack(side='left', padx=(0, 8))
        var_selected_title = tk.StringVar(value=initial_title)
        ent_selected_title = ttk.Entry(f_title_input, textvariable=var_selected_title, font=('Consolas', 10, 'bold'))
        ent_selected_title.pack(side='left', fill='x', expand=True, padx=(0, 8))
        HoverTooltip(ent_selected_title, (
            "활성화된 Excel에서 드래그한 열의 제목을 입력합니다.\n\n"
            "예: 계약명, PO번호, PR번호, 계약번호, 공사번호.\n"
            "입력한 제목을 기준으로 소스 파일의 같은 의미 헤더를 자동 탐색하고, 기준 파일의 선택 열과 기존 헤더 매핑 목록을 함께 갱신합니다."
        ))

        basis_box = ttk.LabelFrame(f, text='행 매칭 기준 헤더 수동 선택', padding=10)
        basis_box.pack(fill='x', pady=(0, 12))

        f_src_basis = ttk.Frame(basis_box)
        f_src_basis.pack(fill='x', pady=(0, 6))
        ttk.Label(f_src_basis, text='소스 파일 기준 헤더:', font=('Segoe UI', 9, 'bold'), width=18, anchor='w').pack(side='left')
        var_source_basis = tk.StringVar(value=self.format_source_header_option(source_default_col, source_header_map) if source_default_col else NO_HEADER_MATCH_OPTION)
        cmb_source_basis = ttk.Combobox(f_src_basis, textvariable=var_source_basis, values=source_options, font=('Consolas', 9), width=58)
        cmb_source_basis.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_source_basis, (
            "소스 파일에서 행을 찾을 때 사용할 기준 헤더입니다.\n\n"
            "계약명 기준이면 공사/입찰명 또는 계약명을 선택합니다.\n"
            "PO번호 기준이면 소스 파일의 PO번호 헤더와 Col 좌표를 선택합니다.\n"
            "선택한 값은 실행부의 소스 레코드 키로 저장되며, 같은 헤더/좌표가 데이터 쓰기 매핑에 있으면 로그에 연계 확인이 표시됩니다."
        ))

        f_tgt_basis = ttk.Frame(basis_box)
        f_tgt_basis.pack(fill='x', pady=(0, 4))
        ttk.Label(f_tgt_basis, text='기준 파일 기준 헤더:', font=('Segoe UI', 9, 'bold'), width=18, anchor='w').pack(side='left')
        var_target_basis = tk.StringVar(value=self.format_target_header_option(target_default_col, target_header_map))
        cmb_target_basis = ttk.Combobox(f_tgt_basis, textvariable=var_target_basis, values=target_options, font=('Consolas', 9), width=58)
        cmb_target_basis.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_target_basis, (
            "발주내역 파일에서 대상 기준으로 사용할 헤더입니다.\n\n"
            "Excel에서 AA열 PO번호를 드래그했다면 여기서도 PO번호 / AA열 / Col 27을 선택합니다.\n"
            "이 선택값이 최종 range_address와 col_idx로 저장되므로, 보고서와 빨간색 누락 표시도 이 열 기준으로 동작합니다."
        ))

        lbl_basis_preview = ttk.Label(basis_box, text='', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0')
        lbl_basis_preview.pack(anchor='w', pady=(4, 0))

        lbl_preview = ttk.Label(f, text='', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0')
        lbl_preview.pack(anchor='w', pady=(0, 15))

        def get_selected_basis():
            source_col = self.parse_optional_col_num_from_option(var_source_basis.get())
            target_col = self.parse_optional_col_num_from_option(var_target_basis.get())
            source_header = self.extract_header_name_from_option(var_source_basis.get(), source_header_map, source_col)
            target_header = self.extract_header_name_from_option(var_target_basis.get(), target_header_map, target_col)
            return source_col, source_header, target_col, target_header

        def autofill_title_from_range(force=False):
            range_target_col, _ = extract_range_target_col(var_input.get())
            detected_title = target_header_map.get(range_target_col, '')
            if detected_title and (force or not var_selected_title.get().strip()):
                var_selected_title.set(detected_title)

        def sync_basis_from_title(commit=False):
            range_target_col, _ = extract_range_target_col(var_input.get())
            title = var_selected_title.get().strip()
            source_col, source_header, target_col, target_header = self.build_auto_basis_sync_result(
                title, range_target_col, source_header_map, target_header_map
            )
            var_source_basis.set(self.format_source_header_option(source_col, source_header_map) if source_col else NO_HEADER_MATCH_OPTION)
            var_target_basis.set(self.format_target_header_option(target_col, target_header_map) if target_col else NO_HEADER_MATCH_OPTION)
            updated_mappings = []
            if commit:
                updated_mappings = self.update_linked_b2b_mappings_for_basis(target_header, source_col, target_col)
            return source_col, source_header, target_col, target_header, updated_mappings

        def on_auto_sync_click():
            source_col, source_header, target_col, target_header, _ = sync_basis_from_title(commit=False)
            target_letter = col_num_to_letter(target_col) if target_col else '?'
            source_col_text = f'Col {source_col}' if source_col else NO_HEADER_MATCH_OPTION
            target_col_text = f'{target_letter}열 Col {target_col}' if target_col else NO_HEADER_MATCH_OPTION
            self.log('INFO', f'선택 열 제목 기준 자동 연계 준비: 소스 [{source_header}/{source_col_text}] ↔ 기준 [{target_header}/{target_col_text}]')
            on_parse_preview()

        btn_auto_sync = ttk.Button(f_title_input, text='🔗 제목 기준 자동 연계', command=on_auto_sync_click)
        btn_auto_sync.pack(side='right')
        HoverTooltip(btn_auto_sync, (
            "입력한 선택 열 제목으로 소스/기준 헤더 콤보박스를 자동 갱신합니다.\n\n"
            "예: 선택 열 제목을 PO번호로 입력하면 소스 PO번호 Col과 기준 PO번호 열을 찾아 연계 기준으로 설정합니다. 최종 매핑 목록 갱신은 [범위 확정] 때 저장됩니다."
        ))

        def on_parse_preview(*args):
            autofill_title_from_range(force=False)
            val = var_input.get().strip().upper().replace('$', '')
            m = re.findall(r'\d+', val)
            _, _, target_col, target_header = get_selected_basis()
            target_letter = col_num_to_letter(target_col) if target_col else '?'
            if len(m) >= 2:
                r1, r2 = int(m[0]), int(m[-1])
                s_r, e_r = min(r1, r2), max(r1, r2)
                cnt = e_r - s_r + 1
                lbl_preview.config(text=f'✔ 감지 결과: {target_letter}{s_r}:{target_letter}{e_r} ({s_r}행 ~ {e_r}행, 총 {cnt}개 행 지정)')
            elif len(m) == 1:
                r1 = int(m[0])
                lbl_preview.config(text=f'✔ 감지 결과: {target_letter}{r1} ({r1}행 단일 행 지정)')
            else:
                lbl_preview.config(text='⚠️ 올바른 행 범위(예: X4942:X4963 또는 4942:4963)를 입력하세요.')
            source_col, source_header, _, _ = get_selected_basis()
            source_col_text = f'Col {source_col}' if source_col else NO_HEADER_MATCH_OPTION
            target_col_text = f'{target_letter}열 Col {target_col}' if target_col else NO_HEADER_MATCH_OPTION
            lbl_basis_preview.config(
                text=f'연계 기준: 소스 [{source_header} / {source_col_text}] ↔ 기준 [{target_header} / {target_col_text}]'
            )

        var_input.trace_add('write', on_parse_preview)
        var_source_basis.trace_add('write', on_parse_preview)
        var_target_basis.trace_add('write', on_parse_preview)
        var_selected_title.trace_add('write', lambda *_args: (sync_basis_from_title(commit=False), on_parse_preview()))
        sync_basis_from_title(commit=False)
        on_parse_preview()

        def on_confirm():
            val = var_input.get().strip().upper().replace('$', '')
            m = re.findall(r'\d+', val)
            if not m:
                messagebox.showwarning('범위 확인', '올바른 범위(예: X4942:X4963)를 입력해 주세요.')
                return

            if len(m) >= 2:
                r1, r2 = int(m[0]), int(m[-1])
                start_row, end_row = min(r1, r2), max(r1, r2)
                item_count = end_row - start_row + 1
                range_address = f'X{start_row}:X{end_row}'
            else:
                start_row = int(m[0])
                end_row = start_row
                item_count = 1
                range_address = f'X{start_row}'

            source_col, source_header, target_col, target_header = get_selected_basis()
            if not source_col or is_no_header_match_option(source_header):
                messagebox.showwarning('기준 헤더 확인', '소스 파일 기준 헤더가 자동으로 매칭되지 않았습니다.\n[소스 파일 기준 헤더]에서 같은 의미의 열을 직접 선택한 뒤 다시 확정하세요.', parent=popup)
                return
            if not target_col or is_no_header_match_option(target_header):
                messagebox.showwarning('기준 헤더 확인', '기준 파일 기준 헤더가 자동으로 매칭되지 않았습니다.\n[기준 파일 기준 헤더]에서 같은 의미의 열을 직접 선택한 뒤 다시 확정하세요.', parent=popup)
                return
            auto_updated_mappings = self.update_linked_b2b_mappings_for_basis(target_header, source_col, target_col)
            target_letter = col_num_to_letter(target_col)
            if not source_header or source_header == f'Col {source_col}':
                messagebox.showwarning('기준 헤더 확인', f'소스 파일에서 선택한 Col {source_col}의 헤더명이 확인되지 않습니다.\n소스 파일 분석 결과를 확인하거나 다른 헤더를 선택하세요.', parent=popup)
                return
            if not target_header or target_header == f'Col {target_col}':
                messagebox.showwarning('기준 헤더 확인', f'기준 파일에서 선택한 Col {target_col}의 헤더명이 확인되지 않습니다.\n기준 파일 분석 결과를 확인하거나 다른 헤더를 선택하세요.', parent=popup)
                return

            if len(m) >= 2:
                range_address = f'{target_letter}{start_row}:{target_letter}{end_row}'
            else:
                range_address = f'{target_letter}{start_row}'

            sheet_name = '집계'
            self.source_match_col = source_col
            self.source_match_header = source_header
            self.target_match_col = target_col
            self.target_match_header = target_header
            self.target_range_info = {
                'sheet_name': sheet_name,
                'range_address': range_address,
                'start_row': start_row,
                'end_row': end_row,
                'col_idx': target_col,
                'item_count': item_count,
                'source_match_col': source_col,
                'source_match_header': source_header,
                'target_match_col': target_col,
                'target_match_header': target_header,
            }

            self.selected_sheet_name.set(sheet_name)
            self.selected_range_address.set(range_address)

            display_msg = f'시트: [{sheet_name}] | 범위: [{range_address}] | 기준: 소스[{source_header}/Col {source_col}] ↔ 발주[{target_header}/Col {target_col}]'
            self.lbl_target_range.config(text=display_msg, foreground='#2B6CB0', font=('Segoe UI', 9, 'bold'))
            linked_mappings = []
            source_norm = normalize_text(source_header)
            target_norm = normalize_text(target_header)
            for mapping in self.header_mappings:
                source_names = [mapping.get('source_name', '')] + mapping.get('source_aliases', [])
                target_names = [mapping.get('target_name', '')]
                source_linked = any(source_norm and source_norm == normalize_text(name) for name in source_names)
                target_linked = any(target_norm and target_norm == normalize_text(name) for name in target_names)
                if source_linked or target_linked or int(mapping.get('target_col', 0) or 0) == target_col:
                    linked_mappings.append(mapping.get('source_name', f'Col {mapping.get("target_col", "?")}'))
            if linked_mappings:
                self.log('INFO', f'동일 헤더/좌표 데이터 매핑 연계 확인: {", ".join(dict.fromkeys(linked_mappings))}')
            else:
                self.log('WARN', f'선택한 행 매칭 기준 [{source_header} ↔ {target_header}]과 동일한 데이터 쓰기 매핑은 없습니다. 행 매칭 기준으로만 적용됩니다.')
            if auto_updated_mappings:
                self.log('SUCCESS', f'선택 열 제목 [{target_header}] 기준으로 기존 데이터 매핑 {len(auto_updated_mappings)}건 자동 갱신 완료')
            else:
                self.log('WARN', f'선택 열 제목 [{target_header}]과 직접 연결된 기존 데이터 쓰기 매핑이 없어 행 매칭 기준만 자동 갱신했습니다.')
            self.log('SUCCESS', f'엑셀 범위 지정 확정 성공: {display_msg}')
            on_close_popup()

        def on_close_popup():
            self.range_popup = None
            self.release_interactive_excel_session('범위 확정/팝업 닫기')
            popup.destroy()

        popup.protocol("WM_DELETE_WINDOW", on_close_popup)

        btn_box = ttk.Frame(f)
        btn_box.pack(fill='x')

        btn_cancel = ttk.Button(btn_box, text='❌ 취소', command=on_close_popup)
        btn_cancel.pack(side='right', padx=(8, 0))
        btn_confirm = ttk.Button(btn_box, text='✅ 범위 확정', style='Primary.TButton', command=on_confirm)
        btn_confirm.pack(side='right')
        HoverTooltip(btn_confirm, (
            "현재 입력된 행 범위와 두 기준 헤더를 저장합니다.\n\n"
            "저장 후 확인할 내용:\n"
            "1. 메인 화면의 지정 범위 라벨에 소스 헤더 ↔ 기준 헤더가 표시됩니다.\n"
            "2. 실행 확인 창에도 동일 기준이 표시됩니다.\n"
            "3. 데이터 쓰기 열 변경이 필요하면 [헤더 매핑 설정]에서 별도 편집합니다."
        ))

    # =========================================================================
    # 수동 헤더 매핑 편집기 (신규 기능)
    # =========================================================================

    def show_header_mapping_editor(self):
        """소스↔타겟 헤더 매핑 편집 모달 (Canvas 연결선 시각화 + 타겟 헤더명 표시)"""
        popup = tk.Toplevel(self.root)
        popup.title('⚙️ 소스↔타겟 헤더 매핑 편집기 (자동 + 수동)')
        popup.geometry('1180x660')
        popup.grab_set()
        popup.attributes('-topmost', True)

        # 소스 컬럼 인덱스 해석용 맵 구성
        source_col_map = {}
        if self.source_info and 'header_names' in self.source_info:
            for idx_h, name in enumerate(self.source_info['header_names'], 1):
                source_col_map[name] = idx_h

        def resolve_src_col(mapping):
            for alias in mapping.get('source_aliases', [mapping['source_name']]):
                for k, v in source_col_map.items():
                    if alias in k or k in alias:
                        return v
            return mapping.get('default_source_col', '?')

        def resolve_tgt_name(mapping):
            col = mapping['target_col']
            if self.target_info and 'header_names' in self.target_info and col in self.target_info['header_names']:
                return self.target_info['header_names'][col]
            return mapping.get('target_name', mapping['source_name'])

        main_f = ttk.Frame(popup, padding=12)
        main_f.pack(fill='both', expand=True)

        lbl_editor_title = ttk.Label(main_f, text='⚙️ 소스(B2B) ↔ 타겟(발주내역) 컬럼 매핑 설정', font=('Segoe UI', 13, 'bold'), foreground=self.COLOR_PRIMARY)
        lbl_editor_title.pack(anchor='w', pady=(0, 3))
        HoverTooltip(lbl_editor_title, (
            "이 화면은 '행을 어떤 기준으로 찾을지'가 아니라 '매칭된 행에 어떤 값을 쓸지'를 설정합니다.\n\n"
            "워크플로:\n"
            "1. 범위 팝업에서 소스 기준 헤더와 발주내역 기준 헤더를 확정합니다.\n"
            "2. 이 화면에서 쓰기 대상 컬럼을 확인합니다.\n"
            "3. 기존 항목을 편집하거나 신규 매핑을 추가합니다.\n"
            "4. 저장 후 스마트 매핑을 실행합니다.\n\n"
            "예: PO번호를 기준으로 행을 찾고 PO번호도 갱신해야 하면, PO번호 매핑이 소스 PO번호 ↔ 발주내역 PO번호 Col로 연결되어 있는지 확인합니다."
        ))
        ttk.Label(main_f, text='아래 연결선으로 소스와 타겟의 헤더명/컬럼 연계 관계를 확인하세요. 클릭하여 선택 → 더블클릭 또는 버튼으로 편집/추가/삭제할 수 있습니다.', font=('Segoe UI', 9), foreground='#4A5568').pack(anchor='w', pady=(0, 8))

        # === Canvas 영역 ===
        canvas_frame = ttk.Frame(main_f)
        canvas_frame.pack(fill='both', expand=True, pady=(0, 8))

        canvas = tk.Canvas(canvas_frame, bg='#FFFFFF', highlightthickness=1, highlightbackground='#CBD5E0')
        v_scroll = ttk.Scrollbar(canvas_frame, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        v_scroll.pack(side='right', fill='y')

        # 레이아웃 상수 (타겟 헤더명 열 포함 1140px 레이아웃)
        ROW_H = 40
        HDR_Y = 22
        START_Y = 48
        SRC_COL_X = 18
        SRC_NAME_X = 85
        ARROW_X1 = 340
        ARROW_X2 = 500
        TGT_COL_X = 520
        TGT_LETTER_X = 590
        TGT_NAME_X = 670   # ★ 신규 타겟 헤더명 위치 ★
        MODE_X = 1000
        CANVAS_W = 1140

        selected_idx = [None]

        def refresh():
            canvas.delete('all')
            n = len(self.header_mappings)
            total_h = START_Y + n * ROW_H + 30
            canvas.configure(scrollregion=(0, 0, CANVAS_W, max(total_h, 420)))

            # ── 컬럼 헤더 (소스 / 타겟) ──
            # 소스 영역 배경
            canvas.create_rectangle(5, 3, ARROW_X1 - 15, HDR_Y + 12, fill='#F0FFF4', outline='#C6F6D5')
            canvas.create_text(SRC_COL_X, HDR_Y, text='소스 Col', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#276749')
            canvas.create_text(SRC_NAME_X, HDR_Y, text='소스 헤더명 (B2B 컬럼)', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#276749')

            # 연결 영역
            canvas.create_text((ARROW_X1 + ARROW_X2) // 2, HDR_Y, text='연결', anchor='center', font=('Segoe UI', 9, 'bold'), fill='#718096')

            # 타겟 영역 배경
            canvas.create_rectangle(ARROW_X2 + 5, 3, CANVAS_W - 5, HDR_Y + 12, fill='#EBF8FF', outline='#BEE3F8')
            canvas.create_text(TGT_COL_X, HDR_Y, text='타겟 Col', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#2C5282')
            canvas.create_text(TGT_LETTER_X, HDR_Y, text='타겟 열', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#2C5282')
            canvas.create_text(TGT_NAME_X, HDR_Y, text='타겟 헤더명 (발주내역)', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#2C5282')
            canvas.create_text(MODE_X, HDR_Y, text='모드', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#4A5568')

            # 헤더 구분선
            canvas.create_line(5, HDR_Y + 14, CANVAS_W - 5, HDR_Y + 14, fill='#A0AEC0', width=1)

            for idx, m in enumerate(self.header_mappings):
                y = START_Y + idx * ROW_H + ROW_H // 2
                y_top = START_Y + idx * ROW_H + 2
                y_bot = y_top + ROW_H - 4

                # 선택 강조 또는 교차 배경색
                if selected_idx[0] == idx:
                    canvas.create_rectangle(5, y_top, CANVAS_W - 5, y_bot, fill='#BEE3F8', outline='#2B6CB0', width=2)
                else:
                    bg = '#FFFFFF' if idx % 2 == 0 else '#F7FAFC'
                    canvas.create_rectangle(5, y_top, CANVAS_W - 5, y_bot, fill=bg, outline='#EDF2F7')

                # 모드별 색상
                if m['mode'] == '자동':
                    line_color, src_fg, tgt_fg = '#38A169', '#276749', '#2C5282'
                    mode_text = '🤖 자동'
                elif m['mode'] == '신규':
                    line_color, src_fg, tgt_fg = '#3182CE', '#2B6CB0', '#2B6CB0'
                    mode_text = '➕ 신규'
                else:
                    line_color, src_fg, tgt_fg = '#E53E3E', '#C53030', '#C53030'
                    mode_text = '✏️ 수동'

                src_col = resolve_src_col(m)
                src_col_text = f'Col {src_col:>2}' if isinstance(src_col, int) else f'Col {src_col}'
                tgt_letter = col_num_to_letter(m['target_col'])
                tgt_name = resolve_tgt_name(m)

                # ── 소스 측 (좌) ──
                canvas.create_text(SRC_COL_X, y, text=src_col_text, anchor='w', font=('Consolas', 10, 'bold'), fill=src_fg)
                canvas.create_text(SRC_NAME_X, y, text=m['source_name'], anchor='w', font=('Consolas', 10), fill=src_fg)

                # ── 연결선 (화살표) ──
                canvas.create_oval(ARROW_X1 - 5, y - 5, ARROW_X1 + 5, y + 5, fill=line_color, outline=line_color)
                canvas.create_line(ARROW_X1 + 6, y, ARROW_X2 - 6, y, fill=line_color, width=2.5, arrow=tk.LAST, arrowshape=(10, 14, 5))
                canvas.create_rectangle(ARROW_X2 - 5, y - 5, ARROW_X2 + 5, y + 5, fill=line_color, outline=line_color)

                # ── 타겟 측 (우) ──
                canvas.create_text(TGT_COL_X, y, text=f'Col {m["target_col"]}', anchor='w', font=('Consolas', 10, 'bold'), fill=tgt_fg)
                canvas.create_text(TGT_LETTER_X, y, text=f'{tgt_letter}열', anchor='w', font=('Consolas', 11, 'bold'), fill=tgt_fg)
                canvas.create_text(TGT_NAME_X, y, text=tgt_name, anchor='w', font=('Consolas', 10, 'bold'), fill='#1A365D') # ★ 실제 파싱된 타겟 헤더명 표시 ★

                # ── 모드 ──
                canvas.create_text(MODE_X, y, text=mode_text, anchor='w', font=('Segoe UI', 9), fill='#4A5568')

            # 하단 구분선
            if n > 0:
                y_end = START_Y + n * ROW_H + 5
                canvas.create_line(5, y_end, CANVAS_W - 5, y_end, fill='#CBD5E0', width=1)
                leg_y = y_end + 18
                canvas.create_oval(20, leg_y - 4, 28, leg_y + 4, fill='#38A169', outline='#38A169')
                canvas.create_text(35, leg_y, text='자동 매핑', anchor='w', font=('Segoe UI', 8), fill='#718096')
                canvas.create_oval(130, leg_y - 4, 138, leg_y + 4, fill='#E53E3E', outline='#E53E3E')
                canvas.create_text(145, leg_y, text='수동 변경', anchor='w', font=('Segoe UI', 8), fill='#718096')
                canvas.create_oval(240, leg_y - 4, 248, leg_y + 4, fill='#3182CE', outline='#3182CE')
                canvas.create_text(255, leg_y, text='신규 추가', anchor='w', font=('Segoe UI', 8), fill='#718096')

        def on_click(event):
            cy = canvas.canvasy(event.y)
            idx = int((cy - START_Y) / ROW_H)
            if 0 <= idx < len(self.header_mappings):
                selected_idx[0] = idx
            else:
                selected_idx[0] = None
            refresh()

        def on_double_click(event):
            cy = canvas.canvasy(event.y)
            idx = int((cy - START_Y) / ROW_H)
            if 0 <= idx < len(self.header_mappings):
                selected_idx[0] = idx
                self.show_mapping_edit_dialog(popup, idx, resolve_src_col, refresh)

        canvas.bind('<Button-1>', on_click)
        canvas.bind('<Double-1>', on_double_click)
        canvas.bind('<Configure>', lambda e: popup.after(30, refresh))

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind('<MouseWheel>', on_mousewheel)

        refresh()

        # ── 버튼 프레임 ──
        btn_frame = ttk.Frame(main_f)
        btn_frame.pack(fill='x', pady=(0, 8))

        def on_add():
            self.show_add_mapping_dialog(popup, resolve_src_col, refresh)

        def on_edit():
            if selected_idx[0] is None:
                messagebox.showwarning('선택 필요', '편집할 매핑 항목을 먼저 클릭하여 선택하세요.\n(파란색 강조 표시된 행)', parent=popup)
                return
            self.show_mapping_edit_dialog(popup, selected_idx[0], resolve_src_col, refresh)

        def on_delete():
            if selected_idx[0] is None:
                messagebox.showwarning('선택 필요', '삭제할 매핑 항목을 먼저 클릭하여 선택하세요.', parent=popup)
                return
            idx = selected_idx[0]
            mapping = self.header_mappings[idx]
            tgt_l = col_num_to_letter(mapping['target_col'])
            tgt_n = resolve_tgt_name(mapping)
            if messagebox.askyesno('삭제 확인', f'아래 매핑을 삭제하시겠습니까?\n\n[{mapping["source_name"]}] ───▶ [{tgt_n} ({tgt_l}열, Col {mapping["target_col"]})]', parent=popup):
                del self.header_mappings[idx]
                selected_idx[0] = None
                refresh()
                self.log('INFO', f'매핑 삭제됨: {mapping["source_name"]} → Col {mapping["target_col"]}')

        def on_reset():
            if messagebox.askyesno('기본값 복원', '모든 수동 편집을 초기화하고\n자동 기본 7쌍으로 복원하시겠습니까?', parent=popup):
                self.header_mappings = self.get_default_mappings()
                selected_idx[0] = None
                refresh()
                self.log('INFO', '헤더 매핑이 자동 기본값(7쌍)으로 복원되었습니다.')

        ttk.Button(btn_frame, text='➕ 신규 매핑 추가', style='Primary.TButton', command=on_add).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text='✏️ 선택 항목 편집', command=on_edit).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text='🗑️ 선택 삭제', style='Reset.TButton', command=on_delete).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text='🔄 자동 기본값 복원', command=on_reset).pack(side='left')

        # 하단 확정/취소
        bot_frame = ttk.Frame(main_f)
        bot_frame.pack(fill='x')

        def on_save():
            count = len(self.header_mappings)
            manual_count = sum(1 for m in self.header_mappings if m['mode'] != '자동')
            if manual_count > 0:
                mode_text = f'🔧 수동 편집됨 ({count}쌍, 수동 {manual_count}건)'
            else:
                mode_text = f'📋 자동 ({count}쌍 기본 연결)'
            self.lbl_mapping_mode.config(text=f'매핑 모드: {mode_text}')
            self.log('SUCCESS', f'헤더 매핑 설정 저장 완료: {count}쌍 (자동 {count - manual_count}건 + 수동 {manual_count}건)')
            popup.destroy()

        ttk.Button(bot_frame, text='❌ 취소', command=popup.destroy).pack(side='right', padx=(5, 0))
        ttk.Button(bot_frame, text='✅ 확정 저장', style='Primary.TButton', command=on_save).pack(side='right')

    def show_mapping_edit_dialog(self, parent, mapping_idx, resolve_src_col_fn, refresh_callback):
        """개별 매핑 편집 다이얼로그 (소스/타겟 컬럼 콤보박스 + 소스/타겟 헤더명 대칭적 구성)"""
        mapping = self.header_mappings[mapping_idx]
        src_col_resolved = resolve_src_col_fn(mapping)

        popup = tk.Toplevel(parent)
        popup.title(f'✏️ 매핑 편집: {mapping["source_name"]} → {col_num_to_letter(mapping["target_col"])}열')
        popup.geometry('640x480')
        popup.grab_set()
        popup.attributes('-topmost', True)

        f = ttk.Frame(popup, padding=20)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text='✏️ 소스↔타겟 매핑 편집', font=('Segoe UI', 12, 'bold'), foreground=self.COLOR_PRIMARY).pack(anchor='w', pady=(0, 5))

        cur_tgt_l = col_num_to_letter(mapping['target_col'])
        cur_tgt_n = (self.target_info['header_names'].get(mapping['target_col']) if (self.target_info and 'header_names' in self.target_info and mapping['target_col'] in self.target_info['header_names']) else mapping.get('target_name', '낙찰금액' if mapping['target_col'] == 25 else mapping['source_name']))
        cur_info = f'현재: [Col {src_col_resolved}: {mapping["source_name"]}] ───▶ [{cur_tgt_n} ({cur_tgt_l}열, Col {mapping["target_col"]})]'
        ttk.Label(f, text=cur_info, font=('Consolas', 9), foreground='#718096').pack(anchor='w', pady=(0, 12))

        # ── 1. 소스 옵션 구성 ──
        source_header_map = {1: '생성일자', 2: '공사/입찰/계약번호', 4: '계약금액', 5: '계약일자(체결일)', 6: 'PO전송견적금액', 7: 'PO번호', 8: 'PR번호'}
        if self.source_info and 'header_names' in self.source_info:
            for c_idx, h_name in enumerate(self.source_info['header_names'], 1):
                source_header_map[c_idx] = h_name

        src_col_options = []
        max_source_col = max(50, max(source_header_map.keys()) if source_header_map else 50)
        for i in range(1, max_source_col + 1):
            hname = source_header_map.get(i, '')
            if hname:
                src_col_options.append(f'Col {i:02d} - {hname}')
            else:
                src_col_options.append(f'Col {i:02d}')

        # ── 2. 타겟 옵션 구성 ──
        target_header_map = {
            24: '계약명', 25: '낙찰금액', 26: 'PO금액', 27: 'PO번호', 28: 'PR번호',
            29: 'PR금액', 30: 'PR 발생일', 31: '공사/ 유지보수 번호', 32: '공사 요청일',
            33: '공사 완료일', 34: '계약 체결일'
        }
        if self.target_info and 'header_names' in self.target_info:
            for c_idx, h_name in self.target_info['header_names'].items():
                target_header_map[c_idx] = h_name
        for m in self.header_mappings:
            if m['target_col'] not in target_header_map or m.get('mode') != '자동':
                target_header_map[m['target_col']] = m.get('target_name', m['source_name'])

        tgt_col_options = []
        max_target_col = max(52, max(target_header_map.keys()) if target_header_map else 52)
        for i in range(1, max_target_col + 1):
            letter = col_num_to_letter(i)
            hname = target_header_map.get(i, '')
            if hname:
                tgt_col_options.append(f'{letter} (Col {i}) - {hname}')
            else:
                tgt_col_options.append(f'{letter} (Col {i})')

        # ─────────────────────────────────────────────────────────────
        # [소스 (B2B) 영역] - 타겟 영역과 100% 동일한 대칭 구조
        # ─────────────────────────────────────────────────────────────
        # 1-A. 소스 컬럼 선택 (Combobox)
        f_src_col = ttk.Frame(f)
        f_src_col.pack(fill='x', pady=(0, 6))
        ttk.Label(f_src_col, text='📌 소스 컬럼 (B2B):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')

        init_src_val = f'Col {src_col_resolved:02d}' if isinstance(src_col_resolved, int) else 'Col 01'
        for opt in src_col_options:
            if isinstance(src_col_resolved, int) and f'Col {src_col_resolved:02d}' in opt:
                init_src_val = opt
                break
            elif mapping['source_name'] in opt:
                init_src_val = opt
                break

        var_source_col = tk.StringVar(value=init_src_val)
        cmb_source_col = ttk.Combobox(f_src_col, textvariable=var_source_col, values=src_col_options, font=('Consolas', 10), width=34)
        cmb_source_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_source_col, (
            "소스 파일에서 가져올 값의 헤더와 Col 좌표입니다.\n\n"
            "선택 즉시 아래 소스 헤더명이 자동 갱신됩니다.\n"
            "예: PO번호 값을 발주내역에 쓰려면 소스 PO번호가 표시된 Col을 선택합니다."
        ))

        # 1-B. 소스 헤더명 (B2B) 입력 필드 (Entry)
        f_src_name = ttk.Frame(f)
        f_src_name.pack(fill='x', pady=(0, 10))
        ttk.Label(f_src_name, text='🏷️ 소스 헤더명 (B2B):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        var_source_name = tk.StringVar(value=mapping['source_name'])
        ent_source_name = ttk.Entry(f_src_name, textvariable=var_source_name, font=('Consolas', 10), width=34)
        ent_source_name.pack(side='left', fill='x', expand=True)

        # 소스 컬럼 콤보박스 선택 시 소스 헤더명 자동 연동
        def on_source_col_change(*args):
            src_text = var_source_col.get().strip()
            if ' - ' in src_text:
                auto_hname = src_text.split(' - ')[-1].strip()
                var_source_name.set(auto_hname)

        var_source_col.trace_add('write', on_source_col_change)

        # ─────────────────────────────────────────────────────────────
        # [타겟 (발주내역) 영역] - 대칭 구조
        # ─────────────────────────────────────────────────────────────
        # 2-A. 타겟 컬럼 선택 (Combobox)
        f_tgt_col = ttk.Frame(f)
        f_tgt_col.pack(fill='x', pady=(0, 6))
        ttk.Label(f_tgt_col, text='🎯 타겟 컬럼 (발주내역):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')

        init_tgt_val = f'{col_num_to_letter(mapping["target_col"])} (Col {mapping["target_col"]})'
        for opt in tgt_col_options:
            if f'(Col {mapping["target_col"]})' in opt:
                init_tgt_val = opt
                break

        var_target_col = tk.StringVar(value=init_tgt_val)
        cmb_target_col = ttk.Combobox(f_tgt_col, textvariable=var_target_col, values=tgt_col_options, font=('Consolas', 10), width=34)
        cmb_target_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_target_col, (
            "발주내역에서 값을 쓸 대상 헤더와 Col 좌표입니다.\n\n"
            "행 매칭 기준을 PO번호로 바꾸더라도, 쓰기 대상 컬럼은 이 콤보박스에서 별도로 관리합니다.\n"
            "동일한 의미의 헤더가 이미 있으면 기존 매핑을 편집하고, 없으면 신규 매핑으로 추가하세요."
        ))

        # 2-B. 타겟 헤더명 입력 필드 (Entry)
        f_tgt_name = ttk.Frame(f)
        f_tgt_name.pack(fill='x', pady=(0, 12))
        ttk.Label(f_tgt_name, text='🏷️ 타겟 헤더명 (발주내역):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        var_target_name = tk.StringVar(value=cur_tgt_n)
        ent_target_name = ttk.Entry(f_tgt_name, textvariable=var_target_name, font=('Consolas', 10), width=34)
        ent_target_name.pack(side='left', fill='x', expand=True)

        # 타겟 컬럼 콤보박스 선택 시 타겟 헤더명 자동 연동
        def on_target_col_change(*args):
            tgt_text = var_target_col.get().strip()
            if ' - ' in tgt_text:
                auto_hname = tgt_text.split(' - ')[-1].strip()
                var_target_name.set(auto_hname)

        var_target_col.trace_add('write', on_target_col_change)

        # 변경 결과 미리보기
        lbl_preview = ttk.Label(f, text='', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0')
        lbl_preview.pack(anchor='w', pady=(0, 12))

        def update_preview(*args):
            src_n = var_source_name.get().strip()
            src_c_text = var_source_col.get().strip()
            tgt_n = var_target_name.get().strip()
            tgt_c_text = var_target_col.get().strip()

            m_sc = re.search(r'Col\s*\d+', src_c_text)
            sc_str = m_sc.group(0) if m_sc else src_c_text
            m_tc = re.search(r'^[A-Z]+\s*\(Col\s*\d+\)', tgt_c_text)
            tc_str = m_tc.group(0) if m_tc else tgt_c_text

            lbl_preview.config(text=f'변경 결과: [{sc_str}: {src_n}] ───▶ [{tgt_n} ({tc_str})]')

        var_source_col.trace_add('write', update_preview)
        var_source_name.trace_add('write', update_preview)
        var_target_col.trace_add('write', update_preview)
        var_target_name.trace_add('write', update_preview)
        update_preview()

        btn_f = ttk.Frame(f)
        btn_f.pack(fill='x')

        def on_save():
            new_source = var_source_name.get().strip()
            src_c_text = var_source_col.get().strip()
            tgt_c_text = var_target_col.get().strip()
            new_target_name = var_target_name.get().strip()

            if not new_source:
                messagebox.showwarning('입력 확인', '소스 헤더명을 입력하거나 선택하세요.', parent=popup)
                return
            m_sc_nums = re.findall(r'\d+', src_c_text)
            default_sc = int(m_sc_nums[0]) if m_sc_nums else 1

            m_tc_nums = re.findall(r'\d+', tgt_c_text)
            if not m_tc_nums:
                messagebox.showwarning('입력 확인', '올바른 타겟 컬럼을 선택하세요.', parent=popup)
                return
            new_target_col = int(m_tc_nums[0])
            if not new_target_name:
                new_target_name = new_source

            self.header_mappings[mapping_idx] = {
                'source_name': new_source,
                'source_aliases': [new_source],
                'target_col': new_target_col,
                'target_name': new_target_name,
                'default_source_col': default_sc,
                'mode': '수동'
            }
            refresh_callback()
            self.log('INFO', f'매핑 편집 완료: [Col {default_sc}: {new_source}] → [{new_target_name} ({col_num_to_letter(new_target_col)}열, Col {new_target_col})]')
            popup.destroy()

        ttk.Button(btn_f, text='❌ 취소', command=popup.destroy).pack(side='right', padx=(5, 0))
        btn_save = ttk.Button(btn_f, text='✅ 저장', style='Primary.TButton', command=on_save)
        btn_save.pack(side='right')
        HoverTooltip(btn_save, (
            "선택한 소스 헤더/타겟 헤더/Col 좌표를 현재 매핑 항목에 저장합니다.\n\n"
            "저장 후 매핑 편집기 목록의 연결선과 실행 확인 창에 반영됩니다."
        ))

    def show_add_mapping_dialog(self, parent, resolve_src_col_fn, refresh_callback):
        """신규 매핑 추가 다이얼로그 (소스/타겟 컬럼 콤보박스 + 소스/타겟 헤더명 대칭적 구성)"""
        popup = tk.Toplevel(parent)
        popup.title('➕ 신규 매핑 추가')
        popup.geometry('640x480')
        popup.grab_set()
        popup.attributes('-topmost', True)

        f = ttk.Frame(popup, padding=20)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text='➕ 신규 소스↔타겟 매핑 추가', font=('Segoe UI', 12, 'bold'), foreground=self.COLOR_PRIMARY).pack(anchor='w', pady=(0, 12))

        # ── 1. 소스 옵션 구성 ──
        source_header_map = {1: '생성일자', 2: '공사/입찰/계약번호', 4: '계약금액', 5: '계약일자(체결일)', 6: 'PO전송견적금액', 7: 'PO번호', 8: 'PR번호'}
        if self.source_info and 'header_names' in self.source_info:
            for c_idx, h_name in enumerate(self.source_info['header_names'], 1):
                source_header_map[c_idx] = h_name

        src_col_options = []
        max_source_col = max(50, max(source_header_map.keys()) if source_header_map else 50)
        for i in range(1, max_source_col + 1):
            hname = source_header_map.get(i, '')
            if hname:
                src_col_options.append(f'Col {i:02d} - {hname}')
            else:
                src_col_options.append(f'Col {i:02d}')

        # ── 2. 타겟 옵션 구성 ──
        target_header_map = {
            24: '계약명', 25: '낙찰금액', 26: 'PO금액', 27: 'PO번호', 28: 'PR번호',
            29: 'PR금액', 30: 'PR 발생일', 31: '공사/ 유지보수 번호', 32: '공사 요청일',
            33: '공사 완료일', 34: '계약 체결일'
        }
        if self.target_info and 'header_names' in self.target_info:
            for c_idx, h_name in self.target_info['header_names'].items():
                target_header_map[c_idx] = h_name
        for m in self.header_mappings:
            if m['target_col'] not in target_header_map or m.get('mode') != '자동':
                target_header_map[m['target_col']] = m.get('target_name', m['source_name'])

        tgt_col_options = []
        max_target_col = max(52, max(target_header_map.keys()) if target_header_map else 52)
        for i in range(1, max_target_col + 1):
            letter = col_num_to_letter(i)
            hname = target_header_map.get(i, '')
            if hname:
                tgt_col_options.append(f'{letter} (Col {i}) - {hname}')
            else:
                tgt_col_options.append(f'{letter} (Col {i})')

        # ─────────────────────────────────────────────────────────────
        # [소스 (B2B) 영역]
        # ─────────────────────────────────────────────────────────────
        f_src_col = ttk.Frame(f)
        f_src_col.pack(fill='x', pady=(0, 6))
        ttk.Label(f_src_col, text='📌 소스 컬럼 (B2B):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        var_source_col = tk.StringVar()
        cmb_source_col = ttk.Combobox(f_src_col, textvariable=var_source_col, values=src_col_options, font=('Consolas', 10), width=34)
        cmb_source_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_source_col, "신규 매핑으로 추가할 소스 파일 헤더와 Col 좌표를 선택합니다.")

        f_src_name = ttk.Frame(f)
        f_src_name.pack(fill='x', pady=(0, 10))
        ttk.Label(f_src_name, text='🏷️ 소스 헤더명 (B2B):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        var_source_name = tk.StringVar()
        ent_source_name = ttk.Entry(f_src_name, textvariable=var_source_name, font=('Consolas', 10), width=34)
        ent_source_name.pack(side='left', fill='x', expand=True)

        def on_source_col_change(*args):
            src_text = var_source_col.get().strip()
            if ' - ' in src_text:
                auto_hname = src_text.split(' - ')[-1].strip()
                var_source_name.set(auto_hname)

        var_source_col.trace_add('write', on_source_col_change)

        # ─────────────────────────────────────────────────────────────
        # [타겟 (발주내역) 영역]
        # ─────────────────────────────────────────────────────────────
        f_tgt_col = ttk.Frame(f)
        f_tgt_col.pack(fill='x', pady=(0, 6))
        ttk.Label(f_tgt_col, text='🎯 타겟 컬럼 (발주내역):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        var_target_col = tk.StringVar()
        cmb_target_col = ttk.Combobox(f_tgt_col, textvariable=var_target_col, values=tgt_col_options, font=('Consolas', 10), width=34)
        cmb_target_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_target_col, "신규 매핑으로 값을 쓸 발주내역 헤더와 Col 좌표를 선택합니다.")

        f_tgt_name = ttk.Frame(f)
        f_tgt_name.pack(fill='x', pady=(0, 12))
        ttk.Label(f_tgt_name, text='🏷️ 타겟 헤더명 (발주내역):', font=('Segoe UI', 10, 'bold'), width=22, anchor='w').pack(side='left')
        var_target_name = tk.StringVar()
        ent_target_name = ttk.Entry(f_tgt_name, textvariable=var_target_name, font=('Consolas', 10), width=34)
        ent_target_name.pack(side='left', fill='x', expand=True)

        def on_target_col_change(*args):
            tgt_text = var_target_col.get().strip()
            if ' - ' in tgt_text:
                auto_hname = tgt_text.split(' - ')[-1].strip()
                var_target_name.set(auto_hname)

        var_target_col.trace_add('write', on_target_col_change)

        # 미리보기
        lbl_preview = ttk.Label(f, text='소스 컬럼과 타겟 컬럼을 선택하세요.', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0')
        lbl_preview.pack(anchor='w', pady=(0, 12))

        def update_preview(*args):
            src_n = var_source_name.get().strip()
            src_c_text = var_source_col.get().strip()
            tgt_n = var_target_name.get().strip()
            tgt_c_text = var_target_col.get().strip()

            if src_n and tgt_c_text:
                m_sc = re.search(r'Col\s*\d+', src_c_text)
                sc_str = m_sc.group(0) if m_sc else src_c_text
                m_tc = re.search(r'^[A-Z]+\s*\(Col\s*\d+\)', tgt_c_text)
                tc_str = m_tc.group(0) if m_tc else tgt_c_text
                lbl_preview.config(text=f'신규 매핑: [{sc_str}: {src_n}] ───▶ [{tgt_n} ({tc_str})]')
            else:
                lbl_preview.config(text='소스 컬럼과 타겟 컬럼을 선택하세요.')

        var_source_col.trace_add('write', update_preview)
        var_source_name.trace_add('write', update_preview)
        var_target_col.trace_add('write', update_preview)
        var_target_name.trace_add('write', update_preview)

        btn_f = ttk.Frame(f)
        btn_f.pack(fill='x')

        def on_add():
            new_source = var_source_name.get().strip()
            src_c_text = var_source_col.get().strip()
            tgt_c_text = var_target_col.get().strip()
            new_target_name = var_target_name.get().strip()

            if not new_source:
                messagebox.showwarning('입력 확인', '소스 헤더명을 입력하거나 선택하세요.', parent=popup)
                return
            m_sc_nums = re.findall(r'\d+', src_c_text)
            default_sc = int(m_sc_nums[0]) if m_sc_nums else 1

            m_tc_nums = re.findall(r'\d+', tgt_c_text)
            if not m_tc_nums:
                messagebox.showwarning('입력 확인', '올바른 타겟 컬럼을 선택하세요.', parent=popup)
                return
            new_target_col = int(m_tc_nums[0])
            if not new_target_name:
                new_target_name = new_source

            for existing in self.header_mappings:
                if existing['target_col'] == new_target_col:
                    if not messagebox.askyesno('중복 경고', f'타겟 컬럼 {col_num_to_letter(new_target_col)} (Col {new_target_col})은\n이미 [{existing["source_name"]}]에 매핑되어 있습니다.\n그래도 추가하시겠습니까?', parent=popup):
                        return

            self.header_mappings.append({
                'source_name': new_source,
                'source_aliases': [new_source],
                'target_col': new_target_col,
                'target_name': new_target_name,
                'default_source_col': default_sc,
                'mode': '신규'
            })
            refresh_callback()
            self.log('INFO', f'신규 매핑 추가됨: [Col {default_sc}: {new_source}] → [{new_target_name} ({col_num_to_letter(new_target_col)}열, Col {new_target_col})]')
            popup.destroy()

        ttk.Button(btn_f, text='❌ 취소', command=popup.destroy).pack(side='right', padx=(5, 0))
        btn_add = ttk.Button(btn_f, text='➕ 추가', style='Primary.TButton', command=on_add)
        btn_add.pack(side='right')
        HoverTooltip(btn_add, (
            "신규 소스↔발주내역 쓰기 매핑을 추가합니다.\n\n"
            "이미 같은 타겟 Col이 있으면 중복 경고가 표시됩니다. 같은 의미의 헤더라면 기존 항목 편집을 우선 사용하세요."
        ))

    # =========================================================================
    # 매핑 실행 (범위 정밀 타격 + 동적 헤더 매핑)
    # =========================================================================

    def on_run_mapping(self):
        if self.active_tasks['source'] or self.active_tasks['target']:
            self.queued_action = self.on_run_mapping
            self.log('INFO', '⏳ 백그라운드 분석이 진행 중입니다. 완료되는 즉시 [스마트 매핑]을 자동으로 실행합니다.')
            return

        source_path = self.source_file_path.get()
        target_path = self.target_file_path.get()

        if not source_path or not os.path.exists(source_path):
            messagebox.showwarning('입력 확인', '1단계: 올바른 소스 파일(.xlsx/.xlsb/.xls)을 선택하세요.')
            return

        if not target_path or not os.path.exists(target_path):
            messagebox.showwarning('입력 확인', '2단계: 올바른 타겟 파일(.xlsx/.xlsb/.xls)을 선택하세요.')
            return

        if not self.target_range_info:
            messagebox.showwarning('입력 확인', '2단계: "엑셀 창 열기 & 마우스 범위 확정" 버튼을 통해 매핑 범위를 선택하세요.')
            return

        if not self.header_mappings:
            messagebox.showwarning('입력 확인', '헤더 매핑이 비어있습니다. "⚙️ 헤더 매핑 설정" 버튼에서 매핑을 설정하세요.')
            return

        mapping_desc = '\n'.join([f'  • {m["source_name"]} → {col_num_to_letter(m["target_col"])}열 (Col {m["target_col"]}) [{m["mode"]}]' for m in self.header_mappings])
        source_basis = self.target_range_info.get('source_match_header', '공사/입찰명')
        source_basis_col = self.target_range_info.get('source_match_col', 3)
        target_basis = self.target_range_info.get('target_match_header', self.target_range_info.get('basis_header', '계약명'))
        target_basis_col = self.target_range_info.get('col_idx', 24)
        confirm_msg = (
            f'매핑을 실행하시겠습니까?\n\n'
            f'• 소스 파일: {os.path.basename(source_path)}\n'
            f'• 타겟 파일: {os.path.basename(target_path)}\n'
            f'• 타겟 시트: {self.target_range_info["sheet_name"]}\n'
            f'• 선택 범위: {self.target_range_info["range_address"]} ({self.target_range_info["start_row"]}~{self.target_range_info["end_row"]}행, 총 {self.target_range_info["item_count"]}개 행)\n'
            f'• 행 매칭 기준: 소스 [{source_basis} / Col {source_basis_col}] ↔ 기준 [{target_basis} / Col {target_basis_col}]\n'
            f'• 헤더 매핑 ({len(self.header_mappings)}쌍):\n{mapping_desc}\n\n'
            f'⚠️ 정밀 타격 모드: 위 범위 외 셀은 절대 변경되지 않습니다.\n'
            f'※ 작업 실행 전 자동 백업본(.bak)이 생성됩니다.'
        )
        if not messagebox.askyesno('매핑 실행 확인', confirm_msg):
            return

        self.btn_run.config(state='disabled')
        self.progress['value'] = 0
        self.update_mapping_status(0, '⏳ 매핑 프로세스를 시작합니다...')

        threading.Thread(target=self.execute_mapping_process, daemon=True).start()

    def execute_mapping_process(self):
        ensure_excel_com_available()
        pythoncom.CoInitialize()
        source_path = os.path.abspath(self.source_file_path.get())
        target_path = os.path.abspath(self.target_file_path.get())
        range_info = self.target_range_info
        mappings = list(self.header_mappings)

        excel = None
        excel_state = None
        wb_source = None
        wb_target = None
        target_opened_by_script = False
        excel_created_by_script = False

        try:
            # 1단계 백업
            self.update_mapping_status(5, '💾 [1/6단계] 안전 백업 파일(.bak)을 생성하고 있습니다...')
            backup_path = create_timestamped_backup(target_path)
            self.log('SUCCESS', f'안전 백업 파일 생성 완료: {os.path.basename(backup_path)}')

            # 2단계 엑셀 연결
            self.update_mapping_status(10, '📂 [2/6단계] Excel COM 엔진 연결 및 B2B 소스 파일 로드 중...')
            def is_usable_excel_app(app):
                if app is None:
                    return False
                try:
                    _ = int(app.Hwnd)
                    _ = app.Workbooks.Count
                    return True
                except Exception:
                    return False

            try:
                opened_target = win32com.client.GetObject(target_path)
                if is_usable_excel_app(opened_target.Application):
                    wb_target = opened_target
                    excel = opened_target.Application
            except Exception:
                pass

            if excel is None:
                try:
                    candidate = win32com.client.GetActiveObject('Excel.Application')
                    if is_usable_excel_app(candidate):
                        excel = candidate
                except Exception:
                    pass

            if excel is None:
                excel = win32com.client.DispatchEx('Excel.Application')
                excel_created_by_script = True

            try:
                excel.Visible = not excel_created_by_script
            except Exception:
                pass
            excel_state = set_excel_fast_mode(excel)

            self.log('INFO', 'B2BIEAMS 소스 데이터를 메모리에 로드 중...')
            wb_source = excel.Workbooks.Open(source_path, ReadOnly=True)
            ws_source = wb_source.Worksheets(1)
            
            source_used = ws_source.UsedRange
            src_rows = source_used.Rows.Count
            src_cols = source_used.Columns.Count

            header_offset = self.source_info['header_rows'] if (self.source_info and 'header_rows' in self.source_info) else 1
            
            # 소스 헤더 → 컬럼 인덱스 매핑 구성
            col_map = {}
            for r_h in range(1, header_offset + 1):
                for c in range(1, src_cols + 1):
                    val = safe_cell_value(ws_source, r_h, c)
                    if val:
                        s_val = str(val).strip()
                        if s_val not in col_map:
                            col_map[s_val] = c

            def find_col_idx(possible_names, default_fallback):
                for name in possible_names:
                    for k, idx in col_map.items():
                        if name in k or k in name:
                            return idx
                return default_fallback

            source_match_col = int(range_info.get('source_match_col') or 0)
            source_match_header = range_info.get('source_match_header', '')
            if not source_match_col:
                source_match_col = find_col_idx(['공사/입찰명', '공사명', '계약명', '입찰명', '공사/입찰'], 3)
                source_match_header = source_match_header or '공사/입찰명'

            target_match_col = int(range_info.get('col_idx') or range_info.get('target_match_col') or 24)
            target_match_header = range_info.get('target_match_header') or range_info.get('basis_header') or '계약명'
            identifier_match_mode = self.is_identifier_match_basis(source_match_header, target_match_header)

            # 행 매칭 기준 컬럼: 기본 계약명 또는 사용자가 선택한 PO/PR/계약번호 등
            idx_contract_name = source_match_col

            # 동적 매핑: 각 헤더 매핑의 소스 컬럼 인덱스 해석
            resolved_mappings = []
            for m in mappings:
                src_col = find_col_idx(m.get('source_aliases', [m['source_name']]), m.get('default_source_col', 1))
                resolved_mappings.append({
                    'source_name': m['source_name'],
                    'source_col': src_col,
                    'target_col': m['target_col'],
                    'mode': m['mode'],
                })

            mapping_summary = ', '.join([f'{rm["source_name"]}(Col{rm["source_col"]})→{col_num_to_letter(rm["target_col"])}' for rm in resolved_mappings])
            self.log('INFO', f'동적 헤더 매핑 해석 완료 ({len(resolved_mappings)}쌍): {mapping_summary}')

            # 3단계 소스 레코드 파싱 (동적 매핑 기반)
            self.update_mapping_status(25, f'🔍 [3/6단계] B2B 소스 레코드 {src_rows:,}개 행 정밀 파싱 중...')
            source_start_row = header_offset + 1
            source_needed_cols = [idx_contract_name] + [rm['source_col'] for rm in resolved_mappings]
            source_col_data = read_excel_columns(ws_source, source_start_row, src_rows, source_needed_cols)
            source_records = []
            for row_offset, r in enumerate(range(source_start_row, src_rows + 1)):
                if r % 100 == 0 or r == src_rows:
                    prog_p = 25 + int(((r - header_offset) / max(1, src_rows - header_offset)) * 20)
                    self.update_mapping_status(prog_p, f'🔍 [3/6단계] B2B 소스 데이터 파싱 중... ({r:,} / {src_rows:,}행 읽음, 유효 {len(source_records):,}건)')

                c_name_val = source_col_data.get(idx_contract_name, [None])[row_offset]
                c_name = str(c_name_val).strip() if c_name_val else ''
                
                if not c_name and not identifier_match_mode:
                    for alt_c in range(1, src_cols + 1):
                        alt_val = safe_cell_value(ws_source, r, alt_c)
                        if alt_val and any('\uac00' <= ch <= '\ud7a3' for ch in str(alt_val)):
                            sv = str(alt_val).strip()
                            if not sv.replace(',', '').replace('.', '').isdigit():
                                c_name = sv
                                break

                if not c_name: continue

                room_key = extract_room_key(c_name)
                floors = extract_floors(c_name)
                discipline = extract_discipline(c_name)
                round_num = extract_round_num(c_name)
                norm_name = normalize_text(c_name)

                # 동적으로 각 매핑 컬럼의 데이터 읽기
                record = {
                    'row': r,
                    'name': c_name,
                    'room_key': room_key,
                    'floors': floors,
                    'discipline': discipline,
                    'round': round_num,
                    'norm': norm_name,
                }
                for rm in resolved_mappings:
                    record[f'col_{rm["target_col"]}'] = source_col_data.get(rm['source_col'], [None])[row_offset]

                source_records.append(record)

            self.log('INFO', f'소스 총 {len(source_records):,}건 유효 데이터 레코드 파싱 완료.')
            wb_source.Close(False)
            wb_source = None

            # 4단계 타겟 워크북 로드
            self.update_mapping_status(48, f'🎯 [4/6단계] 타겟 발주내역 시트[{range_info["sheet_name"]}] 로딩 중...')
            if not wb_target:
                for open_wb in excel.Workbooks:
                    if open_wb.FullName.lower() == target_path.lower():
                        wb_target = open_wb
                        break
            if not wb_target:
                wb_target = excel.Workbooks.Open(target_path)
                target_opened_by_script = True

            ws_target = wb_target.Worksheets(range_info['sheet_name'])
            
            target_start = range_info['start_row']
            target_end = range_info['end_row']
            target_col = target_match_col
            target_letter = col_num_to_letter(target_col)

            total_target_count = target_end - target_start + 1
            matched_count = 0
            highlighted_cells = 0

            audit_item_list = []
            used_source_rows = set()  # ★ 1대1 매핑 전용 트래킹 (다중 매핑 완전 차단) ★

            # ===== 범위 정밀 타격 감사 로그 =====
            self.log('INFO', f'🎯 정밀 타격 범위 확정: [{range_info["sheet_name"]}]!{range_info["range_address"]} ({target_start}행~{target_end}행)')
            self.log('INFO', f'🔒 범위 외 셀 접근 절대 차단 (행 경계: {target_start} ≤ row ≤ {target_end})')
            self.log('INFO', f'📌 행 매칭 기준: 소스 [{source_match_header} / Col {idx_contract_name}] ↔ 기준 [{target_match_header} / {target_letter}열 Col {target_col}]')
            self.log('INFO', f'📌 타겟 기준 읽기: {target_letter}열(Col {target_col}) 전용 — 다른 열 fallback 스캔 비활성화')
            self.log('INFO', f'🔒 1대1 매핑 엄격 적용 (동일 소스 레코드 중복 매핑 절대 금지)')

            # 5단계 스마트 매핑 연산
            target_col_values = read_excel_columns(ws_target, target_start, target_end, [target_col]).get(target_col, [])
            source_by_room = {}
            for source in source_records:
                if source['room_key']:
                    source_by_room.setdefault(source['room_key'], []).append(source)

            for idx, r in enumerate(range(target_start, target_end + 1)):
                seq_no = idx + 1
                prog_curr = 50 + int(((idx + 1) / total_target_count) * 42)
                self.update_mapping_status(prog_curr, f'⚡ [5/6단계] 스마트 매핑 진행 중: {idx + 1} / {total_target_count}건 진행 완료 (Row {r:4d} 연산 중...)')

                # === 범위 정밀 타격: 선택된 기준 열에서만 매칭 값 읽기 ===
                target_str = ''
                cand_val = target_col_values[idx] if idx < len(target_col_values) else None
                if cand_val is not None:
                    cand_text = str(cand_val).strip()
                    is_numeric_only = cand_text.replace(',', '').replace('.', '').isdigit()
                    if identifier_match_mode:
                        target_str = cand_text
                    elif cand_text and not is_numeric_only:
                        target_str = cand_text

                # 선택 기준 열이 비었으면 해당 행은 FAIL — 다른 열 스캔 절대 금지
                if not target_str:
                    audit_item_list.append({
                        'no': seq_no,
                        'row': r,
                        'target_str': f'({target_letter}열 빈 셀)',
                        'status': 'FAIL',
                        'source_name': '-',
                        'reason': f'⚠️ 지정 {target_letter}열(Col {target_col})에 {target_match_header} 값 없음'
                    })
                    self.log('WARN', f'No.{seq_no:02d} Row {r:4d} 매칭 건너뜀 ({target_letter}열 {target_match_header} 값 미발견 — 범위 외 스캔 차단)')
                    continue

                t_room_key = extract_room_key(target_str)
                t_floors = extract_floors(target_str)
                t_discipline = extract_discipline(target_str)
                t_round = extract_round_num(target_str)
                t_norm = normalize_text(target_str)

                best_match = None
                match_reason = ''
                
                # ★ 1대1 매핑 엄격 적용: 이미 매핑된 소스 레코드 제외 ★
                avail_sources = [s for s in source_records if s['row'] not in used_source_rows]

                # 매칭 1단계: 호수/동 및 층수 정밀 일치
                if t_room_key:
                    candidates = [s for s in source_by_room.get(t_room_key, []) if s['row'] not in used_source_rows and are_floors_compatible(t_floors, s['floors'])]
                    if candidates:
                        if t_discipline:
                            c_disc = [c for c in candidates if c['discipline'] == t_discipline]
                            if c_disc: candidates = c_disc
                        if t_round:
                            c_rnd = [c for c in candidates if c['round'] == t_round]
                            if c_rnd: candidates = c_rnd
                        best_match = candidates[0]
                        match_reason = f'🎯 호수/동 정밀 일치 ({t_room_key})'

                # 매칭 2단계: 층수/공종/차수 100% 호환 정밀 일치 (호수 상충 금지, 층수 불일치 엄격 차단)
                if not best_match:
                    cands = [s for s in avail_sources if are_floors_compatible(t_floors, s['floors'])]
                    if t_discipline:
                        cands = [c for c in cands if c['discipline'] == t_discipline]
                    if t_round:
                        cands = [c for c in cands if c['round'] == t_round]
                    
                    if cands:
                        for cand in cands:
                            # 호실 키가 둘 다 존재하는데 상충되는 경우 매칭 불허
                            if t_room_key and cand['room_key'] and cand['room_key'] != t_room_key:
                                continue
                            if cand['norm'] in t_norm or t_norm in cand['norm']:
                                best_match = cand
                                match_reason = f'🎯 공종({t_discipline or "일반"}) + 차수({t_round or "일반"}) + 층수 일치'
                                break

                # 매칭 3단계: 엄격한 텍스트 유사도 매칭 (호실/층수 상충 절대 금지)
                if not best_match:
                    for cand in avail_sources:
                        if t_room_key and cand['room_key'] and cand['room_key'] != t_room_key:
                            continue
                        if not are_floors_compatible(t_floors, cand['floors']):
                            continue
                        if t_discipline and cand['discipline'] and cand['discipline'] != t_discipline:
                            continue
                        if t_round and cand['round'] and cand['round'] != t_round:
                            continue
                            
                        if len(cand['norm']) > 3 and (cand['norm'] in t_norm or t_norm in cand['norm']):
                            best_match = cand
                            match_reason = '🔍 텍스트 유사도 매칭'
                            break

                if best_match:
                    used_source_rows.add(best_match['row'])  # ★ 1대1 매핑 등록 (중복 사용 완전 차단) ★
                    matched_count += 1

                    # === 범위 정밀 타격: 쓰기 전 행 번호 이중 검증 ===
                    if r < target_start or r > target_end:
                        self.log('ERROR', f'🚨 [범위 위반 차단] Row {r}은 지정 범위({target_start}~{target_end}) 밖입니다! 쓰기를 거부합니다.')
                        continue

                    # 쓰기 보류 - 검증 후 커밋 단계에서 작성
                    audit_item_list.append({
                        'no': seq_no,
                        'row': r,
                        'target_str': target_str,
                        'status': 'SUCCESS',
                        'source_name': best_match['name'],
                        'reason': match_reason,
                        'best_match': best_match
                    })
                    self.log('SUCCESS', f'No.{seq_no:02d} Row {r:4d} 매칭 성공 (반영 대기): [{target_str[:22]}] ↔ [{best_match["name"][:22]}] ({match_reason})')
                else:
                    if t_room_key:
                        fail_reason = f'❌ B2B 소스 파일 내 호실/동[{t_room_key}] 데이터 미존재 (누락)'
                    elif t_discipline or t_round:
                        fail_reason = f'❌ B2B 소스 내 [{t_discipline} {t_round}] 미발생 (누락)'
                    else:
                        fail_reason = '❌ B2B 소스 파일 내 미존재 (누락)'
                    audit_item_list.append({
                        'no': seq_no,
                        'row': r,
                        'target_str': target_str,
                        'status': 'FAIL',
                        'source_name': '-',
                        'reason': fail_reason,
                    })
                    self.log('WARN', f'No.{seq_no:02d} Row {r:4d} 매칭 실패 (누락): [{target_str}] -> ({fail_reason})')

            # 6단계 결과 저장 및 리포트 표시
            self.update_mapping_status(95, '💾 [6/6단계] 매핑 검증 연산 완료 (엑셀 반영 대기 중...)')
            self.log('SUCCESS', f'🎯 가상 매핑 연산 완료! (총 {matched_count}/{total_target_count}건 성공)')
            self.log('SUCCESS', f'🔒 범위 정밀 타격 완료: [{range_info["sheet_name"]}]!{range_info["range_address"]} 범위 내 연산')
            
            report_data = {
                'target_path': target_path,
                'total_selected': total_target_count,
                'matched_count': matched_count,
                'missing_count': total_target_count - matched_count,
                'highlighted_cells': highlighted_cells,
                'backup_path': backup_path,
                'range_address': range_info['range_address'],
                'sheet_name': range_info['sheet_name'],
                'source_match_col': idx_contract_name,
                'source_match_header': source_match_header,
                'target_match_col': target_col,
                'target_match_header': target_match_header,
                'items': audit_item_list,
                'mappings_used': resolved_mappings,
            }

            self.update_mapping_status(100, f'✅ 매핑 연산 완료! (성공: {matched_count}건 / 누락: {total_target_count - matched_count}건)')
            self.root.after(0, lambda: self.show_final_audit_report_dialog(report_data))

        except Exception as e:
            err_trace = traceback.format_exc()
            err_msg_str = str(e)
            if err_msg_str == 'None' or not err_msg_str:
                err_msg_str = f'Excel COM 셀 연산 예외 ({repr(e)})'
            self.update_mapping_status(0, f'❌ 매핑 중 오류 발생: {err_msg_str}')
            self.log('ERROR', f'매핑 연산 중 오류 발생:\n{err_trace}')
            self.root.after(0, lambda: messagebox.showerror('매핑 오류', f'매핑 연산 중 오류가 발생했습니다:\n{err_msg_str}'))
        finally:
            if wb_source:
                try: wb_source.Close(False)
                except Exception: pass
            if wb_target:
                try:
                    if target_opened_by_script:
                        wb_target.Close(False)
                except Exception: pass
            restore_excel_mode(excel, excel_state)
            if excel and excel_created_by_script:
                try: excel.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self.btn_run.config(state='normal'))

    def show_final_audit_report_dialog(self, report):
        popup = tk.Toplevel(self.root)
        popup.title('📊 B2B 데이터 매핑 최종 종합 검증 보고서')
        popup.geometry('1180x780')
        popup.grab_set()

        main_f = ttk.Frame(popup, padding=15)
        main_f.pack(fill='both', expand=True)

        top_cards = ttk.Frame(main_f, style='Card.TFrame', padding=12)
        top_cards.pack(fill='x', pady=(0, 8))

        total = report['total_selected']
        success = report['matched_count']
        missing = report['missing_count']
        rate = (success / total * 100.0) if total > 0 else 0.0
        source_basis_title = report.get('source_match_header', '소스 기준값')
        target_basis_title = report.get('target_match_header', '타겟 기준값')
        target_basis_col = report.get('target_match_col', 24)
        target_basis_letter = col_num_to_letter(target_basis_col)

        stat_lbl = (
            f"🎯 [범위]: {report['sheet_name']}!{report['range_address']}   |   "
            f"기준: {source_basis_title} ↔ {target_basis_title}({target_basis_letter}열)   |   "
            f"전체 선택: {total}건   |   "
            f"✅ 매핑 성공: {success}건   |   "
            f"❌ 매핑 누락: {missing}건   |   "
            f"성공률: {rate:.1f}%   |   "
            f"노란색 셀: {report['highlighted_cells']}개"
        )
        
        f_title = ttk.Frame(top_cards)
        f_title.pack(fill='x', pady=(0, 4))
        ttk.Label(f_title, text='📊 매핑 실행 종합 현황 리포트 (교차 배경색으로 시각적 식별 편의성 극대화)', font=('Segoe UI', 12, 'bold'), foreground='#1A365D').pack(side='left')
        
        def copy_full_report_to_clipboard():
            tsv_lines = [
                f"[매핑 종합 검증 보고서]",
                stat_lbl,
                "",
                f"No\t행번호\t판정상태\t선택한 타겟 {target_basis_title}\t매핑된 B2B 원본 {source_basis_title}\t매핑 판정 근거 및 사유"
            ]
            for it in report['items']:
                if it['status'] == 'SUCCESS': st_tag = '성공(일치)'
                elif it['status'] == 'MISMATCH': st_tag = '불일치'
                else: st_tag = '누락'
                tsv_lines.append(f"{it['no']:02d}\tRow {it['row']}\t{st_tag}\t{it['target_str']}\t{it['source_name']}\t{it['reason']}")
            
            full_tsv = '\n'.join(tsv_lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(full_tsv)
            self.root.update()
            messagebox.showinfo('클립보드 복사 완료', '전체 보고서 내용(탭 구분)이 클립보드에 복사되었습니다!\n[Ctrl+V]를 누르면 엑셀, 메모장, 메신저에 바로 붙여넣기할 수 있습니다.', parent=popup)

        btn_copy_all = ttk.Button(f_title, text='📋 전체 결과 엑셀용 클립보드 복사', style='Copy.TButton', command=copy_full_report_to_clipboard)
        btn_copy_all.pack(side='right')

        def toggle_status(event=None):
            selected = tree.selection()
            if not selected:
                if event is None:
                    messagebox.showwarning('선택 안됨', '상태를 변경할 항목을 표에서 선택해주세요.', parent=popup)
                return
            for item_id in selected:
                values = list(tree.item(item_id, 'values'))
                no_str = values[0]
                for it in report['items']:
                    if f"{it['no']:02d}" == no_str:
                        # 양방향 상태 전환 (성공(일치) ↔ 불일치/누락)
                        if it['status'] == 'SUCCESS':
                            it['status'] = 'MISMATCH'
                            it['reason'] = '⚠️ 사용자 검열: 불일치 판정 (제외됨)'
                        else:
                            it['status'] = 'SUCCESS'
                            it['reason'] = '✅ 사용자 검열: 성공(일치) 복원'
                        break
            apply_filter()

        try:
            ttk.Style().configure('Danger.TButton', background='#E53E3E', foreground='black')
        except Exception: pass
        btn_toggle = ttk.Button(f_title, text='🔄 선택 항목 [성공(일치) ↔ 불일치] 양방향 상태 전환', style='Danger.TButton', command=toggle_status)
        btn_toggle.pack(side='right', padx=(0, 10))

        lbl_top_stat = ttk.Label(top_cards, text=stat_lbl, font=('Consolas', 10, 'bold'), foreground='#2B6CB0')
        lbl_top_stat.pack(anchor='w')

        # ----------------------------------------------------------------------
        # 툴바: 실시간 필터링, 열 틀고정, 열 관리, 클립 복사
        # ----------------------------------------------------------------------
        tool_bar = ttk.Frame(main_f, padding=(0, 0, 0, 6))
        tool_bar.pack(fill='x')

        ttk.Label(tool_bar, text='🔍 실시간 필터링:', font=('Segoe UI', 9, 'bold')).pack(side='left', padx=(0, 4))
        var_filter = tk.StringVar()
        ent_filter = ttk.Entry(tool_bar, textvariable=var_filter, width=24, font=('Segoe UI', 9))
        ent_filter.pack(side='left', padx=(0, 10))

        lbl_filter_count = ttk.Label(tool_bar, text='표시 중: 0 / 0 건', font=('Segoe UI', 9), foreground='#4A5568')
        lbl_filter_count.pack(side='left', padx=(0, 15))

        ttk.Label(tool_bar, text='📌 고정할 열 수:', font=('Segoe UI', 9, 'bold')).pack(side='left', padx=(8, 4))
        var_freeze_count = tk.IntVar(value=2)
        spn_freeze = ttk.Spinbox(tool_bar, from_=1, to=6, textvariable=var_freeze_count, width=4, font=('Segoe UI', 9, 'bold'))
        spn_freeze.pack(side='left', padx=(0, 6))

        btn_freeze_apply = ttk.Button(tool_bar, text='📌 틀고정 적용', command=lambda: apply_freeze_columns())
        btn_freeze_apply.pack(side='left', padx=(0, 10))

        def copy_selected_rows_to_clipboard(event=None):
            selected = tree.selection()
            if not selected:
                if event is None:
                    messagebox.showwarning('선택 없음', '클립보드에 복사할 항목을 표에서 선택해 주세요.', parent=popup)
                return
            lines = [f"No\t행번호\t판정상태\t선택한 타겟 {target_basis_title}\t매핑된 B2B 원본 {source_basis_title}\t매핑 판정 근거 및 사유"]
            for s_id in selected:
                vals = tree.item(s_id, 'values')
                lines.append('\t'.join(str(v) for v in vals))
            text_data = '\n'.join(lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(text_data)
            self.root.update()
            if event is None:
                messagebox.showinfo('복사 완료', f'선택한 {len(selected)}개 항목이 클립보드(탭 구분)에 복사되었습니다!\n엑셀에 Ctrl+V로 붙여넣으세요.', parent=popup)

        btn_copy_sel = ttk.Button(tool_bar, text='✂️ 선택 행 복사', command=copy_selected_rows_to_clipboard)
        btn_copy_sel.pack(side='left', padx=(0, 10))

        btn_reset_all = ttk.Button(tool_bar, text='🔄 열 설정/필터 초기화', command=lambda: reset_column_layout_and_filter())
        btn_reset_all.pack(side='right', padx=(4, 0))

        btn_col_manage = ttk.Button(tool_bar, text='⚙️ 열 배치/숨기기 관리...', command=lambda: open_column_manager_dialog())
        btn_col_manage.pack(side='right', padx=(4, 0))

        notebook = ttk.Notebook(main_f)
        notebook.pack(fill='both', expand=True, pady=(0, 8))

        # 탭 1: 표 형식 보기 (Zebra 교차 배경색)
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text=' 📋 표 형식 보기 (행별 교차 배경색 적용) ')

        DEFAULT_COLS = ('no', 'row', 'status', 'target_str', 'source_name', 'reason')
        COL_TITLES = {
            'no': 'No.',
            'row': '행번호',
            'status': '판정 상태',
            'target_str': f'선택한 타겟 {target_basis_title}',
            'source_name': f'매핑된 B2B 원본 {source_basis_title}',
            'reason': '매핑 판정 근거 및 사유'
        }
        DEFAULT_WIDTHS = {
            'no': 50, 'row': 65, 'status': 90,
            'target_str': 270, 'source_name': 280, 'reason': 220
        }

        tree = ttk.Treeview(tab1, columns=DEFAULT_COLS, show='headings', height=14)

        tree.tag_configure('SUCCESS_EVEN', background='#F0FFF4', foreground='#2F855A')
        tree.tag_configure('SUCCESS_ODD',  background='#E6FFFA', foreground='#234E52')
        tree.tag_configure('FAIL_EVEN',    background='#FFF5F5', foreground='#C53030')
        tree.tag_configure('FAIL_ODD',     background='#FED7D7', foreground='#9B2C2C')
        tree.tag_configure('MISMATCH_EVEN', background='#FFE4E1', foreground='#FF0000', font=('Segoe UI', 9, 'bold'))
        tree.tag_configure('MISMATCH_ODD',  background='#FFE4E1', foreground='#FF0000', font=('Segoe UI', 9, 'bold'))

        scrollbar1 = ttk.Scrollbar(tab1, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar1.set)

        tree.pack(side='left', fill='both', expand=True)
        scrollbar1.pack(side='right', fill='y')

        sort_states = {c: False for c in DEFAULT_COLS}
        active_display_columns = list(DEFAULT_COLS)

        def sort_by_column(col):
            reverse = not sort_states[col]
            sort_states[col] = reverse

            items = [(tree.set(k, col), k) for k in tree.get_children('')]

            def sort_key(item):
                val = item[0]
                m = re.findall(r'\d+', str(val))
                if m:
                    try: return (0, float(m[0]))
                    except ValueError: pass
                return (1, str(val))

            items.sort(key=sort_key, reverse=reverse)

            for index, (val, k) in enumerate(items):
                tree.move(k, '', index)

            for c in DEFAULT_COLS:
                base_title = COL_TITLES[c]
                if c == col:
                    indicator = ' ▲' if not reverse else ' ▼'
                    tree.heading(c, text=base_title + indicator)
                else:
                    tree.heading(c, text=base_title)

        for c in DEFAULT_COLS:
            title = COL_TITLES[c]
            tree.heading(c, text=title, command=lambda _c=c: sort_by_column(_c))
            tree.column(c, width=DEFAULT_WIDTHS[c], anchor='center' if c in ('no', 'row', 'status') else 'w')

        def apply_filter(*args):
            query = var_filter.get().strip().lower()
            tree.delete(*tree.get_children())

            cur_success = sum(1 for it in report['items'] if it['status'] == 'SUCCESS')
            cur_missing = sum(1 for it in report['items'] if it['status'] in ('MISMATCH', 'FAIL'))
            cur_total = len(report['items'])
            cur_rate = (cur_success / cur_total * 100.0) if cur_total > 0 else 0.0

            lbl_top_stat.config(text=(
                f"🎯 [범위]: {report['sheet_name']}!{report['range_address']}   |   "
                f"전체 선택: {cur_total}건   |   "
                f"✅ 매핑 성공: {cur_success}건   |   "
                f"❌ 매핑 불일치/누락: {cur_missing}건   |   "
                f"성공률: {cur_rate:.1f}%"
            ))

            visible_count = 0
            for idx, item in enumerate(report['items']):
                if item['status'] == 'SUCCESS': status_tag_text = '✅ 성공'
                elif item['status'] == 'MISMATCH': status_tag_text = '🚫 불일치'
                else: status_tag_text = '❌ 누락'

                row_values = (
                    f"{item['no']:02d}",
                    f"Row {item['row']}",
                    status_tag_text,
                    item['target_str'],
                    item['source_name'],
                    item['reason']
                )

                row_str = ' '.join(str(v) for v in row_values).lower()
                if not query or query in row_str:
                    is_even = (visible_count % 2 == 0)
                    tag_name = f"{item['status']}_{'EVEN' if is_even else 'ODD'}"
                    tree.insert('', 'end', values=row_values, tags=(tag_name,))
                    visible_count += 1

            lbl_filter_count.config(text=f'표시 중: {visible_count} / {len(report["items"])} 건')

        var_filter.trace_add('write', apply_filter)

        def apply_freeze_columns():
            try:
                cnt = int(var_freeze_count.get())
            except Exception:
                cnt = 2
            cnt = max(1, min(cnt, len(active_display_columns)))
            var_freeze_count.set(cnt)

            frozen_cols = active_display_columns[:cnt]
            rem_cols = active_display_columns[cnt:]
            tree['displaycolumns'] = frozen_cols + rem_cols

        apply_freeze_columns()
        apply_filter()

        tree.bind('<Double-1>', toggle_status)
        tree.bind('<Control-c>', copy_selected_rows_to_clipboard)

        def open_column_manager_dialog():
            mgr_pop = tk.Toplevel(popup)
            mgr_pop.title('⚙️ 테이블 열 배치 및 순서 관리')
            mgr_pop.geometry('480x440')
            mgr_pop.attributes('-topmost', True)

            f_mgr = ttk.Frame(mgr_pop, padding=15)
            f_mgr.pack(fill='both', expand=True)

            ttk.Label(f_mgr, text='⚙️ 화면 표시 열 순서 지정', font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0, 6))

            manageable_cols = list(active_display_columns)

            lbl_box = tk.Listbox(f_mgr, font=('Segoe UI', 10), selectmode='single', height=10)
            lbl_box.pack(fill='both', expand=True, pady=(0, 10))

            for c in manageable_cols:
                lbl_box.insert('end', f'{COL_TITLES[c]} ({c})')

            f_ctrl = ttk.Frame(f_mgr)
            f_ctrl.pack(fill='x', pady=(0, 10))

            def move_up():
                sel = lbl_box.curselection()
                if not sel or sel[0] == 0: return
                idx = sel[0]
                manageable_cols[idx], manageable_cols[idx-1] = manageable_cols[idx-1], manageable_cols[idx]
                lbl_box.delete(0, 'end')
                for c in manageable_cols:
                    lbl_box.insert('end', f'{COL_TITLES[c]} ({c})')
                lbl_box.selection_set(idx-1)

            def move_down():
                sel = lbl_box.curselection()
                if not sel or sel[0] == len(manageable_cols) - 1: return
                idx = sel[0]
                manageable_cols[idx], manageable_cols[idx+1] = manageable_cols[idx+1], manageable_cols[idx]
                lbl_box.delete(0, 'end')
                for c in manageable_cols:
                    lbl_box.insert('end', f'{COL_TITLES[c]} ({c})')
                lbl_box.selection_set(idx+1)

            ttk.Button(f_ctrl, text='▲ 위로 (좌측으로 이동)', command=move_up).pack(side='left', padx=(0, 6))
            ttk.Button(f_ctrl, text='▼ 아래로 (우측으로 이동)', command=move_down).pack(side='left')

            def apply_col_changes():
                nonlocal active_display_columns
                active_display_columns = list(manageable_cols)
                tree['displaycolumns'] = active_display_columns
                apply_freeze_columns()
                mgr_pop.destroy()

            f_mgr_btns = ttk.Frame(f_mgr)
            f_mgr_btns.pack(fill='x', side='bottom')

            ttk.Button(f_mgr_btns, text='☑ 설정 적용', style='Primary.TButton', command=apply_col_changes).pack(side='right', padx=(6, 0))
            ttk.Button(f_mgr_btns, text='취소', command=mgr_pop.destroy).pack(side='right')

        def reset_column_layout_and_filter():
            nonlocal active_display_columns
            var_filter.set('')
            var_freeze_count.set(2)
            active_display_columns = list(DEFAULT_COLS)
            tree['displaycolumns'] = DEFAULT_COLS

            for c in DEFAULT_COLS:
                sort_states[c] = False
                tree.heading(c, text=COL_TITLES[c])
                tree.column(c, width=DEFAULT_WIDTHS[c])

            apply_freeze_columns()
            apply_filter()

        # 탭 2: 텍스트 자유 선택/복사 모드
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text=' 📝 텍스트 자유 선택/복사 모드 ')

        txt_audit_report = scrolledtext.ScrolledText(tab2, font=('Consolas', 10), bg='#FFFFFF', fg='#1A202C', selectbackground='#3182CE', selectforeground='#FFFFFF')
        txt_audit_report.pack(fill='both', expand=True)

        txt_audit_report.tag_config('LINE_EVEN', background='#FFFFFF')
        txt_audit_report.tag_config('LINE_ODD', background='#F7FAFC')

        txt_audit_report.insert(tk.END, f"===========================================================================================\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"                    📊 B2B 데이터 매핑 최종 종합 검증 보고서\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"===========================================================================================\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"{stat_lbl}\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"-------------------------------------------------------------------------------------------\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"No | 행번호  | 상태   | 타겟 {target_basis_title:<28s} | B2B 원본 {source_basis_title:<28s} | 판정 근거 및 사유\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"-------------------------------------------------------------------------------------------\n", 'LINE_EVEN')

        for idx, it in enumerate(report['items']):
            line_tag = 'LINE_EVEN' if (idx % 2 == 0) else 'LINE_ODD'
            if it['status'] == 'SUCCESS': st_tag = '✅성공'
            elif it['status'] == 'MISMATCH': st_tag = '🚫불일치'
            else: st_tag = '❌누락'
            line_str = f"{it['no']:02d} | Row {it['row']:4d} | {st_tag} | {it['target_str']:<35s} | {it['source_name']:<35s} | {it['reason']}\n"
            txt_audit_report.insert(tk.END, line_str, line_tag)

        txt_audit_report.insert(tk.END, f"===========================================================================================\n", 'LINE_EVEN')

        # 탭 3: 사용된 헤더 매핑 정보
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text=' ⚙️ 사용된 헤더 매핑 정보 ')

        txt_mapping_info = scrolledtext.ScrolledText(tab3, font=('Consolas', 10), bg='#FFFFFF', fg='#1A202C')
        txt_mapping_info.pack(fill='both', expand=True)

        txt_mapping_info.insert(tk.END, f"===========================================================================================\n")
        txt_mapping_info.insert(tk.END, f"                    ⚙️ 이번 매핑에 사용된 헤더 매핑 설정\n")
        txt_mapping_info.insert(tk.END, f"===========================================================================================\n")
        txt_mapping_info.insert(tk.END, f"No | 소스 헤더명 (B2B)                  | 소스 Col | 타겟 열 | 타겟 Col | 모드\n")
        txt_mapping_info.insert(tk.END, f"-------------------------------------------------------------------------------------------\n")

        if 'mappings_used' in report:
            for idx, rm in enumerate(report['mappings_used']):
                txt_mapping_info.insert(tk.END, f"{idx+1:02d} | {rm['source_name']:<35s} | Col {rm['source_col']:3d} | {col_num_to_letter(rm['target_col']):>5s}열 | Col {rm['target_col']:3d} | {rm['mode']}\n")

        txt_mapping_info.insert(tk.END, f"===========================================================================================\n")
        txt_mapping_info.config(state='disabled')

        bot_f = ttk.Frame(popup, padding=(0, 5, 0, 0))
        bot_f.pack(fill='x')
        
        lbl_bak = ttk.Label(bot_f, text=f"💾 자동 백업 생성 위치: {os.path.basename(report['backup_path'])}", font=('Segoe UI', 8), foreground='#718096')
        lbl_bak.pack(side='left', padx=15)

        btn_copy_bot = ttk.Button(bot_f, text='📋 전체 텍스트 클립보드 복사', style='Copy.TButton', command=copy_full_report_to_clipboard)
        btn_copy_bot.pack(side='right', padx=(0, 10))

        def commit_and_close():
            result = messagebox.askyesno('엑셀 최종 반영', '현재 지정된 상태(성공 및 불일치)를 엑셀 파일에 최종 반영하시겠습니까?\n\n- 성공: 매핑 값 쓰기 및 노란색 칠하기\n- 불일치: 값 쓰기 제외, 붉은색 표시\n\n이 작업은 엑셀 파일을 변경합니다.', parent=popup)
            if not result: return
            
            self.update_mapping_status(10, '💾 엑셀에 최종 반영 중 (잠시만 기다려주세요)...')
            popup.destroy()
            
            def do_commit():
                ensure_excel_com_available()
                pythoncom.CoInitialize()
                wb_target_commit = None
                excel = None
                excel_state = None
                try:
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel_state = set_excel_fast_mode(excel)
                    
                    target_path = report['target_path']
                    wb_target_commit = excel.Workbooks.Open(target_path)
                    ws_target_commit = wb_target_commit.Worksheets(report['sheet_name'])
                    
                    YELLOW_COLOR = 65535
                    RED_COLOR = 255
                    
                    highlighted_cells = 0
                    mismatch_cells = 0
                    
                    for it in report['items']:
                        r = it['row']
                        if it['status'] == 'SUCCESS':
                            best_match = it['best_match']
                            for rm in report['mappings_used']:
                                val = best_match.get(f'col_{rm["target_col"]}')
                                if val is not None:
                                    cell = ws_target_commit.Cells(r, rm['target_col'])
                                    cell.Value = val
                                    cell.Interior.Color = YELLOW_COLOR
                                    highlighted_cells += 1
                        elif it['status'] in ('MISMATCH', 'FAIL'):
                            cell = ws_target_commit.Cells(r, target_basis_col)
                            cell.Interior.Color = RED_COLOR
                            mismatch_cells += 1
                            
                    wb_target_commit.Save()
                    self.update_mapping_status(100, f'✅ 최종 엑셀 반영 완료! (성공 반영 {highlighted_cells}셀, 불일치 반영 {mismatch_cells}건)')
                    
                    self.root.after(0, lambda: messagebox.showinfo('반영 완료', f'엑셀 파일에 매핑 결과가 성공적으로 반영되었습니다.\n\n적용된 매핑: {highlighted_cells}셀\n불일치 표시(빨간색): {mismatch_cells}건'))
                except Exception as e:
                    err_msg = str(e)
                    self.root.after(0, lambda msg=err_msg: messagebox.showerror('반영 오류', f'엑셀 반영 중 오류 발생:\n{msg}'))
                finally:
                    if wb_target_commit:
                        try: wb_target_commit.Close(False)
                        except Exception: pass
                    restore_excel_mode(excel, excel_state)
                    if excel:
                        try: excel.Quit()
                        except Exception: pass
                    pythoncom.CoUninitialize()

            threading.Thread(target=do_commit, daemon=True).start()

        btn_close = ttk.Button(bot_f, text='💾 검증 완료 및 엑셀에 최종 반영', style='Primary.TButton', command=commit_and_close)
        btn_close.pack(side='right', padx=(0, 15))

        self.log('SUCCESS', f'최종 종합 검증 보고서 팝업 표시 완료 (성공 {success}건 / 누락 {missing}건)')

def main():
    root = tk.Tk()
    B2BMappingAutomationApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
