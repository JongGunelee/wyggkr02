# -*- coding: utf-8 -*-
"""
================================================================================
🚀 B2B 3자간(발주내역 ↔ 매입월정산 ↔ 매출월정산) 데이터 교차 검증 및 정제 자동화 솔루션 v7.2
================================================================================
- 계층 구조: Clean Architecture 3-Layer Pattern
  1. DOMAIN LAYER (Engine): Strict TextMatcher, Dynamic IssueType Extractor, ReconciliationEngine, 3-Way Header Mapping Resolver, Real-time Dynamic Excel Header Scanner
  2. PRESENTATION LAYER (View): Tkinter GUI (3자간 전용 헤더 매핑 시각화 Canvas 편집기, 업로드 엑셀 파일 실시간 동적 구조 파악 연동, 교차 색상 Zebra Striping, 사용자 지정 열 틀고정)
  3. APPLICATION LAYER (Controller): Smart Action Queue, COM Isolation, Background Threading

- 주요 개편 사항 (v7.2 엑셀 파일 업로드 시 자동 구조 파악 실시간 동적 연동):
  - [100% 동적 엑셀 헤더 자동 스캐너 엔진 (`auto_scan_headers_from_file`)]:
    - 하드코딩 의존성을 완전히 탈피하여, 사용자가 엑셀 파일(발주/매입/매출)을 업로드/선택할 때마다 백그라운드 스레드에서 해당 엑셀 파일의 1~40번 전체 컬럼 실데이터 헤더를 즉시 자동 파싱 및 실시간 업데이트
    - 새로 변경된 엑셀 파일이나 다음 달 정산 파일을 선택하더라도 실시간으로 해당 파일의 실제 헤더 구조를 자동 파악하여 드롭다운 옵션에 100% 반영
  - [스마트 초기 폴백(Fallback) 사전 구축]:
    - 파일 선택 이전 초기 상태에서는 사전 검증된 실데이터 헤더를 기본값으로 제공하며, 파일 선택 시 즉시 실시간 동적 헤더 정보로 자동 교체
  - [발주/매입/매출 3자간 완전 독립 동적 옵션 연동]:
    - 드롭다운 목록(`X (Col 24) - [실시간 파싱된 헤더]`)이 각 엑셀 파일별로 실시간 동적 생성
================================================================================
"""
import os
import sys
import time

# [가이드라인 준수] CP949 표준 출력 콘솔 강제 UTF-8 인코딩 통제
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    else:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except Exception:
    pass

import re
import difflib
import gc
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
    import win32process
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

EXCEL_CALCULATION_MANUAL = -4135
NO_HEADER_MATCH_OPTION = '매칭 되지 않음'


def is_no_header_match_option(option_str):
    return str(option_str or '').strip().startswith(NO_HEADER_MATCH_OPTION)


def ensure_excel_com_available():
    """Excel만 설치된 PC에서도 pywin32가 없으면 1회 자동 설치를 시도한다."""
    global COM_AVAILABLE, win32com, pythoncom, win32gui, win32con, win32process
    if COM_AVAILABLE:
        return True
    try:
        import win32com.client as _win32_client
        import pythoncom as _pythoncom
        import win32gui as _win32gui
        import win32con as _win32con
        import win32process as _win32process
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
            import win32process as _win32process
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
    win32process = _win32process
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


def normalize_file_identity(path):
    """같은 파일이 여러 방식의 경로 문자열로 들어와도 하나의 식별자로 비교한다."""
    if not path:
        return ''
    try:
        return os.path.normcase(os.path.abspath(os.path.expanduser(str(path))))
    except Exception:
        return os.path.normcase(str(path))


def classify_reconciliation_file_selection(file_paths, current_purchase='', current_sales=''):
    """
    다중 선택된 파일을 기본 매입/매출 1개씩과 무제한 n자 파일로 분류한다.

    규칙:
    - 선택 묶음 안에서 첫 번째 매입 파일만 기본 매입 슬롯에 배정한다.
    - 선택 묶음 안에서 첫 번째 매출 파일만 기본 매출 슬롯에 배정한다.
    - 같은 분류의 두 번째 이후 파일과 기타 파일은 모두 n자로 누적한다.
    - 파일명에 매입/매출 키워드가 없는 파일은 기본 슬롯이 비어 있어도 n자로 누적한다.
    - 파일명에 매입과 매출이 모두 들어간 파일은 오분류 방지를 위해 n자로 누적한다.
    """
    selected_paths = []
    seen = set()
    for fp in file_paths or []:
        file_id = normalize_file_identity(fp)
        if not file_id or file_id in seen:
            continue
        seen.add(file_id)
        selected_paths.append(fp)

    purchase_path = ''
    sales_path = ''
    overflow_paths = []

    for fp in selected_paths:
        fn = os.path.basename(fp)
        is_purchase = '매입' in fn
        is_sales = '매출' in fn

        if is_purchase and not is_sales and not purchase_path:
            purchase_path = fp
        elif is_sales and not is_purchase and not sales_path:
            sales_path = fp
        else:
            overflow_paths.append(fp)

    return purchase_path, sales_path, overflow_paths


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


def delete_rows_or_raise(ws, rows, label):
    failed_rows = []
    for r in sorted(rows, reverse=True):
        try:
            ws.Rows(r).Delete()
        except Exception:
            failed_rows.append(r)
    if failed_rows:
        failed_text = ', '.join(str(r) for r in failed_rows[:20])
        if len(failed_rows) > 20:
            failed_text += ', ...'
        raise RuntimeError(f'{label} 행 삭제 실패: {failed_text}')
    return len(rows)


def extract_issue_from_cached_row(row_values, preferred_cols=None):
    target_cols = preferred_cols if preferred_cols else [19, 22]
    for c_idx in target_cols:
        val = row_values.get(c_idx)
        if val:
            s_val = str(val).strip()
            if s_val in ['역발행', '정발행', '미발행', '직접발행']:
                return s_val
    for c_idx in range(1, 36):
        val = row_values.get(c_idx)
        if val:
            s_val = str(val).strip()
            if s_val in ['역발행', '정발행', '미발행', '직접발행']:
                return s_val
    return '-'


# ==============================================================================
# HELPER FUNCTIONS & DEFAULT COLUMN HEADERS DICTIONARIES
# ==============================================================================

def col_num_to_letter(col_num):
    """숫자 컬럼 번호를 엑셀 알파벳 열 이름으로 변환 (예: 1 -> A, 24 -> X, 28 -> AB)"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def col_letter_to_num(letter):
    """엑셀 열 문자(A, B, ..., AA, AB, ...)를 컬럼 번호(1-based)로 변환"""
    result = 0
    for ch in str(letter).upper():
        if 'A' <= ch <= 'Z':
            result = result * 26 + (ord(ch) - 64)
    return result or 1


class HoverTooltip:
    """Tkinter 위젯에 상세 사용 가이드를 지연 표시하는 경량 툴팁."""

    def __init__(self, widget, text, delay=450, wraplength=680):
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


def clean_po_val(val):
    """PO/PR번호 문자열 정제 (float .0 제거 및 공백 정리)"""
    if not val: return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2].strip()
    return s

def is_valid_po_no(text):
    """실제 PO/PR/공사번호 검증 (날짜 2026-08-31 및 긴 한글 문장 제외)"""
    if not text: return False
    s = str(text).strip()
    # 1. 날짜 형태(YYYY-MM-DD 또는 시분초 포함) 제외
    if re.search(r'\d{4}-\d{2}-\d{2}', s) or ':' in s:
        return False
    # 2. 15자 이상의 긴 한글 문장(작업 설명 내용) 제외
    if len(s) > 15 and any('\uac00' <= ch <= '\ud7a3' for ch in s):
        return False
    # 3. 최소한 알파벳이나 숫자가 하나 이상 포함되어 있으면 유효한 PO/PR 번호로 인정
    if any(ch.isalnum() for ch in s):
        return True
    return False


# 초기 파일 미선택 시 기본 폴백(Fallback) 헤더 사전
DEFAULT_REF_COL_HEADERS = {
    1: 'Ev. OFF', 2: '평균(연도별)', 3: '정산담당', 4: '회계연도', 5: '매입 마감(월)',
    6: '회계 마감', 7: '특이사항', 8: '발주처', 9: '비용배분', 10: 'S&I 센터',
    11: '작업유형', 12: '장소', 13: '건물동', 14: '건물층', 15: '호실',
    16: '주요위험', 17: '공종(대)', 18: '공종(중)', 19: '공종(소)', 20: '협력사(표준)',
    21: '협력사(공구기구)', 22: '협력사 담당자', 23: '협력사 연락처', 24: '계약명', 25: '낙찰금액',
    26: 'PO금액', 27: 'PO번호', 28: 'PR번호', 29: 'PR금액', 30: 'PR 발생일',
    31: '공사/유지보수 번호', 32: '공사 요청일', 33: '공사 완료일', 34: '계약 체결일', 35: '담당'
}

DEFAULT_PUR_COL_HEADERS = {
    1: 'No.', 2: '정산상태', 3: '상태메시지', 4: '정산전표', 5: 'WBS 정보/WBS',
    6: 'WBS명', 7: '기성정보/업체', 8: '업체명', 9: '서비스유형', 10: '계정코드',
    11: '정산일자', 12: '기성계획금액', 13: '증감금액', 14: '정산금액', 15: '세액',
    16: '총금액', 17: '인지세차감금액', 18: '고용보험료차감금액', 19: '발행구분', 20: '미확정사유',
    21: '회원ID', 22: '회원명', 23: '국세청승인번호', 24: '발행상태', 25: '정산취소 품의번호',
    26: '공사번호', 27: '계약번호', 28: '계약명', 29: '계약담당자', 30: '부서',
    31: '매입계약유형', 32: 'FM전표유형', 33: 'FM전표번호', 34: '결산조정전표', 35: '정산주기'
}

DEFAULT_SAL_COL_HEADERS = {
    1: 'No.', 2: '정산상태', 3: '상태메시지', 4: '정산전표', 5: '세금계산서통합',
    6: 'WBS 정보/WBS', 7: 'WBS명', 8: '청구정보/고객', 9: '고객명', 10: '서비스유형',
    11: '세부항목', 12: '적요', 13: '정산일자', 14: '청구계획금액', 15: '증감금액',
    16: '정산금액', 17: '세액', 18: '총금액', 19: '미확정사유', 20: '회원ID',
    21: '회원명', 22: '발행구분', 23: '국세청승인번호', 24: '발행상태', 25: '공사번호',
    26: '계약번호', 27: '계약명', 28: '계약담당자', 29: '부서', 30: '매출계약유형',
    31: 'FM전표유형', 32: 'FM전표번호', 33: '결산조정전표', 34: '정산주기'
}

def parse_col_num_from_str(option_str):
    """옵션 텍스트에서 컬럼 숫자 추출 (예: 'X (Col 24) - 계약명' -> 24)"""
    if not option_str: return 1
    if is_no_header_match_option(option_str):
        return 1
    m = re.search(r'Col\s*(\d+)', option_str)
    if m:
        return int(m.group(1))
    m2 = re.findall(r'\d+', option_str)
    if m2:
        return int(m2[0])
    return 1


def parse_optional_col_num_from_str(option_str):
    if not option_str or is_no_header_match_option(option_str):
        return None
    m = re.search(r'Col\s*(\d+)', str(option_str))
    if m:
        return int(m.group(1))
    m2 = re.findall(r'\d+', str(option_str))
    return int(m2[0]) if m2 else None


# ==============================================================================
# LAYER 1: DOMAIN ENGINE (비즈니스 로직 & 엑셀 제어 도우미)
# ==============================================================================

def safe_cell_value(ws, r, c):
    """엑셀 셀 값을 안전하게 추출 (Value / Value2 2단계 fallback)"""
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

def extract_row_issue_type(ws, r, preferred_cols=None):
    """
    행 내부의 발행구분(역발행/정발행)을 동적으로 안전하게 추출
    """
    target_cols = preferred_cols if preferred_cols else [19, 22]
    for c_idx in target_cols:
        val = safe_cell_value(ws, r, c_idx)
        if val:
            s_val = str(val).strip()
            if s_val in ['역발행', '정발행', '미발행', '직접발행']:
                return s_val

    for c_idx in range(1, 36):
        val = safe_cell_value(ws, r, c_idx)
        if val:
            s_val = str(val).strip()
            if s_val in ['역발행', '정발행', '미발행', '직접발행']:
                return s_val

    return '-'

def extract_room_key(text):
    """동/호수 패턴 정밀 추출"""
    if not text: return ''
    s = str(text).strip()
    m = re.search(r'(\d+)\s*동\s*(\d+)\s*호?', s)
    if m:
        return f'{m.group(1)}-{m.group(2)}'
    m2 = re.search(r'(\d+)[\s\-_/]+(\d+)', s)
    if m2:
        return f'{m2.group(1)}-{m2.group(2)}'
    m3 = re.search(r'(\d+)\s*동', s)
    if m3:
        return f'{m3.group(1)}동'
    return ''

def extract_discipline(text):
    """공종명 키워드 추출"""
    if not text: return ''
    s = str(text)
    disciplines = ['기계', '건축', '전기', '토목', '소방', '통신', '조경', '수장', '설비', '위생', '가스', '소화기', '덕트', '배관']
    for d in disciplines:
        if d in s:
            return d
    return ''

def extract_round_num(text):
    """차수 추출"""
    if not text: return ''
    m = re.search(r'(\d+)\s*차', str(text))
    if m:
        return f'{m.group(1)}차'
    return ''

def normalize_text(text):
    """텍스트 정규화"""
    if not text: return ''
    t = re.sub(r'\[[^\]]+\]', '', str(text))
    t = t.replace(' ', '').replace('\n', '').replace('\r', '').lower()
    t = t.replace('수장작업', '수장공사').replace('누수수리', '누수보수')
    t = t.replace('작업', '공사').replace('보수', '수리')
    return t

def is_text_matched(target_text, ref_text):
    """엄격하고 정밀한 텍스트 유사성 매칭 알고리즘"""
    if not target_text or not ref_text:
        return False

    norm_t = normalize_text(target_text)
    norm_r = normalize_text(ref_text)

    if not norm_t or not norm_r:
        return False

    if norm_t == norm_r:
        return True

    if norm_t in norm_r or norm_r in norm_t:
        return True

    room_t = extract_room_key(target_text)
    room_r = extract_room_key(ref_text)
    if room_t and room_r:
        if room_t != room_r:
            return False
        disc_t = extract_discipline(target_text)
        disc_r = extract_discipline(ref_text)
        if disc_t and disc_r and disc_t != disc_r:
            return False
        return True

    return False


# ==============================================================================
# LAYER 2 & 3: APPLICATION CONTROLLER & PRESENTATION VIEW
# ==============================================================================

class B2BThreeWayReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title('B2B 3자간(발주-매입-매출) 데이터 교차 검증 및 정제 자동화 v7.2')
        self.root.geometry('1260x980')
        self.root.minsize(1080, 860)

        # 경로 및 설정 변수
        self.ref_file_path = tk.StringVar()
        self.purchase_file_path = tk.StringVar()
        self.sales_file_path = tk.StringVar()
        self.selected_range_address = tk.StringVar()
        self.extra_parties = []
        self.next_extra_party_no = 1
        
        self.ref_range_info = None
        self.ref_items = []
        self.range_popup = None
        self.excel_activation_busy = False
        self.interactive_excel_com_initialized = False
        self.interactive_excel_app = None
        self.interactive_target_wb = None

        # 실시간 엑셀 동적 헤더 파싱 사전 (초기값: 폴백 사전)
        self.ref_headers = dict(DEFAULT_REF_COL_HEADERS)
        self.pur_headers = dict(DEFAULT_PUR_COL_HEADERS)
        self.sal_headers = dict(DEFAULT_SAL_COL_HEADERS)

        # 3자간 전용 헤더 매핑 설정
        self.header_mappings = self.get_default_mappings()

        # 백그라운드 상태 및 스마트 액션 큐
        self.active_tasks = {'ref': False, 'purchase': False, 'sales': False, 'extra': False}
        self.queued_action = None

        self.setup_styles()
        self.build_ui()

        # [가이드라인 준수] GUI 가시성 및 윈도우 포커스 하드닝 고정 시퀀스
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(500, lambda: self.root.attributes('-topmost', False))
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

    # --------------------------------------------------------------------------
    # 🎯 [100% 동적 엑셀 헤더 실시간 파싱 엔진]
    # --------------------------------------------------------------------------
    def auto_scan_headers_from_file(self, file_path, tag):
        """
        사용자가 엑셀 파일(발주/매입/매출)을 선택하면 순수 파이썬(pyxlsb/openpyxl) 기반으로
        엑셀 COM 프로세스 충돌이나 닫힘 현상 전혀 없이 헤더 구조를 즉시 자동 파싱합니다.
        """
        if not file_path or not os.path.exists(file_path):
            return

        def _scan_thread():
            scanned_dict = {}
            abs_p = os.path.abspath(file_path)
            ext = os.path.splitext(abs_p)[1].lower()

            # 1. xlsb 지원 (pyxlsb - 순수 파이썬 라이브러리, COM 닫힘 버그 0%)
            if ext == '.xlsb':
                try:
                    import pyxlsb
                    with pyxlsb.open_workbook(abs_p) as wb:
                        sheet_name = None
                        for s_name in wb.sheets:
                            if '집계' in s_name:
                                sheet_name = s_name
                                break
                        if not sheet_name:
                            sheet_name = wb.sheets[0]
                        with wb.get_sheet(sheet_name) as sheet:
                            r1, r2 = None, None
                            for idx, row in enumerate(sheet.rows()):
                                if idx == 0:
                                    r1 = [str(cell.v or '').strip() if cell.v is not None else '' for cell in row]
                                elif idx == 1:
                                    r2 = [str(cell.v or '').strip() if cell.v is not None else '' for cell in row]
                                    break
                            max_c = max(len(r1 or []), len(r2 or []))
                            for c in range(1, min(max_c + 1, 121)):
                                v1 = r1[c-1] if r1 and c-1 < len(r1) else ''
                                v2 = r2[c-1] if r2 and c-1 < len(r2) else ''
                                scanned_dict[c] = self.compose_header_label(v1, v2)
                except Exception:
                    pass

            # 2. xlsx/xlsm 지원 (openpyxl)
            if not scanned_dict and ext in ('.xlsx', '.xlsm'):
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(abs_p, read_only=True, data_only=True)
                    sheet_name = None
                    for s_name in wb.sheetnames:
                        if '집계' in s_name:
                            sheet_name = s_name
                            break
                    if not sheet_name:
                        sheet_name = wb.sheetnames[0]
                    ws = wb[sheet_name]
                    rows = list(ws.iter_rows(max_row=2, values_only=True))
                    r1 = list(rows[0]) if len(rows) > 0 else []
                    r2 = list(rows[1]) if len(rows) > 1 else []
                    max_c = max(len(r1), len(r2))
                    for c in range(1, min(max_c + 1, 121)):
                        v1 = str(r1[c-1] or '').strip() if c-1 < len(r1) else ''
                        v2 = str(r2[c-1] or '').strip() if c-1 < len(r2) else ''
                        scanned_dict[c] = self.compose_header_label(v1, v2)
                    wb.close()
                except Exception:
                    pass

            # 3. 폴백: win32com (만약 위 파이썬 전용 라이브러리로 못 읽은 경우 - excel.Quit() 절대 실행 금지!)
            if not scanned_dict:
                ensure_excel_com_available()
                pythoncom.CoInitialize()
                excel = None
                excel_state = None
                try:
                    excel = win32com.client.DispatchEx('Excel.Application')
                    excel.Visible = False
                    excel_state = set_excel_fast_mode(excel)
                    wb = excel.Workbooks.Open(abs_p, ReadOnly=True)
                    ws = None
                    if tag == 'ref':
                        for s in wb.Worksheets:
                            if '집계' in s.Name:
                                ws = s
                                break
                    if not ws: ws = wb.Worksheets(1)

                    for c in range(1, 121):
                        v1 = str(ws.Cells(1, c).Value or '').strip()
                        v2 = str(ws.Cells(2, c).Value or '').strip()
                        scanned_dict[c] = self.compose_header_label(v1, v2)
                    wb.Close(False)
                except Exception: pass
                finally:
                    restore_excel_mode(excel, excel_state)
                    if excel:
                        try: excel.Quit()
                        except Exception: pass
                    pythoncom.CoUninitialize()

            def _update_headers_in_ui():
                fn = os.path.basename(file_path)
                if tag == 'ref' and scanned_dict:
                    self.ref_headers.update(scanned_dict)
                    self.log('SUCCESS', f'🔍 [발주내역] 엑셀 실데이터 구조 자동 파악 완료 ({fn})')
                elif tag == 'pur' and scanned_dict:
                    self.pur_headers.update(scanned_dict)
                    self.log('SUCCESS', f'🔍 [매입월정산] 엑셀 실데이터 구조 자동 파악 완료 ({fn})')
                elif tag == 'sal' and scanned_dict:
                    self.sal_headers.update(scanned_dict)
                    self.log('SUCCESS', f'🔍 [매출월정산] 엑셀 실데이터 구조 자동 파악 완료 ({fn})')
                elif str(tag).startswith('extra:') and scanned_dict:
                    party_id = str(tag).split(':', 1)[1]
                    party = self.get_extra_party_by_id(party_id)
                    if party:
                        party['headers'].clear()
                        party['headers'].update(scanned_dict)
                        if not party.get('mapping_manual'):
                            party['mapping'].update(self.infer_party_column_mapping(scanned_dict, fallback_contract_col=party['mapping'].get('contract_col', 11)))
                        self.log('SUCCESS', f'🔍 [{party["label"]}] 엑셀 실데이터 구조 자동 파악 완료 ({fn})')

            self.root.after(0, _update_headers_in_ui)

        threading.Thread(target=_scan_thread, daemon=True).start()

    def get_ref_col_option_str(self, c):
        let = col_num_to_letter(c)
        hdr = self.ref_headers.get(c, '')
        return f"{let} (Col {c}) - {hdr}" if hdr else f"{let} (Col {c})"

    def get_pur_col_option_str(self, c):
        let = col_num_to_letter(c)
        hdr = self.pur_headers.get(c, '')
        return f"{let} (Col {c}) - {hdr}" if hdr else f"{let} (Col {c})"

    def get_sal_col_option_str(self, c):
        let = col_num_to_letter(c)
        hdr = self.sal_headers.get(c, '')
        return f"{let} (Col {c}) - {hdr}" if hdr else f"{let} (Col {c})"

    def get_extra_col_option_str(self, party, c):
        let = col_num_to_letter(c)
        hdr = party.get('headers', {}).get(c, '')
        return f"{let} (Col {c}) - {hdr}" if hdr else f"{let} (Col {c})"

    def get_extra_party_by_id(self, party_id):
        for party in self.extra_parties:
            if party['id'] == party_id:
                return party
        return None

    def is_file_already_registered(self, file_path):
        file_id = normalize_file_identity(file_path)
        if not file_id:
            return False
        def get_var_text(var):
            try:
                return var.get() if hasattr(var, 'get') else str(var or '')
            except Exception:
                return ''

        registered_ids = {
            normalize_file_identity(get_var_text(getattr(self, 'purchase_file_path', ''))),
            normalize_file_identity(get_var_text(getattr(self, 'sales_file_path', ''))),
        }
        for party in getattr(self, 'extra_parties', []):
            path_var = party.get('file_path_var')
            registered_ids.add(normalize_file_identity(get_var_text(path_var)))
        registered_ids.discard('')
        return file_id in registered_ids

    def infer_col_by_keywords(self, headers, keyword_groups, fallback):
        normalized_headers = {c: normalize_text(self.clean_basis_title(h)) for c, h in headers.items() if h}
        for keywords in keyword_groups:
            normalized_keywords = [normalize_text(k) for k in keywords]
            for c, header in normalized_headers.items():
                if all(k in header for k in normalized_keywords):
                    return c
        for keywords in keyword_groups:
            normalized_keywords = [normalize_text(k) for k in keywords]
            for c, header in normalized_headers.items():
                if any(k in header for k in normalized_keywords):
                    return c
        return fallback

    def find_header_col_by_title(self, headers, title, fallback=None):
        normalized_title = normalize_text(self.clean_basis_title(title))
        if not normalized_title:
            return fallback
        normalized_headers = {c: normalize_text(self.clean_basis_title(h)) for c, h in headers.items() if h}
        for c, header in normalized_headers.items():
            if header == normalized_title:
                return c
        for c, header in normalized_headers.items():
            if normalized_title in header or header in normalized_title:
                return c
        return fallback

    def infer_field_key_from_title(self, title):
        clean_title = self.clean_basis_title(title)
        if not clean_title:
            return None
        raw = str(clean_title or '').upper()
        normalized = normalize_text(clean_title)
        if any(token in normalized for token in ['마감', '월말', '정산월', '회계월', '회계마감']):
            return None
        if 'PR' in raw or 'pr' in normalized:
            return 'pr_no'
        if 'PO' in raw or 'po' in normalized or 'WBS' in raw:
            return 'po_no'
        if '번호' in str(clean_title or ''):
            return 'po_no'
        if any(token in str(clean_title or '') for token in ['발행', '상태', '구분']):
            return 'issue_type'
        if any(token in str(clean_title or '') for token in ['유형', '계정', '서비스']):
            return 'contract_type'
        if any(token in str(clean_title or '') for token in ['금액', '총액', '비용', '세액', '단가', '정산액', '낙찰액']):
            return 'amount'
        if any(token in str(clean_title or '') for token in ['계약명', '공사명', '입찰명', '유지관리명', '관리명', '적요', '세부항목', '작업명', '업무명', '대상명', '시설명', 'WBS명']):
            return 'contract_name'
        return None

    def get_keyword_groups_for_title(self, title):
        field_key = self.infer_field_key_from_title(title)
        if field_key == 'po_no':
            return [['PO번호'], ['PO'], ['계약번호'], ['공사번호'], ['WBS']]
        if field_key == 'pr_no':
            return [['PR번호'], ['PR'], ['계약번호'], ['공사번호']]
        if field_key == 'issue_type':
            return [['발행구분'], ['발행상태'], ['정산상태'], ['상태']]
        if field_key == 'contract_type':
            return [['계정코드'], ['서비스유형'], ['계약유형'], ['정산유형'], ['유형']]
        if field_key == 'amount':
            return [
                ['낙찰금액'],
                ['정산금액'],
                ['계약금액'],
                ['청구계획금액'],
                ['기성계획금액'],
                ['PO금액'],
                ['PR금액'],
                ['총금액'],
                ['금액'],
                ['세액'],
            ]
        if field_key == 'contract_name':
            return [['계약명'], ['유지관리명'], ['관리명'], ['적요'], ['세부항목'], ['공사명'], ['입찰명'], ['작업명'], ['업무명'], ['대상명'], ['시설명'], ['WBS명']]
        return []

    def get_similarity_terms_for_field(self, field_key):
        if not field_key:
            return ([], [])
        if field_key == 'po_no':
            return (
                ['PO번호', 'PO', '계약번호', '공사번호', 'WBS', '번호'],
                ['계약명', '유지관리명', '적요', '세부항목', '위험', '여부', '금액', '일자', '상태']
            )
        if field_key == 'pr_no':
            return (
                ['PR번호', 'PR', '계약번호', '공사번호', '번호'],
                ['계약명', '유지관리명', '적요', '위험', '여부', '금액', '일자', '상태']
            )
        if field_key == 'issue_type':
            return (
                ['발행구분', '발행상태', '정산상태', '상태', '구분'],
                ['계약명', '유지관리명', '적요', '번호', '금액', '일자', '위험']
            )
        if field_key == 'contract_type':
            return (
                ['계정코드', '서비스유형', '계약유형', '정산유형', '유형', '계정', '서비스'],
                ['계약명', '유지관리명', '적요', '번호', '금액', '일자', '위험', '여부']
            )
        if field_key == 'amount':
            return (
                ['낙찰금액', '정산금액', '계약금액', '청구계획금액', '기성계획금액', 'PO금액', 'PR금액', '총금액', '금액', '세액', '총액', '비용', '단가'],
                ['WBS명', '계약명', '유지관리명', '관리명', '적요', '세부항목', '공사명', '작업명', '업무명', '번호', '발행', '상태', '유형', '계정', '담당', '위험', '여부', '일자']
            )
        return (
            ['계약명', '유지관리명', '유지관리', '관리명', '적요', '세부항목', '공사명', '작업명', '업무명', '명칭', '대상명', '시설명', 'WBS명'],
            ['위험', '여부', '상태', '발행', '유형', '계정', '금액', '일자', '번호', '담당', '부서', '연락처', '비용', '세액', '승인']
        )

    def score_header_similarity(self, title, header):
        clean_title = self.clean_basis_title(title)
        clean_header = self.clean_basis_title(header)
        normalized_title = normalize_text(clean_title)
        normalized_header = normalize_text(clean_header)
        if not normalized_title or not normalized_header:
            return -1000

        field_key = self.infer_field_key_from_title(clean_title)
        if not field_key:
            return -1000
        if normalized_header == normalized_title:
            return 1000

        score = int(difflib.SequenceMatcher(None, normalized_title, normalized_header).ratio() * 100)
        if normalized_title in normalized_header or normalized_header in normalized_title:
            score += 220

        positive_terms, negative_terms = self.get_similarity_terms_for_field(field_key)
        for term in positive_terms:
            normalized_term = normalize_text(term)
            if normalized_term and normalized_term in normalized_header:
                score += 60 + min(len(normalized_term), 8) * 4
                if normalized_term in normalized_title or normalized_title in normalized_term:
                    score += 80

        for term in negative_terms:
            normalized_term = normalize_text(term)
            if normalized_term and normalized_term in normalized_header:
                score -= 85

        if field_key == 'contract_name':
            if normalized_header.endswith('명') or '명칭' in clean_header:
                score += 35
            if any(term in clean_header for term in ['유지관리', '작업', '공사', '세부항목', '적요']):
                score += 70
            if any(term in clean_header for term in ['위험', '여부']):
                score -= 160

        if field_key == 'amount':
            if any(term in clean_header for term in ['금액', '총액', '비용', '세액', '단가']):
                score += 120
            else:
                score -= 260
            if any(term in clean_header for term in ['WBS명', '계약명', '공사명', '적요', '세부항목', '유지관리명']):
                score -= 220

        if field_key in ('po_no', 'pr_no') and '번호' in clean_header:
            score += 45

        return score

    def infer_col_by_similarity(self, headers, title, fallback=None):
        field_key = self.infer_field_key_from_title(title)
        if not field_key:
            return fallback
        best_col = fallback
        best_score = -1000
        for col, header in headers.items():
            score = self.score_header_similarity(title, header)
            if score > best_score:
                best_col = col
                best_score = score
        minimum_score = 95 if field_key == 'contract_name' else 80
        if best_score >= minimum_score:
            return best_col
        return fallback

    def infer_col_for_title(self, headers, title, fallback):
        field_key = self.infer_field_key_from_title(title)
        if not field_key:
            return fallback
        exact_col = self.find_header_col_by_title(headers, title)
        if exact_col:
            return exact_col
        keyword_col = self.infer_col_by_keywords(headers, self.get_keyword_groups_for_title(title), None)
        if keyword_col:
            return keyword_col
        return self.infer_col_by_similarity(headers, title, fallback)

    def apply_3way_mapping_group(self, field_key, field_label, ref_col, pur_col, sal_col, mode='수동-기준연계'):
        ref_hdr_name = self.ref_headers.get(ref_col, f'Col {ref_col}')
        pur_hdr_name = self.pur_headers.get(pur_col, f'Col {pur_col}')
        sal_hdr_name = self.sal_headers.get(sal_col, f'Col {sal_col}')
        updated_mapping = {
            'field_key': field_key,
            'field_label': field_label,
            'ref_col': ref_col,
            'ref_header': f'{ref_hdr_name} (Col {ref_col} / {col_num_to_letter(ref_col)}열)',
            'pur_col': pur_col,
            'pur_header': f'{pur_hdr_name} (Col {pur_col} / {col_num_to_letter(pur_col)}열)',
            'sal_col': sal_col,
            'sal_header': f'{sal_hdr_name} (Col {sal_col} / {col_num_to_letter(sal_col)}열)',
            'mode': mode
        }
        for idx, mapping in enumerate(self.header_mappings):
            if mapping.get('field_key') == field_key:
                self.header_mappings[idx] = updated_mapping
                return 'updated'
        self.header_mappings.append(updated_mapping)
        return 'added'

    def auto_apply_ref_basis_title(self, header_title, ref_col, pur_col_override=None, sal_col_override=None, extra_col_overrides=None):
        title = self.clean_basis_title(header_title) or self.clean_basis_title(self.ref_headers.get(ref_col, f'Col {ref_col}'))
        self.ref_headers[int(ref_col)] = title
        inferred_key = self.infer_field_key_from_title(title)

        pur_fallbacks = {
            'contract_name': 28,
            'po_no': 7,
            'pr_no': 8,
            'issue_type': 19,
            'contract_type': 10,
            'amount': 14,
        }
        sal_fallbacks = {
            'contract_name': 11,
            'po_no': 27,
            'pr_no': 28,
            'issue_type': 22,
            'contract_type': 10,
            'amount': 16,
        }
        current_primary = next((m for m in self.header_mappings if m.get('field_key') == 'contract_name'), {})

        def current_or_default(col_key, default_col):
            try:
                return int(current_primary.get(col_key) or default_col)
            except Exception:
                return int(default_col)

        if inferred_key:
            pur_col = int(
                pur_col_override
                or self.infer_col_for_title(self.pur_headers, title, None)
                or current_or_default('pur_col', pur_fallbacks.get(inferred_key, 28))
            )
            sal_col = int(
                sal_col_override
                or self.infer_col_for_title(self.sal_headers, title, None)
                or current_or_default('sal_col', sal_fallbacks.get(inferred_key, 11))
            )
        else:
            pur_col = int(pur_col_override or current_or_default('pur_col', 28))
            sal_col = int(sal_col_override or current_or_default('sal_col', 11))

        updated = []
        primary_label = f'{title} 기준 매칭' if inferred_key != 'contract_name' else '계약명 / 적요'
        self.apply_3way_mapping_group('contract_name', primary_label, int(ref_col), int(pur_col), int(sal_col))
        updated.append('contract_name')

        linked_labels = {
            'po_no': 'PO번호',
            'pr_no': 'PR번호 / 담당자',
            'issue_type': '발행구분',
            'contract_type': '계약방식 / 정산유형',
            'amount': '금액',
        }
        if inferred_key != 'contract_name':
            if inferred_key:
                self.apply_3way_mapping_group(inferred_key, linked_labels.get(inferred_key, title), int(ref_col), int(pur_col), int(sal_col))
                updated.append(inferred_key)

        extra_col_overrides = extra_col_overrides or {}
        for party in self.extra_parties:
            headers = party.get('headers', {})
            if not headers and party['id'] not in extra_col_overrides:
                continue
            if party['id'] in extra_col_overrides:
                extra_col = int(extra_col_overrides[party['id']])
            elif inferred_key:
                inferred_extra_col = self.infer_col_for_title(headers, title, None)
                if not inferred_extra_col:
                    continue
                extra_col = int(inferred_extra_col)
            else:
                continue
            party['mapping']['contract_col'] = int(extra_col)
            if inferred_key == 'po_no':
                party['mapping']['po_col'] = int(extra_col)
            elif inferred_key == 'issue_type':
                party['mapping']['issue_col'] = int(extra_col)
            elif inferred_key == 'contract_type':
                party['mapping']['type_col'] = int(extra_col)
            elif inferred_key == 'amount':
                party['mapping']['amount_col'] = int(extra_col)
            party['mapping_manual'] = True

        return {
            'title': title,
            'field_key': inferred_key or 'custom_basis',
            'ref_col': int(ref_col),
            'pur_col': int(pur_col),
            'sal_col': int(sal_col),
            'updated_groups': updated,
        }

    def infer_party_column_mapping(self, headers, fallback_contract_col=11):
        return {
            'contract_col': self.infer_col_by_keywords(headers, [['계약명'], ['적요'], ['세부항목'], ['WBS명']], fallback_contract_col),
            'po_col': self.infer_col_by_keywords(headers, [['PO번호'], ['계약번호'], ['공사번호'], ['WBS']], 27),
            'issue_col': self.infer_col_by_keywords(headers, [['발행구분'], ['발행상태'], ['정산상태']], 22),
            'type_col': self.infer_col_by_keywords(headers, [['계정코드'], ['서비스유형'], ['정산유형']], 10),
        }

    def add_extra_party_file(self, file_path, label=None):
        if not file_path:
            return None
        if self.is_file_already_registered(file_path):
            self.log('WARN', f'n자 파일 중복 등록 건너뜀: {os.path.basename(file_path)}')
            return None
        party_id = f'extra_{self.next_extra_party_no}_{int(time.time() * 1000)}'
        party_label = label or f'n자_{self.next_extra_party_no:02d}'
        self.next_extra_party_no += 1
        party = {
            'id': party_id,
            'label': party_label,
            'file_path_var': tk.StringVar(value=file_path),
            'headers': {},
            'mapping': {
                'contract_col': 11,
                'po_col': 27,
                'issue_col': 22,
                'type_col': 10,
            },
            'mapping_manual': False,
        }
        self.extra_parties.append(party)
        self.render_extra_party_inputs()
        self.log('INFO', f'{party_label} 파일 추가됨: {os.path.basename(file_path)}')
        self.auto_scan_headers_from_file(file_path, f'extra:{party_id}')
        return party

    def render_extra_party_inputs(self):
        if not hasattr(self, 'extra_party_frame'):
            return
        for child in self.extra_party_frame.winfo_children():
            child.destroy()
        visible_rows = min(max(len(self.extra_parties), 0), 4)
        canvas_height = max(1, visible_rows * 36 + (4 if visible_rows else 0))
        if hasattr(self, 'extra_party_canvas'):
            self.extra_party_canvas.configure(height=canvas_height)
        if hasattr(self, 'extra_party_vscroll'):
            if len(self.extra_parties) > visible_rows and visible_rows:
                self.extra_party_vscroll.grid()
            else:
                self.extra_party_vscroll.grid_remove()
        for idx, party in enumerate(self.extra_parties, start=1):
            row = ttk.Frame(self.extra_party_frame, style='Card.TFrame')
            row.pack(fill='x', pady=(2, 4))
            ttk.Label(row, text=f'• {party["label"]}:', font=('Segoe UI', 9, 'bold'), width=12).pack(side='left')
            ttk.Entry(row, textvariable=party['file_path_var'], font=('Consolas', 9)).pack(side='left', fill='x', expand=True, padx=(0, 8))

            def remove_party(p=party):
                self.extra_parties = [x for x in self.extra_parties if x['id'] != p['id']]
                self.render_extra_party_inputs()
                self.update_files_status_label()

            ttk.Button(row, text='헤더 매칭...', command=lambda p=party: self.show_extra_party_mapping_dialog(p)).pack(side='right', padx=(0, 4))
            ttk.Button(row, text='삭제', command=remove_party).pack(side='right')
        if hasattr(self, 'extra_party_canvas'):
            self.extra_party_canvas.update_idletasks()
            self.extra_party_canvas.configure(scrollregion=self.extra_party_canvas.bbox('all'))

    def show_extra_party_mapping_dialog(self, party):
        if party['file_path_var'].get():
            self.auto_scan_headers_from_file(party['file_path_var'].get(), f'extra:{party["id"]}')

        popup = tk.Toplevel(self.root)
        popup.title(f'⚙️ {party["label"]} 헤더 매칭 설정')
        popup.geometry('880x560')
        popup.grab_set()

        f = ttk.Frame(popup, padding=16)
        f.pack(fill='both', expand=True)
        ttk.Label(f, text=f'{party["label"]} 헤더 매칭 설정', font=('Segoe UI', 11, 'bold'), foreground=self.COLOR_PRIMARY).pack(anchor='w', pady=(0, 8))
        ttk.Label(f, text='기본 3자 기준 헤더를 발주 내역, 매입, 매출로 분리해 확인한 뒤 n자 대상 컬럼을 매칭합니다.', font=('Segoe UI', 9), foreground='#4A5568').pack(anchor='w', pady=(0, 10))

        def fmt_header(headers, col):
            col = int(col or 1)
            letter = col_num_to_letter(col)
            return f'{letter} (Col {col}) - {headers.get(col, f"Col {col}")}'

        base_box = ttk.LabelFrame(f, text='기본 3자 기준 헤더 그룹', padding=8)
        base_box.pack(fill='x', pady=(0, 12))

        header_row = ttk.Frame(base_box)
        header_row.pack(fill='x', pady=(0, 4))
        for title, width in [('매칭 항목', 18), ('발주 내역', 28), ('매입', 28), ('매출', 28)]:
            ttk.Label(header_row, text=title, font=('Segoe UI', 9, 'bold'), width=width, anchor='w').pack(side='left', padx=(0, 6))

        base_fields = [
            ('contract_name', '계약명/적요'),
            ('po_no', 'PO/계약/공사번호'),
            ('issue_type', '발행구분/상태'),
            ('contract_type', '유형/계정/서비스'),
        ]
        for field_key, label in base_fields:
            m = next((x for x in self.header_mappings if x.get('field_key') == field_key), None)
            if not m:
                continue
            row = ttk.Frame(base_box)
            row.pack(fill='x', pady=2)
            ttk.Label(row, text=label, font=('Segoe UI', 9, 'bold'), width=18, anchor='w').pack(side='left', padx=(0, 6))
            ttk.Label(row, text=fmt_header(self.ref_headers, m.get('ref_col')), font=('Consolas', 9), width=28, anchor='w').pack(side='left', padx=(0, 6))
            ttk.Label(row, text=fmt_header(self.pur_headers, m.get('pur_col')), font=('Consolas', 9), width=28, anchor='w').pack(side='left', padx=(0, 6))
            ttk.Label(row, text=fmt_header(self.sal_headers, m.get('sal_col')), font=('Consolas', 9), width=28, anchor='w').pack(side='left')

        target_box = ttk.LabelFrame(f, text=f'{party["label"]} 대상 헤더 매칭', padding=8)
        target_box.pack(fill='both', expand=True, pady=(0, 12))
        HoverTooltip(target_box, (
            "n자 대상 파일의 컬럼을 기본 3자 실행 그룹에 맞춰 연결합니다.\n\n"
            "워크플로:\n"
            "1. 기본 3자 기준 헤더 그룹에서 발주/매입/매출의 현재 기준을 확인합니다.\n"
            "2. 아래 n자 컬럼을 같은 의미의 헤더로 선택합니다.\n"
            "3. 예를 들어 주 기준을 PO번호로 운용한다면 n자 계약명/적요 컬럼도 실제 PO 또는 계약번호 계열 컬럼으로 맞춥니다.\n"
            "4. 저장하면 최종 보고서 우측 n자 그룹에도 해당 기준으로 표시됩니다."
        ))

        max_extra_col = max(52, max(party.get('headers', {}).keys()) if party.get('headers') else 52)
        options = [self.get_extra_col_option_str(party, i) for i in range(1, max_extra_col + 1)]
        mapping = party['mapping']

        fields = [
            ('contract_col', f'{party["label"]} 계약명/적요 컬럼', mapping.get('contract_col', 11)),
            ('po_col', f'{party["label"]} PO/계약/공사번호 컬럼', mapping.get('po_col', 27)),
            ('issue_col', f'{party["label"]} 발행구분/상태 컬럼', mapping.get('issue_col', 22)),
            ('type_col', f'{party["label"]} 유형/계정/서비스 컬럼', mapping.get('type_col', 10)),
        ]
        vars_by_key = {}
        for key, label, default_col in fields:
            row = ttk.Frame(target_box)
            row.pack(fill='x', pady=(0, 8))
            ttk.Label(row, text=label, font=('Segoe UI', 9, 'bold'), width=30, anchor='w').pack(side='left')
            var = tk.StringVar(value=self.get_extra_col_option_str(party, default_col))
            vars_by_key[key] = var
            cmb_extra = ttk.Combobox(row, textvariable=var, values=options, font=('Consolas', 9), width=56)
            cmb_extra.pack(side='left', fill='x', expand=True)
            HoverTooltip(cmb_extra, (
                f"{label}에 대응되는 {party['label']} 파일의 실제 헤더와 Col 좌표를 선택합니다.\n\n"
                "기본 3자에서 선택한 논리 그룹과 같은 의미의 컬럼을 골라야 최종 보고서의 n자 비교 결과가 일관되게 표시됩니다."
            ))

        def save_mapping():
            for key, var in vars_by_key.items():
                party['mapping'][key] = parse_col_num_from_str(var.get())
            party['mapping_manual'] = True
            self.log('SUCCESS', f'{party["label"]} 헤더 매칭 저장 완료: 계약 Col {party["mapping"]["contract_col"]}, 발행구분 Col {party["mapping"]["issue_col"]}')
            popup.destroy()

        btns = ttk.Frame(f)
        btns.pack(fill='x', side='bottom')
        ttk.Button(btns, text='저장', style='Primary.TButton', command=save_mapping).pack(side='right', padx=(6, 0))
        ttk.Button(btns, text='취소', command=popup.destroy).pack(side='right')

    def update_files_status_label(self):
        extra_text = f' | n자: {len(self.extra_parties)}개' if self.extra_parties else ''
        self.lbl_files_status.config(
            text=f'✅ 선택 완료 | 매입: {os.path.basename(self.purchase_file_path.get()) if self.purchase_file_path.get() else "미선택"} | 매출: {os.path.basename(self.sales_file_path.get()) if self.sales_file_path.get() else "미선택"}{extra_text}',
            foreground='#2F855A'
        )

    def get_default_mappings(self):
        """
        3자간 (발주내역 ↔ 매입월정산 ↔ 매출월정산) 헤더 및 컬럼 매핑 표준 정의
        """
        return [
            {
                'field_key': 'contract_name',
                'field_label': '계약명 / 적요',
                'ref_col': 24,
                'ref_header': '계약명 (Col 24 / X열)',
                'pur_col': 28,
                'pur_header': '계약명 (Col 28 / AB열)',
                'sal_col': 11,
                'sal_header': '세부항목 (Col 11 / K열)',
                'mode': '자동'
            },
            {
                'field_key': 'po_no',
                'field_label': 'PO번호',
                'ref_col': 27,
                'ref_header': 'PO번호 (Col 27 / AA열)',
                'pur_col': 7,
                'pur_header': '기성정보/업체 (Col 7 / G열)',
                'sal_col': 27,
                'sal_header': '계약명 (Col 27 / AA열)',
                'mode': '자동'
            },
            {
                'field_key': 'issue_type',
                'field_label': '발행구분',
                'ref_col': 19,
                'ref_header': '공종(소) (Col 19 / S열)',
                'pur_col': 19,
                'pur_header': '발행구분 (Col 19 / S열)',
                'sal_col': 22,
                'sal_header': '발행구분 (Col 22 / V열)',
                'mode': '자동'
            },
            {
                'field_key': 'contract_type',
                'field_label': '계약방식 / 정산유형',
                'ref_col': 10,
                'ref_header': 'S&I 센터 (Col 10 / J열)',
                'pur_col': 10,
                'pur_header': '계정코드 (Col 10 / J열)',
                'sal_col': 10,
                'sal_header': '서비스유형 (Col 10 / J열)',
                'mode': '자동'
            },
            {
                'field_key': 'pr_no',
                'field_label': 'PR번호 / 담당자',
                'ref_col': 28,
                'ref_header': 'PR번호 (Col 28 / AB열)',
                'pur_col': 8,
                'pur_header': '업체명 (Col 8 / H열)',
                'sal_col': 28,
                'sal_header': '계약담당자 (Col 28 / AB열)',
                'mode': '자동'
            }
        ]

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
        self.style.configure('Header.TLabel', background=self.COLOR_PRIMARY, foreground='#FFFFFF', font=('Segoe UI', 15, 'bold'))
        self.style.configure('SubHeader.TLabel', background=self.COLOR_PRIMARY, foreground='#CBD5E0', font=('Segoe UI', 9))
        self.style.configure('Title.TLabel', background=self.COLOR_CARD, foreground=self.COLOR_PRIMARY, font=('Segoe UI', 11, 'bold'))
        self.style.configure('Status.TLabel', background=self.COLOR_CARD, foreground='#4A5568', font=('Segoe UI', 9))
        
        self.style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'), background=self.COLOR_SECONDARY, foreground='#FFFFFF')
        self.style.map('Primary.TButton', background=[('active', '#2C5282')])

        self.style.configure('Accent.TButton', font=('Segoe UI', 11, 'bold'), background=self.COLOR_ACCENT, foreground='#FFFFFF')
        self.style.map('Accent.TButton', background=[('active', '#B7791F')])

        self.style.configure('Multi.TButton', font=('Segoe UI', 10, 'bold'), background='#805AD5', foreground='#FFFFFF')
        self.style.map('Multi.TButton', background=[('active', '#6B46C1')])

        self.style.configure('Config.TButton', font=('Segoe UI', 10, 'bold'), background='#2D3748', foreground='#FFFFFF')
        self.style.map('Config.TButton', background=[('active', '#1A202C')])

        self.style.configure('Reset.TButton', font=('Segoe UI', 9, 'bold'), background='#E53E3E', foreground='#FFFFFF')
        self.style.map('Reset.TButton', background=[('active', '#C53030')])

        self.style.configure('Treeview', font=('Consolas', 9), rowheight=28)
        self.style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'), background='#E2E8F0', foreground='#1A365D')

    def build_ui(self):
        # 상단 타이틀
        header_frame = tk.Frame(self.root, bg=self.COLOR_PRIMARY, height=75)
        header_frame.pack(fill='x', side='top')
        
        lbl_title = ttk.Label(header_frame, text='B2B 3자간(발주내역 ↔ 매입월정산 ↔ 매출월정산) 정밀 교차 검증 솔루션 v7.2', style='Header.TLabel')
        lbl_title.pack(anchor='w', padx=20, pady=(10, 2))
        
        lbl_sub = ttk.Label(header_frame, text='⚡ 엑셀 파일 업로드 시 자동 구조 파악 실시간 동적 연동 & 사용자 지정 열 틀고정(기본 2번열) 및 교차 색상(Zebra Striping)', style='SubHeader.TLabel')
        lbl_sub.pack(anchor='w', padx=20, pady=(0, 8))

        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill='both', expand=True)

        # 카드 1: 1단계 기준 파일 선택 및 마우스 범위 지정
        card1 = ttk.Frame(main_container, style='Card.TFrame', padding=12)
        card1.pack(fill='x', pady=(0, 10))

        f1_hdr = ttk.Frame(card1, style='Card.TFrame')
        f1_hdr.pack(fill='x', pady=(0, 6))

        ttk.Label(f1_hdr, text='📌 1단계: 기준 파일 선택 (발주내역_(기숙사 및 사택).xlsb)', style='Title.TLabel').pack(side='left')

        # 3자간 전용 헤더 매핑 설정 버튼
        btn_config = ttk.Button(f1_hdr, text='⚙️ 3자간 헤더 매핑 설정 (발주-매입-매출)...', style='Config.TButton', command=self.show_header_mapping_config_dialog)
        btn_config.pack(side='right')
        HoverTooltip(btn_config, (
            "[3자 교차 검증] 헤더 매핑 통합 설정 워크플로\n\n"
            "1. 기준 발주내역, 매입월정산, 매출월정산 파일을 먼저 선택하고 자동 헤더 스캔 완료 로그를 확인합니다.\n"
            "2. 기본 3자 그룹은 실행부와 직접 연결됩니다.\n"
            "   - [계약명 / 적요]: 주 매칭 기준입니다. 계약명 대신 PO번호를 주 기준으로 쓰려면 이 행의 발주/매입/매출 컬럼을 각각 PO 계열 컬럼으로 바꿉니다.\n"
            "   - [PO번호]: PO번호 검증/표시 기준입니다. PO 대조값도 함께 맞추려면 이 행도 확인합니다.\n"
            "   - [발행구분], [계약방식]: 상태/유형 판정 기준입니다.\n"
            "3. n자 파일을 추가한 경우 n자 헤더 매칭 창에서 같은 논리 그룹에 맞는 n자 컬럼을 선택합니다.\n"
            "4. 저장 미리보기의 헤더명과 Col 좌표가 맞는지 확인한 뒤 저장합니다.\n\n"
            "범위 선택은 행 범위를 확정하는 단계이고, 실제 교차 검증 기준 컬럼은 이 통합 설정 화면의 논리 그룹이 결정합니다."
        ))

        f1_input = ttk.Frame(card1, style='Card.TFrame')
        f1_input.pack(fill='x')

        self.txt_ref = ttk.Entry(f1_input, textvariable=self.ref_file_path, font=('Consolas', 9))
        self.txt_ref.pack(side='left', fill='x', expand=True, padx=(0, 8))

        btn_browse_ref = ttk.Button(f1_input, text='📂 기준 파일 찾기...', command=self.on_select_ref_file)
        btn_browse_ref.pack(side='right')

        f1_interactive = ttk.Frame(card1, style='Card.TFrame')
        f1_interactive.pack(fill='x', pady=(8, 0))

        self.btn_drag_select = ttk.Button(f1_interactive, text='🖥️ 엑셀 창 열기 & 마우스 범위 확정', style='Primary.TButton', command=self.on_interactive_drag_select)
        self.btn_drag_select.pack(side='left', padx=(0, 8))
        HoverTooltip(self.btn_drag_select, (
            "[3자 교차 검증] 발주내역 대상 범위 선택 워크플로\n\n"
            "1. 기준 발주내역 파일을 선택합니다.\n"
            "2. 이 버튼을 눌러 Excel을 전면으로 활성화합니다.\n"
            "3. 계약명 X열에 한정하지 말고, 이번 검증의 대상 기준으로 삼을 열의 행 범위를 드래그합니다.\n"
            "   예: 계약명 X4942:X4963, PO번호 AA4942:AA4963.\n"
            "4. 팝업에서 [영역 읽기]를 누르고 감지된 범위를 확인합니다.\n"
            "5. 계약명이 아닌 PO번호/PR번호 등을 주 기준으로 쓸 때는 반드시 [3자간 헤더 매핑 설정]에서 [계약명 / 적요] 실행 그룹의 발주/매입/매출 컬럼도 같은 기준으로 바꿉니다.\n\n"
            "주의: 여기서 저장되는 것은 기준 발주내역의 행 범위입니다. 실제 매칭 기준 컬럼은 통합 헤더 매핑 설정이 결정합니다."
        ))

        self.btn_reset_range = ttk.Button(f1_interactive, text='🔄 범위 선택 초기화', style='Reset.TButton', command=self.on_reset_range_selection)
        self.btn_reset_range.pack(side='left', padx=(0, 10))

        self.lbl_target_range = ttk.Label(f1_interactive, text='지정된 범위: 미선택 (엑셀 창에서 마우스로 범위를 선택하세요)', style='Status.TLabel')
        self.lbl_target_range.pack(side='left', fill='x', expand=True)

        # 카드 2: 2단계 매입 & 매출 정산 파일 선택
        card2 = ttk.Frame(main_container, style='Card.TFrame', padding=12)
        card2.pack(fill='x', pady=(0, 10))

        f2_title = ttk.Frame(card2, style='Card.TFrame')
        f2_title.pack(fill='x', pady=(0, 6))
        ttk.Label(f2_title, text='🎯 2단계: 정산 대상 파일 선택 (매입월정산.xlsx & 매출월정산.xlsx)', style='Title.TLabel').pack(side='left')
        
        btn_multi = ttk.Button(f2_title, text='📂 [추천] 매입/매출 + n자 파일 다중 선택...', style='Multi.TButton', command=self.on_select_multi_files)
        btn_multi.pack(side='right')
        HoverTooltip(btn_multi, (
            "정산 대상 파일을 한 번에 여러 개 선택합니다.\n\n"
            "분류 규칙:\n"
            "1. 선택 묶음 안의 첫 번째 매입 파일은 기본 [매입월정산]으로 배정됩니다.\n"
            "2. 선택 묶음 안의 첫 번째 매출 파일은 기본 [매출월정산]으로 배정됩니다.\n"
            "3. 추가 매입/매출 파일과 기타 파일은 모두 n자_01, n자_02 순서로 누적됩니다.\n"
            "4. 파일명에 매입/매출 키워드가 없으면 매입/매출 입력창은 비워두고 n자로만 등록됩니다.\n"
            "5. 같은 파일을 다시 선택하면 중복 등록하지 않습니다."
        ))

        f2_pur = ttk.Frame(card2, style='Card.TFrame')
        f2_pur.pack(fill='x', pady=(2, 4))
        ttk.Label(f2_pur, text='• 매입월정산:', font=('Segoe UI', 9, 'bold'), width=12).pack(side='left')
        self.txt_pur = ttk.Entry(f2_pur, textvariable=self.purchase_file_path, font=('Consolas', 9))
        self.txt_pur.pack(side='left', fill='x', expand=True, padx=(0, 8))
        btn_browse_pur = ttk.Button(f2_pur, text='📂 개별 선택...', command=self.on_select_purchase_file)
        btn_browse_pur.pack(side='right')

        f2_sal = ttk.Frame(card2, style='Card.TFrame')
        f2_sal.pack(fill='x', pady=(2, 4))
        ttk.Label(f2_sal, text='• 매출월정산:', font=('Segoe UI', 9, 'bold'), width=12).pack(side='left')
        self.txt_sal = ttk.Entry(f2_sal, textvariable=self.sales_file_path, font=('Consolas', 9))
        self.txt_sal.pack(side='left', fill='x', expand=True, padx=(0, 8))
        btn_browse_sal = ttk.Button(f2_sal, text='📂 개별 선택...', command=self.on_select_sales_file)
        btn_browse_sal.pack(side='right')

        self.extra_party_scroll_outer = ttk.Frame(card2, style='Card.TFrame')
        self.extra_party_scroll_outer.pack(fill='x', pady=(2, 4))
        self.extra_party_canvas = tk.Canvas(self.extra_party_scroll_outer, height=1, highlightthickness=0, bg=self.COLOR_CARD)
        self.extra_party_vscroll = ttk.Scrollbar(self.extra_party_scroll_outer, orient='vertical', command=self.extra_party_canvas.yview)
        self.extra_party_canvas.configure(yscrollcommand=self.extra_party_vscroll.set)
        self.extra_party_canvas.grid(row=0, column=0, sticky='ew')
        self.extra_party_vscroll.grid(row=0, column=1, sticky='ns')
        self.extra_party_vscroll.grid_remove()
        self.extra_party_scroll_outer.columnconfigure(0, weight=1)

        self.extra_party_frame = ttk.Frame(self.extra_party_canvas, style='Card.TFrame')
        self.extra_party_canvas_window = self.extra_party_canvas.create_window((0, 0), window=self.extra_party_frame, anchor='nw')

        def _sync_extra_party_scrollregion(event=None):
            self.extra_party_canvas.configure(scrollregion=self.extra_party_canvas.bbox('all'))
            self.extra_party_canvas.itemconfigure(self.extra_party_canvas_window, width=self.extra_party_canvas.winfo_width())

        self.extra_party_frame.bind('<Configure>', _sync_extra_party_scrollregion)
        self.extra_party_canvas.bind('<Configure>', _sync_extra_party_scrollregion)

        self.lbl_files_status = ttk.Label(card2, text='ⓘ 파일명에 매입/매출 키워드가 있는 첫 파일만 기본 슬롯에 배정되고, 나머지는 n자로 누적됩니다.', style='Status.TLabel')
        self.lbl_files_status.pack(anchor='w', pady=(4, 0))

        # 카드 3: 3단계 실행 및 실시간 진행 현황
        card3 = ttk.Frame(main_container, style='Card.TFrame', padding=12)
        card3.pack(fill='both', expand=True, pady=(0, 5))

        f3_top = ttk.Frame(card3, style='Card.TFrame')
        f3_top.pack(fill='x', pady=(0, 4))

        ttk.Label(f3_top, text='🚀 3단계: 3자 정산 교정 실행 및 실시간 진행 현황', style='Title.TLabel').pack(side='left')
        
        self.btn_run = ttk.Button(f3_top, text='▶️ 3자 교차 검증 및 데이터 자동 정제 실행', style='Accent.TButton', command=self.on_run_reconciliation)
        self.btn_run.pack(side='right')

        self.lbl_mapping_status = ttk.Label(card3, text='ⓘ 준비: 버튼을 누르면 3자 교차 검증 및 정제 분석이 시작됩니다.', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0', background=self.COLOR_CARD)
        self.lbl_mapping_status.pack(anchor='w', pady=(0, 4))

        self.progress = ttk.Progressbar(card3, mode='determinate')
        self.progress.pack(fill='x', pady=(0, 6))

        self.log_area = scrolledtext.ScrolledText(card3, height=9, font=('Consolas', 9), bg='#1E1E1E', fg='#D4D4D4', insertbackground='#FFFFFF')
        self.log_area.pack(fill='both', expand=True)

        self.log_area.tag_config('INFO', foreground='#4EC9B0')
        self.log_area.tag_config('SUCCESS', foreground='#6A9955', font=('Consolas', 9, 'bold'))
        self.log_area.tag_config('WARN', foreground='#CE9178')
        self.log_area.tag_config('ERROR', foreground='#F44747', font=('Consolas', 9, 'bold'))

        self.log('INFO', '시스템이 준비되었습니다. 3자간 헤더 매핑 설정을 확인 후 기준 파일과 매입/매출 파일을 순서대로 선택하세요.')
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

    def update_dynamic_ui_state(self):
        is_busy = any(self.active_tasks.values())
        if is_busy:
            self.btn_drag_select.config(text='⏳ 백그라운드 분석 중 (예약 가능)')
            self.btn_run.config(text='⏳ 분석 대기 중 (자동 예약)')
        else:
            self.btn_drag_select.config(text='🖥️ 엑셀 창 열기 & 마우스 범위 확정')
            self.btn_run.config(text='▶️ 3자 교차 검증 및 데이터 자동 정제 실행')

    def on_select_ref_file(self):
        file_path = filedialog.askopenfilename(
            title='기준 발주내역 엑셀 파일 선택',
            filetypes=[('Excel Files', '*.xlsb;*.xlsx;*.xls'), ('All Files', '*.*')]
        )
        if not file_path: return
        self.ref_file_path.set(file_path)
        self.log('INFO', f'기준 파일 선택됨: {os.path.basename(file_path)}')
        self.auto_scan_headers_from_file(file_path, 'ref')

    def on_select_multi_files(self):
        file_paths = filedialog.askopenfilenames(
            title='매입월정산, 매출월정산 및 n자 엑셀 파일 한꺼번에 선택',
            filetypes=[('Excel Files', '*.xlsx;*.xlsb;*.xls'), ('All Files', '*.*')]
        )
        if not file_paths: return

        pur_path, sal_path, extra_paths = classify_reconciliation_file_selection(
            file_paths,
            current_purchase=self.purchase_file_path.get(),
            current_sales=self.sales_file_path.get(),
        )

        if pur_path:
            self.purchase_file_path.set(pur_path)
            self.log('INFO', f'매입 파일 다중 선택 감지: {os.path.basename(pur_path)}')
            self.auto_scan_headers_from_file(pur_path, 'pur')
        if sal_path:
            self.sales_file_path.set(sal_path)
            self.log('INFO', f'매출 파일 다중 선택 감지: {os.path.basename(sal_path)}')
            self.auto_scan_headers_from_file(sal_path, 'sal')

        added_extra_count = 0
        for extra_path in extra_paths:
            if self.add_extra_party_file(extra_path):
                added_extra_count += 1

        self.log(
            'SUCCESS',
            f'다중 선택 분류 완료: 기본 매입 {1 if pur_path else 0}개, 기본 매출 {1 if sal_path else 0}개, 신규 n자 {added_extra_count}개 / 전체 n자 {len(self.extra_parties)}개'
        )

        self.update_files_status_label()

    def on_select_purchase_file(self):
        file_path = filedialog.askopenfilename(
            title='매입월정산 엑셀 파일 개별 선택',
            filetypes=[('Excel Files', '*.xlsx;*.xlsb;*.xls'), ('All Files', '*.*')]
        )
        if not file_path: return
        self.purchase_file_path.set(file_path)
        self.log('INFO', f'매입 파일 개별 선택됨: {os.path.basename(file_path)}')
        self.auto_scan_headers_from_file(file_path, 'pur')
        self.update_files_status_label()

    def on_select_sales_file(self):
        file_path = filedialog.askopenfilename(
            title='매출월정산 엑셀 파일 개별 선택',
            filetypes=[('Excel Files', '*.xlsx;*.xlsb;*.xls'), ('All Files', '*.*')]
        )
        if not file_path: return
        self.sales_file_path.set(file_path)
        self.log('INFO', f'매출 파일 개별 선택됨: {os.path.basename(file_path)}')
        self.auto_scan_headers_from_file(file_path, 'sal')
        self.update_files_status_label()

    # --------------------------------------------------------------------------
    # 3자간 (발주내역 ↔ 매입월정산 ↔ 매출월정산) 전용 헤더 매핑 시각화 다이얼로그
    # --------------------------------------------------------------------------
    def show_header_mapping_config_dialog(self):
        # 다이얼로그 오픈 시 선택된 엑셀 파일이 있다면 헤더 자동 동적 스캔 트리거
        if self.ref_file_path.get():
            self.auto_scan_headers_from_file(self.ref_file_path.get(), 'ref')
        if self.purchase_file_path.get():
            self.auto_scan_headers_from_file(self.purchase_file_path.get(), 'pur')
        if self.sales_file_path.get():
            self.auto_scan_headers_from_file(self.sales_file_path.get(), 'sal')
        for party in self.extra_parties:
            extra_path_var = party.get('file_path_var')
            try:
                extra_path = extra_path_var.get() if hasattr(extra_path_var, 'get') else str(extra_path_var or '')
            except Exception:
                extra_path = ''
            if extra_path:
                self.auto_scan_headers_from_file(extra_path, f'extra:{party["id"]}')

        popup = tk.Toplevel(self.root)
        popup.title('⚙️ 3자간 (발주내역 ↔ 매입월정산 ↔ 매출월정산) 동적 헤더 및 컬럼 매핑 통합 설정')
        popup.geometry('1280x650')
        popup.grab_set()

        main_f = ttk.Frame(popup, padding=12)
        main_f.pack(fill='both', expand=True)

        lbl_config_title = ttk.Label(main_f, text='⚙️ 3자간 (발주내역 ↔ 매입월정산 ↔ 매출월정산) 컬럼 매핑 통합 설정', font=('Segoe UI', 13, 'bold'), foreground=self.COLOR_PRIMARY)
        lbl_config_title.pack(anchor='w', pady=(0, 3))
        HoverTooltip(lbl_config_title, (
            "3자 교차 검증의 실행 기준을 통합 관리하는 화면입니다.\n\n"
            "워크플로:\n"
            "1. 기준/매입/매출 파일 선택 후 자동 헤더 스캔을 완료합니다.\n"
            "2. 범위 팝업에서 발주내역 대상 행 범위를 확정합니다.\n"
            "3. 이 화면에서 실행 그룹별 컬럼을 확인합니다.\n"
            "4. 계약명 대신 PO번호를 주 기준으로 쓰려면 [계약명 / 적요] 행을 편집하여 발주/매입/매출 컬럼을 PO 계열로 맞춥니다.\n"
            "5. PO번호 자체의 검증/표시도 필요하면 [PO번호] 행도 함께 확인합니다.\n"
            "6. n자 파일은 우측 그룹과 n자 헤더 매칭 창에서 같은 논리 그룹으로 확장됩니다."
        ))
        ttk.Label(main_f, text='3개 엑셀 파일(기준 발주내역, 매입월정산, 매출월정산)의 실데이터 컬럼 연계 상태를 확인하고 편집/추가/삭제할 수 있습니다.', font=('Segoe UI', 9), foreground='#4A5568').pack(anchor='w', pady=(0, 8))

        canvas_frame = ttk.Frame(main_f)
        canvas_frame.pack(fill='both', expand=True, pady=(0, 8))

        canvas = tk.Canvas(canvas_frame, bg='#FFFFFF', highlightthickness=1, highlightbackground='#CBD5E0')
        v_scroll = ttk.Scrollbar(canvas_frame, orient='vertical', command=canvas.yview)
        h_scroll = ttk.Scrollbar(canvas_frame, orient='horizontal', command=canvas.xview)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        ROW_H = 42
        HDR_Y = 22
        START_Y = 48

        REF_COL_X = 20
        REF_NAME_X = 90
        ARROW1_X1 = 330
        ARROW1_X2 = 410
        PUR_COL_X = 430
        PUR_NAME_X = 500
        ARROW2_X1 = 740
        ARROW2_X2 = 820
        SAL_COL_X = 840
        SAL_NAME_X = 910
        SAL_GROUP_RIGHT = 1135
        EXTRA_START_X = 1150
        EXTRA_GROUP_W = 285
        EXTRA_COL_OFFSET = 18
        EXTRA_NAME_OFFSET = 86
        MODE_W = 95

        selected_idx = [None]

        def get_canvas_width():
            return EXTRA_START_X + len(self.extra_parties) * EXTRA_GROUP_W + MODE_W + 25

        def get_mode_x():
            return EXTRA_START_X + len(self.extra_parties) * EXTRA_GROUP_W + 15

        def get_extra_col_for_mapping(party, mapping):
            field_key = mapping.get('field_key', '')
            field_label = mapping.get('field_label', '')
            party_mapping = party.get('mapping', {})
            if field_key == 'po_no' or 'PO' in field_label or '번호' in field_label:
                return party_mapping.get('po_col', 27)
            if field_key == 'issue_type' or '발행' in field_label or '상태' in field_label:
                return party_mapping.get('issue_col', 22)
            if field_key == 'contract_type' or '유형' in field_label or '계정' in field_label or '서비스' in field_label:
                return party_mapping.get('type_col', 10)
            return party_mapping.get('contract_col', 11)

        def refresh():
            canvas.delete('all')
            n = len(self.header_mappings)
            total_h = START_Y + n * ROW_H + 30
            canvas_w = get_canvas_width()
            mode_x = get_mode_x()
            canvas.configure(scrollregion=(0, 0, canvas_w, max(total_h, 440)))

            # ── 1구획: 기준 발주내역 (파란색 계열) ──
            canvas.create_rectangle(5, 3, ARROW1_X1 - 15, HDR_Y + 12, fill='#EBF8FF', outline='#BEE3F8')
            canvas.create_text(REF_COL_X, HDR_Y, text='[1] 발주 Col', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#2C5282')
            canvas.create_text(REF_NAME_X, HDR_Y, text='기준 발주내역 컬럼명', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#2C5282')

            # 연결선 1
            canvas.create_text((ARROW1_X1 + ARROW1_X2) // 2, HDR_Y, text='연계', anchor='center', font=('Segoe UI', 9, 'bold'), fill='#718096')

            # ── 2구획: 매입월정산 (주황색 계열) ──
            canvas.create_rectangle(ARROW1_X2 + 5, 3, ARROW2_X1 - 15, HDR_Y + 12, fill='#FFFAF0', outline='#FEEBC8')
            canvas.create_text(PUR_COL_X, HDR_Y, text='[2] 매입 Col', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#C05621')
            canvas.create_text(PUR_NAME_X, HDR_Y, text='매입월정산 대조 컬럼명', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#C05621')

            # 연결선 2
            canvas.create_text((ARROW2_X1 + ARROW2_X2) // 2, HDR_Y, text='연계', anchor='center', font=('Segoe UI', 9, 'bold'), fill='#718096')

            # ── 3구획: 매출월정산 (초록색 계열) ──
            canvas.create_rectangle(ARROW2_X2 + 5, 3, SAL_GROUP_RIGHT, HDR_Y + 12, fill='#F0FFF4', outline='#C6F6D5')
            canvas.create_text(SAL_COL_X, HDR_Y, text='[3] 매출 Col', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#276749')
            canvas.create_text(SAL_NAME_X, HDR_Y, text='매출월정산 대조 컬럼명', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#276749')

            for extra_idx, extra_party in enumerate(self.extra_parties, start=4):
                base_x = EXTRA_START_X + (extra_idx - 4) * EXTRA_GROUP_W
                right_x = base_x + EXTRA_GROUP_W - 20
                canvas.create_text(base_x - 35, HDR_Y, text='연계', anchor='center', font=('Segoe UI', 9, 'bold'), fill='#718096')
                canvas.create_rectangle(base_x, 3, right_x, HDR_Y + 12, fill='#F7FAFC', outline='#D6BCFA')
                canvas.create_text(base_x + EXTRA_COL_OFFSET, HDR_Y, text=f'[{extra_idx}] {extra_party["label"]} Col', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#553C9A')
                canvas.create_text(base_x + EXTRA_NAME_OFFSET, HDR_Y, text=f'{extra_party["label"]} 대조 컬럼명', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#553C9A')

            canvas.create_rectangle(mode_x - 10, 3, canvas_w - 5, HDR_Y + 12, fill='#F7FAFC', outline='#E2E8F0')
            canvas.create_text(mode_x, HDR_Y, text='모드', anchor='w', font=('Segoe UI', 9, 'bold'), fill='#4A5568')

            canvas.create_line(5, HDR_Y + 14, canvas_w - 5, HDR_Y + 14, fill='#A0AEC0', width=1)

            for idx, m in enumerate(self.header_mappings):
                y = START_Y + idx * ROW_H + ROW_H // 2
                y_top = START_Y + idx * ROW_H + 2
                y_bot = y_top + ROW_H - 4

                if selected_idx[0] == idx:
                    canvas.create_rectangle(5, y_top, canvas_w - 5, y_bot, fill='#BEE3F8', outline='#2B6CB0', width=2)
                else:
                    bg = '#FFFFFF' if idx % 2 == 0 else '#F7FAFC'
                    canvas.create_rectangle(5, y_top, canvas_w - 5, y_bot, fill=bg, outline='#EDF2F7')

                mode = m.get('mode', '자동')
                if mode == '자동':
                    line_color = '#38A169'
                    mode_text = '🤖 자동'
                elif mode == '신규':
                    line_color = '#3182CE'
                    mode_text = '➕ 신규'
                else:
                    line_color = '#E53E3E'
                    mode_text = '✏️ 수동'

                ref_let = col_num_to_letter(m['ref_col'])
                pur_let = col_num_to_letter(m['pur_col'])
                sal_let = col_num_to_letter(m['sal_col'])

                ref_hdr_name = self.ref_headers.get(m['ref_col'], f'Col {m["ref_col"]}')
                pur_hdr_name = self.pur_headers.get(m['pur_col'], f'Col {m["pur_col"]}')
                sal_hdr_name = self.sal_headers.get(m['sal_col'], f'Col {m["sal_col"]}')

                # ── [1] 발주내역 컬럼 ──
                canvas.create_text(REF_COL_X, y, text=f'{ref_let:>2} ({m["ref_col"]})', anchor='w', font=('Consolas', 10, 'bold'), fill='#2C5282')
                canvas.create_text(REF_NAME_X, y, text=f'{ref_hdr_name} (Col {m["ref_col"]} / {ref_let}열)', anchor='w', font=('Segoe UI', 9), fill='#1A365D')

                # ── 화살표 1 ──
                canvas.create_oval(ARROW1_X1 - 4, y - 4, ARROW1_X1 + 4, y + 4, fill=line_color, outline=line_color)
                canvas.create_line(ARROW1_X1 + 5, y, ARROW1_X2 - 5, y, fill=line_color, width=2, arrow=tk.LAST, arrowshape=(8, 11, 4))

                # ── [2] 매입월정산 컬럼 ──
                canvas.create_text(PUR_COL_X, y, text=f'{pur_let:>2} ({m["pur_col"]})', anchor='w', font=('Consolas', 10, 'bold'), fill='#C05621')
                canvas.create_text(PUR_NAME_X, y, text=f'{pur_hdr_name} (Col {m["pur_col"]} / {pur_let}열)', anchor='w', font=('Segoe UI', 9), fill='#742A2A')

                # ── 화살표 2 ──
                canvas.create_oval(ARROW2_X1 - 4, y - 4, ARROW2_X1 + 4, y + 4, fill=line_color, outline=line_color)
                canvas.create_line(ARROW2_X1 + 5, y, ARROW2_X2 - 5, y, fill=line_color, width=2, arrow=tk.LAST, arrowshape=(8, 11, 4))

                # ── [3] 매출월정산 컬럼 ──
                canvas.create_text(SAL_COL_X, y, text=f'{sal_let:>2} ({m["sal_col"]})', anchor='w', font=('Consolas', 10, 'bold'), fill='#276749')
                canvas.create_text(SAL_NAME_X, y, text=f'{sal_hdr_name} (Col {m["sal_col"]} / {sal_let}열)', anchor='w', font=('Segoe UI', 9), fill='#22543D')

                for extra_idx, extra_party in enumerate(self.extra_parties, start=4):
                    base_x = EXTRA_START_X + (extra_idx - 4) * EXTRA_GROUP_W
                    arrow_x1 = base_x - 80
                    arrow_x2 = base_x - 15
                    extra_col = get_extra_col_for_mapping(extra_party, m)
                    extra_let = col_num_to_letter(extra_col)
                    extra_hdr_name = extra_party.get('headers', {}).get(extra_col, f'Col {extra_col}')
                    canvas.create_oval(arrow_x1 - 4, y - 4, arrow_x1 + 4, y + 4, fill=line_color, outline=line_color)
                    canvas.create_line(arrow_x1 + 5, y, arrow_x2, y, fill=line_color, width=2, arrow=tk.LAST, arrowshape=(8, 11, 4))
                    canvas.create_text(base_x + EXTRA_COL_OFFSET, y, text=f'{extra_let:>2} ({extra_col})', anchor='w', font=('Consolas', 10, 'bold'), fill='#553C9A')
                    canvas.create_text(base_x + EXTRA_NAME_OFFSET, y, text=f'{extra_hdr_name} (Col {extra_col} / {extra_let}열)', anchor='w', font=('Segoe UI', 9), fill='#44337A')

                # ── 모드 ──
                canvas.create_text(mode_x, y, text=mode_text, anchor='w', font=('Segoe UI', 9), fill='#4A5568')

            if n > 0:
                y_end = START_Y + n * ROW_H + 5
                canvas.create_line(5, y_end, canvas_w - 5, y_end, fill='#CBD5E0', width=1)

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
                self.show_3way_mapping_edit_dialog(popup, idx, refresh)

        canvas.bind('<Button-1>', on_click)
        canvas.bind('<Double-1>', on_double_click)

        btn_bar = ttk.Frame(main_f)
        btn_bar.pack(fill='x', pady=(6, 0))

        def on_edit():
            if selected_idx[0] is not None and 0 <= selected_idx[0] < len(self.header_mappings):
                self.show_3way_mapping_edit_dialog(popup, selected_idx[0], refresh)
            else:
                messagebox.showwarning('선택 확인', '편집할 매핑 항목을 클릭해 선택하세요.', parent=popup)

        def on_add():
            self.show_3way_add_mapping_dialog(popup, refresh)

        def on_delete():
            if selected_idx[0] is not None and 0 <= selected_idx[0] < len(self.header_mappings):
                m = self.header_mappings[selected_idx[0]]
                if messagebox.askyesno('삭제 확인', f'매핑 [{m["field_label"]}] 항목을 삭제하시겠습니까?', parent=popup):
                    del self.header_mappings[selected_idx[0]]
                    selected_idx[0] = None
                    refresh()
            else:
                messagebox.showwarning('선택 확인', '삭제할 매핑 항목을 선택하세요.', parent=popup)

        def on_reset():
            if messagebox.askyesno('초기화 확인', '모든 3자간 매핑을 기본 표준 설정값(자동)으로 복원하시겠습니까?', parent=popup):
                self.header_mappings = self.get_default_mappings()
                selected_idx[0] = None
                refresh()

        btn_edit = ttk.Button(btn_bar, text='✏️ 선택 매핑 편집...', command=on_edit)
        btn_edit.pack(side='left', padx=(0, 6))
        HoverTooltip(btn_edit, (
            "선택한 실행 그룹의 발주/매입/매출 컬럼을 수동 변경합니다.\n\n"
            "계약명 대신 PO번호를 주 매칭 기준으로 쓰려면 [계약명 / 적요] 행을 선택하고 이 버튼을 누른 뒤, 세 파일의 대응 컬럼을 PO 계열 헤더로 맞추세요."
        ))
        btn_add = ttk.Button(btn_bar, text='➕ 신규 매핑 추가...', command=on_add)
        btn_add.pack(side='left', padx=(0, 6))
        HoverTooltip(btn_add, (
            "보고서/표시용 보조 매핑 그룹을 추가합니다.\n\n"
            "주의: 기본 실행 핵심 기준을 바꾸려면 신규 추가보다 기존 [계약명 / 적요], [PO번호], [발행구분], [계약방식] 행을 편집하는 것이 맞습니다."
        ))
        ttk.Button(btn_bar, text='🗑️ 선택 매핑 삭제', command=on_delete).pack(side='left', padx=(0, 6))
        ttk.Button(btn_bar, text='🔄 기본값 복원', command=on_reset).pack(side='left')

        ttk.Button(btn_bar, text='☑ 매핑 저장 및 닫기', style='Primary.TButton', command=popup.destroy).pack(side='right')

        refresh()

    def show_3way_mapping_edit_dialog(self, parent, mapping_idx, refresh_callback):
        mapping = self.header_mappings[mapping_idx]
        popup = tk.Toplevel(parent)
        popup.title(f'✏️ 3자간 매핑 편집 [{mapping["field_label"]}]')
        popup.geometry('760x520')
        popup.grab_set()

        f = ttk.Frame(popup, padding=16)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text=f'✏️ [{mapping["field_label"]}] 3자간 컬럼 매핑 수정', font=('Segoe UI', 11, 'bold'), foreground=self.COLOR_PRIMARY).pack(anchor='w', pady=(0, 10))
        ttk.Label(
            f,
            text=f'실행 연계 그룹: {mapping.get("field_key", "custom")} | 선택한 헤더명과 Col 좌표가 저장 후 교차 검증 실행부에 그대로 적용됩니다.',
            font=('Segoe UI', 9),
            foreground='#4A5568'
        ).pack(anchor='w', pady=(0, 10))

        # 업로드/파싱된 엑셀 실데이터 구조 기반 실시간 동적 옵션 생성
        max_ref_col = max(52, max(self.ref_headers.keys()) if self.ref_headers else 52)
        max_pur_col = max(52, max(self.pur_headers.keys()) if self.pur_headers else 52)
        max_sal_col = max(52, max(self.sal_headers.keys()) if self.sal_headers else 52)
        ref_col_options = [self.get_ref_col_option_str(i) for i in range(1, max_ref_col + 1)]
        pur_col_options = [self.get_pur_col_option_str(i) for i in range(1, max_pur_col + 1)]
        sal_col_options = [self.get_sal_col_option_str(i) for i in range(1, max_sal_col + 1)]

        # 1. 매핑 항목 명칭
        f_label = ttk.Frame(f)
        f_label.pack(fill='x', pady=(0, 8))
        ttk.Label(f_label, text='🏷️ 매핑 항목 명칭:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_label = tk.StringVar(value=mapping['field_label'])
        ent_label = ttk.Entry(f_label, textvariable=var_label, font=('Segoe UI', 10), width=38)
        ent_label.pack(side='left', fill='x', expand=True)

        # 2. [1] 기준 발주내역 컬럼
        f_ref = ttk.Frame(f)
        f_ref.pack(fill='x', pady=(0, 8))
        ttk.Label(f_ref, text='📌 [1] 발주내역 컬럼:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_ref_col = tk.StringVar(value=self.get_ref_col_option_str(mapping["ref_col"]))
        cmb_ref_col = ttk.Combobox(f_ref, textvariable=var_ref_col, values=ref_col_options, font=('Consolas', 9), width=38)
        cmb_ref_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_ref_col, (
            "기준 발주내역에서 이 실행 그룹에 사용할 헤더와 Col 좌표입니다.\n\n"
            "예: PO번호를 주 기준으로 쓰려면 [계약명 / 적요] 그룹에서 발주 컬럼을 PO번호(예: AA Col 27)로 선택합니다."
        ))

        # 3. [2] 매입월정산 컬럼
        f_pur = ttk.Frame(f)
        f_pur.pack(fill='x', pady=(0, 8))
        ttk.Label(f_pur, text='🎯 [2] 매입월정산 컬럼:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_pur_col = tk.StringVar(value=self.get_pur_col_option_str(mapping["pur_col"]))
        cmb_pur_col = ttk.Combobox(f_pur, textvariable=var_pur_col, values=pur_col_options, font=('Consolas', 9), width=38)
        cmb_pur_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_pur_col, (
            "매입월정산에서 발주 기준과 같은 의미로 대조할 헤더입니다.\n\n"
            "발주 기준을 PO번호로 바꾸면 매입 쪽도 PO/계약번호/공사번호 등 실제 대응되는 식별 컬럼으로 맞춥니다."
        ))

        # 4. [3] 매출월정산 컬럼
        f_sal = ttk.Frame(f)
        f_sal.pack(fill='x', pady=(0, 12))
        ttk.Label(f_sal, text='🎯 [3] 매출월정산 컬럼:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_sal_col = tk.StringVar(value=self.get_sal_col_option_str(mapping["sal_col"]))
        cmb_sal_col = ttk.Combobox(f_sal, textvariable=var_sal_col, values=sal_col_options, font=('Consolas', 9), width=38)
        cmb_sal_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_sal_col, (
            "매출월정산에서 발주 기준과 같은 의미로 대조할 헤더입니다.\n\n"
            "발주 기준을 PO번호로 바꾸면 매출 쪽도 PO/계약번호/공사번호 등 실제 대응되는 식별 컬럼으로 맞춥니다."
        ))

        lbl_preview = ttk.Label(f, text='', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0', wraplength=700, justify='left')
        lbl_preview.pack(anchor='w', pady=(2, 12))

        def update_preview(*args):
            ref_col = parse_col_num_from_str(var_ref_col.get())
            pur_col = parse_col_num_from_str(var_pur_col.get())
            sal_col = parse_col_num_from_str(var_sal_col.get())
            ref_hdr = self.ref_headers.get(ref_col, f'Col {ref_col}')
            pur_hdr = self.pur_headers.get(pur_col, f'Col {pur_col}')
            sal_hdr = self.sal_headers.get(sal_col, f'Col {sal_col}')
            linked_extras = len(self.extra_parties)
            extra_msg = f' | n자 그룹 {linked_extras}개도 동일 실행 그룹 기준으로 표시 연계' if linked_extras else ''
            lbl_preview.config(
                text=(
                    f'저장 미리보기: 발주 [{ref_hdr} / {col_num_to_letter(ref_col)}열 Col {ref_col}] ↔ '
                    f'매입 [{pur_hdr} / {col_num_to_letter(pur_col)}열 Col {pur_col}] ↔ '
                    f'매출 [{sal_hdr} / {col_num_to_letter(sal_col)}열 Col {sal_col}]{extra_msg}'
                )
            )

        var_ref_col.trace_add('write', update_preview)
        var_pur_col.trace_add('write', update_preview)
        var_sal_col.trace_add('write', update_preview)
        update_preview()

        def on_save():
            new_label = var_label.get().strip()
            ref_c_text = var_ref_col.get().strip()
            pur_c_text = var_pur_col.get().strip()
            sal_c_text = var_sal_col.get().strip()

            if not new_label:
                messagebox.showwarning('입력 확인', '매핑 항목 명칭을 입력해 주세요.', parent=popup)
                return

            new_ref_col = parse_col_num_from_str(ref_c_text)
            new_pur_col = parse_col_num_from_str(pur_c_text)
            new_sal_col = parse_col_num_from_str(sal_c_text)

            ref_hdr_name = self.ref_headers.get(new_ref_col, f'Col {new_ref_col}')
            pur_hdr_name = self.pur_headers.get(new_pur_col, f'Col {new_pur_col}')
            sal_hdr_name = self.sal_headers.get(new_sal_col, f'Col {new_sal_col}')

            self.header_mappings[mapping_idx] = {
                'field_key': mapping.get('field_key', 'custom'),
                'field_label': new_label,
                'ref_col': new_ref_col,
                'ref_header': f'{ref_hdr_name} (Col {new_ref_col} / {col_num_to_letter(new_ref_col)}열)',
                'pur_col': new_pur_col,
                'pur_header': f'{pur_hdr_name} (Col {new_pur_col} / {col_num_to_letter(new_pur_col)}열)',
                'sal_col': new_sal_col,
                'sal_header': f'{sal_hdr_name} (Col {new_sal_col} / {col_num_to_letter(new_sal_col)}열)',
                'mode': '수동'
            }
            refresh_callback()
            linked_same_group = sum(
                1 for idx, m in enumerate(self.header_mappings)
                if idx != mapping_idx and m.get('field_key') == mapping.get('field_key')
            )
            if linked_same_group:
                self.log('INFO', f'동일 실행 연계 그룹 {linked_same_group}건 확인됨: [{new_label}] 저장값이 같은 field_key 기준으로 함께 참조됩니다.')
            else:
                self.log('INFO', f'동일 실행 연계 그룹 추가 항목 없음: [{new_label}] 현재 항목 기준으로 실행부에 적용됩니다.')
            self.log('INFO', f'3자 매핑 저장 완료: [{new_label}] 발주:{col_num_to_letter(new_ref_col)}({ref_hdr_name}) ──▶ 매입:{col_num_to_letter(new_pur_col)}({pur_hdr_name}) ──▶ 매출:{col_num_to_letter(new_sal_col)}({sal_hdr_name})')
            popup.destroy()

        f_btns = ttk.Frame(f)
        f_btns.pack(fill='x', side='bottom')
        btn_save = ttk.Button(f_btns, text='☑ 저장', style='Primary.TButton', command=on_save)
        btn_save.pack(side='right', padx=(6, 0))
        HoverTooltip(btn_save, (
            "현재 실행 그룹에 선택한 발주/매입/매출 헤더명과 Col 좌표를 저장합니다.\n\n"
            "저장 미리보기에 세 파일의 헤더명이 모두 같은 의미로 표시되는지 확인한 뒤 저장하세요."
        ))
        ttk.Button(f_btns, text='취소', command=popup.destroy).pack(side='right')

    def show_3way_add_mapping_dialog(self, parent, refresh_callback):
        popup = tk.Toplevel(parent)
        popup.title('➕ 신규 3자간 헤더 매핑 추가')
        popup.geometry('760x520')
        popup.grab_set()

        f = ttk.Frame(popup, padding=16)
        f.pack(fill='both', expand=True)

        ttk.Label(f, text='➕ 신규 3자간 컬럼 매핑 추가', font=('Segoe UI', 11, 'bold'), foreground=self.COLOR_PRIMARY).pack(anchor='w', pady=(0, 10))

        max_ref_col = max(52, max(self.ref_headers.keys()) if self.ref_headers else 52)
        max_pur_col = max(52, max(self.pur_headers.keys()) if self.pur_headers else 52)
        max_sal_col = max(52, max(self.sal_headers.keys()) if self.sal_headers else 52)
        ref_col_options = [self.get_ref_col_option_str(i) for i in range(1, max_ref_col + 1)]
        pur_col_options = [self.get_pur_col_option_str(i) for i in range(1, max_pur_col + 1)]
        sal_col_options = [self.get_sal_col_option_str(i) for i in range(1, max_sal_col + 1)]

        f_label = ttk.Frame(f)
        f_label.pack(fill='x', pady=(0, 8))
        ttk.Label(f_label, text='🏷️ 매핑 항목 명칭:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_label = tk.StringVar(value='')
        ent_label = ttk.Entry(f_label, textvariable=var_label, font=('Segoe UI', 10), width=38)
        ent_label.pack(side='left', fill='x', expand=True)

        f_ref = ttk.Frame(f)
        f_ref.pack(fill='x', pady=(0, 8))
        ttk.Label(f_ref, text='📌 [1] 발주내역 컬럼:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_ref_col = tk.StringVar(value=self.get_ref_col_option_str(24))
        cmb_ref_col = ttk.Combobox(f_ref, textvariable=var_ref_col, values=ref_col_options, font=('Consolas', 9), width=38)
        cmb_ref_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_ref_col, "신규 보조 매핑 그룹의 발주내역 헤더와 Col 좌표를 선택합니다.")

        f_pur = ttk.Frame(f)
        f_pur.pack(fill='x', pady=(0, 8))
        ttk.Label(f_pur, text='🎯 [2] 매입월정산 컬럼:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_pur_col = tk.StringVar(value=self.get_pur_col_option_str(28))
        cmb_pur_col = ttk.Combobox(f_pur, textvariable=var_pur_col, values=pur_col_options, font=('Consolas', 9), width=38)
        cmb_pur_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_pur_col, "신규 보조 매핑 그룹의 매입월정산 헤더와 Col 좌표를 선택합니다.")

        f_sal = ttk.Frame(f)
        f_sal.pack(fill='x', pady=(0, 12))
        ttk.Label(f_sal, text='🎯 [3] 매출월정산 컬럼:', font=('Segoe UI', 10, 'bold'), width=24, anchor='w').pack(side='left')
        var_sal_col = tk.StringVar(value=self.get_sal_col_option_str(11))
        cmb_sal_col = ttk.Combobox(f_sal, textvariable=var_sal_col, values=sal_col_options, font=('Consolas', 9), width=38)
        cmb_sal_col.pack(side='left', fill='x', expand=True)
        HoverTooltip(cmb_sal_col, "신규 보조 매핑 그룹의 매출월정산 헤더와 Col 좌표를 선택합니다.")

        lbl_preview = ttk.Label(f, text='', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0', wraplength=700, justify='left')
        lbl_preview.pack(anchor='w', pady=(2, 12))

        def update_preview(*args):
            ref_col = parse_col_num_from_str(var_ref_col.get())
            pur_col = parse_col_num_from_str(var_pur_col.get())
            sal_col = parse_col_num_from_str(var_sal_col.get())
            ref_hdr = self.ref_headers.get(ref_col, f'Col {ref_col}')
            pur_hdr = self.pur_headers.get(pur_col, f'Col {pur_col}')
            sal_hdr = self.sal_headers.get(sal_col, f'Col {sal_col}')
            lbl_preview.config(
                text=(
                    f'신규 저장 미리보기: 발주 [{ref_hdr} / {col_num_to_letter(ref_col)}열 Col {ref_col}] ↔ '
                    f'매입 [{pur_hdr} / {col_num_to_letter(pur_col)}열 Col {pur_col}] ↔ '
                    f'매출 [{sal_hdr} / {col_num_to_letter(sal_col)}열 Col {sal_col}]'
                )
            )

        var_ref_col.trace_add('write', update_preview)
        var_pur_col.trace_add('write', update_preview)
        var_sal_col.trace_add('write', update_preview)
        update_preview()

        def on_add_save():
            new_label = var_label.get().strip()
            ref_c_text = var_ref_col.get().strip()
            pur_c_text = var_pur_col.get().strip()
            sal_c_text = var_sal_col.get().strip()

            if not new_label:
                messagebox.showwarning('입력 확인', '매핑 항목 명칭을 입력하세요.', parent=popup)
                return

            new_ref_col = parse_col_num_from_str(ref_c_text)
            new_pur_col = parse_col_num_from_str(pur_c_text)
            new_sal_col = parse_col_num_from_str(sal_c_text)

            ref_hdr_name = self.ref_headers.get(new_ref_col, f'Col {new_ref_col}')
            pur_hdr_name = self.pur_headers.get(new_pur_col, f'Col {new_pur_col}')
            sal_hdr_name = self.sal_headers.get(new_sal_col, f'Col {new_sal_col}')

            self.header_mappings.append({
                'field_key': 'custom_' + str(len(self.header_mappings)),
                'field_label': new_label,
                'ref_col': new_ref_col,
                'ref_header': f'{ref_hdr_name} (Col {new_ref_col} / {col_num_to_letter(new_ref_col)}열)',
                'pur_col': new_pur_col,
                'pur_header': f'{pur_hdr_name} (Col {new_pur_col} / {col_num_to_letter(new_pur_col)}열)',
                'sal_col': new_sal_col,
                'sal_header': f'{sal_hdr_name} (Col {new_sal_col} / {col_num_to_letter(new_sal_col)}열)',
                'mode': '신규'
            })
            refresh_callback()
            self.log('INFO', f'신규 매핑 [{new_label}]은 custom 그룹으로 추가되었습니다. 계약명/PO/발행구분처럼 실행 핵심 기준을 바꾸려면 기존 핵심 그룹 행을 편집하세요.')
            self.log('INFO', f'신규 3자 매핑 추가 완료: [{new_label}] 발주:{col_num_to_letter(new_ref_col)}({ref_hdr_name}) ──▶ 매입:{col_num_to_letter(new_pur_col)}({pur_hdr_name}) ──▶ 매출:{col_num_to_letter(new_sal_col)}({sal_hdr_name})')
            popup.destroy()

        f_btns = ttk.Frame(f)
        f_btns.pack(fill='x', side='bottom')
        btn_add_save = ttk.Button(f_btns, text='☑ 추가', style='Primary.TButton', command=on_add_save)
        btn_add_save.pack(side='right', padx=(6, 0))
        HoverTooltip(btn_add_save, (
            "신규 보조 매핑 그룹을 추가합니다.\n\n"
            "주 매칭 기준을 바꾸는 작업은 신규 추가가 아니라 기존 [계약명 / 적요] 실행 그룹 편집으로 처리하세요."
        ))
        ttk.Button(f_btns, text='취소', command=popup.destroy).pack(side='right')

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
        target_wb = None

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
                for wb in app.Workbooks:
                    try:
                        wb_full_name = str(wb.FullName or '').lower()
                        wb_name = str(wb.Name or '').lower()
                        if wb_full_name == abs_target.lower() or wb_name == fn_lower:
                            return wb
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
                app.WindowState = -4143  # xlNormal
                _ = int(app.Hwnd)
                self.interactive_excel_app = app
                self.persist_excel_app = app
                return True
            except Exception:
                return False

        # 기존 자동화 인스턴스가 ROT에 죽은 객체로 남는 경우가 있어, 반드시 Hwnd/Workbooks 접근으로 검증한다.
        # 파일 경로 기반 GetObject는 닫힌 대용량 파일을 여는 부작용이 있어 전면화 속도를 늦추므로 사용하지 않는다.
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
                target_wb = candidate_wb
                break
            if excel is None:
                excel = candidate

        if excel:
            target_wb = target_wb or find_target_workbook(excel)

        if not target_wb:
            self.log('INFO', '대상 파일이 현재 Excel에 열려 있지 않아 Windows 기본 열기로 빠르게 활성화합니다.')
            try:
                os.startfile(abs_target)
            except Exception as exc:
                self.log('WARN', f'OS 기본 열기로 Excel 파일을 열지 못했습니다. COM 열기로 재시도합니다: {exc}')

            deadline = time.perf_counter() + 6.0
            while time.perf_counter() < deadline and not target_wb:
                try:
                    candidate = win32com.client.GetActiveObject('Excel.Application')
                    if is_usable_excel_app(candidate) and configure_visible_excel(candidate):
                        excel = candidate
                        target_wb = find_target_workbook(excel)
                        if target_wb:
                            break
                except Exception:
                    pass
                time.sleep(0.15)

        if not target_wb:
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
                    target_wb = excel.Workbooks.Open(abs_target)
                except Exception as exc:
                    self.log('WARN', f'기준 발주내역 파일을 Excel COM으로 열지 못했습니다: {exc}')

        if excel:
            if target_wb:
                self.interactive_target_wb = target_wb
                self.persist_target_wb = target_wb
                try:
                    target_wb.Activate()
                except Exception: pass
                try:
                    target_wb.Windows(1).Activate()
                except Exception: pass
        elif not target_wb:
            try:
                os.startfile(abs_target)
            except Exception as exc:
                self.log('ERROR', f'Excel 실행 실패: {exc}')

        # Windows API SetWindowPos + Alt-Key + AttachThreadInput으로 엑셀 창 화면 물리적 실체화 (전체화면 아님)
        if COM_AVAILABLE:
            import win32api
            target_hwnd = None

            if excel:
                for hwnd_getter in (
                    lambda: int(excel.Hwnd),
                    lambda: int(target_wb.Windows(1).Hwnd) if target_wb else None,
                ):
                    try:
                        candidate_hwnd = hwnd_getter()
                        if candidate_hwnd and win32gui.IsWindow(candidate_hwnd):
                            target_hwnd = candidate_hwnd
                            break
                    except Exception:
                        pass

            def enum_windows_callback(hwnd, extra):
                nonlocal target_hwnd
                if win32gui.IsWindowVisible(hwnd):
                    if win32gui.GetClassName(hwnd) == 'XLMAIN':
                        target_hwnd = hwnd
                        return False
                return True

            if not target_hwnd:
                try:
                    win32gui.EnumWindows(enum_windows_callback, None)
                except Exception: pass

            if not target_hwnd:
                target_hwnd = win32gui.FindWindow('XLMAIN', None)

            if target_hwnd and win32gui.IsWindow(target_hwnd):
                try:
                    # 1. 엑셀이 작업표시줄에 최소화되어 있을 경우 일반(Restored) 창으로 화면 복원
                    if win32gui.IsIconic(target_hwnd):
                        win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                    else:
                        win32gui.ShowWindow(target_hwnd, win32con.SW_SHOW)

                    # 2. SWP_NOMOVE | SWP_NOSIZE로 크기/위치 유지하면서 최상위 데스크톱 레이어로 물리적 표시
                    win32gui.SetWindowPos(
                        target_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                    )
                    win32gui.SetWindowPos(
                        target_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                    )

                    # 3. Windows LockSetForegroundWindow 방지용 Alt 키 시뮬레이션 및 Thread Attach
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

                    # 4. COM AppActivate 호출로 엑셀 활성화
                    if excel:
                        try:
                            shell = win32com.client.Dispatch('WScript.Shell')
                            app_titles = []
                            if target_wb:
                                app_titles.append(target_wb.Name)
                            app_titles.extend([fn, excel.Caption, 'Microsoft Excel', 'Excel'])
                            for title in app_titles:
                                try:
                                    if title and shell.AppActivate(title):
                                        break
                                except Exception:
                                    pass
                        except Exception: pass
                except Exception: pass

        # 현재 스레드의 COM 객체에서 마우스 드래그 선택 영역 주소 정밀 파싱
        if excel:
            for prop in ['Selection', 'ActiveWindow.RangeSelection', 'ActiveCell']:
                if sel_addr: break
                try:
                    obj = excel
                    for p in prop.split('.'):
                        obj = getattr(obj, p)
                    if obj and hasattr(obj, 'Address'):
                        raw = str(obj.Address).replace('$', '').strip()
                        if raw:
                            sel_addr = raw
                except Exception: pass

        self.log('INFO', f'Excel 활성화 처리 완료: {time.perf_counter() - activation_started_at:.2f}초')
        return sel_addr

    def on_interactive_drag_select(self):
        if any(self.active_tasks.values()):
            self.queued_action = self.on_interactive_drag_select
            self.log('INFO', '⏳ 백그라운드 작업 중입니다. 완료되는 즉시 [마우스 범위 확정] 창을 자동으로 엽니다.')
            return

        ref_path = self.ref_file_path.get()
        if not ref_path or not os.path.exists(ref_path):
            messagebox.showwarning('입력 확인', '먼저 올바른 기준 발주내역 엑셀 파일(.xlsb/.xlsx)을 선택하세요.')
            return

        if self.excel_activation_busy:
            self.log('INFO', 'Excel 전면 활성화가 이미 진행 중입니다. 잠시 후 다시 시도하세요.')
            return

        if self.range_popup is not None and self.range_popup.winfo_exists():
            if not self.confirm_excel_cleanup_before_activation(parent=self.range_popup, show_when_visible_only=False):
                return
            self.excel_activation_busy = True
            try:
                self.range_popup.lift()
                self.range_popup.attributes('-topmost', False)
                self.bring_excel_to_front_and_get_selection(ref_path)
                self.root.lower()
                self.range_popup.attributes('-topmost', True)
                self.range_popup.lift()
            except Exception as exc:
                self.log('WARN', f'⚠️ Excel 창 전면 활성화 재시도 실패: {exc}')
            finally:
                self.excel_activation_busy = False
            return

        if not self.confirm_excel_cleanup_before_activation(parent=self.root):
            return

        self.log('INFO', '기준 발주내역 Excel 화면을 전면 활성화합니다...')
        self.excel_activation_busy = True
        try:
            self.root.attributes('-topmost', False)
            self.root.update_idletasks()
            curr_sel = self.bring_excel_to_front_and_get_selection(ref_path)
            self.root.lower()
        except Exception as exc:
            curr_sel = ''
            self.log('WARN', f'⚠️ Excel 창 전면 활성화 실패: {exc}')
        finally:
            self.excel_activation_busy = False
        self.root.after(300, lambda: self.show_range_confirmation_dialog(ref_path, curr_sel))

    def show_range_confirmation_dialog(self, ref_path, curr_sel):
        if self.range_popup is not None and self.range_popup.winfo_exists():
            self.range_popup.lift()
            return

        popup = tk.Toplevel(self.root)
        self.range_popup = popup
        popup.title('🎯 기준 발주내역 마우스 대상 행 범위 확정')
        popup.geometry('800x540')

        # 안내창은 보이게 유지하되 포커스는 Excel에 돌려 사용자가 바로 드래그할 수 있게 한다.
        popup.attributes('-topmost', True)
        popup.lift()

        f = ttk.Frame(popup, padding=20)
        f.pack(fill='both', expand=True)

        lbl_title = ttk.Label(f, text='🎯 기준 발주내역 대상 행 범위 확정 (사용자 선택 기준)', font=('Segoe UI', 12, 'bold'), foreground=self.COLOR_PRIMARY)
        lbl_title.pack(anchor='w', pady=(0, 8))
        HoverTooltip(lbl_title, (
            "이 창은 기준 발주내역에서 검증 대상 행 범위를 저장합니다.\n\n"
            "계약명 X열만 사용할 필요는 없습니다. PO번호, PR번호, 계약번호 등 사용자가 기준으로 삼을 열을 드래그할 수 있습니다.\n"
            "다만 실제 발주↔매입↔매출 교차 검증 기준 컬럼은 [3자간 헤더 매핑 설정]의 논리 그룹이 결정합니다."
        ))

        msg_desc = (
            '엑셀 창이 전면에 열렸습니다! 엑셀 화면에서 기준으로 사용할 열의 대상 행 범위\n'
            '(예: X4942:X4963, AA4942:AA4963 또는 4942:4963)를 드래그한 뒤 [영역 읽기]를 누르세요.'
        )
        ttk.Label(f, text=msg_desc, justify='left', font=('Segoe UI', 9)).pack(anchor='w', pady=(0, 12))

        initial_val = 'X4942:X4963'
        if curr_sel and ':' in curr_sel:
            initial_val = curr_sel
        elif curr_sel:
            m_r = re.findall(r'\d+', curr_sel)
            if m_r and int(m_r[0]) > 10:
                initial_val = curr_sel

        f_entry = ttk.Frame(f)
        f_entry.pack(fill='x', pady=(0, 12))

        ttk.Label(f_entry, text='매핑 행 범위:', font=('Segoe UI', 10, 'bold')).pack(side='left', padx=(0, 8))
        var_input = tk.StringVar(value=initial_val)
        ent_range = ttk.Entry(f_entry, textvariable=var_input, font=('Consolas', 11, 'bold'))
        ent_range.pack(side='left', fill='x', expand=True, padx=(0, 8))
        HoverTooltip(ent_range, (
            "기준 발주내역에서 선택한 행 범위입니다.\n\n"
            "입력 예시:\n"
            "- 계약명 기준: X4942:X4963\n"
            "- PO번호 기준: AA4942:AA4963\n"
            "- 행 번호만 입력: 4942:4963\n\n"
            "열 문자가 없으면 기본 X열로 표시되지만, PO번호 등 다른 열을 기준으로 쓰려면 Excel에서 해당 열 범위를 드래그해 읽는 방식을 권장합니다."
        ))

        lbl_preview = ttk.Label(f, text='', font=('Segoe UI', 9, 'bold'), foreground='#2B6CB0')
        lbl_preview.pack(anchor='w', pady=(0, 15))

        lbl_workflow = ttk.Label(
            f,
            text='ⓘ 계약명이 아닌 열을 선택했다면 [3자간 헤더 매핑 설정]에서 실행 그룹 컬럼을 같은 기준으로 변경하세요.',
            font=('Segoe UI', 9, 'bold'),
            foreground='#805AD5'
        )
        lbl_workflow.pack(anchor='w', pady=(0, 10))
        HoverTooltip(lbl_workflow, (
            "수동 기준 변경 상세 절차\n\n"
            "1. 이 창에서 발주내역의 대상 행 범위를 저장합니다.\n"
            "2. PO번호를 주 기준으로 쓰려면 [3자간 헤더 매핑 설정]을 엽니다.\n"
            "3. [계약명 / 적요] 행을 편집합니다. 이 행은 실제 주 매칭 기준입니다.\n"
            "4. 발주 컬럼을 PO번호(예: AA Col 27)로, 매입/매출 컬럼도 각 파일의 PO 또는 계약번호 계열 컬럼으로 바꿉니다.\n"
            "5. PO 검증값도 별도로 맞추려면 [PO번호] 행도 같은 방식으로 확인합니다.\n"
            "6. n자 파일이 있으면 n자 헤더 매칭 창에서 같은 논리 그룹의 n자 컬럼을 선택합니다."
        ))

        def extract_range_col_letter(range_text):
            m_col = re.search(r'([A-Z]+)\s*\$?\d+', range_text.upper().replace('$', ''))
            if m_col:
                return m_col.group(1)
            return 'X'

        initial_col_letter = extract_range_col_letter(initial_val)
        initial_col_num = col_letter_to_num(initial_col_letter)
        f_title_input = ttk.Frame(f)
        f_title_input.pack(fill='x', pady=(0, 10))
        ttk.Label(f_title_input, text='선택 열 제목:', font=('Segoe UI', 10, 'bold'), width=14, anchor='w').pack(side='left', padx=(0, 8))
        var_selected_title = tk.StringVar(value=self.clean_basis_title(self.ref_headers.get(initial_col_num, f'Col {initial_col_num}')))
        ent_selected_title = ttk.Entry(f_title_input, textvariable=var_selected_title, font=('Consolas', 10, 'bold'))
        ent_selected_title.pack(side='left', fill='x', expand=True, padx=(0, 8))
        HoverTooltip(ent_selected_title, (
            "Excel에서 드래그한 기준 열의 제목을 입력합니다.\n\n"
            "예: 계약명, PO번호, PR번호, 계약번호, 공사번호.\n"
            "확정 시 이 제목을 기준으로 [계약명 / 적요] 주 매칭 그룹과 같은 의미의 연계 그룹이 자동 갱신됩니다."
        ))

        def get_auto_apply_preview():
            col_letter = extract_range_col_letter(var_input.get())
            col_num = col_letter_to_num(col_letter)
            title = self.clean_basis_title(var_selected_title.get()) or self.clean_basis_title(self.ref_headers.get(col_num, f'Col {col_num}'))
            field_key = self.infer_field_key_from_title(title)
            pur_col = self.infer_col_for_title(self.pur_headers, title, None)
            sal_col = self.infer_col_for_title(self.sal_headers, title, None)
            return title, field_key, col_num, pur_col, sal_col

        def autofill_title_from_range(force=False):
            col_letter = extract_range_col_letter(var_input.get())
            col_num = col_letter_to_num(col_letter)
            detected_title = self.ref_headers.get(col_num, '')
            if detected_title and (force or not var_selected_title.get().strip()):
                var_selected_title.set(self.clean_basis_title(detected_title))

        def on_auto_sync_click():
            set_linked_selectors_from_title()
            title, field_key, ref_col, pur_col, sal_col = get_auto_apply_preview()
            field_label = field_key or NO_HEADER_MATCH_OPTION
            pur_text = f'Col {pur_col}' if pur_col else NO_HEADER_MATCH_OPTION
            sal_text = f'Col {sal_col}' if sal_col else NO_HEADER_MATCH_OPTION
            self.log('INFO', f'선택 열 제목 기준 자동 연계 준비: [{title}] field_key={field_label}, 발주 Col {ref_col}, 매입 {pur_text}, 매출 {sal_text}')
            on_parse_preview()

        btn_auto_sync = ttk.Button(f_title_input, text='🔗 제목 기준 자동 연계', command=on_auto_sync_click)
        btn_auto_sync.pack(side='right')
        HoverTooltip(btn_auto_sync, (
            "입력한 선택 열 제목으로 자동 갱신될 3자 실행 그룹을 미리 계산합니다.\n\n"
            "실제 저장은 [범위 확정 및 저장]을 누를 때 수행됩니다."
        ))

        linked_box = ttk.LabelFrame(f, text='매입 / 매출 / n자 연계 열 제목 자동 선정 및 수동 변경', padding=8)
        linked_box.pack(fill='x', pady=(0, 10))

        max_pur_col = max(52, max(self.pur_headers.keys()) if self.pur_headers else 52)
        max_sal_col = max(52, max(self.sal_headers.keys()) if self.sal_headers else 52)
        pur_options = [NO_HEADER_MATCH_OPTION] + [self.get_pur_col_option_str(i) for i in range(1, max_pur_col + 1)]
        sal_options = [NO_HEADER_MATCH_OPTION] + [self.get_sal_col_option_str(i) for i in range(1, max_sal_col + 1)]

        def add_link_row(parent, label_text, variable, options, tooltip):
            row = ttk.Frame(parent)
            row.pack(fill='x', pady=(0, 4))
            ttk.Label(row, text=label_text, font=('Segoe UI', 9, 'bold'), width=16, anchor='w').pack(side='left', padx=(0, 6))
            combo = ttk.Combobox(row, textvariable=variable, values=options, font=('Consolas', 9), width=68)
            combo.pack(side='left', fill='x', expand=True)
            HoverTooltip(combo, tooltip)
            return combo

        var_pur_link = tk.StringVar()
        var_sal_link = tk.StringVar()
        add_link_row(
            linked_box,
            '매입 연계 열:',
            var_pur_link,
            pur_options,
            '선택 열 제목과 같은 의미로 매입월정산에서 대조할 컬럼입니다. 자동 선정값이 맞지 않으면 직접 변경하세요.'
        )
        add_link_row(
            linked_box,
            '매출 연계 열:',
            var_sal_link,
            sal_options,
            '선택 열 제목과 같은 의미로 매출월정산에서 대조할 컬럼입니다. 자동 선정값이 맞지 않으면 직접 변경하세요.'
        )

        extra_link_vars = {}
        for party in self.extra_parties:
            max_extra_col = max(52, max(party.get('headers', {}).keys()) if party.get('headers') else 52)
            extra_options = [NO_HEADER_MATCH_OPTION] + [self.get_extra_col_option_str(party, i) for i in range(1, max_extra_col + 1)]
            var_extra = tk.StringVar()
            extra_link_vars[party['id']] = (party, var_extra)
            add_link_row(
                linked_box,
                f'{party["label"]}:',
                var_extra,
                extra_options,
                f'{party["label"]} 파일에서 선택 열 제목과 같은 의미로 대조할 컬럼입니다. 자동 선정값이 맞지 않으면 직접 변경하세요.'
            )

        def set_linked_selectors_from_title():
            title, _field_key, _ref_col, pur_col, sal_col = get_auto_apply_preview()
            var_pur_link.set(self.get_pur_col_option_str(pur_col) if pur_col else NO_HEADER_MATCH_OPTION)
            var_sal_link.set(self.get_sal_col_option_str(sal_col) if sal_col else NO_HEADER_MATCH_OPTION)
            for party_id, (party, variable) in extra_link_vars.items():
                extra_col = self.infer_col_for_title(
                    party.get('headers', {}),
                    title,
                    None
                )
                variable.set(self.get_extra_col_option_str(party, extra_col) if extra_col else NO_HEADER_MATCH_OPTION)

        set_linked_selectors_from_title()

        def on_parse_preview(*args):
            autofill_title_from_range(force=False)
            val = var_input.get().strip().upper().replace('$', '')
            m = re.findall(r'\d+', val)
            col_letter = extract_range_col_letter(val)
            col_num = col_letter_to_num(col_letter)
            header_name = self.clean_basis_title(var_selected_title.get()) or self.clean_basis_title(self.ref_headers.get(col_num, f'Col {col_num}'))
            title, field_key, _, pur_col, sal_col = get_auto_apply_preview()
            field_label = field_key or NO_HEADER_MATCH_OPTION
            pur_text = f'Col {pur_col}' if pur_col else NO_HEADER_MATCH_OPTION
            sal_text = f'Col {sal_col}' if sal_col else NO_HEADER_MATCH_OPTION
            auto_msg = f' | 자동연계: {field_label}, 매입 {pur_text}, 매출 {sal_text}'
            if len(m) >= 2:
                r1, r2 = int(m[0]), int(m[-1])
                s_r, e_r = min(r1, r2), max(r1, r2)
                cnt = e_r - s_r + 1
                lbl_preview.config(text=f'✔ 감지 결과: {col_letter}{s_r}:{col_letter}{e_r} ({header_name} / Col {col_num}, {s_r}행 ~ {e_r}행, 총 {cnt}개 지정){auto_msg}')
            elif len(m) == 1:
                r1 = int(m[0])
                lbl_preview.config(text=f'✔ 감지 결과: {col_letter}{r1} ({header_name} / Col {col_num}, {r1}행 단일 지정){auto_msg}')
            else:
                lbl_preview.config(text='⚠️ 올바른 행 범위(예: X4942:X4963 또는 AA4942:AA4963)를 입력하세요.')

        var_input.trace_add('write', lambda *_args: (set_linked_selectors_from_title(), on_parse_preview()))
        var_selected_title.trace_add('write', lambda *_args: (set_linked_selectors_from_title(), on_parse_preview()))

        def bring_and_refresh():
            btn_act_excel.config(state='disabled')
            try:
                if not self.confirm_excel_cleanup_before_activation(parent=popup, show_when_visible_only=False):
                    return
                new_sel = self.bring_excel_to_front_and_get_selection(ref_path)
                if new_sel:
                    formatted = new_sel
                    m = re.findall(r'\d+', new_sel)
                    if len(m) >= 2 and not re.search(r'[A-Z]+\$?\d+', new_sel.upper().replace('$', '')):
                        formatted = f'X{m[0]}:X{m[-1]}'
                    elif len(m) == 1 and not re.search(r'[A-Z]+\$?\d+', new_sel.upper().replace('$', '')):
                        formatted = f'X{m[0]}'

                    var_input.set(formatted)
                    autofill_title_from_range(force=True)
                    set_linked_selectors_from_title()
                    on_parse_preview()
                    self.log('SUCCESS', f'🎯 엑셀 드래그 마우스 선택 영역 자동 읽기 성공: {formatted}')
                else:
                    self.log('WARN', '⚠️ 선택 영역을 읽지 못했습니다. 엑셀에서 범위를 선택 후 다시 눌러주세요.')
            finally:
                btn_act_excel.config(state='normal')

        btn_act_excel = ttk.Button(f_entry, text='🖥️ 엑셀 창 열기 & 영역 읽기', command=bring_and_refresh)
        btn_act_excel.pack(side='right')
        HoverTooltip(btn_act_excel, (
            "활성화된 기준 발주내역 Excel에서 현재 선택한 범위를 다시 읽습니다.\n\n"
            "절차:\n"
            "1. Excel에서 원하는 기준 열의 행 범위를 드래그합니다.\n"
            "2. 이 버튼을 누릅니다.\n"
            "3. 감지 결과에 헤더명, 열 문자, Col 번호가 맞게 표시되는지 확인합니다.\n"
            "4. 계약명 외 기준이면 저장 후 [3자간 헤더 매핑 설정]에서 실행 그룹 컬럼까지 맞춥니다."
        ))

        on_parse_preview()

        def on_close_popup():
            self.range_popup = None
            self.release_interactive_excel_session('범위 확정/팝업 닫기')
            popup.destroy()

        def on_confirm():
            val = var_input.get().strip().upper().replace('$', '')
            m = re.findall(r'\d+', val)
            if not m:
                messagebox.showwarning('범위 확인', '올바른 범위(예: X4942:X4963 또는 AA4942:AA4963)를 입력해 주세요.')
                return

            if len(m) >= 2:
                r1, r2 = int(m[0]), int(m[-1])
                start_row, end_row = min(r1, r2), max(r1, r2)
            else:
                start_row = int(m[0])
                end_row = start_row
            col_letter = extract_range_col_letter(val)
            col_num = col_letter_to_num(col_letter)
            header_name = self.clean_basis_title(var_selected_title.get()) or self.clean_basis_title(self.ref_headers.get(col_num, f'Col {col_num}'))
            range_address = f'{col_letter}{start_row}:{col_letter}{end_row}' if start_row != end_row else f'{col_letter}{start_row}'
            pur_override = parse_optional_col_num_from_str(var_pur_link.get())
            sal_override = parse_optional_col_num_from_str(var_sal_link.get())
            missing_targets = []
            if not pur_override:
                missing_targets.append('매입 연계 열')
            if not sal_override:
                missing_targets.append('매출 연계 열')
            extra_overrides = {}
            for party_id, (party, variable) in extra_link_vars.items():
                extra_col = parse_optional_col_num_from_str(variable.get())
                if not extra_col:
                    missing_targets.append(f'{party["label"]} 연계 열')
                else:
                    extra_overrides[party_id] = extra_col
            selected_link_count = int(bool(pur_override)) + int(bool(sal_override)) + len(extra_overrides)
            if selected_link_count == 0:
                messagebox.showwarning(
                    '연계 열 선택 필요',
                    '자동 유사성 매칭 또는 수동 선택된 연계 열이 없습니다.\n\n'
                    '매입/매출/n자 중 최소 1개 이상의 같은 의미 열을 직접 선택한 뒤 다시 확정하세요.',
                    parent=popup
                )
                return
            if missing_targets:
                self.log(
                    'WARN',
                    '일부 연계 열이 매칭되지 않아 해당 항목은 기존 설정을 유지합니다: '
                    + ', '.join(missing_targets)
                )
            auto_result = self.auto_apply_ref_basis_title(
                header_name,
                col_num,
                pur_col_override=pur_override,
                sal_col_override=sal_override,
                extra_col_overrides=extra_overrides,
            )

            self.ref_range_info = {
                'start_row': start_row,
                'end_row': end_row,
                'range_address': range_address,
                'basis_col': col_num,
                'basis_letter': col_letter,
                'basis_header': header_name,
                'basis_field_key': auto_result['field_key'],
            }
            self.selected_range_address.set(range_address)
            
            status_text = f'지정된 범위: {range_address} ({header_name} / Col {col_num}, {start_row}행~{end_row}행 확정)'
            self.lbl_target_range.config(text=status_text, foreground='#2F855A', font=('Segoe UI', 9, 'bold'))
            self.log('SUCCESS', f'🎯 기준 발주내역 대상 행 범위 확정 완료: {range_address} ({header_name} / Col {col_num})')
            self.log('SUCCESS', f'선택 열 제목 [{header_name}] 기준 자동 연계 갱신 완료: {", ".join(auto_result["updated_groups"])} / 매입 Col {auto_result["pur_col"]}, 매출 Col {auto_result["sal_col"]}')
            on_close_popup()

        popup.protocol("WM_DELETE_WINDOW", on_close_popup)

        f_btns = ttk.Frame(f)
        f_btns.pack(fill='x', side='bottom')

        btn_confirm = ttk.Button(f_btns, text='☑ 범위 확정 및 저장', style='Primary.TButton', command=on_confirm)
        btn_confirm.pack(side='right', padx=(8, 0))
        HoverTooltip(btn_confirm, (
            "현재 범위를 기준 발주내역 대상 행 범위로 저장합니다.\n\n"
            "계약명 외 기준을 쓰는 경우 저장 후 다음 작업을 이어가세요:\n"
            "1. [3자간 헤더 매핑 설정] 열기\n"
            "2. [계약명 / 적요] 실행 그룹 편집\n"
            "3. 발주/매입/매출 컬럼을 같은 기준 의미의 헤더로 선택\n"
            "4. 저장 미리보기에서 세 파일의 헤더명과 Col 좌표 확인"
        ))
        ttk.Button(f_btns, text='취소', command=on_close_popup).pack(side='right')

    def on_reset_range_selection(self):
        self.ref_range_info = None
        self.selected_range_address.set('')
        self.lbl_target_range.config(
            text='지정된 범위: 미선택 (엑셀 창에서 마우스로 범위를 선택하세요)',
            foreground='#4A5568',
            font=('Segoe UI', 9)
        )
        self.log('INFO', '기준 마우스 범위 선택이 초기화되었습니다.')

    def scan_party_against_ref(self, excel, party_spec, ref_contract_list):
        path = os.path.abspath(party_spec['file_path'])
        label = party_spec['label']
        contract_col = party_spec['contract_col']
        po_col = party_spec.get('po_col', 27)
        issue_col = party_spec.get('issue_col', 22)
        type_col = party_spec.get('type_col', 10)
        candidate_defaults = party_spec.get('candidate_defaults', [11, 12, 28, 27, 24, 6, 7])

        wb = excel.Workbooks.Open(path, ReadOnly=True)
        try:
            ws = wb.Worksheets(1)
            total_rows = ws.UsedRange.Rows.Count
            candidate_cols = list(dict.fromkeys([contract_col] + candidate_defaults + [po_col]))
            needed_cols = list(dict.fromkeys(candidate_cols + [issue_col, 19, 22, type_col] + list(range(1, 36))))
            col_data = read_excel_columns(ws, 3, total_rows, needed_cols)

            rows_to_keep = []
            rows_to_delete = []
            matched_indices = set()
            extracted_info = {}

            for offset, r in enumerate(range(3, total_rows + 1)):
                row_values = {c: col_data.get(c, [None])[offset] for c in needed_cols}
                row_texts = [str(row_values.get(c) or '').strip() for c in candidate_cols if row_values.get(c)]
                if not row_texts:
                    continue

                issue_val = extract_issue_from_cached_row(row_values, preferred_cols=[issue_col, 22, 19])
                type_val = str(row_values.get(type_col) or '-').strip()

                matched = False
                for idx, ref_item in enumerate(ref_contract_list):
                    if idx in matched_indices:
                        continue
                    for cell_t in row_texts:
                        if is_text_matched(cell_t, ref_item['contract_name']):
                            matched = True
                            matched_indices.add(idx)
                            extracted_info[idx] = {
                                'matched_text': cell_t,
                                'issue_type': issue_val,
                                'contract_type': type_val,
                            }
                            break
                    if matched:
                        break

                if matched:
                    rows_to_keep.append(r)
                else:
                    rows_to_delete.append(r)

            missing_items = [item for idx, item in enumerate(ref_contract_list) if idx not in matched_indices]
            return {
                'id': party_spec['id'],
                'label': label,
                'file_path': path,
                'contract_col': contract_col,
                'rows_to_keep': rows_to_keep,
                'rows_to_delete': rows_to_delete,
                'matched_indices': matched_indices,
                'extracted_info': extracted_info,
                'missing_items': missing_items,
            }
        finally:
            wb.Close(False)

    # --------------------------------------------------------------------------
    # 3자 교차 검증 및 데이터 정제 실행 엔진 (v7.2 3자 동적 헤더 매핑 100% 연동)
    # --------------------------------------------------------------------------
    def on_run_reconciliation(self):
        if any(self.active_tasks.values()):
            self.queued_action = self.on_run_reconciliation
            self.log('INFO', '⏳ 백그라운드 파싱이 진행 중입니다. 완료되는 즉시 3자 교차 검증을 자동 실행합니다.')
            return

        ref_path = self.ref_file_path.get()
        pur_path = self.purchase_file_path.get()
        sal_path = self.sales_file_path.get()

        if not ref_path or not os.path.exists(ref_path):
            messagebox.showwarning('입력 확인', '1단계: 올바른 기준 발주내역 엑셀 파일을 선택하세요.')
            return
        if not pur_path or not os.path.exists(pur_path):
            messagebox.showwarning('입력 확인', '2단계: 올바른 매입월정산 엑셀 파일을 선택하세요.')
            return
        if not sal_path or not os.path.exists(sal_path):
            messagebox.showwarning('입력 확인', '2단계: 올바른 매출월정산 엑셀 파일을 선택하세요.')
            return
        for party in self.extra_parties:
            extra_path = party['file_path_var'].get()
            if not extra_path or not os.path.exists(extra_path):
                messagebox.showwarning('입력 확인', f'{party["label"]}: 올바른 n자 엑셀 파일을 선택하세요.')
                return
        if not self.ref_range_info:
            messagebox.showwarning('입력 확인', '1단계: "엑셀 창 열기 & 마우스 범위 확정" 버튼으로 기준 범위를 지정을 하세요.')
            return

        self.log('INFO', f'🚀 3자간 기본 검증 + n자 {len(self.extra_parties)}개 교차 검증 및 정제 분석 스레드를 가동합니다...')
        threading.Thread(target=self.execute_reconciliation_process, daemon=True).start()

    def execute_reconciliation_process(self):
        """3자 교차 대조 핵심 엔진 (v7.2 3자간 동적 컬럼 매핑 적용)"""
        ensure_excel_com_available()
        pythoncom.CoInitialize()
        excel = None
        excel_state = None
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel_state = set_excel_fast_mode(excel)

            ref_path = os.path.abspath(self.ref_file_path.get())
            pur_path = os.path.abspath(self.purchase_file_path.get())
            sal_path = os.path.abspath(self.sales_file_path.get())

            # 3자 동적 매핑에서 컬럼 인덱스 해석
            col_contract_ref = 24
            col_po_ref = 27

            col_contract_pur = 28
            col_po_pur = 7
            col_issue_pur = 19
            col_type_pur = 10

            col_contract_sal = 11
            col_po_sal = 27
            col_issue_sal = 22

            for m in self.header_mappings:
                key = m.get('field_key')
                if key == 'contract_name' or '계약명' in m.get('field_label', ''):
                    col_contract_ref = m.get('ref_col', 24)
                    col_contract_pur = m.get('pur_col', 28)
                    col_contract_sal = m.get('sal_col', 11)
                elif key == 'po_no' or 'PO' in m.get('field_label', ''):
                    col_po_ref = m.get('ref_col', 27)
                    col_po_pur = m.get('pur_col', 7)
                    col_po_sal = m.get('sal_col', 27)
                elif key == 'issue_type' or '발행' in m.get('field_label', ''):
                    col_issue_pur = m.get('pur_col', 19)
                    col_issue_sal = m.get('sal_col', 22)
                elif key == 'contract_type' or '방식' in m.get('field_label', ''):
                    col_type_pur = m.get('pur_col', 10)

            # 1. 기준 발주내역 파일 데이터 추출
            self.root.after(0, lambda: self.progress.config(value=10))
            self.log('INFO', f'1/4: 기준 발주내역 엑셀에서 계약명(Col {col_contract_ref}) 및 PO번호(Col {col_po_ref}) 추출 중...')
            
            wb_ref = excel.Workbooks.Open(ref_path, ReadOnly=True)
            ws_ref = None
            for sheet in wb_ref.Worksheets:
                if '집계' in sheet.Name:
                    ws_ref = sheet
                    break
            if not ws_ref:
                ws_ref = wb_ref.Worksheets(1)

            start_row = self.ref_range_info['start_row']
            end_row = self.ref_range_info['end_row']

            ref_cols = [col_contract_ref, col_po_ref]
            if col_po_ref != 27:
                ref_cols.append(27)
            ref_col_data = read_excel_columns(ws_ref, start_row, end_row, ref_cols)
            ref_contract_list = []
            for offset, r in enumerate(range(start_row, end_row + 1)):
                contract_name = ref_col_data.get(col_contract_ref, [None])[offset]
                
                # 발주내역 Col 27(PO번호) 셀 직접 1:1 대조 읽기 (옆 컬럼 무단 우회 탐색 100% 삭제)
                raw_po_val = ref_col_data.get(col_po_ref, [None])[offset]
                if not raw_po_val and col_po_ref != 27:
                    raw_po_val = ref_col_data.get(27, [None])[offset]

                cleaned_po = clean_po_val(raw_po_val)
                if cleaned_po:
                    po_no = cleaned_po
                else:
                    po_no = '빈셀임'

                if contract_name:
                    ref_contract_list.append({
                        'ref_row': r,
                        'contract_name': str(contract_name).strip(),
                        'po_no': po_no,
                        'room_key': extract_room_key(contract_name),
                        'discipline': extract_discipline(contract_name),
                        'round_num': extract_round_num(contract_name)
                    })
            wb_ref.Close(False)

            self.log('SUCCESS', f'기준 발주내역 계약명 {len(ref_contract_list)}건 추출 완료')

            # 2. 매입월정산 파일 탐색 및 대조
            self.root.after(0, lambda: self.progress.config(value=35))
            self.log('INFO', f'2/4: 매입월정산 엑셀 파싱 및 매칭 중 (계약명: Col {col_contract_pur}, 발행구분: Col {col_issue_pur})...')
            
            wb_pur = excel.Workbooks.Open(pur_path, ReadOnly=True)
            ws_pur = wb_pur.Worksheets(1)

            pur_used = ws_pur.UsedRange
            pur_total_rows = pur_used.Rows.Count

            pur_candidate_cols = list(dict.fromkeys([col_contract_pur, 28, 11, 12, col_po_pur, 27, 24, 6, 7]))
            pur_needed_cols = list(dict.fromkeys(pur_candidate_cols + [col_issue_pur, 19, 22, col_type_pur] + list(range(1, 36))))
            pur_col_data = read_excel_columns(ws_pur, 3, pur_total_rows, pur_needed_cols)

            pur_rows_to_keep = []
            pur_rows_to_delete = []
            pur_matched_ref_indices = set()
            pur_extracted_info = {}

            for offset, r in enumerate(range(3, pur_total_rows + 1)):
                row_values = {c: pur_col_data.get(c, [None])[offset] for c in pur_needed_cols}
                row_texts = [str(row_values.get(c) or '').strip() for c in pur_candidate_cols if row_values.get(c)]
                if not row_texts:
                    continue

                issue_val = extract_issue_from_cached_row(row_values, preferred_cols=[col_issue_pur, 19, 22])
                type_val = str(row_values.get(col_type_pur) or '-').strip()

                matched = False
                for idx, ref_item in enumerate(ref_contract_list):
                    if idx in pur_matched_ref_indices:
                        continue
                    for cell_t in row_texts:
                        if is_text_matched(cell_t, ref_item['contract_name']):
                            matched = True
                            pur_matched_ref_indices.add(idx)
                            pur_extracted_info[idx] = {
                                'matched_text': cell_t,
                                'issue_type': issue_val,
                                'contract_type': type_val
                            }
                            break
                    if matched: break

                if matched:
                    pur_rows_to_keep.append(r)
                else:
                    pur_rows_to_delete.append(r)

            wb_pur.Close(False)

            # 3. 매출월정산 파일 탐색 및 대조
            self.root.after(0, lambda: self.progress.config(value=65))
            self.log('INFO', f'3/4: 매출월정산 엑셀 파싱 및 매칭 중 (적요: Col {col_contract_sal}, 발행구분: Col {col_issue_sal})...')

            wb_sal = excel.Workbooks.Open(sal_path, ReadOnly=True)
            ws_sal = wb_sal.Worksheets(1)

            sal_used = ws_sal.UsedRange
            sal_total_rows = sal_used.Rows.Count

            sal_candidate_cols = list(dict.fromkeys([col_contract_sal, 11, 12, 28, col_po_sal, 27, 24, 6, 7]))
            sal_needed_cols = list(dict.fromkeys(sal_candidate_cols + [col_issue_sal, 22, 19] + list(range(1, 36))))
            sal_col_data = read_excel_columns(ws_sal, 3, sal_total_rows, sal_needed_cols)

            sal_rows_to_keep = []
            sal_rows_to_delete = []
            sal_matched_ref_indices = set()
            sal_extracted_info = {}

            for offset, r in enumerate(range(3, sal_total_rows + 1)):
                row_values = {c: sal_col_data.get(c, [None])[offset] for c in sal_needed_cols}
                row_texts = [str(row_values.get(c) or '').strip() for c in sal_candidate_cols if row_values.get(c)]
                if not row_texts:
                    continue

                issue_val = extract_issue_from_cached_row(row_values, preferred_cols=[col_issue_sal, 22, 19])

                matched = False
                for idx, ref_item in enumerate(ref_contract_list):
                    if idx in sal_matched_ref_indices:
                        continue
                    for cell_t in row_texts:
                        if is_text_matched(cell_t, ref_item['contract_name']):
                            matched = True
                            sal_matched_ref_indices.add(idx)
                            sal_extracted_info[idx] = {
                                'matched_text': cell_t,
                                'issue_type': issue_val
                            }
                            break
                    if matched: break

                if matched:
                    sal_rows_to_keep.append(r)
                else:
                    sal_rows_to_delete.append(r)

            wb_sal.Close(False)

            # 4. n자 추가 파일 탐색 및 대조
            extra_party_results = []
            if self.extra_parties:
                self.log('INFO', f'추가 n자 파일 {len(self.extra_parties)}개 파싱 및 매칭 중...')
            for party in self.extra_parties:
                if not party.get('mapping_manual'):
                    inferred = self.infer_party_column_mapping(party.get('headers', {}), fallback_contract_col=party['mapping'].get('contract_col', 11))
                    party['mapping'].update(inferred)
                party_spec = {
                    'id': party['id'],
                    'label': party['label'],
                    'file_path': party['file_path_var'].get(),
                    'contract_col': party['mapping'].get('contract_col', 11),
                    'po_col': party['mapping'].get('po_col', 27),
                    'issue_col': party['mapping'].get('issue_col', 22),
                    'type_col': party['mapping'].get('type_col', 10),
                    'candidate_defaults': [11, 12, 28, 27, 24, 6, 7],
                }
                result = self.scan_party_against_ref(excel, party_spec, ref_contract_list)
                extra_party_results.append(result)
                self.log('SUCCESS', f'[{party["label"]}] 매칭 {len(result["matched_indices"])}/{len(ref_contract_list)}건 완료')

            # 5. n자 확장 대조 통합 감사 데이터 구조 생성
            self.root.after(0, lambda: self.progress.config(value=90))
            self.log('INFO', '통합 감사 결과 수집 및 n자 확장 리포트 준비 중...')

            reconciliation_report = []
            pur_missing_items = []
            sal_missing_items = []

            for idx, ref_item in enumerate(ref_contract_list):
                pur_ok = idx in pur_matched_ref_indices
                sal_ok = idx in sal_matched_ref_indices

                pur_info = pur_extracted_info.get(idx, {})
                sal_info = sal_extracted_info.get(idx, {})

                if not pur_ok:
                    pur_missing_items.append(ref_item)
                if not sal_ok:
                    sal_missing_items.append(ref_item)

                extra_statuses = {}
                all_extra_ok = True
                for party_result in extra_party_results:
                    extra_ok = idx in party_result['matched_indices']
                    if not extra_ok:
                        all_extra_ok = False
                    extra_info = party_result['extracted_info'].get(idx, {})
                    extra_statuses[party_result['id']] = {
                        'label': party_result['label'],
                        'matched_text': extra_info.get('matched_text', '-'),
                        'status': '매칭 성공' if extra_ok else '미매칭(누락)',
                        'issue_type': extra_info.get('issue_type', '-'),
                        'matched': extra_ok,
                    }

                final_judge = '성공' if (pur_ok and sal_ok and all_extra_ok) else '불일치'

                reconciliation_report.append({
                    'idx': idx + 1,
                    'ref_contract': ref_item['contract_name'],
                    'ref_po_no': ref_item['po_no'],
                    'sep1': '│',
                    'pur_matched_text': pur_info.get('matched_text', '-'),
                    'pur_status': '매칭 성공' if pur_ok else '미매칭(누락)',
                    'pur_issue_type': pur_info.get('issue_type', '-'),
                    'pur_contract_type': pur_info.get('contract_type', '-'),
                    'sep2': '│',
                    'sal_matched_text': sal_info.get('matched_text', '-'),
                    'sal_status': '매칭 성공' if sal_ok else '미매칭(누락)',
                    'sal_issue_type': sal_info.get('issue_type', '-'),
                    'sep3': '│',
                    'final_judge': final_judge,
                    'pur_matched': pur_ok,
                    'sal_matched': sal_ok,
                    'extra_statuses': extra_statuses,
                })

            audit_context = {
                'report': reconciliation_report,
                'pur_rows_to_delete': pur_rows_to_delete,
                'sal_rows_to_delete': sal_rows_to_delete,
                'pur_missing_items': pur_missing_items,
                'sal_missing_items': sal_missing_items,
                'col_pur_contract': col_contract_pur,
                'col_sal_contract': col_contract_sal,
                'extra_party_results': extra_party_results,
            }

            self.root.after(0, lambda: self.progress.config(value=100))
            self.log('SUCCESS', f'🎉 3자 교차 대조 완료! 리포트 창을 생성합니다.')
            
            self.root.after(0, lambda: self.show_final_audit_report_dialog(audit_context))

        except Exception as e:
            err_text = str(e)
            err_msg = f'3자 교차 검증 중 오류 발생: {err_text}\n{traceback.format_exc()}'
            self.log('ERROR', err_msg)
            self.root.after(0, lambda msg=err_text: messagebox.showerror('교차 검증 오류', f'검증 중 오류가 발생했습니다:\n{msg}'))
        finally:
            restore_excel_mode(excel, excel_state)
            if excel:
                try: excel.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()

    # --------------------------------------------------------------------------
    # 3자 통합 검증 리포트 대화상자 (v7.2 - 사용자 지정 열 틀고정 & 교차 색상 Zebra Striping)
    # --------------------------------------------------------------------------
    def show_final_audit_report_dialog(self, context):
        popup = tk.Toplevel(self.root)
        popup.title('📋 3자간(발주-매입-매출) 교차 검증 및 정제 최종 보고서 (v7.2 완벽 고도화)')
        popup.geometry('1480x850')
        popup.grab_set()

        main_f = ttk.Frame(popup, padding=15)
        main_f.pack(fill='both', expand=True)

        report_data = context['report']
        extra_party_results = context.get('extra_party_results', [])
        total_cnt = len(report_data)
        success_cnt = sum(1 for x in report_data if x['final_judge'] == '성공')
        mismatch_cnt = total_cnt - success_cnt
        rate = (success_cnt / total_cnt * 100.0) if total_cnt > 0 else 0.0

        top_cards = ttk.Frame(main_f, style='Card.TFrame', padding=12)
        top_cards.pack(fill='x', pady=(0, 8))

        stat_lbl = (
            f"전체 계약 {total_cnt}건 중   |   "
            f"✔ 완전 매칭: {success_cnt}건   |   "
            f"⚠️ 미매칭/불일치: {mismatch_cnt}건   |   "
            f"성공률: {rate:.1f}%"
        )

        f_title = ttk.Frame(top_cards)
        f_title.pack(fill='x', pady=(0, 4))
        ttk.Label(f_title, text='📊 n자 확장 교차 검증 종합 현황 리포트 (기본 3자 + 사용자 추가 n자)', font=('Segoe UI', 12, 'bold'), foreground='#1A365D').pack(side='left')

        def copy_full_report_to_clipboard():
            extra_headers = []
            for party in extra_party_results:
                extra_headers.extend([f'[{party["label"]}] 실제 매칭명', f'{party["label"]} 상태', f'{party["label"]} 발행구분'])
            tsv_lines = [
                f"[n자 확장 교차 검증 종합 보고서]",
                stat_lbl,
                "",
                "\t".join(["No", "발주 계약명 (기준)", "PO번호", "[매입] 실제 매칭 계약명", "매입 상태", "매입 발행구분", "매입 계정코드", "[매출] 실제 매칭 적요/계약명", "매출 상태", "매출 발행구분"] + extra_headers + ["최종 판정"])
            ]
            for it in report_data:
                extra_values = []
                for party in extra_party_results:
                    status = it.get('extra_statuses', {}).get(party['id'], {})
                    extra_values.extend([
                        status.get('matched_text', '-'),
                        status.get('status', '미매칭(누락)'),
                        status.get('issue_type', '-'),
                    ])
                row_values = [
                    it['idx'], it['ref_contract'], it['ref_po_no'],
                    it['pur_matched_text'], it['pur_status'], it['pur_issue_type'], it['pur_contract_type'],
                    it['sal_matched_text'], it['sal_status'], it['sal_issue_type'],
                    *extra_values,
                    it['final_judge'],
                ]
                tsv_lines.append('\t'.join(str(v) for v in row_values))

            full_tsv = '\n'.join(tsv_lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(full_tsv)
            self.root.update()
            messagebox.showinfo('클립보드 복사 완료', '전체 3자간 교차 검증 보고서(탭 구분)가 클립보드에 복사되었습니다!\n[Ctrl+V]를 누르면 엑셀, 메모장, 메신저에 바로 붙여넣기할 수 있습니다.', parent=popup)

        btn_copy_all = ttk.Button(f_title, text='📋 전체 결과 엑셀용 클립보드 복사', style='Copy.TButton', command=copy_full_report_to_clipboard)
        btn_copy_all.pack(side='right')

        def toggle_status(event=None):
            selected = tree.selection()
            if not selected:
                if event is None:
                    messagebox.showwarning('선택 안됨', '판정 상태를 변경할 항목을 표에서 선택해주세요.', parent=popup)
                return
            for item_id in selected:
                curr_vals = list(tree.item(item_id, 'values'))
                idx_num = int(curr_vals[0]) - 1
                curr_judge = report_data[idx_num]['final_judge']
                new_judge = '불일치' if curr_judge == '성공' else '성공'
                report_data[idx_num]['final_judge'] = new_judge
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
        ent_filter = ttk.Entry(tool_bar, textvariable=var_filter, width=26, font=('Segoe UI', 9))
        ent_filter.pack(side='left', padx=(0, 10))

        lbl_filter_count = ttk.Label(tool_bar, text='표시 중: 0 / 0 건', font=('Segoe UI', 9), foreground='#4A5568')
        lbl_filter_count.pack(side='left', padx=(0, 15))

        ttk.Label(tool_bar, text='📌 고정할 열 수:', font=('Segoe UI', 9, 'bold')).pack(side='left', padx=(8, 4))
        var_freeze_count = tk.IntVar(value=2)
        spn_freeze = ttk.Spinbox(tool_bar, from_=1, to=10, textvariable=var_freeze_count, width=4, font=('Segoe UI', 9, 'bold'))
        spn_freeze.pack(side='left', padx=(0, 6))

        btn_freeze_apply = ttk.Button(tool_bar, text='📌 틀고정 적용', command=lambda: apply_freeze_columns())
        btn_freeze_apply.pack(side='left', padx=(0, 10))

        def copy_selected_rows_to_clipboard(event=None):
            selected = tree.selection()
            if not selected:
                if event is None:
                    messagebox.showwarning('선택 없음', '클립보드에 복사할 항목을 표에서 선택해 주세요.', parent=popup)
                return
            lines = ['\t'.join(COL_TITLES[c] for c in active_display_columns if not (c.startswith('sep') or c.endswith('_sep')))]
            for s_id in selected:
                vals = [v for v in tree.item(s_id, 'values') if v != '│']
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

        # 탭 1: 표 형식 보기
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text=' 📋 표 형식 보기 (3자간 매칭 대조 표) ')

        banner_frame = tk.Frame(tab1, bg='#1A365D', height=28)
        banner_frame.pack(fill='x', pady=(0, 4))

        tk.Label(banner_frame, text='[1] 발주내역 (대상 기준)', bg='#1A365D', fg='#63B3ED', font=('Segoe UI', 9, 'bold'), width=36, anchor='w').pack(side='left', padx=(10, 0))
        tk.Label(banner_frame, text='│', bg='#1A365D', fg='#CBD5E0', font=('Segoe UI', 11, 'bold')).pack(side='left')
        tk.Label(banner_frame, text='[2] 매입월정산 (실제 매칭 대조)', bg='#1A365D', fg='#F6AD55', font=('Segoe UI', 9, 'bold'), width=48, anchor='center').pack(side='left')
        tk.Label(banner_frame, text='│', bg='#1A365D', fg='#CBD5E0', font=('Segoe UI', 11, 'bold')).pack(side='left')
        tk.Label(banner_frame, text='[3] 매출월정산 (실제 매칭 대조)', bg='#1A365D', fg='#68D391', font=('Segoe UI', 9, 'bold'), width=40, anchor='center').pack(side='left')
        tk.Label(banner_frame, text='│', bg='#1A365D', fg='#CBD5E0', font=('Segoe UI', 11, 'bold')).pack(side='left')
        tk.Label(banner_frame, text='[4] 판정', bg='#1A365D', fg='#FC8181', font=('Segoe UI', 9, 'bold'), width=12, anchor='center').pack(side='left')

        DEFAULT_COLS = (
            'idx', 'ref_contract', 'ref_po_no', 'sep1',
            'pur_matched_text', 'pur_status', 'pur_issue_type', 'pur_contract_type', 'sep2',
            'sal_matched_text', 'sal_status', 'sal_issue_type', 'sep3',
            *[col for party in extra_party_results for col in (f'extra_{party["id"]}_text', f'extra_{party["id"]}_status', f'extra_{party["id"]}_issue', f'extra_{party["id"]}_sep')],
            'final_judge'
        )

        COL_TITLES = {
            'idx': 'No',
            'ref_contract': '발주 계약명 (기준)',
            'ref_po_no': 'PO번호',
            'sep1': '│',
            'pur_matched_text': '[매입] 실제 매칭 계약명',
            'pur_status': '매입 상태',
            'pur_issue_type': '매입 발행구분',
            'pur_contract_type': '매입 계정코드',
            'sep2': '│',
            'sal_matched_text': '[매출] 실제 매칭 적요/계약명',
            'sal_status': '매출 상태',
            'sal_issue_type': '매출 발행구분',
            'sep3': '│',
            'final_judge': '최종 판정'
        }
        for party in extra_party_results:
            COL_TITLES[f'extra_{party["id"]}_text'] = f'[{party["label"]}] 실제 매칭명'
            COL_TITLES[f'extra_{party["id"]}_status'] = f'{party["label"]} 상태'
            COL_TITLES[f'extra_{party["id"]}_issue'] = f'{party["label"]} 발행구분'
            COL_TITLES[f'extra_{party["id"]}_sep'] = '│'

        DEFAULT_WIDTHS = {
            'idx': 40, 'ref_contract': 190, 'ref_po_no': 120, 'sep1': 15,
            'pur_matched_text': 210, 'pur_status': 85, 'pur_issue_type': 90, 'pur_contract_type': 95, 'sep2': 15,
            'sal_matched_text': 210, 'sal_status': 85, 'sal_issue_type': 90, 'sep3': 15,
            'final_judge': 85
        }
        for party in extra_party_results:
            DEFAULT_WIDTHS[f'extra_{party["id"]}_text'] = 210
            DEFAULT_WIDTHS[f'extra_{party["id"]}_status'] = 95
            DEFAULT_WIDTHS[f'extra_{party["id"]}_issue'] = 95
            DEFAULT_WIDTHS[f'extra_{party["id"]}_sep'] = 15

        tree_frame = ttk.Frame(tab1)
        tree_frame.pack(fill='both', expand=True, pady=(0, 6))

        tree = ttk.Treeview(tree_frame, columns=DEFAULT_COLS, show='headings', height=16)

        tree.tag_configure('EVEN_SUCCESS', background='#FFFFFF', foreground='#1A365D')
        tree.tag_configure('ODD_SUCCESS', background='#F0FFF4', foreground='#22543D')
        tree.tag_configure('EVEN_MISMATCH', background='#FFF5F5', foreground='#9B2C2C', font=('Segoe UI', 9, 'bold'))
        tree.tag_configure('ODD_MISMATCH', background='#FED7D7', foreground='#742A2A', font=('Segoe UI', 9, 'bold'))

        y_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        x_scrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        tree.grid(row=0, column=0, sticky='nsew')
        y_scrollbar.grid(row=0, column=1, sticky='ns')
        x_scrollbar.grid(row=1, column=0, sticky='ew')
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        sort_states = {c: False for c in DEFAULT_COLS}
        active_display_columns = list(DEFAULT_COLS)
        manageable_order = [c for c in DEFAULT_COLS if not (c.startswith('sep') or c.endswith('_sep'))]
        visible_columns = {c: True for c in manageable_order}
        column_groups = {
            '기준 발주내역': ['idx', 'ref_contract', 'ref_po_no'],
            '기본 3자 - 매입월정산': ['pur_matched_text', 'pur_status', 'pur_issue_type', 'pur_contract_type'],
            '기본 3자 - 매출월정산': ['sal_matched_text', 'sal_status', 'sal_issue_type'],
            '최종 판정': ['final_judge'],
        }
        for party in extra_party_results:
            column_groups[f'n자 - {party["label"]}'] = [
                f'extra_{party["id"]}_text',
                f'extra_{party["id"]}_status',
                f'extra_{party["id"]}_issue',
            ]

        def separator_after(col):
            if col == 'ref_po_no':
                return 'sep1'
            if col == 'pur_contract_type':
                return 'sep2'
            if col == 'sal_issue_type':
                return 'sep3'
            for party in extra_party_results:
                if col == f'extra_{party["id"]}_issue':
                    return f'extra_{party["id"]}_sep'
            return None

        def rebuild_display_columns():
            nonlocal active_display_columns
            columns = []
            for c in manageable_order:
                if not visible_columns.get(c, True):
                    continue
                columns.append(c)
                sep = separator_after(c)
                if sep and sep in DEFAULT_COLS:
                    columns.append(sep)
            active_display_columns = columns
            tree['displaycolumns'] = tuple(active_display_columns)
            apply_freeze_columns()

        def sort_by_column(col):
            if col.startswith('sep') or col.endswith('_sep'): return
            reverse = not sort_states[col]
            sort_states[col] = reverse

            items = [(tree.set(k, col), k) for k in tree.get_children('')]

            def sort_key(item):
                val = item[0]
                try:
                    return (0, float(val))
                except ValueError:
                    return (1, str(val))

            items.sort(key=sort_key, reverse=reverse)

            for index, (val, k) in enumerate(items):
                tree.move(k, '', index)

            for c in DEFAULT_COLS:
                if c.startswith('sep') or c.endswith('_sep'): continue
                base_title = COL_TITLES[c]
                if c == col:
                    indicator = ' ▲' if not reverse else ' ▼'
                    tree.heading(c, text=base_title + indicator)
                else:
                    tree.heading(c, text=base_title)

        for c in DEFAULT_COLS:
            title = COL_TITLES[c]
            tree.heading(c, text=title, command=lambda _c=c: sort_by_column(_c))
            tree.column(c, width=DEFAULT_WIDTHS[c], anchor='center' if ('status' in c or 'type' in c or 'issue' in c or c in ('idx', 'ref_po_no', 'sep1', 'sep2', 'sep3', 'final_judge') or c.endswith('_sep')) else 'w')

        lbl_summary = None

        def apply_filter(*args):
            nonlocal lbl_summary
            query = var_filter.get().strip().lower()
            tree.delete(*tree.get_children())

            cur_success = sum(1 for x in report_data if x['final_judge'] == '성공')
            cur_mismatch = sum(1 for x in report_data if x['final_judge'] != '성공')
            cur_total = len(report_data)
            cur_rate = (cur_success / cur_total * 100.0) if cur_total > 0 else 0.0

            lbl_top_stat.config(text=f"전체 계약 {cur_total}건 중   |   ✔ 완전 매칭: {cur_success}건   |   ⚠️ 미매칭/불일치: {cur_mismatch}건   |   성공률: {cur_rate:.1f}%")
            if lbl_summary is not None:
                lbl_summary.config(text=f'전체 계약 {cur_total}건 중  |  ✔ 완전 매칭: {cur_success}건  |  ⚠️ 미매칭/불일치: {cur_mismatch}건')

            visible_count = 0
            for item in report_data:
                row_str = ' '.join(str(v) for v in item.values()).lower()
                if not query or query in row_str:
                    is_even = (visible_count % 2 == 0)
                    if item['final_judge'] == '성공':
                        tag = 'EVEN_SUCCESS' if is_even else 'ODD_SUCCESS'
                    else:
                        tag = 'EVEN_MISMATCH' if is_even else 'ODD_MISMATCH'

                    extra_values = []
                    for party in extra_party_results:
                        status = item.get('extra_statuses', {}).get(party['id'], {})
                        extra_values.extend([
                            status.get('matched_text', '-'),
                            status.get('status', '미매칭(누락)'),
                            status.get('issue_type', '-'),
                            '│',
                        ])

                    tree.insert('', 'end', values=(
                        item['idx'],
                        item['ref_contract'],
                        item['ref_po_no'],
                        '│',
                        item['pur_matched_text'],
                        item['pur_status'],
                        item['pur_issue_type'],
                        item['pur_contract_type'],
                        '│',
                        item['sal_matched_text'],
                        item['sal_status'],
                        item['sal_issue_type'],
                        '│',
                        *extra_values,
                        item['final_judge']
                    ), tags=(tag,))
                    visible_count += 1

            lbl_filter_count.config(text=f'표시 중: {visible_count} / {len(report_data)} 건')

        var_filter.trace_add('write', apply_filter)

        def apply_freeze_columns():
            try:
                cnt = int(var_freeze_count.get())
            except Exception:
                cnt = 2

            cnt = max(1, min(cnt, 8))
            var_freeze_count.set(cnt)

            non_sep_cols = [c for c in active_display_columns if not (c.startswith('sep') or c.endswith('_sep'))]
            frozen_data_cols = non_sep_cols[:cnt]

            frozen_layout = []
            for c in active_display_columns:
                if c in frozen_data_cols or ((c.startswith('sep') or c.endswith('_sep')) and frozen_layout):
                    frozen_layout.append(c)
                    if len([x for x in frozen_layout if not (x.startswith('sep') or x.endswith('_sep'))]) == cnt:
                        break

            remaining_layout = [c for c in active_display_columns if c not in frozen_layout]
            tree['displaycolumns'] = tuple(frozen_layout + remaining_layout)

        apply_freeze_columns()
        apply_filter()

        tree.bind('<Double-1>', toggle_status)
        tree.bind('<Control-c>', copy_selected_rows_to_clipboard)

        def open_column_manager_dialog():
            mgr_pop = tk.Toplevel(popup)
            mgr_pop.title('⚙️ 테이블 열 숨기기/보이기 & 순서 배치 관리')
            mgr_pop.geometry('680x620')
            mgr_pop.attributes('-topmost', True)

            f_mgr = ttk.Frame(mgr_pop, padding=15)
            f_mgr.pack(fill='both', expand=True)

            ttk.Label(f_mgr, text='⚙️ 그룹별 열 숨기기/보이기 및 순서 지정', font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(0, 6))

            checkbox_outer = ttk.Frame(f_mgr)
            checkbox_outer.pack(fill='both', expand=True, pady=(0, 10))

            cb_canvas = tk.Canvas(checkbox_outer, highlightthickness=0)
            cb_scroll = ttk.Scrollbar(checkbox_outer, orient='vertical', command=cb_canvas.yview)
            cb_canvas.configure(yscrollcommand=cb_scroll.set)
            cb_canvas.pack(side='left', fill='both', expand=True)
            cb_scroll.pack(side='right', fill='y')

            cb_frame = ttk.Frame(cb_canvas)
            cb_window = cb_canvas.create_window((0, 0), window=cb_frame, anchor='nw')

            def _on_cb_configure(event=None):
                cb_canvas.configure(scrollregion=cb_canvas.bbox('all'))
                cb_canvas.itemconfigure(cb_window, width=cb_canvas.winfo_width())

            cb_frame.bind('<Configure>', _on_cb_configure)
            cb_canvas.bind('<Configure>', _on_cb_configure)

            original_visible_columns = visible_columns.copy()
            original_manageable_order = list(manageable_order)
            temp_visible_vars = {c: tk.BooleanVar(value=visible_columns.get(c, True)) for c in manageable_order}

            def preview_visibility_changes():
                for col_key, var in temp_visible_vars.items():
                    visible_columns[col_key] = bool(var.get())
                rebuild_display_columns()

            def cancel_col_changes():
                nonlocal manageable_order
                visible_columns.clear()
                visible_columns.update(original_visible_columns)
                manageable_order = list(original_manageable_order)
                rebuild_display_columns()
                mgr_pop.destroy()

            mgr_pop.protocol('WM_DELETE_WINDOW', cancel_col_changes)

            for group_name, cols in column_groups.items():
                group_box = ttk.LabelFrame(cb_frame, text=group_name, padding=8)
                group_box.pack(fill='x', pady=(0, 8))
                for c in cols:
                    if c in temp_visible_vars:
                        ttk.Checkbutton(group_box, text=f'{COL_TITLES[c]} ({c})', variable=temp_visible_vars[c], command=preview_visibility_changes).pack(anchor='w', pady=1)

            ttk.Label(f_mgr, text='↕ 열 순서 조정 (선택한 열을 좌/우 방향으로 이동)', font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(2, 4))
            manageable_cols = list(manageable_order)

            lbl_box = tk.Listbox(f_mgr, font=('Segoe UI', 10), selectmode='single', height=12)
            lbl_box.pack(fill='both', expand=False, pady=(0, 10))

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
                nonlocal manageable_order
                manageable_order = list(manageable_cols)
                for c, var in temp_visible_vars.items():
                    visible_columns[c] = bool(var.get())
                rebuild_display_columns()
                mgr_pop.destroy()

            f_mgr_btns = ttk.Frame(f_mgr)
            f_mgr_btns.pack(fill='x', side='bottom')

            ttk.Button(f_mgr_btns, text='☑ 설정 적용', style='Primary.TButton', command=apply_col_changes).pack(side='right', padx=(6, 0))
            ttk.Button(f_mgr_btns, text='취소', command=cancel_col_changes).pack(side='right')

        def reset_column_layout_and_filter():
            nonlocal active_display_columns, manageable_order
            var_filter.set('')
            var_freeze_count.set(2)
            manageable_order = [c for c in DEFAULT_COLS if not (c.startswith('sep') or c.endswith('_sep'))]
            for c in visible_columns:
                visible_columns[c] = True
            active_display_columns = list(DEFAULT_COLS)
            tree['displaycolumns'] = tuple(DEFAULT_COLS)

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
        txt_audit_report.insert(tk.END, f"                    📋 n자 확장 교차 검증 최종 보고서\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"===========================================================================================\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"{stat_lbl}\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"-------------------------------------------------------------------------------------------\n", 'LINE_EVEN')
        extra_text_headers = ''.join(f" | {party['label']} 매칭명                     | {party['label']}상태" for party in extra_party_results)
        txt_audit_report.insert(tk.END, f"No | 발주 계약명 (기준)                 | PO번호       | 매입 계약명                     | 매입상태 | 매출 계약명                     | 매출상태{extra_text_headers} | 최종판정\n", 'LINE_EVEN')
        txt_audit_report.insert(tk.END, f"-------------------------------------------------------------------------------------------\n", 'LINE_EVEN')

        for idx, it in enumerate(report_data):
            line_tag = 'LINE_EVEN' if (idx % 2 == 0) else 'LINE_ODD'
            st_p = '✅성공' if it['pur_matched'] else '❌누락'
            st_s = '✅성공' if it['sal_matched'] else '❌누락'
            extra_text_values = ''
            for party in extra_party_results:
                status = it.get('extra_statuses', {}).get(party['id'], {})
                st_x = '✅성공' if status.get('matched') else '❌누락'
                extra_text_values += f" | {status.get('matched_text', '-'):<30s} | {st_x}"
            line_str = f"{it['idx']:02d} | {it['ref_contract']:<30s} | {it['ref_po_no']:<12s} | {it['pur_matched_text']:<30s} | {st_p} | {it['sal_matched_text']:<30s} | {st_s}{extra_text_values} | {it['final_judge']}\n"
            txt_audit_report.insert(tk.END, line_str, line_tag)

        txt_audit_report.insert(tk.END, f"===========================================================================================\n", 'LINE_EVEN')

        # 탭 3: 사용된 헤더 매핑 정보
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text=' ⚙️ 사용된 헤더 매핑 정보 ')

        txt_mapping_info = scrolledtext.ScrolledText(tab3, font=('Consolas', 10), bg='#FFFFFF', fg='#1A202C')
        txt_mapping_info.pack(fill='both', expand=True)

        txt_mapping_info.insert(tk.END, f"===========================================================================================\n")
        txt_mapping_info.insert(tk.END, f"                    ⚙️ 이번 3자 교차 대조에 사용된 헤더 매핑 설정\n")
        txt_mapping_info.insert(tk.END, f"===========================================================================================\n")
        txt_mapping_info.insert(tk.END, f"No | 항목 명칭           | 발주 Col | 매입 Col | 매출 Col | 모드\n")
        txt_mapping_info.insert(tk.END, f"-------------------------------------------------------------------------------------------\n")

        for idx, m in enumerate(self.header_mappings):
            ref_c = m.get('ref_col', 1)
            pur_c = m.get('pur_col', 1)
            sal_c = m.get('sal_col', 1)
            lbl = m.get('field_label', f"Custom {idx+1}")
            mode_str = m.get('mode', '자동')
            txt_mapping_info.insert(tk.END, f"{idx+1:02d} | {lbl:<20s} | Col {ref_c:3d}  | Col {pur_c:3d}  | Col {sal_c:3d}  | {mode_str}\n")

        if extra_party_results:
            txt_mapping_info.insert(tk.END, f"-------------------------------------------------------------------------------------------\n")
            txt_mapping_info.insert(tk.END, f"n자 자동 헤더 매칭 현황\n")
            for party in self.extra_parties:
                mapping = party.get('mapping', {})
                txt_mapping_info.insert(
                    tk.END,
                    f"- {party['label']}: 계약/적요 Col {mapping.get('contract_col', 11)}, PO/식별 Col {mapping.get('po_col', 27)}, 발행구분 Col {mapping.get('issue_col', 22)}, 유형 Col {mapping.get('type_col', 10)}\n"
                )

        txt_mapping_info.insert(tk.END, f"===========================================================================================\n")
        txt_mapping_info.config(state='disabled')

        f_bottom = ttk.Frame(main_f)
        f_bottom.pack(fill='x', pady=(4, 0))

        lbl_summary = ttk.Label(f_bottom, text=f'전체 계약 {total_cnt}건 중  |  ✔ 완전 매칭: {success_cnt}건  |  ⚠️ 미매칭/불일치: {mismatch_cnt}건', font=('Segoe UI', 10, 'bold'), foreground=self.COLOR_PRIMARY)
        lbl_summary.pack(side='left')

        def commit_and_close():
            if messagebox.askyesno('엑셀 커밋 승인', '매입 및 매출 엑셀 파일의 비대상 행 삭제 및 누락 항목(PO번호 입력 - 노란색 셀 바로 좌측) 커밋을 실행하시겠습니까?'):
                popup.destroy()
                self.log('INFO', '🚀 엑셀 자동 데이터 가공 및 커밋 트랜잭션을 시작합니다...')
                threading.Thread(target=self.commit_to_excel_process, args=(context,), daemon=True).start()

        ttk.Button(f_bottom, text='💾 엑셀 자동 데이터 가공 및 커밋 실행', style='Primary.TButton', command=commit_and_close).pack(side='right', padx=(10, 0))
        ttk.Button(f_bottom, text='닫기', command=popup.destroy).pack(side='right')

    # --------------------------------------------------------------------------
    # 엑셀 물리적 데이터 가공 & 커밋 엔진 (v7.2 PO번호 노란색 셀 바로 좌측 삽입)
    # --------------------------------------------------------------------------
    def commit_to_excel_process(self, context):
        ensure_excel_com_available()
        pythoncom.CoInitialize()
        excel = None
        excel_state = None
        wb_pur = None
        wb_sal = None
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
            excel.Visible = False
            excel_state = set_excel_fast_mode(excel)

            pur_path = os.path.abspath(self.purchase_file_path.get())
            sal_path = os.path.abspath(self.sales_file_path.get())

            self.log('INFO', '매입월정산 엑셀 파일 가공 중...')
            pur_backup_path = create_timestamped_backup(pur_path)
            self.log('SUCCESS', f'매입월정산 백업 생성 완료: {os.path.basename(pur_backup_path)}')
            
            wb_pur = excel.Workbooks.Open(pur_path)
            ws_pur = wb_pur.Worksheets(1)

            pur_del_rows = sorted(context['pur_rows_to_delete'], reverse=True)
            deleted_pur_count = delete_rows_or_raise(ws_pur, pur_del_rows, '매입월정산')
            
            self.log('SUCCESS', f'매입월정산 비대상 행 {deleted_pur_count}개 역방향 삭제 완료')

            col_pur_c = context['col_pur_contract']
            pur_start_row = ws_pur.UsedRange.Rows.Count + 1

            for idx, missing in enumerate(context['pur_missing_items']):
                target_r = pur_start_row + idx
                cell_c = ws_pur.Cells(target_r, col_pur_c)
                cell_c.Value = missing['contract_name']
                cell_c.Interior.Color = 65535 # Yellow

                if col_pur_c > 1:
                    cell_p = ws_pur.Cells(target_r, col_pur_c - 1) # 바로 좌측 셀
                    po_val = clean_po_val(missing.get('po_no', ''))
                    cell_p.Value = str(po_val if po_val else '빈셀임')
                    cell_p.Interior.Color = 65535 # Yellow

            wb_pur.Save()
            wb_pur.Close(True)
            self.log('SUCCESS', f'매입월정산 미매칭 계약명 {len(context["pur_missing_items"])}건 하단 추가 (노란색 셀 바로 좌측 PO번호 삽입) 완료')

            self.log('INFO', '매출월정산 엑셀 파일 가공 중...')
            sal_backup_path = create_timestamped_backup(sal_path)
            self.log('SUCCESS', f'매출월정산 백업 생성 완료: {os.path.basename(sal_backup_path)}')

            wb_sal = excel.Workbooks.Open(sal_path)
            ws_sal = wb_sal.Worksheets(1)

            sal_del_rows = sorted(context['sal_rows_to_delete'], reverse=True)
            deleted_sal_count = delete_rows_or_raise(ws_sal, sal_del_rows, '매출월정산')

            self.log('SUCCESS', f'매출월정산 비대상 행 {deleted_sal_count}개 역방향 삭제 완료')

            col_sal_c = context['col_sal_contract']
            sal_start_row = ws_sal.UsedRange.Rows.Count + 1

            for idx, missing in enumerate(context['sal_missing_items']):
                target_r = sal_start_row + idx
                cell_c = ws_sal.Cells(target_r, col_sal_c)
                cell_c.Value = missing['contract_name']
                cell_c.Interior.Color = 65535 # Yellow

                if col_sal_c > 1:
                    cell_p = ws_sal.Cells(target_r, col_sal_c - 1) # 바로 좌측 셀
                    po_val = clean_po_val(missing.get('po_no', ''))
                    cell_p.Value = str(po_val if po_val else '빈셀임')
                    cell_p.Interior.Color = 65535 # Yellow

            wb_sal.Save()
            wb_sal.Close(True)
            self.log('SUCCESS', f'매출월정산 미매칭 계약명 {len(context["sal_missing_items"])}건 하단 추가 (노란색 셀 바로 좌측 PO번호 삽입) 완료')

            self.log('SUCCESS', '🎉🎉 모든 매입/매출월정산 엑셀 파일의 자동 데이터 정제 및 커밋(PO번호 바로 좌측 입력)이 성공적으로 완수되었습니다!')
            self.root.after(0, lambda: messagebox.showinfo('커밋 완료', '매입/매출월정산 엑셀 파일의 정제 가공 및 불일치 PO번호 노란색 셀 바로 좌측 입력이 완벽하게 완료되었습니다!'))

        except Exception as e:
            err_text = str(e)
            err_msg = f'엑셀 커밋 가공 중 오류 발생: {err_text}\n{traceback.format_exc()}'
            self.log('ERROR', err_msg)
            self.root.after(0, lambda msg=err_text: messagebox.showerror('커밋 오류', f'엑셀 반영 중 오류가 발생했습니다:\n{msg}'))
        finally:
            for wb in (wb_pur, wb_sal):
                if wb:
                    try: wb.Close(False)
                    except Exception: pass
            restore_excel_mode(excel, excel_state)
            if excel:
                try: excel.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()


def main():
    root = tk.Tk()
    B2BThreeWayReconciliationApp(root)
    root.mainloop()

if __name__ == '__main__':
    try:
        main()
    except Exception:
        traceback.print_exc()
        input('오류가 발생했습니다. Enter 키를 눌러 종료하세요...')
