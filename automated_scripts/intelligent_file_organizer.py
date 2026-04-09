"""
================================================================================
 [지능형 파일 정리기 (Intelligent File Organizer) v34.1.16] (Ultimate Master)
================================================================================
- 아키텍처: Clean Layer Architecture (Domain / Presentation / Application)
- 주요 기능: 파일명 패턴 매칭, 일괄 변경, 충돌 방지 미리보기 기반 자동 정리
- 가이드라인 준수: 00 PRD 가이드.md | AI_CODING_GUIDELINES_2026.md
- 무결성 보증: 실행 전 중복 검증 및 원자적 이름 변경 시뮬레이션 지원
================================================================================
"""
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import fnmatch
import shutil
import csv

USER_MANUAL_TEXT = """# 📘 지능형 파일 관리 시스템 - 사용 매뉴얼

> 버전: v34.1.25 | 최종 수정: 2026-04-08

---

## 🎯 이 프로그램은 뭘 하는 건가요?
파일 이름을 한꺼번에 바꿔주는 도구입니다.

예를 들어, 이런 파일들이 있다면:
보고서_최종.xlsx
보고서_수정본.xlsx
회의록_0401.docx

이렇게 바꿀 수 있습니다:
01_보고서_최종.xlsx
02_보고서_수정본.xlsx
04_회의록_0401.docx

---

## 🚀 기본 사용법 (5단계)
1. 파일 불러오기: [폴더 선택] 또는 [개별 파일 선택]으로 목록을 채웁니다.
2. 규칙 선택: 오른쪽 패널에서 '1:형식', '2:교체' 등 원하는 규칙을 고릅니다.
3. 파라미터 설정: 아래 나타나는 옵션(찾을 단어, 붙일 번호 등)을 채웁니다.
4. 규칙 추가 → 미리보기: [개별 규칙 추가]를 누르고, [미리보기 업데이트]로 리스트에서 어떻게 바뀔지 확인합니다.
5. 일괄 변경 실행: 결과가 만족스럽다면 [일괄 변경 실행]을 눌러 실제 파일 이름을 바꿉니다.

---

## 📖 규칙별 상세 사용법

### 1:형식 (번호 매기기)
* 언제? 연속된 번호를 파일 앞에 붙일 때
* 옵션:
  - 시작 번호, 자리수(예: 01, 02) 지정
* 기호 사용:
  - /n = 원래 이름
  - /01 = 순서 번호
  - /Y, /M, /D = 연, 월, 일
* 예시: 형식 "/01_/n" (시작 1, 자리 2)
  보고서.xlsx → 01_보고서.xlsx

### 2:교체 (글자 바꾸기)
* 언제? "최종"을 "완료"로 바꾸거나, "(복사본)"을 지울 때
* 예시: 찾기="_최종", 바꾸기="_완료"

### 3:삽입 / 4:삭제
* 지정된 위치(앞에서 0번째 등)에 글자를 끼워넣거나 잘라냅니다.

### ⭐ 6:접두사(이동/복사)
* 언제? 기존 VBA 기능 대체용으로, 파일 첫 숫자를 기준으로 자동으로 같은 숫자의 폴더를 찾아 정리할 때.
* 작동방식: 
  "03_소방자재.xlsx" -> [목적지 폴더]\03 소방공사\01 자료\표준단가_소방자재.xlsx 로 이동/복사
* 설정 항목:
  - 목적지(Base): "01 건축공사", "02 전기공사" 등 번호 폴더들이 속한 최상위 폴더
  - 하위경로(Sub): 이동할 안쪽 경로(예: "01 자료")
  - 추가 접두사: 복사될 때 앞에 새로 붙일 텍스트(예: "표준_")
"""

# ═══════════════════════════════════════════════════════════
# LAYER 1: DOMAIN (Engine - Pure Business Logic)
# ═══════════════════════════════════════════════════════════

class FileOrganizerEngine:
    def __init__(self):
        self.stop_requested = False

    def scan_files(self, directory, recursive=True, pattern="*", callback=None):
        """
        Scan directory for files based on pattern.
        """
        results = []
        try:
            for root, dirs, files in os.walk(directory):
                if not recursive and root != directory:
                    continue
                
                for file in files:
                    if fnmatch.fnmatch(file, pattern):
                        full_path = os.path.join(root, file)
                        stats = os.stat(full_path)
                        results.append({
                            'name': file,
                            'path': root,
                            'full_path': full_path,
                            'ext': os.path.splitext(file)[1],
                            'size': stats.st_size,
                            'mtime': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                            'ctime': datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
                        })
                
                if self.stop_requested:
                    break
                    
            if callback:
                callback('scan_complete', results)
            return results
        except Exception as e:
            if callback:
                callback('error', str(e))
            return []

    def detect_common_patterns(self, files):
        """
        Intelligently detect common substrings that could be used as markers.
        """
        if not files or len(files) < 2:
            return []
            
        patterns = set()
        for f in files:
            matches = re.findall(r'(\([^\)]+)', f['name'])
            for m in matches:
                patterns.add(m)
            
            date_match = re.search(r'\(\d{2}\.\d{2}\.\d{2}\)', f['name'])
            if date_match:
                patterns.add(date_match.group())
                
        sorted_patterns = sorted(list(patterns), key=len, reverse=True)
        return sorted_patterns[:10]

    def preview_rename(self, files, rules):
        """
        Apply a sequence of rules to files and return previews with caching optimizations.
        """
        previews = []
        planned_names = {}  # {new_full_path: original_name}
        file_counter = 0
        
        # [최적화] 목적지 폴더 캐시 (분류 규칙 사용 시 디스크 I/O 절감)
        dest_cache = {} # {base_dest: [folder_entries]}

        for file_info in files:
            if self.stop_requested:
                break
                
            current_name, ext = os.path.splitext(file_info['name'])
            file_counter += 1
            current_rules = [] 
            
            # 이전 분류 정보 초기화
            if '_is_distribute' in file_info: del file_info['_is_distribute']
            if '_dist_dest_base' in file_info: del file_info['_dist_dest_base']
            
            for rule in rules:
                r_type = rule['type']
                params = rule['params']
                
                if r_type == 'marker_truncate':
                    marker = params.get('marker', '')
                    suffix = params.get('suffix', '')
                    if marker and marker in current_name:
                        idx = current_name.find(marker)
                        current_name = current_name[:idx + len(marker)] + suffix
                
                elif r_type == 'regex_replace':
                    pattern = params.get('pattern', '')
                    replacement = params.get('replacement', '')
                    try:
                        current_name = re.sub(pattern, replacement, current_name)
                    except: pass
                
                elif r_type == 'simple_replace':
                    old_str = params.get('old_str', '')
                    new_str = params.get('new_str', '')
                    current_name = current_name.replace(old_str, new_str)

                elif r_type == 'normalize':
                    current_name = re.sub(r'[\\/:*?"<>|]', '', current_name)
                    if params.get('space_to_under'):
                        current_name = current_name.replace(' ', '_')
                    if params.get('to_upper'):
                        current_name = current_name.upper()

                elif r_type == 'format_pattern':
                    pattern_str = params.get('pattern', '/n')
                    start_num = params.get('start', 1)
                    digits = params.get('digits', 2)
                    increment = params.get('increment', 1)
                    zero_pad = params.get('zero_pad', True)
                    pos_type = params.get('pos_type', 'replace') 
                    
                    current_num = start_num + (file_counter - 1) * increment
                    num_str = str(current_num).zfill(digits) if zero_pad else str(current_num)
                    
                    result = pattern_str
                    has_explicit_n = '/n' in result
                    
                    result = result.replace('/n', current_name)
                    result = result.replace('/e', ext[1:] if ext else '')
                    now = datetime.now()
                    result = result.replace('/YMD', now.strftime('%Y%m%d'))
                    result = result.replace('/HMS', now.strftime('%H%M%S'))
                    result = re.sub(r'/(\d+)', num_str, result)
                    
                    if pos_type == 'prefix' and not has_explicit_n:
                         current_name = result + current_name
                    elif pos_type == 'suffix' and not has_explicit_n:
                         current_name = current_name + result
                    else:
                         current_name = result

                elif r_type == 'replace':
                    find_str = params.get('find', '')
                    replace_str = params.get('replace', '')
                    ignore_case = params.get('ignore_case', False)
                    max_count = params.get('max_count', 0)
                    
                    if find_str:
                        if ignore_case:
                            pattern = re.compile(re.escape(find_str), re.IGNORECASE)
                            if max_count > 0:
                                current_name = pattern.sub(replace_str, current_name, count=max_count)
                            else:
                                current_name = pattern.sub(replace_str, current_name)
                        else:
                            if max_count > 0:
                                current_name = current_name.replace(find_str, replace_str, max_count)
                            else:
                                current_name = current_name.replace(find_str, replace_str)

                elif r_type == 'insert':
                    insert_str = params.get('text', '')
                    position = params.get('position', 0)
                    from_end = params.get('from_end', False)
                    
                    if insert_str:
                        if from_end:
                            actual_pos = max(0, len(current_name) - position)
                        else:
                            actual_pos = min(position, len(current_name))
                        current_name = current_name[:actual_pos] + insert_str + current_name[actual_pos:]

                elif r_type == 'delete':
                    position = params.get('position', 0)
                    length = params.get('length', 1)
                    from_end = params.get('from_end', False)
                    
                    if from_end:
                        actual_start = max(0, len(current_name) - position - length)
                        actual_end = max(0, len(current_name) - position)
                    else:
                        actual_start = min(position, len(current_name))
                        actual_end = min(position + length, len(current_name))
                    current_name = current_name[:actual_start] + current_name[actual_end:]

                rule_name_map = {
                    'format_pattern': '1:형식',
                    'replace': '2:교체',
                    'insert': '3:삽입',
                    'delete': '4:삭제',
                    'case_convert': '5:대소',
                    'prefix_distribute': '6:분류',
                    'marker_truncate': '자르기'
                }
                current_rules.append(rule_name_map.get(r_type, r_type))

                if r_type == 'prefix_distribute':
                    c_prefix = params.get('prefix', '')
                    num_len = params.get('num_len', 2)
                    
                    numeric_prefix = ""
                    for char in current_name:
                        if char.isdigit():
                            numeric_prefix += char
                            if len(numeric_prefix) == num_len: break
                    
                    if len(numeric_prefix) == num_len:
                        idx = current_name.find(numeric_prefix)
                        current_name = c_prefix + current_name[idx + num_len:]
                        
                        # [최적화] 지능형 폴더 매칭 시뮬레이션 (캐시 사용)
                        dest_base = params.get('base_dest', '')
                        if dest_base and os.path.exists(dest_base):
                            if dest_base not in dest_cache:
                                try:
                                    dest_cache[dest_base] = [e.name for e in os.scandir(dest_base) if e.is_dir()]
                                except:
                                    dest_cache[dest_base] = []
                            
                            for folder_name in dest_cache[dest_base]:
                                if folder_name.startswith(numeric_prefix):
                                    file_info['_dist_dest_base'] = folder_name
                                    break

                    file_info['_is_distribute'] = True
                    if '_dist_dest_base' not in file_info:
                        file_info['_dist_dest_base'] = f"{numeric_prefix}???"

                elif r_type == 'case_convert':
                    target = params.get('target', 'name')
                    case_type = params.get('case_type', 'lower')
                    ignore_chars = params.get('ignore_chars', '')
                    
                    def convert_case(s, ctype, ignore):
                        if not s: return s
                        result = []
                        for char in s:
                            if char in ignore:
                                result.append(char)
                            elif ctype == 'lower':
                                result.append(char.lower())
                            elif ctype == 'upper':
                                result.append(char.upper())
                            else:
                                result.append(char)
                        converted = ''.join(result)
                        if ctype == 'title':
                            converted = converted.title()
                        return converted
                    
                    if target == 'name' or target == 'all':
                        current_name = convert_case(current_name, case_type, ignore_chars)

            new_name = current_name + ext
            new_full_path = os.path.join(file_info['path'], new_name)
            
            status = " + ".join(current_rules) if current_rules else "[OK]"
            
            if file_info.get('_is_distribute'):
                dest_info = file_info.get('_dist_dest_base', '??')
                status = f"6:분류 -> {dest_info}"
            elif new_name == file_info['name']:
                status = '[UNCHANGED]'
            elif new_full_path in planned_names:
                status = '[WARN] 충돌(중복)'
            elif os.path.exists(new_full_path):
                status = '[WARN] 충돌(기존파일)'

            planned_names[new_full_path] = file_info['name']
            
            previews.append({
                'original_path': file_info['full_path'],
                'original_name': file_info['name'],
                'new_name': new_name,
                'status': status,
                'size_kb': round(file_info.get('size', 0) / 1024, 2),
                'mtime': file_info.get('mtime', ''),
                'ctime': file_info.get('ctime', ''),
                'is_distribute': file_info.get('_is_distribute', False),
                'dist_dest': file_info.get('_dist_dest_base', '')  
            })
            
        return previews

    def perform_distribute(self, rename_list, rules, callback=None):
        """
        [VBA 이식] 숫자 접두사 기반 하위 폴더 자동 매칭 및 분류 실행
        """
        success_count = 0
        fail_count = 0
        import shutil

        # 'prefix_distribute' 규칙 찾기
        dist_rule = next((r for r in rules if r['type'] == 'prefix_distribute'), None)
        if not dist_rule:
             if callback: callback('rename_complete', {'success': 0, 'failed': 0, 'message': '분류 규칙이 없습니다.'})
             return

        params = dist_rule['params']
        base_dest = params.get('base_dest', '')
        sub_path = params.get('sub_path', '')
        num_len = params.get('num_len', 2)
        mode = params.get('mode', 'copy')
        custom_prefix = params.get('prefix', '')

        if not base_dest or not os.path.exists(base_dest):
            if callback: callback('rename_complete', {'success': 0, 'failed': 0, 'message': '목적지 경로가 유효하지 않습니다.'})
            return

        for item in rename_list:
            if self.stop_requested: break
            
            filename = item['original_name']
            
            # 숫자 접두사 추출
            numeric_prefix = ""
            for char in filename:
                if char.isdigit():
                    numeric_prefix += char
                    if len(numeric_prefix) == num_len: break
            
            if len(numeric_prefix) < num_len:
                fail_count += 1
                if callback: callback('item_fail', {'name': filename, 'error': f'숫자 접두사({num_len}자리) 미달'})
                continue

            # 하위 폴더 검색
            target_subfolder = None
            try:
                for entry in os.scandir(base_dest):
                    if entry.is_dir() and entry.name.startswith(numeric_prefix):
                        target_subfolder = entry.path
                        break
            except Exception as e:
                fail_count += 1
                if callback: callback('item_fail', {'name': filename, 'error': f'폴더 스캔 실패: {str(e)}'})
                continue

            if not target_subfolder:
                fail_count += 1
                if callback: callback('item_fail', {'name': filename, 'error': '일치하는 숫자 폴더 없음'})
                continue

            # 최종 경로 조합
            final_dest = os.path.join(target_subfolder, sub_path) if sub_path else target_subfolder
            
            if not os.path.exists(final_dest):
                fail_count += 1
                if callback: callback('item_fail', {'name': filename, 'error': f'하위 경로 없음: {sub_path}'})
                continue

            # 새 파일명 생성 (VBA 로직: 커스텀 접두사 + 원본에서 숫자 접두사 제외한 나머지)
            num_idx = filename.find(numeric_prefix)
            new_file_name = custom_prefix + filename[num_idx + num_len:]
            final_path = os.path.join(final_dest, new_file_name)

            # 충돌 방지
            if os.path.exists(final_path):
                base, ext = os.path.splitext(final_path)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                final_path = f"{base}_{counter}{ext}"
                new_file_name = os.path.basename(final_path)

            try:
                if mode == 'move':
                    shutil.move(item['original_path'], final_path)
                else:
                    shutil.copy2(item['original_path'], final_path)
                
                success_count += 1
                mode_verb = "복사" if mode == 'copy' else "이동"
                if callback: callback('item_success', {'old': filename, 'new': f'[{mode_verb} 완료] -> {os.path.basename(final_dest)}\\{new_file_name}'})
            except Exception as e:
                fail_count += 1
                if callback: callback('item_fail', {'name': filename, 'error': str(e)})

        if callback:
            callback('rename_complete', {'success': success_count, 'failed': fail_count})

    def perform_rename(self, rename_list, callback=None):
        """
        Execute the renaming.
        """
        success_count = 0
        fail_count = 0
        
        for item in rename_list:
            if self.stop_requested:
                break
                
            old_path = item['original_path']
            dir_path = os.path.dirname(old_path)
            new_path = os.path.join(dir_path, item['new_name'])
            
            try:
                if old_path != new_path:
                    final_path = new_path
                    if os.path.exists(final_path):
                        base, ext = os.path.splitext(new_path)
                        counter = 1
                        while os.path.exists(f"{base}_{counter}{ext}"):
                            counter += 1
                        final_path = f"{base}_{counter}{ext}"
                    
                    os.rename(old_path, final_path)
                    success_count += 1
                    if callback:
                        # [무결성] original_path를 명시적으로 전달하여 UI 동기화 속도 및 정확도 향상
                        callback('item_success', {
                            'path': item['original_path'],
                            'old': item['original_name'], 
                            'new': os.path.basename(final_path)
                        })
                else:
                    success_count += 1
            except Exception as e:
                fail_count += 1
                if callback:
                    callback('item_fail', {
                        'path': item['original_path'],
                        'name': item['original_name'], 
                        'error': str(e)
                    })
        
        if callback:
            callback('rename_complete', {'success': success_count, 'failed': fail_count})

# ═══════════════════════════════════════════════════════════
# LAYER 2: PRESENTATION (View - Modern Clean UI)
# ═══════════════════════════════════════════════════════════

class FileOrganizerView:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[지능형 파일 정리기 v34.1.16]")
        self.root.geometry("1400x950")
        self.root.configure(bg="#F5F7FA")
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)  # 상태바 행
        self.root.columnconfigure(0, weight=1)
        
        self.COLORS = {
            'primary': '#2D3436',
            'accent': '#0984E3',
            'bg': '#F5F7FA',
            'card': '#FFFFFF',
            'text': '#2D3436',
            'text_light': '#636E72',
            'success': '#00B894',
            'error': '#D63031'
        }
        
        self._set_style()
        self._build_ui()

    def _set_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", background="white", foreground=self.COLORS['text'], rowheight=30, fieldbackground="white", font=('Malgun Gothic', 10))
        style.map("Treeview", background=[('selected', self.COLORS['accent'])])
        style.configure("Treeview.Heading", font=('Malgun Gothic', 10, 'bold'), background="#E1E5EB")
        style.configure("Action.TButton", font=('Malgun Gothic', 10, 'bold'), padding=10, background=self.COLORS['accent'], foreground="white")
        style.map("Action.TButton", background=[('active', '#0773C5')])

    def _build_ui(self):
        # 전체를 감싸는 root의 row/col 설정은 __init__에서 완료됨
        main_frame = tk.Frame(self.root, bg=self.COLORS['bg'], padx=20, pady=10)
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # 상하 조절을 위한 PanedWindow 도입
        self.paned = tk.PanedWindow(main_frame, orient=tk.VERTICAL, bg=self.COLORS['bg'], sashwidth=4, sashrelief=tk.FLAT)
        self.paned.grid(row=0, column=0, sticky='nsew')

        # 상단 영역 (Header + Top Card + Content Frame)
        self.upper_pane = tk.Frame(self.paned, bg=self.COLORS['bg'])
        self.paned.add(self.upper_pane, stretch="always")
        
        self.upper_pane.columnconfigure(0, weight=1)
        self.upper_pane.rowconfigure(2, weight=1)
        
        main_frame_ref = self.upper_pane # 기존 main_frame 역할을 upper_pane이 수행

        header_frame = tk.Frame(main_frame_ref, bg=self.COLORS['bg'])
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        tk.Label(header_frame, text="[지능형 파일 관리 시스템]", font=('Malgun Gothic', 20, 'bold'), fg=self.COLORS['primary'], bg=self.COLORS['bg']).pack(side=tk.LEFT)
        tk.Label(header_frame, text="Intelligent File Orchestrator", font=('Segoe UI', 10), fg=self.COLORS['text_light'], bg=self.COLORS['bg']).pack(side=tk.LEFT, padx=15, pady=(10, 0))
        # 도움말 버튼 추가
        ttk.Button(header_frame, text="📘 도움말 / 사용 매뉴얼", style="Action.TButton", command=self.controller.show_manual).pack(side=tk.RIGHT, padx=5)

        top_card = tk.Frame(main_frame_ref, bg=self.COLORS['card'], padx=15, pady=15, highlightthickness=1, highlightbackground="#E1E5EB")
        top_card.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        tk.Label(top_card, text="[DIR] 대상 폴더:", font=("Malgun Gothic", 9), bg=self.COLORS['card'], fg=self.COLORS['text']).grid(row=0, column=0, sticky='w')
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(top_card, textvariable=self.path_var, font=('Malgun Gothic', 10), width=80, relief=tk.FLAT, bg="#F0F2F5")
        self.path_entry.grid(row=0, column=1, padx=10, pady=5)
        # 텍스트 직접 입력 후 엔터 시 스캔 실행
        self.path_entry.bind("<Return>", lambda e: self.controller.handle_scan())
        
        ttk.Button(top_card, text="폴더 선택", command=self.controller.handle_browse).grid(row=0, column=2, padx=5)
        ttk.Button(top_card, text="개별 파일 추가", command=self.controller.handle_file_select).grid(row=0, column=3, padx=5)
        ttk.Button(top_card, text="목록 초기화", command=self.controller.handle_clear_list).grid(row=0, column=4, padx=5)
        self.recursive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(top_card, text="하위 폴더 포함", variable=self.recursive_var, bg=self.COLORS['card'], activebackground=self.COLORS['card'], font=('Malgun Gothic', 9)).grid(row=1, column=1, sticky='w', padx=10)

        content_frame = tk.Frame(main_frame_ref, bg=self.COLORS['bg'])
        content_frame.grid(row=2, column=0, sticky='nsew')
        content_frame.columnconfigure(0, weight=4) # Treeview 더 넓게
        content_frame.columnconfigure(1, weight=3) # Rules 패널도 넓게 (약 4:3 비율)
        content_frame.rowconfigure(0, weight=1)

        tree_frame = tk.Frame(content_frame, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB")
        tree_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        header = tk.Frame(tree_frame, bg="#1A237E")
        header.pack(fill="x")
        tk.Label(header, text="[FILE] 지능형 파일 고속 정리기 v34.1.16", font=("Malgun Gothic", 14, "bold"), bg="#1A237E", fg="white").pack(pady=5)
        
        # [NEW] 파일 선택 제어 및 확장자 필터 바 추가
        filter_bar = tk.Frame(tree_frame, bg="#F0F2F5", padx=5, pady=5)
        filter_bar.pack(fill="x")
        
        tk.Label(filter_bar, text="선택 제어:", font=('Malgun Gothic', 9, 'bold'), bg="#F0F2F5").pack(side=tk.LEFT, padx=(5,2))
        sb_all = tk.Button(filter_bar, text="전체선택", font=('Malgun Gothic', 8), command=self.controller.handle_select_all)
        sb_all.pack(side=tk.LEFT, padx=2)
        sb_none = tk.Button(filter_bar, text="전체해제", font=('Malgun Gothic', 8), command=self.controller.handle_deselect_all)
        sb_none.pack(side=tk.LEFT, padx=2)
        sb_inv = tk.Button(filter_bar, text="선택반전", font=('Malgun Gothic', 8), command=self.controller.handle_invert_selection)
        sb_inv.pack(side=tk.LEFT, padx=2)
        
        tk.Label(filter_bar, text="| 확장자 필터:", font=('Malgun Gothic', 9, 'bold'), bg="#F0F2F5").pack(side=tk.LEFT, padx=(15,2))
        self.ext_filter_entry = tk.Entry(filter_bar, font=('Malgun Gothic', 9), width=8)
        self.ext_filter_entry.insert(0, ".xlsx")
        self.ext_filter_entry.pack(side=tk.LEFT, padx=2)
        
        tk.Button(filter_bar, text="이 확장자만 선택", font=('Malgun Gothic', 8), command=lambda: self.controller.handle_filter_by_ext(self.ext_filter_entry.get(), "select")).pack(side=tk.LEFT, padx=2)
        tk.Button(filter_bar, text="이 확장자 제외", font=('Malgun Gothic', 8), command=lambda: self.controller.handle_filter_by_ext(self.ext_filter_entry.get(), "deselect")).pack(side=tk.LEFT, padx=2)
        
        cols = ('check', 'original', 'arrow', 'preview', 'status', 'result')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', selectmode='extended')
        self.tree.heading('check', text='선택', command=lambda: self.controller.handle_toggle_all())
        self.tree.heading('original', text='현재 파일명', command=lambda: self._treeview_sort_column('original', False))
        self.tree.heading('arrow', text='→')
        self.tree.heading('preview', text='변경될 파일명', command=lambda: self._treeview_sort_column('preview', False))
        self.tree.heading('status', text='상태', command=lambda: self._treeview_sort_column('status', False))
        self.tree.heading('result', text='작업 결과', command=lambda: self._treeview_sort_column('result', False))
        self.tree.column('check', width=50, anchor='center')
        self.tree.column('original', width=180)
        self.tree.column('arrow', width=30, anchor='center')
        self.tree.column('preview', width=220)
        self.tree.column('status', width=120, anchor='center')
        self.tree.column('result', width=120, anchor='w')
        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 이벤트 바인딩 개선
        self.tree.bind("<Button-1>", self._on_tree_click) # 단일 클릭으로 체크박스 즉시 처리
        self.tree.bind("<Double-1>", self._on_tree_edit)
        self.tree.bind("<space>", lambda e: self.controller.handle_multi_toggle()) # 스페이스바로 다중 선택 토글
        self.edit_entry = None # 현재 편집 중인 엔트리 박스

        rules_container = tk.Frame(content_frame, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB")
        rules_container.grid(row=0, column=1, sticky='nsew')
        rules_container.rowconfigure(0, weight=1)
        rules_container.columnconfigure(0, weight=1)

        rules_canvas = tk.Canvas(rules_container, bg=self.COLORS['card'], highlightthickness=0)
        rsb = ttk.Scrollbar(rules_container, orient="vertical", command=rules_canvas.yview)
        rules_frame = tk.Frame(rules_canvas, bg=self.COLORS['card'], padx=15, pady=15)
        rules_frame.bind("<Configure>", lambda e: rules_canvas.configure(scrollregion=rules_canvas.bbox("all")))
        rc_win = rules_canvas.create_window((0, 0), window=rules_frame, anchor="nw")
        rules_canvas.bind("<Configure>", lambda e: rules_canvas.itemconfig(rc_win, width=e.width))
        rules_canvas.configure(yscrollcommand=rsb.set)
        rules_canvas.grid(row=0, column=0, sticky='nsew')
        rsb.grid(row=0, column=1, sticky='ns')

        tk.Label(rules_frame, text="[관리 규칙 설정]", font=('Malgun Gothic', 11, 'bold'), bg=self.COLORS['card'], fg=self.COLORS['primary'], pady=5).pack()
        tk.Label(rules_frame, text="[패턴 제안]:", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold'), fg=self.COLORS['accent']).pack(anchor='w', pady=(5, 0))
        self.pattern_listbox = tk.Listbox(rules_frame, height=3, font=('Malgun Gothic', 9), bg="#F0F2F5", relief=tk.FLAT)
        self.pattern_listbox.pack(fill=tk.X, pady=(2, 5))
        self.pattern_listbox.bind('<<ListboxSelect>>', self.controller.handle_pattern_select)

        tk.Label(rules_frame, text="[적용 규칙 목록]:", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold')).pack(anchor='w', pady=(5, 0))
        self.rule_listbox = tk.Listbox(rules_frame, height=4, font=('Malgun Gothic', 9), bg="#F0F2F5", relief=tk.FLAT)
        self.rule_listbox.pack(fill=tk.X, pady=(2, 2))
        
        rb_frame = tk.Frame(rules_frame, bg=self.COLORS['card'])
        rb_frame.pack(fill=tk.X)
        ttk.Button(rb_frame, text="규칙 추가", width=10, command=self.controller.handle_add_rule).pack(side=tk.LEFT, padx=2)
        ttk.Button(rb_frame, text="규칙 초기화", width=10, command=self.controller.handle_clear_rules).pack(side=tk.LEFT, padx=2)

        tk.Label(rules_frame, text="[새 규칙 정의]", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold'), fg=self.COLORS['text_light']).pack(anchor='w', pady=(10, 0))
        self.rule_type = tk.StringVar(value="format_pattern")
        types = [
            ("1:형식", "format_pattern"), 
            ("2:교체", "replace"), 
            ("3:삽입", "insert"), 
            ("4:삭제", "delete"), 
            ("5:대/소", "case_convert"),
            ("6:접두사(이동/복사)", "prefix_distribute")
        ]
        t_frame = tk.Frame(rules_frame, bg=self.COLORS['card'])
        t_frame.pack(fill=tk.X, pady=5)
        # 2열 그리드로 배치하여 폭 절약 및 가독성 확보
        for i, (text, val) in enumerate(types):
            at = tk.Radiobutton(t_frame, text=text, variable=self.rule_type, value=val, 
                                bg=self.COLORS['card'], font=('Malgun Gothic', 9), 
                                indicatoron=0, width=12, padx=5, pady=5, 
                                selectcolor="#E1F5FE", command=self._on_rule_change)
            at.grid(row=i//3, column=i%3, sticky='ew', padx=2, pady=2)
        t_frame.columnconfigure((0,1,2), weight=1)

        self.params_frame = tk.Frame(rules_frame, bg=self.COLORS['card'], pady=15)
        self.params_frame.pack(fill=tk.X)

        act_frame = tk.Frame(rules_container, bg="#F8FAFC", pady=10, padx=15, highlightthickness=1, highlightbackground="#E1E5EB")
        act_frame.grid(row=1, column=0, sticky='ew')
        ttk.Button(act_frame, text="미리보기 업데이트", command=self.controller.handle_preview).pack(fill=tk.X, pady=2)
        
        # [NEW] 중단 버튼 추가
        self.stop_btn = tk.Button(act_frame, text="⛔ 작업 중단 (Stop)", font=('Malgun Gothic', 10, 'bold'), bg="#FFEAA7", fg="#D63031", relief=tk.GROOVE, command=self.controller.handle_stop)
        self.stop_btn.pack(fill=tk.X, pady=2)
        
        self.apply_btn = ttk.Button(act_frame, text="[일괄 변경 실행]", style="Action.TButton", command=self.controller.handle_run)
        self.apply_btn.pack(fill=tk.X, pady=(10, 0))

        # 하단 로그 영역 (PanedWindow 하단에 추가)
        self.lower_pane = tk.Frame(self.paned, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB", padx=10, pady=5)
        self.paned.add(self.lower_pane, stretch="never", minsize=50) # 하단은 기본적으로 늘어나지 않도록 설정
        
        self.log_text = tk.Text(self.lower_pane, font=('Consolas', 9), bg="#F8FAFC", fg="#2D3436", relief=tk.FLAT, height=4)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.status_var = tk.StringVar(value="준비됨")
        tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, font=('Malgun Gothic', 9), bg="#E1E5EB").grid(row=1, column=0, sticky='ew')
        self._on_rule_change()

    def _set_format_preset(self, val):
        self.format_pattern_entry.delete(0, tk.END)
        self.format_pattern_entry.insert(0, val)
        self.controller.handle_preview()

    def _on_rule_change(self):
        for w in self.params_frame.winfo_children(): w.destroy()
        
        rt = self.rule_type.get()
        if rt == "format_pattern":
            presets = [
                ("번호_이름", "/01_/n"), ("이름_번호", "/n_/01"),
                ("번호이름", "/01/n"), ("이름번호", "/n/01"),
                ("날짜_이름", "/YMD_/n"), ("이름_날짜", "/n_/YMD"),
                ("날짜이름", "/YMD/n"), ("이름날짜", "/n/YMD"),
                ("번호만", "/01"), ("날짜만", "/YMD")
            ]
            
            # 공간 효율성 개선: 3개씩 배치하고 패딩을 줄여 전체 높이 최적화
            row_frame = None
            for i, (txt, val) in enumerate(presets):
                if i % 3 == 0: # 3개씩 한 줄로 배치
                    row_frame = tk.Frame(self.params_frame, bg=self.COLORS['card'])
                    row_frame.pack(fill=tk.X, pady=1)
                    for c in range(3): row_frame.columnconfigure(c, weight=1)
                
                btn = tk.Button(row_frame, text=txt, font=('Malgun Gothic', 9), 
                                bg="#F8F9FA", fg=self.COLORS['primary'],
                                relief=tk.GROOVE, pady=2,
                                command=lambda v=val: self._set_format_preset(v))
                btn.grid(row=0, column=i%3, sticky='ew', padx=2)

            tk.Label(self.params_frame, text="형식(O): (직접 수정 가능)", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold')).pack(anchor='w')
            self.format_pattern_entry = tk.Entry(self.params_frame, font=('Consolas', 10), bg="#F0F2F5", relief=tk.FLAT)
            self.format_pattern_entry.insert(0, "/01_/n") # 기본값을 더 흔한 번호_이름으로 변경
            self.format_pattern_entry.pack(fill=tk.X, pady=2)
            
            # 기호 설명 가이드
            guide_lbl = tk.Label(self.params_frame, text="/n: 원본이름  /01: 번호  /YMD: 날짜", 
                                font=('Malgun Gothic', 8), fg=self.COLORS['text_light'], bg=self.COLORS['card'])
            guide_lbl.pack(anchor='w')

            f_opt = tk.Frame(self.params_frame, bg=self.COLORS['card'])
            f_opt.pack(fill=tk.X, pady=5)
            tk.Label(f_opt, text="시작 번호:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_start_var = tk.IntVar(value=1)
            tk.Spinbox(f_opt, from_=0, to=999, width=5, textvariable=self.format_start_var).pack(side=tk.LEFT, padx=5)
            tk.Label(f_opt, text="자리수:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_digits_var = tk.IntVar(value=2)
            tk.Spinbox(f_opt, from_=1, to=10, width=5, textvariable=self.format_digits_var).pack(side=tk.LEFT, padx=5)
            
            # 위치 옵션 추가
            pos_frame = tk.Frame(self.params_frame, bg=self.COLORS['card'])
            pos_frame.pack(fill=tk.X, pady=5)
            tk.Label(pos_frame, text="적용 위치:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_pos_var = tk.StringVar(value="replace")
            for t_txt, t_val in [("앞에 추가", "prefix"), ("뒤에 추가", "suffix"), ("전체 대체", "replace")]:
                tk.Radiobutton(pos_frame, text=t_txt, variable=self.format_pos_var, value=t_val, bg=self.COLORS['card'], font=('Malgun Gothic', 8)).pack(side=tk.LEFT, padx=2)

            self.format_inc_var = tk.IntVar(value=1)
            self.format_zeropad_var = tk.BooleanVar(value=True)
        elif rt == "replace":
            tk.Label(self.params_frame, text="찾을 문자열:", bg=self.COLORS['card']).pack(anchor='w')
            self.replace_find_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.replace_find_entry.pack(fill=tk.X, pady=2)
            tk.Label(self.params_frame, text="바꿀 문자열:", bg=self.COLORS['card']).pack(anchor='w')
            self.replace_to_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.replace_to_entry.pack(fill=tk.X, pady=2)
            self.replace_ignorecase_var = tk.BooleanVar(value=False)
            tk.Checkbutton(self.params_frame, text="대소문자 무시", variable=self.replace_ignorecase_var, bg=self.COLORS['card']).pack(anchor='w')
            self.replace_max_var = tk.IntVar(value=0)
        elif rt == "insert":
            tk.Label(self.params_frame, text="삽입할 문자열:", bg=self.COLORS['card']).pack(anchor='w')
            self.insert_text_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.insert_text_entry.pack(fill=tk.X, pady=2)
            self.insert_position_var = tk.IntVar(value=0)
            tk.Spinbox(self.params_frame, from_=0, to=999, textvariable=self.insert_position_var).pack(anchor='w')
            self.insert_fromend_var = tk.BooleanVar(value=False)
            tk.Checkbutton(self.params_frame, text="뒤에서부터", variable=self.insert_fromend_var, bg=self.COLORS['card']).pack(anchor='w')
        elif rt == "delete":
            self.delete_position_var = tk.IntVar(value=0)
            tk.Spinbox(self.params_frame, from_=0, to=999, textvariable=self.delete_position_var).pack(anchor='w')
            self.delete_length_var = tk.IntVar(value=1)
            tk.Spinbox(self.params_frame, from_=1, to=999, textvariable=self.delete_length_var).pack(anchor='w')
            self.delete_fromend_var = tk.BooleanVar(value=False)
            tk.Checkbutton(self.params_frame, text="뒤에서부터", variable=self.delete_fromend_var, bg=self.COLORS['card']).pack(anchor='w')
        elif rt == "case_convert":
            self.case_target_var = tk.StringVar(value="name")
            self.case_type_var = tk.StringVar(value="lower")
            for t in ["lower", "upper", "title"]: tk.Radiobutton(self.params_frame, text=t, variable=self.case_type_var, value=t, bg=self.COLORS['card']).pack(anchor='w')
            self.case_ignore_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.case_ignore_entry.pack(fill=tk.X, pady=2)
        elif rt == "prefix_distribute":
            # VBA 스타일 숫자 접두사 분류 UI
            f_dest = tk.LabelFrame(self.params_frame, text="분류 설정 (VBA 이식)", bg=self.COLORS['card'], padx=5, pady=5)
            f_dest.pack(fill=tk.X, pady=5)
            
            tk.Label(f_dest, text="목적지(Base):", bg=self.COLORS['card']).grid(row=0, column=0, sticky='w')
            self.dist_dest_var = tk.StringVar()
            tk.Entry(f_dest, textvariable=self.dist_dest_var, font=('Malgun Gothic', 9)).grid(row=0, column=1, sticky='ew')
            tk.Button(f_dest, text="..", command=lambda: self.dist_dest_var.set(filedialog.askdirectory())).grid(row=0, column=2)
            
            tk.Label(f_dest, text="하위경로(Sub):", bg=self.COLORS['card']).grid(row=1, column=0, sticky='w')
            self.dist_sub_var = tk.StringVar(value="01 자료")
            tk.Entry(f_dest, textvariable=self.dist_sub_var, font=('Malgun Gothic', 9)).grid(row=1, column=1, columnspan=2, sticky='ew')
            
            f_opt = tk.Frame(f_dest, bg=self.COLORS['card'])
            f_opt.grid(row=2, column=0, columnspan=3, sticky='w', pady=5)
            tk.Label(f_opt, text="숫자길이:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.dist_num_len = tk.IntVar(value=2)
            tk.Spinbox(f_opt, from_=1, to=10, width=3, textvariable=self.dist_num_len).pack(side=tk.LEFT, padx=5)
            
            tk.Label(f_opt, text="추가 접두사:", bg=self.COLORS['card']).pack(side=tk.LEFT, padx=(10,0))
            self.dist_prefix_var = tk.StringVar(value="표준단가_")
            tk.Entry(f_opt, textvariable=self.dist_prefix_var, width=10).pack(side=tk.LEFT, padx=5)
            
            self.dist_mode_var = tk.StringVar(value="copy")
            tk.Radiobutton(f_dest, text="복사", variable=self.dist_mode_var, value="copy", bg=self.COLORS['card']).grid(row=3, column=0)
            tk.Radiobutton(f_dest, text="이동", variable=self.dist_mode_var, value="move", bg=self.COLORS['card']).grid(row=3, column=1)
            
            f_dest.columnconfigure(1, weight=1)

    def log(self, message, m_type="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {message}\n")
        self.log_text.see(tk.END)

    def _on_tree_click(self, event):
        """단일 클릭 시 체크박스 영역이면 즉시 토글"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        if column == "#1": # '선택' 열 클릭 시 즉시 토글
            # item_id가 iid(original_path)이므로 이를 이용해 개별 토글 처리
            self.controller.handle_item_toggle_by_id(item_id)
            return "break"

    def _on_tree_edit(self, event):
        """더블 클릭 시 '변경될 파일명' 편집 엔트리 생성"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        if column != "#4": return # '변경될 파일명' 열 인덱스 (선택열 추가로 #4)
        
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        # 이전 편집창이 있다면 제거
        if self.edit_entry: self.edit_entry.destroy()
        
        x, y, w, h = self.tree.bbox(item_id, column)
        # vals[3]이 변경될 파일명임 (check:0, original:1, arrow:2, preview:3)
        old_val = self.tree.item(item_id, 'values')[3] 
        
        self.edit_entry = tk.Entry(self.tree, font=('Malgun Gothic', 10))
        self.edit_entry.insert(0, old_val)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus_set()
        
        self.edit_entry.bind("<Return>", lambda e: self._save_tree_edit(item_id))
        self.edit_entry.bind("<FocusOut>", lambda e: self._save_tree_edit(item_id))
        self.edit_entry.bind("<Escape>", lambda e: self.edit_entry.destroy())
        self.edit_entry.place(x=x, y=y, width=w, height=h)

    def _save_tree_edit(self, item_id):
        """편집된 내용을 저장하고 Treeview 및 컨트롤러 데이터 업데이트"""
        if not self.edit_entry: return
        new_val = self.edit_entry.get().strip()
        self.edit_entry.destroy()
        self.edit_entry = None
        
        # 트리뷰 업데이트 (논리적 컬럼 인덱스 정합성 확보)
        vals = list(self.tree.item(item_id, 'values'))
        if vals[3] == new_val: return # 변경사항 없음 (Index 3: 변경될 파일명)
        
        vals[3] = new_val      # Index 3: 변경될 파일명에 수정값 반영
        vals[4] = "[수정됨]"   # Index 4: 상태 열에 수정 표시
        self.tree.item(item_id, values=vals, tags=('edited',))
        self.tree.tag_configure('edited', foreground="#E67E22", font=('Malgun Gothic', 9, 'bold'))
        
        # iid(original_path)를 통해 컨트롤러 데이터 업데이트
        self.controller.handle_manual_name_change_by_id(item_id, new_val)

    def update_tree(self, data_list):
        self.tree.delete(*self.tree.get_children())
        for item in data_list:
            tag = 'error' if '[WARN]' in item['status'] else ('silent' if '[UNCHANGED]' in item['status'] else '')
            if item['status'] == "[수정됨]": tag = 'edited'
            
            # 체크 여부 표시 (V / -)
            chk = "✅" if item.get('selected', True) else "⬜"
            # [무결성 보장] 정렬 시 인덱스 꼬임 방지를 위해 iid를 파일의 전체 경로(Unique)로 지정
            self.tree.insert('', tk.END, iid=item['original_path'], values=(chk, item['original_name'], "→", item['new_name'], item['status'], ""), tags=(tag,))
        
        self.tree.tag_configure('error', foreground=self.COLORS['error'])
        self.tree.tag_configure('silent', foreground=self.COLORS['text_light'])
        self.tree.tag_configure('edited', foreground="#E67E22", font=('Malgun Gothic', 9, 'bold'))
        self.tree.tag_configure('fail_red', foreground="#FF0000", font=('Malgun Gothic', 9, 'bold'))
        self.tree.tag_configure('success_blue', foreground="#0000FF")

    def update_tree_result(self, original_path, result_text, is_success=True):
        """[최적화] 루프 없이 iid(path)를 통해 즉시 작업 결과 업데이트"""
        if self.tree.exists(original_path):
            vals = list(self.tree.item(original_path, 'values'))
            vals[5] = result_text # Index 5: '작업 결과' 열에 기록
            tag = 'success_blue' if is_success else 'fail_red'
            self.tree.item(original_path, values=vals, tags=(tag,))

    def _treeview_sort_column(self, col, reverse):
        """헤더 클릭 시 열 정렬 (오름차순/내림차순 토글)"""
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        l.sort(reverse=reverse)

        # 재정렬 후 트리뷰 갱신
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        # 다음 클릭 시 반전되도록 헤더 바인딩 갱신
        self.tree.heading(col, command=lambda: self._treeview_sort_column(col, not reverse))

# ═══════════════════════════════════════════════════════════
# LAYER 3: APPLICATION (Controller)
# ═══════════════════════════════════════════════════════════

class FileOrganizerController:
    def __init__(self, root):
        self.root = root
        self.engine = FileOrganizerEngine()
        self.view = FileOrganizerView(root, self)
        self.scanned_files = []
        self.rules = []
        self.preview_list = []
        self.manual_names = {} # [무결성] 수동으로 수정한 이름들을 보존하기 위한 맵 {path: new_name}
        self.all_selected = True

    def handle_stop(self):
        """진행 중인 모든 비동기 작업(스캔, 미리보기, 실행) 중단"""
        self.engine.stop_requested = True
        self.view.log("🛑 사용자에 의해 작업 중단 요청됨", "error")
        self.view.status_var.set("작업 중단됨")

    def handle_toggle_all(self):
        """제목 클릭 시 전체 선택 / 전체 해제 토글"""
        self.all_selected = not self.all_selected
        for item in self.preview_list:
            item['selected'] = self.all_selected
        self.view.update_tree(self.preview_list)

    def handle_select_all(self):
        for item in self.preview_list: item['selected'] = True
        self.view.update_tree(self.preview_list)
        self.view.log("전체 파일이 선택되었습니다.")

    def handle_deselect_all(self):
        for item in self.preview_list: item['selected'] = False
        self.view.update_tree(self.preview_list)
        self.view.log("전체 파일이 선택 해제되었습니다.")

    def handle_invert_selection(self):
        for item in self.preview_list:
            item['selected'] = not item.get('selected', True)
        self.view.update_tree(self.preview_list)
        self.view.log("전체 선택이 반전되었습니다.")

    def handle_multi_toggle(self):
        """드래그로 선택된 여러 행을 일괄 토글 (스페이스바 연동)"""
        selected_ids = self.view.tree.selection()
        if not selected_ids: return
        
        # 첫 번째 선택 항목의 반대 상태로 모두 통일
        idx0 = self.view.tree.index(selected_ids[0])
        new_state = not self.preview_list[idx0].get('selected', True)
        
        for mid in selected_ids:
            idx = self.view.tree.index(mid)
            self.preview_list[idx]['selected'] = new_state
        
        self.view.update_tree(self.preview_list)
        # 이전 선택 영역 유지
        for mid in selected_ids: self.view.tree.selection_add(mid)

    def handle_item_toggle_by_id(self, original_path):
        """[무결성] iid(path)를 기반으로 개별 아이템 선택 상태 토글"""
        for item in self.preview_list:
            if item['original_path'] == original_path:
                item['selected'] = not item.get('selected', True)
                self.view.update_tree(self.preview_list)
                break

    def handle_item_toggle(self, idx):
        """개별 아이템 선택 상태 토글 (인덱스 기반 - 레거시)"""
        if 0 <= idx < len(self.preview_list):
            curr = self.preview_list[idx].get('selected', True)
            self.preview_list[idx]['selected'] = not curr
            self.view.update_tree(self.preview_list)

    def handle_filter_by_ext(self, ext_pattern, mode="select"):
        """확장자별로 선택 또는 제외 (예: .txt, .jpg ...)"""
        if not ext_pattern.startswith('.'): ext_pattern = '.' + ext_pattern
        count = 0
        for item in self.preview_list:
            if item['original_name'].lower().endswith(ext_pattern.lower()):
                item['selected'] = (mode == "select")
                count += 1
        self.view.update_tree(self.preview_list)
        self.view.log(f"확장자 [{ext_pattern}] 대상 {count}건 {'선택' if mode=='select' else '제외'} 완료")

    def show_manual(self):
        manual_win = tk.Toplevel(self.root)
        manual_win.title("도움말 / 사용 매뉴얼")
        manual_win.geometry("900x700")
        manual_win.configure(bg="#F5F7FA")
        
        text_widget = tk.Text(manual_win, font=('Malgun Gothic', 10), bg="#FFFFFF", fg="#2D3436", relief=tk.FLAT, padx=20, pady=20)
        scrollbar = ttk.Scrollbar(manual_win, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        text_widget.insert(tk.END, USER_MANUAL_TEXT)
        text_widget.configure(state='disabled') # 읽기 전용

    def handle_browse(self):
        d = filedialog.askdirectory()
        if d: 
            self.view.path_var.set(d)
            self.view.log(f"폴더 선택됨: {d}")
            self.handle_scan() # 폴더 선택 시 자동 스캔 실행

    def handle_file_select(self):
        fs = filedialog.askopenfilenames()
        if fs:
            new_files = [{'name': os.path.basename(f), 'path': os.path.dirname(f), 'full_path': f, 'ext': os.path.splitext(f)[1]} for f in fs]
            self._append_unique_files(new_files)
            self.view.log(f"개별 파일 {len(fs)}개 추가됨 (총 {len(self.scanned_files)}개)")
            self.handle_preview()

    def handle_scan(self):
        p = self.view.path_var.get()
        if not p or not os.path.exists(p): return
        def run():
            # 깊은 바닥까지 완전히 탐색
            new_files = self.engine.scan_files(p, recursive=self.view.recursive_var.get())
            def finalize():
                self._append_unique_files(new_files)
                self.view.log(f"스캔 완료: {len(new_files)}개 파일 추가 (총 {len(self.scanned_files)}개)")
                pts = self.engine.detect_common_patterns(self.scanned_files)
                self.view.pattern_listbox.delete(0, tk.END)
                for pt in pts: self.view.pattern_listbox.insert(tk.END, pt)
                self.handle_preview()
            self.root.after(0, finalize)
        threading.Thread(target=run, daemon=True).start()

    def _append_unique_files(self, new_files):
        seen = {f['full_path'] for f in self.scanned_files}
        for f in new_files:
            if f['full_path'] not in seen:
                self.scanned_files.append(f)
                seen.add(f['full_path'])

    def handle_clear_list(self):
        self.scanned_files = []
        self.preview_list = []
        self.view.update_tree([])
        self.view.log("대상 파일 목록이 초기화되었습니다.")
        self.view.status_var.set("목록 초기화됨")

    def handle_preview(self):
        if not self.scanned_files:
            self.view.update_tree([])
            self.view.status_var.set("준비됨")
            return

        self.view.status_var.set("⏳ 미리보기 계산 중...")
        self.engine.stop_requested = False # 신규 계산 시작 시 중단 플래그 초기화

        def run_preview_task():
            # [무결성 보완] 기존 선택 상태(selected)를 보존하여 규칙 변경 시 초기화 방지
            selection_map = {item['original_path']: item.get('selected', True) for item in self.preview_list}
            
            # 엔진에서 미리보기 리스트 생성 (최적화된 버전 호출)
            calculated_previews = self.engine.preview_rename(self.scanned_files, self.rules)
            
            if self.engine.stop_requested: return

            # [무결성 보완] 이전 선택 상태 및 수동 수정 내역 복원
            for item in calculated_previews:
                path = item['original_path']
                if path in selection_map:
                    item['selected'] = selection_map[path]
                if path in self.manual_names:
                    item['new_name'] = self.manual_names[path]
                    item['status'] = "[수정됨]"
            
            # 메인 스레드에서 UI 업데이트
            def finalize():
                self.preview_list = calculated_previews
                self.view.update_tree(self.preview_list)
                self.view.status_var.set(f"미리보기: {len(self.preview_list)}건")
            
            self.root.after(0, finalize)

        threading.Thread(target=run_preview_task, daemon=True).start()

    def handle_manual_name_change_by_id(self, original_path, new_name):
        """정렬 시에도 안전하게 ID(Path) 기반으로 수동 수정 반영 및 보존"""
        for item in self.preview_list:
            if item['original_path'] == original_path:
                item['new_name'] = new_name
                item['status'] = "[수정됨]"
                # 수동 수정 내역 저장맵에 기록 (무결성 보송용)
                self.manual_names[original_path] = new_name
                self.view.log(f"수동 수정: {item['original_name']} -> {new_name}")
                break

    def handle_add_rule(self):
        rt = self.view.rule_type.get()
        p = {}
        if rt == "format_pattern":
            p = {
                'pattern': self.view.format_pattern_entry.get(), 
                'start': self.view.format_start_var.get(), 
                'digits': self.view.format_digits_var.get(), 
                'increment': self.view.format_inc_var.get(), 
                'zero_pad': self.view.format_zeropad_var.get(),
                'pos_type': self.view.format_pos_var.get()
            }
        elif rt == "replace":
            p = {'find': self.view.replace_find_entry.get(), 'replace': self.view.replace_to_entry.get(), 'ignore_case': self.view.replace_ignorecase_var.get(), 'max_count': self.view.replace_max_var.get()}
        elif rt == "insert":
            p = {'text': self.view.insert_text_entry.get(), 'position': self.view.insert_position_var.get(), 'from_end': self.view.insert_fromend_var.get()}
        elif rt == "delete":
            p = {'position': self.view.delete_position_var.get(), 'length': self.view.delete_length_var.get(), 'from_end': self.view.delete_fromend_var.get()}
        elif rt == "case_convert":
            p = {'target': self.view.case_target_var.get(), 'case_type': self.view.case_type_var.get(), 'ignore_chars': self.view.case_ignore_entry.get()}
        elif rt == "prefix_distribute":
            p = {
                'base_dest': self.view.dist_dest_var.get(),
                'sub_path': self.view.dist_sub_var.get(),
                'num_len': self.view.dist_num_len.get(),
                'mode': self.view.dist_mode_var.get(),
                'prefix': self.view.dist_prefix_var.get()
            }
        self.rules.append({'type': rt, 'params': p})
        self.view.rule_listbox.insert(tk.END, f"{rt}(...)")
        self.handle_preview()

    def _generate_task_csv(self, target_list=None):
        """작업 이력 리포트를 CSV 파일로 생성 (안전성 확보용)"""
        # 개별 리스트가 없으면 현재 미리보기의 선택된 대상을 기준으로 생성
        if target_list is None:
            target_list = [f for f in self.preview_list if f.get('selected', True)]
            
        if not target_list:
            self.view.log("리포트를 생성할 대상이 없습니다.", "error")
            return

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"file_organizer_report_{ts}.csv",
            title="작업 리포트 저장 위치 선택"
        )
        if not save_path: return
        
        try:
            with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 헤더 정의: 향후 원복(Undo) 또는 재배치 시 필요한 모든 메타데이터 포함
                writer.writerow([
                    '순번', '기존 파일명', '변경될 파일명', '작업유형', 
                    '원본 전체경로', '이동/변경될 예상경로', 
                    '파일크기(KB)', '원본 수정일', '원본 생성일'
                ])
                
                for i, item in enumerate(target_list, 1):
                    # 이동 경로 계산 로직 (분류 규칙이 있는 경우 반영)
                    target_path = "원본과 동일"
                    work_type = "이름변경"
                    
                    if item.get('is_distribute'):
                        work_type = "분류(이동/복사)"
                        if item.get('dist_dest'):
                            target_path = f"목적지폴더: {item['dist_dest']}"
                    
                    writer.writerow([
                        i,
                        item['original_name'],
                        item['new_name'],
                        work_type,
                        item['original_path'],
                        target_path,
                        item.get('size_kb', 0),
                        item.get('mtime', ''),
                        item.get('ctime', '')
                    ])
            self.view.log(f"심층 추적 리포트 생성 완료: {save_path}", "success")
        except Exception as e:
            self.view.log(f"CSV 생성 중 오류: {str(e)}", "error")

    def handle_clear_rules(self):
        self.rules = []
        self.view.rule_listbox.delete(0, tk.END)
        self.handle_preview()

    def handle_pattern_select(self, e):
        sel = self.view.pattern_listbox.curselection()
        if sel:
            pt = self.view.pattern_listbox.get(sel[0])
            self.rules.append({'type': 'marker_truncate', 'params': {'marker': pt, 'suffix': ')'}})
            self.view.rule_listbox.insert(tk.END, f"truncate({pt})")
            self.handle_preview()

    def handle_run(self):
        # [Critical Fix] 전수 점검 결과, 선택된 파일만 필터링하여 처리하도록 로직을 무결성 기반으로 수정
        active_list = [f for f in self.preview_list if f.get('selected', True)]
        
        if not active_list:
            messagebox.showwarning("대상의 부재", "작업 대상으로 선택된 파일이 하나도 없습니다.\n미리보기 창 좌측의 '선택' 열을 확인하세요.")
            return

        # 적용된 규칙 중 분류 규칙이 있는지 확인
        dist_rule = next((r for r in self.rules if r['type'] == 'prefix_distribute'), None)
        is_move_mode = dist_rule and dist_rule['params'].get('mode') == 'move'
        
        num_active = len(active_list)
        if dist_rule:
            mode_str = "복사" if not is_move_mode else "이동"
            confirm_msg = f"{num_active}개 파일을 목적지로 {mode_str}하시겠습니까?"
            if is_move_mode:
                confirm_msg = f"⚠️ [안전 경고] {num_active}개 파일을 목적지로 '이동'합니다.\n이동 완료 후 원본 경로에서 파일이 사라지므로 위치를 찾기 어려울 수 있습니다.\n\n계속하시겠습니까?"
        else:
            confirm_msg = f"{num_active}개 파일 이름을 변경하시겠습니까?"
            
        if not messagebox.askyesno("확인", confirm_msg): return
        
        # [안전성 보완] CSV 작업 이력 리포트 생성 시 필터링된 목록(active_list) 전달
        if messagebox.askyesno("안전 백업", "작업 전 원본과 목적지 정보를 담은 CSV 리포트를 생성하시겠습니까?\n이동(Move) 작업 시 이 이력이 있으면 안전한 복원이 가능합니다."):
            self._generate_task_csv(target_list=active_list)
            messagebox.showinfo("완료", "작업 리포트(CSV)가 생성되었습니다. 확인 후 실제 작업을 진행합니다.")

        def run():
            def cb(ev, val):
                if ev == 'item_success': 
                    self.root.after(0, lambda: self.view.log(f"성공: {val['old']} {val['new']}", "success"))
                    self.root.after(0, lambda: self.view.update_tree_result(val['path'], "성공 ✅", True))
                elif ev == 'item_fail': 
                    err_msg = f"실패: {val['name']} ({val['error']})"
                    self.root.after(0, lambda: self.view.log(err_msg, "error"))
                    self.root.after(0, lambda: self.view.update_tree_result(val['path'], f"실패 ❌ ({val['error']})", False))
                    # [VBA 무결성 요구사항] server_log.txt에 상세 기록
                    try:
                        with open("server_log.txt", "a", encoding="utf-8") as f:
                            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [DIST_FAIL] {err_msg}\n")
                    except: pass
                elif ev == 'rename_complete': 
                    op_name = "분류" if dist_rule else "이름 변경"
                    msg = f"{op_name} 작업 완료 (성공: {val['success']}, 실패: {val['failed']})"
                    if val.get('message'): msg += f"\n메시지: {val['message']}"
                    self.root.after(0, lambda: messagebox.showinfo("완료", msg))

            # [Critical Fix] 전체 목록이 아닌 'active_list'를 전달하여 선택된 파일만 처리 보장
            has_dist = any(r['type'] == 'prefix_distribute' for r in self.rules)
            if has_dist:
                self.engine.perform_distribute(active_list, self.rules, callback=cb)
            else:
                self.engine.perform_rename(active_list, callback=cb)
        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    
    # [v34.1.22] 창이 표시되지 않는 문제 해결: 최소 크기 지정 및 강제 업데이트
    root.update_idletasks()
    root.minsize(1200, 800)
    
    app = FileOrganizerController(root)
    
    # 윈도우 환경에서 창을 모든 화면의 제일 앞으로 강제 활성화
    root.attributes('-topmost', True)
    root.lift()
    root.focus_force()
    # 0.5초(500ms) 뒤 최상단 고정 속성 해제 (사용자가 다른 창을 띄울 수 있도록)
    root.after(500, lambda: root.attributes('-topmost', False))
    
    root.mainloop()
