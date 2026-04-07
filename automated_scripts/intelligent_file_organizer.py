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
                        results.append({
                            'name': file,
                            'path': root,
                            'full_path': full_path,
                            'ext': os.path.splitext(file)[1]
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
        Apply a sequence of rules to files and return previews.
        """
        previews = []
        planned_names = {}  # {new_full_path: original_name}
        file_counter = 0

        for file_info in files:
            current_name, ext = os.path.splitext(file_info['name'])
            file_counter += 1
            
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
                    
                    current_num = start_num + (file_counter - 1) * increment
                    num_str = str(current_num).zfill(digits) if zero_pad else str(current_num)
                    
                    result = pattern_str
                    result = result.replace('/n', current_name)
                    result = result.replace('/e', ext[1:] if ext else '')
                    result = re.sub(r'/(\d+)', num_str, result)
                    now = datetime.now()
                    result = result.replace('/Y', now.strftime('%Y'))
                    result = result.replace('/M', now.strftime('%m'))
                    result = result.replace('/D', now.strftime('%d'))
                    result = result.replace('/h', now.strftime('%H'))
                    result = result.replace('/m', now.strftime('%M'))
                    result = result.replace('/s', now.strftime('%S'))
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
            
            status = '[OK]'
            if new_name == file_info['name']:
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
                'status': status
            })
            
        return previews

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
                        callback('item_success', {'old': item['original_name'], 'new': os.path.basename(final_path)})
                else:
                    success_count += 1
            except Exception as e:
                fail_count += 1
                if callback:
                    callback('item_fail', {'name': item['original_name'], 'error': str(e)})
        
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
        self.root.geometry("1100x900")
        self.root.configure(bg="#F5F7FA")
        
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
        main_frame = tk.Frame(self.root, bg=self.COLORS['bg'], padx=20, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        header_frame = tk.Frame(main_frame, bg=self.COLORS['bg'])
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        tk.Label(header_frame, text="[지능형 파일 관리 시스템]", font=('Malgun Gothic', 20, 'bold'), fg=self.COLORS['primary'], bg=self.COLORS['bg']).pack(side=tk.LEFT)
        tk.Label(header_frame, text="Intelligent File Orchestrator", font=('Segoe UI', 10), fg=self.COLORS['text_light'], bg=self.COLORS['bg']).pack(side=tk.LEFT, padx=15, pady=(10, 0))

        top_card = tk.Frame(main_frame, bg=self.COLORS['card'], padx=15, pady=15, highlightthickness=1, highlightbackground="#E1E5EB")
        top_card.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        tk.Label(top_card, text="[DIR] 대상 폴더:", font=("Malgun Gothic", 9), bg=self.COLORS['card'], fg=self.COLORS['text']).grid(row=0, column=0, sticky='w')
        self.path_var = tk.StringVar()
        tk.Entry(top_card, textvariable=self.path_var, font=('Malgun Gothic', 10), width=80, relief=tk.FLAT, bg="#F0F2F5").grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(top_card, text="폴더 선택", command=self.controller.handle_browse).grid(row=0, column=2, padx=5)
        ttk.Button(top_card, text="개별 파일 선택", command=self.controller.handle_file_select).grid(row=0, column=3, padx=5)
        ttk.Button(top_card, text="분석 실행", style="Action.TButton", command=self.controller.handle_scan).grid(row=0, column=4, padx=10)
        self.recursive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(top_card, text="하위 폴더 포함", variable=self.recursive_var, bg=self.COLORS['card'], activebackground=self.COLORS['card'], font=('Malgun Gothic', 9)).grid(row=1, column=1, sticky='w', padx=10)

        content_frame = tk.Frame(main_frame, bg=self.COLORS['bg'])
        content_frame.grid(row=2, column=0, sticky='nsew')
        content_frame.columnconfigure(0, weight=3)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)

        tree_frame = tk.Frame(content_frame, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB")
        tree_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        header = tk.Frame(tree_frame, bg="#1A237E")
        header.pack(fill="x")
        tk.Label(header, text="[FILE] 지능형 파일 고속 정리기 v34.1.16", font=("Malgun Gothic", 14, "bold"), bg="#1A237E", fg="white").pack()
        
        cols = ('original', 'arrow', 'preview', 'status')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', selectmode='extended')
        self.tree.heading('original', text='현재 파일명')
        self.tree.heading('arrow', text='→')
        self.tree.heading('preview', text='변경될 파일명')
        self.tree.heading('status', text='상태')
        self.tree.column('original', width=280)
        self.tree.column('arrow', width=30, anchor='center')
        self.tree.column('preview', width=280)
        self.tree.column('status', width=100, anchor='center')
        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

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
        types = [("1:형식", "format_pattern"), ("2:교체", "replace"), ("3:삽입", "insert"), ("4:삭제", "delete"), ("5:대/소", "case_convert")]
        t_frame = tk.Frame(rules_frame, bg=self.COLORS['card'])
        t_frame.pack(fill=tk.X, pady=5)
        for text, val in types:
            tk.Radiobutton(t_frame, text=text, variable=self.rule_type, value=val, bg=self.COLORS['card'], font=('Malgun Gothic', 8), indicatoron=0, width=8, padx=2, command=self._on_rule_change).pack(side=tk.LEFT)

        self.params_frame = tk.Frame(rules_frame, bg=self.COLORS['card'], pady=15)
        self.params_frame.pack(fill=tk.X)

        act_frame = tk.Frame(rules_container, bg="#F8FAFC", pady=10, padx=15, highlightthickness=1, highlightbackground="#E1E5EB")
        act_frame.grid(row=1, column=0, sticky='ew')
        ttk.Button(act_frame, text="미리보기 업데이트", command=self.controller.handle_preview).pack(fill=tk.X, pady=2)
        self.apply_btn = ttk.Button(act_frame, text="[일괄 변경 실행]", style="Action.TButton", command=self.controller.handle_run)
        self.apply_btn.pack(fill=tk.X, pady=(10, 0))

        l_frame = tk.Frame(main_frame, bg=self.COLORS['card'], height=80, highlightthickness=1, highlightbackground="#E1E5EB", padx=10, pady=2)
        l_frame.grid(row=3, column=0, sticky='ew', pady=(10, 0))
        l_frame.grid_propagate(False)
        self.log_text = tk.Text(l_frame, font=('Consolas', 9), bg="#F8FAFC", fg="#2D3436", relief=tk.FLAT)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.status_var = tk.StringVar(value="준비됨")
        tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, font=('Malgun Gothic', 9), bg="#E1E5EB").pack(side=tk.BOTTOM, fill=tk.X)
        self._on_rule_change()

    def _on_rule_change(self):
        for w in self.params_frame.winfo_children(): w.destroy()
        rt = self.rule_type.get()
        if rt == "format_pattern":
            tk.Label(self.params_frame, text="형식(O):", bg=self.COLORS['card'], font=('Malgun Gothic', 9)).pack(anchor='w')
            self.format_pattern_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5", relief=tk.FLAT)
            self.format_pattern_entry.insert(0, "/n_/01")
            self.format_pattern_entry.pack(fill=tk.X, pady=2)
            f_opt = tk.Frame(self.params_frame, bg=self.COLORS['card'])
            f_opt.pack(fill=tk.X, pady=2)
            tk.Label(f_opt, text="시작:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_start_var = tk.IntVar(value=1)
            tk.Spinbox(f_opt, from_=0, to=999, width=5, textvariable=self.format_start_var).pack(side=tk.LEFT, padx=5)
            tk.Label(f_opt, text="자리수:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_digits_var = tk.IntVar(value=2)
            tk.Spinbox(f_opt, from_=1, to=10, width=5, textvariable=self.format_digits_var).pack(side=tk.LEFT, padx=5)
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

    def log(self, message, m_type="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {message}\n")
        self.log_text.see(tk.END)

    def update_tree(self, data_list):
        self.tree.delete(*self.tree.get_children())
        for item in data_list:
            tag = 'error' if '[WARN]' in item['status'] else ('silent' if '[UNCHANGED]' in item['status'] else '')
            self.tree.insert('', tk.END, values=(item['original_name'], "→", item['new_name'], item['status']), tags=(tag,))
        self.tree.tag_configure('error', foreground=self.COLORS['error'])
        self.tree.tag_configure('silent', foreground=self.COLORS['text_light'])

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

    def handle_browse(self):
        d = filedialog.askdirectory()
        if d: self.view.path_var.set(d); self.view.log(f"폴더 선택됨: {d}")

    def handle_file_select(self):
        fs = filedialog.askopenfilenames()
        if fs:
            self.scanned_files = [{'name': os.path.basename(f), 'path': os.path.dirname(f), 'full_path': f, 'ext': os.path.splitext(f)[1]} for f in fs]
            self.view.log(f"{len(self.scanned_files)}개 파일 선택됨")
            self.handle_preview()

    def handle_scan(self):
        p = self.view.path_var.get()
        if not p or not os.path.exists(p): return
        def run():
            self.scanned_files = self.engine.scan_files(p, recursive=self.view.recursive_var.get())
            def finalize():
                self.view.log(f"스캔 완료: {len(self.scanned_files)}개")
                pts = self.engine.detect_common_patterns(self.scanned_files)
                self.view.pattern_listbox.delete(0, tk.END)
                for pt in pts: self.view.pattern_listbox.insert(tk.END, pt)
                self.handle_preview()
            self.root.after(0, finalize)
        threading.Thread(target=run, daemon=True).start()

    def handle_preview(self):
        if not self.scanned_files: return
        self.preview_list = self.engine.preview_rename(self.scanned_files, self.rules)
        self.view.update_tree(self.preview_list)
        self.view.status_var.set(f"미리보기: {len(self.preview_list)}건")

    def handle_add_rule(self):
        rt = self.view.rule_type.get()
        p = {}
        if rt == "format_pattern":
            p = {'pattern': self.view.format_pattern_entry.get(), 'start': self.view.format_start_var.get(), 'digits': self.view.format_digits_var.get(), 'increment': self.view.format_inc_var.get(), 'zero_pad': self.view.format_zeropad_var.get()}
        elif rt == "replace":
            p = {'find': self.view.replace_find_entry.get(), 'replace': self.view.replace_to_entry.get(), 'ignore_case': self.view.replace_ignorecase_var.get(), 'max_count': self.view.replace_max_var.get()}
        elif rt == "insert":
            p = {'text': self.view.insert_text_entry.get(), 'position': self.view.insert_position_var.get(), 'from_end': self.view.insert_fromend_var.get()}
        elif rt == "delete":
            p = {'position': self.view.delete_position_var.get(), 'length': self.view.delete_length_var.get(), 'from_end': self.view.delete_fromend_var.get()}
        elif rt == "case_convert":
            p = {'target': self.view.case_target_var.get(), 'case_type': self.view.case_type_var.get(), 'ignore_chars': self.view.case_ignore_entry.get()}
        self.rules.append({'type': rt, 'params': p})
        self.view.rule_listbox.insert(tk.END, f"{rt}(...)")
        self.handle_preview()

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
        if not self.preview_list: return
        if not messagebox.askyesno("확인", f"{len(self.preview_list)}개 파일 이름을 변경하시겠습니까?"): return
        def run():
            def cb(ev, val):
                if ev == 'item_success': self.root.after(0, lambda: self.view.log(f"성공: {val['old']}->{val['new']}", "success"))
                elif ev == 'item_fail': self.root.after(0, lambda: self.view.log(f"실패: {val['name']} ({val['error']})", "error"))
                elif ev == 'rename_complete': self.root.after(0, lambda: messagebox.showinfo("완료", "작업 완료"))
            self.engine.perform_rename(self.preview_list, callback=cb)
        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = FileOrganizerController(root)
    root.mainloop()
