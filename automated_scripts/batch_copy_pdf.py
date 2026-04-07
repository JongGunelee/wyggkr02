"""
================================================================================
 [파일 스마트 복제 매니저 v1.7.1] (Clean Architecture Refactored)
================================================================================
- Clean Layer Architecture: Engine(Domain) / View(Presentation) / Controller(Application)
- 복합 규칙 엔진: 삽입/삭제/교체/치환 통합 관리 시스템
- PRD 준수: 단일 파일 유지, 기존 로직 100% 보존
================================================================================
"""
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import fnmatch


# ==========================================
# 1. DOMAIN LAYER (Core Engine - 비즈니스 로직)
# ==========================================
class FileCopyEngine:
    """파일 복제 및 명칭 편집 핵심 엔진: GUI와 분리된 순수 로직"""
    
    @staticmethod
    def apply_edit(base_name, edit_mode, options):
        """
        파일명 편집 규칙 적용
        edit_mode: '삽입', '삭제', '구간교체', '단어치환'
        options: 각 모드에 필요한 설정값 dict
        """
        if edit_mode == "삽입":
            naming_mode = options.get("naming_mode", "접두사")
            if naming_mode == "접두사":
                txt = options.get("prefix_text", "")
                pos = options.get("prefix_pos", 0)
                if pos >= len(base_name):
                    return base_name + "_" + txt
                return base_name[:pos] + txt + "_" + base_name[pos:]
            else:  # 접미사
                txt = options.get("suffix_text", "")
                pos = options.get("suffix_pos", 0)
                if pos == 0:
                    return base_name + "_" + txt
                if pos >= len(base_name):
                    return txt + "_" + base_name
                return base_name[:-pos] + "_" + txt + base_name[-pos:]
        
        elif edit_mode == "삭제":
            s = options.get("range_start", 0)
            e = options.get("range_end", 1)
            return base_name[:s] + base_name[e:]
        
        elif edit_mode == "구간교체":
            s = options.get("range_start", 0)
            e = options.get("range_end", 1)
            txt = options.get("replace_text", "")
            return base_name[:s] + txt + base_name[e:]
        
        elif edit_mode == "단어치환":
            find = options.get("find_text", "")
            replace = options.get("change_to_text", "")
            return base_name.replace(find, replace)
        
        return base_name

    @staticmethod
    def get_name_list(mode, files=None, folders=None, base_folder=None, filter_text=None):
        """명칭 추출 소스에서 파일명 리스트 반환"""
        name_list = []
        if mode == "파일" and files:
            name_list = sorted([os.path.basename(f) for f in files], key=lambda x: x.lower())
        elif mode == "폴더" and folders:
            name_list = sorted([os.path.basename(d) for d in folders], key=lambda x: x.lower())
        elif mode == "폴더내파일" and base_folder and os.path.exists(base_folder):
            all_files = [f for f in os.listdir(base_folder) if os.path.isfile(os.path.join(base_folder, f))]
            if filter_text:
                # 와일드카드(*) 포함 시 fnmatch 사용, 그 외에는 부분 일치(in) 사용
                if "*" in filter_text or "?" in filter_text:
                    all_files = [f for f in all_files if fnmatch.fnmatch(f.lower(), filter_text.lower())]
                else:
                    all_files = [f for f in all_files if filter_text.lower() in f.lower()]
            name_list = sorted(all_files, key=lambda x: x.lower())
        return name_list

    @staticmethod
    def copy_files(src_path, dest_dir, name_list, edit_mode, options, callback=None):
        """
        파일 복제 실행
        callback: (type, message) 형태로 진행 상황 전달
        """
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)
        
        src_ext = os.path.splitext(src_path)[1]
        success = 0
        errors = []
        
        for name in name_list:
            list_base = os.path.splitext(name)[0]
            new_base = FileCopyEngine.apply_edit(list_base, edit_mode, options)
            target_path = os.path.join(dest_dir, new_base + src_ext)
            
            try:
                shutil.copy2(src_path, target_path)
                success += 1
                if callback:
                    callback("success", f"성공: {new_base}{src_ext}")
            except Exception as e:
                errors.append(f"{new_base}: {str(e)}")
                if callback:
                    callback("error", f"실패: {new_base}")
        
        return {"success": success, "errors": errors}

    @staticmethod
    def transfer_files(base_folder, dest_dir, name_list, transfer_mode, callback=None):
        """
        파일 일괄 복사 또는 이동
        base_folder: 소스 파일들이 있는 폴더
        dest_dir: 대상 폴더
        name_list: 처리할 파일명 리스트
        transfer_mode: '복사' 또는 '이동'
        """
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir)
        
        success = 0
        errors = []
        
        for name in name_list:
            src_path = os.path.join(base_folder, name)
            dest_path = os.path.join(dest_dir, name)
            
            # 소스와 대상이 같으면 생략
            if os.path.abspath(src_path) == os.path.abspath(dest_path):
                if callback:
                    callback("warning", f"동일 위치 생략: {name}")
                continue

            try:
                if transfer_mode == "복사":
                    shutil.copy2(src_path, dest_path)
                else: # 이동
                    shutil.move(src_path, dest_path)
                
                success += 1
                if callback:
                    callback("success", f"{transfer_mode} 완료: {name}")
            except Exception as e:
                errors.append(f"{name}: {str(e)}")
                if callback:
                    callback("error", f"{transfer_mode} 실패: {name}")
        
        return {"success": success, "errors": errors}


# ==========================================
# 2. PRESENTATION LAYER (View - GUI 정의)
# ==========================================
class FileCopyView:
    """GUI 화면 구성: Controller와 분리된 순수 View"""
    
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("파일 스마트 복제 매니저 v1.7.1")
        self.root.geometry("750x950")
        self.root.configure(bg="#F5F5F5")
        self._build_ui()
    
    def _build_ui(self):
        # Header (Fixed at top)
        hdr = tk.Frame(self.root, bg="#37474F", pady=15)
        hdr.pack(fill="x")
        tk.Label(hdr, text="[파일 스마트 복제 매니저 v1.7.1]", font=("Malgun Gothic", 16, "bold"), bg="#37474F", fg="white").pack()
        tk.Label(hdr, text="복합 규칙 엔진: 삽입/삭제/교체/치환 통합 관리 시스템 (Clean Architecture)", font=("Malgun Gothic", 9), bg="#37474F", fg="#CFD8DC").pack()

        # Canvas for scroll
        self.canvas = tk.Canvas(self.root, bg="#F5F5F5", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        
        # Frame for content
        self.scroll_content = tk.Frame(self.canvas, padx=20, pady=10, bg="#F5F5F5")
        
        # Bind scroll
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scroll_content, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # When scroll_content changes size, update scrollregion
        self.scroll_content.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        # Allow mousewheel scroll
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Ensure scroll_content fills canvas width
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))

        main_frame = self.scroll_content

        # Section 1: Source
        box1 = tk.LabelFrame(main_frame, text="1. 기준 원본 파일 선택", font=("Malgun Gothic", 9, "bold"), padx=10, pady=10, bg="white")
        box1.pack(fill="x", pady=5)
        tk.Entry(box1, textvariable=self.controller.source_file, font=("Consolas", 9)).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(box1, text="파일 선택", command=self.controller.handle_browse_source, bg="#E1F5FE", width=12).pack(side="left")

        # Section 2: Name Source
        box2 = tk.LabelFrame(main_frame, text="2. 대상 명칭 추출 소스", font=("Malgun Gothic", 9, "bold"), padx=10, pady=10, bg="white")
        box2.pack(fill="x", pady=5)
        m_frame = tk.Frame(box2, bg="white")
        m_frame.pack(fill="x", pady=(0, 5))
        tk.Radiobutton(m_frame, text="개별 파일들", variable=self.controller.name_source_mode, value="파일", command=self.controller.handle_mode_change, bg="white").pack(side="left")
        tk.Radiobutton(m_frame, text="폴더명들", variable=self.controller.name_source_mode, value="폴더", command=self.controller.handle_mode_change, bg="white").pack(side="left", padx=15)
        tk.Radiobutton(m_frame, text="폴더 내 파일들", variable=self.controller.name_source_mode, value="폴더내파일", command=self.controller.handle_mode_change, bg="white").pack(side="left")

        # 폴더 내 파일 필터링 옵션
        self.filter_frame = tk.Frame(box2, bg="white")
        self.filter_frame.pack(fill="x", pady=(5, 5))
        tk.Label(self.filter_frame, text="└ 필터:", font=("Malgun Gothic", 8), bg="white", fg="#78909C").pack(side="left", padx=(20, 5))
        tk.Radiobutton(self.filter_frame, text="전체", variable=self.controller.folder_filter_mode, value="전체", bg="white", font=("Malgun Gothic", 8)).pack(side="left")
        tk.Radiobutton(self.filter_frame, text="특정 파일(검색):", variable=self.controller.folder_filter_mode, value="필터", bg="white", font=("Malgun Gothic", 8)).pack(side="left", padx=(10, 5))
        self.filter_ent = tk.Entry(self.filter_frame, textvariable=self.controller.folder_filter_text, font=("Consolas", 9), width=15)
        self.filter_ent.pack(side="left")

        self.name_entry_frame = tk.Frame(box2, bg="white")
        self.name_entry_frame.pack(fill="x")
        self.name_ent = tk.Entry(self.name_entry_frame, font=("Consolas", 9))
        self.name_ent.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.name_btn = tk.Button(self.name_entry_frame, text="선택", command=self.controller.handle_browse_name_source, bg="#E1F5FE", width=12)
        self.name_btn.pack(side="left")

        # Section 3: Advanced Editing Rules
        box3 = tk.LabelFrame(main_frame, text="3. 파일명 편집 규칙 설정 (모드 선택 활성화)", font=("Malgun Gothic", 9, "bold"), padx=10, pady=10, bg="white")
        box3.pack(fill="x", pady=5)

        # Tabs for Modes
        self.rule_tabs = ttk.Notebook(box3)
        self.rule_tabs.pack(fill="x", pady=5)
        
        # Tab 1: Insertion
        t1 = tk.Frame(self.rule_tabs, bg="white", padx=10, pady=10)
        self.rule_tabs.add(t1, text=" 정밀 삽입 ")
        tk.Radiobutton(t1, text="접두사(앞)", variable=self.controller.naming_mode, value="접두사", bg="white").grid(row=0, column=0, sticky="w")
        tk.Entry(t1, textvariable=self.controller.prefix_text, width=15).grid(row=0, column=1, padx=5)
        tk.Label(t1, text="위치(N칸):", bg="white").grid(row=0, column=2)
        tk.Spinbox(t1, from_=0, to=100, textvariable=self.controller.prefix_pos, width=5).grid(row=0, column=3, padx=5)
        
        tk.Radiobutton(t1, text="접미사(뒤)", variable=self.controller.naming_mode, value="접미사", bg="white").grid(row=1, column=0, sticky="w")
        tk.Entry(t1, textvariable=self.controller.suffix_text, width=15).grid(row=1, column=1, padx=5)
        tk.Label(t1, text="위치(N칸):", bg="white").grid(row=1, column=2)
        tk.Spinbox(t1, from_=0, to=100, textvariable=self.controller.suffix_pos, width=5).grid(row=1, column=3, padx=5)

        # Tab 2: Deletion
        t2 = tk.Frame(self.rule_tabs, bg="white", padx=10, pady=10)
        self.rule_tabs.add(t2, text=" 구간 삭제 ")
        tk.Label(t2, text="시작 위치(0부터):", bg="white").grid(row=0, column=0)
        tk.Spinbox(t2, from_=0, to=100, textvariable=self.controller.range_start, width=5).grid(row=0, column=1, padx=5)
        tk.Label(t2, text="~ 끝 위치(미포함):", bg="white").grid(row=0, column=2)
        tk.Spinbox(t2, from_=0, to=200, textvariable=self.controller.range_end, width=5).grid(row=0, column=3, padx=5)

        # Tab 3: Position Replace
        t3 = tk.Frame(self.rule_tabs, bg="white", padx=10, pady=10)
        self.rule_tabs.add(t3, text=" 부분 교체 ")
        tk.Label(t3, text="교체 구간:", bg="white").grid(row=0, column=0)
        tk.Spinbox(t3, from_=0, to=100, textvariable=self.controller.range_start, width=5).grid(row=0, column=1, padx=2)
        tk.Label(t3, text="~", bg="white").grid(row=0, column=2)
        tk.Spinbox(t3, from_=0, to=200, textvariable=self.controller.range_end, width=5).grid(row=0, column=3, padx=2)
        tk.Label(t3, text=" ➔ 교체 텍스트:", bg="white").grid(row=0, column=4)
        tk.Entry(t3, textvariable=self.controller.replace_text, width=15).grid(row=0, column=5, padx=5)

        # Tab 4: Global Replace
        t4 = tk.Frame(self.rule_tabs, bg="white", padx=10, pady=10)
        self.rule_tabs.add(t4, text=" 단어 치환 ")
        tk.Label(t4, text="찾을 단어:", bg="white").grid(row=0, column=0)
        tk.Entry(t4, textvariable=self.controller.find_text, width=15).grid(row=0, column=1, padx=5)
        tk.Label(t4, text=" ➔ 바꿀 단어:", bg="white").grid(row=0, column=2)
        tk.Entry(t4, textvariable=self.controller.change_to_text, width=15).grid(row=0, column=3, padx=5)

        # Tab Switch Handler
        self.rule_tabs.bind("<<NotebookTabChanged>>", self.controller.handle_tab_change)

        # Sampling Preview
        tk.Label(box3, text="[실시간 결과 미리보기]", font=("Malgun Gothic", 8, "bold"), bg="white", fg="#D32F2F").pack(anchor="w", pady=(10, 0))
        self.preview_lbl = tk.Label(box3, textvariable=self.controller.sample_preview, font=("Consolas", 10, "bold"), bg="#FFEBEE", fg="#B71C1C", anchor="w", padx=10, pady=8)
        self.preview_lbl.pack(fill="x", pady=(5, 0))

        # Section 4: Target Folder
        box4 = tk.LabelFrame(main_frame, text="4. 결과물 저장 폴더", font=("Malgun Gothic", 9, "bold"), padx=10, pady=10, bg="white")
        box4.pack(fill="x", pady=5)
        
        # 저장 방식 선택
        m4_frame = tk.Frame(box4, bg="white")
        m4_frame.pack(fill="x", pady=(0, 5))
        tk.Radiobutton(m4_frame, text="소스 폴더와 동일", variable=self.controller.target_folder_mode, value="원본동일", bg="white", font=("Malgun Gothic", 8)).pack(side="left")
        tk.Radiobutton(m4_frame, text="하위 임시폴더 생성", variable=self.controller.target_folder_mode, value="임시폴더", bg="white", font=("Malgun Gothic", 8)).pack(side="left", padx=15)
        tk.Radiobutton(m4_frame, text="수동 직접 지정", variable=self.controller.target_folder_mode, value="수동", bg="white", font=("Malgun Gothic", 8)).pack(side="left")

        self.target_entry_frame = tk.Frame(box4, bg="white")
        self.target_entry_frame.pack(fill="x")
        tk.Entry(self.target_entry_frame, textvariable=self.controller.target_dir, font=("Consolas", 9)).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(self.target_entry_frame, text="폴더 선택", command=self.controller.handle_browse_target, bg="#E1F5FE", width=12).pack(side="left")

        # Section 5: Bulk Transfer
        box5 = tk.LabelFrame(main_frame, text="5. 선택 파일 복사 또는 이동", font=("Malgun Gothic", 9, "bold"), padx=10, pady=10, bg="white")
        box5.pack(fill="x", pady=5)
        
        # 5-1: 명칭 소스 선택 (Section 2와 동일 형식)
        s5_frame = tk.Frame(box5, bg="white")
        s5_frame.pack(fill="x", pady=(0, 5))
        tk.Radiobutton(s5_frame, text="개별 파일들", variable=self.controller.name_source_mode, value="파일", command=self.controller.handle_mode_change, bg="white", font=("Malgun Gothic", 8)).pack(side="left")
        tk.Radiobutton(s5_frame, text="폴더명들", variable=self.controller.name_source_mode, value="폴더", command=self.controller.handle_mode_change, bg="white", font=("Malgun Gothic", 8), padx=10).pack(side="left")
        tk.Radiobutton(s5_frame, text="폴더 내 파일들", variable=self.controller.name_source_mode, value="폴더내파일", command=self.controller.handle_mode_change, bg="white", font=("Malgun Gothic", 8)).pack(side="left")

        # 5-2: 필터 (Section 2와 동일 형식)
        f5_frame = tk.Frame(box5, bg="white")
        f5_frame.pack(fill="x", pady=(0, 5))
        tk.Label(f5_frame, text="└ 필터:", font=("Malgun Gothic", 8), bg="white", fg="#78909C").pack(side="left", padx=(5, 5))
        tk.Radiobutton(f5_frame, text="전체", variable=self.controller.transfer_folder_filter_mode, value="전체", bg="white", font=("Malgun Gothic", 8)).pack(side="left")
        tk.Radiobutton(f5_frame, text="특정 파일(검색):", variable=self.controller.transfer_folder_filter_mode, value="필터", bg="white", font=("Malgun Gothic", 8)).pack(side="left", padx=(5, 5))
        tk.Entry(f5_frame, textvariable=self.controller.transfer_folder_filter_text, font=("Consolas", 9), width=15).pack(side="left")

        # 5-3: 소스 폴더 선택 (Section 2와 동일 형식)
        self.transfer_entry_frame = tk.Frame(box5, bg="white")
        self.transfer_entry_frame.pack(fill="x", pady=(5, 5))
        self.transfer_ent = tk.Entry(self.transfer_entry_frame, font=("Consolas", 9))
        self.transfer_ent.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.transfer_btn_sel = tk.Button(self.transfer_entry_frame, text="선택", command=self.controller.handle_browse_transfer_source, bg="#E1F5FE", width=12)
        self.transfer_btn_sel.pack(side="left")

        # 5-4: 작업 유형 및 저장 폴더 모드 (Section 4와 동일 형식)
        m5_frame = tk.Frame(box5, bg="white")
        m5_frame.pack(fill="x", pady=(5, 5))
        
        # 유형 선택 레이어
        type_layer = tk.Frame(m5_frame, bg="white")
        type_layer.pack(fill="x")
        tk.Label(type_layer, text="작업 유형:", font=("Malgun Gothic", 8, "bold"), bg="white", fg="#455A64").pack(side="left", padx=(0, 10))
        tk.Radiobutton(type_layer, text="복사 (Copy)", variable=self.controller.transfer_mode, value="복사", bg="white", font=("Malgun Gothic", 8)).pack(side="left")
        tk.Radiobutton(type_layer, text="이동 (Move)", variable=self.controller.transfer_mode, value="이동", bg="white", font=("Malgun Gothic", 8)).pack(side="left", padx=15)
        
        # 저장 폴더 모드 레이어
        target_layer = tk.Frame(m5_frame, bg="white", pady=5)
        target_layer.pack(fill="x")
        tk.Label(target_layer, text="저장 위치:", font=("Malgun Gothic", 8, "bold"), bg="white", fg="#455A64").pack(side="left", padx=(0, 10))
        tk.Radiobutton(target_layer, text="소스 폴더와 동일", variable=self.controller.transfer_target_folder_mode, value="원본동일", bg="white", font=("Malgun Gothic", 8)).pack(side="left")
        tk.Radiobutton(target_layer, text="하위 임시폴더 생성", variable=self.controller.transfer_target_folder_mode, value="임시폴더", bg="white", font=("Malgun Gothic", 8)).pack(side="left", padx=5)
        tk.Radiobutton(target_layer, text="수동 직접 지정", variable=self.controller.transfer_target_folder_mode, value="수동", bg="white", font=("Malgun Gothic", 8)).pack(side="left", padx=5)

        # 5-5: 전송 대상 저장 경로 및 실행 버튼
        self.transfer_exec_frame = tk.Frame(box5, bg="white")
        self.transfer_exec_frame.pack(fill="x", pady=(5, 0))
        tk.Entry(self.transfer_exec_frame, textvariable=self.controller.transfer_target_dir, font=("Consolas", 8), bg="#F1F8E9").pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(self.transfer_exec_frame, text="폴더 선택", command=self.controller.handle_browse_transfer_target, bg="#E1F5FE", width=10).pack(side="left", padx=(0, 5))
        self.transfer_btn = tk.Button(self.transfer_exec_frame, text="[실행]", font=("Malgun Gothic", 10, "bold"), 
                                     bg="#0288D1", fg="white", padx=15, command=self.controller.handle_transfer)
        self.transfer_btn.pack(side="right")

        # Log
        tk.Label(main_frame, text="실행 로그", font=("Malgun Gothic", 8, "bold"), bg="#F5F5F5", fg="#546E7A").pack(anchor="w", pady=(10, 0))
        self.log_txt = tk.Text(main_frame, height=8, font=("Consolas", 9), bg="#ECEFF1")
        self.log_txt.pack(fill="both", expand=True, pady=5)

        self.run_btn = tk.Button(main_frame, text="[일괄 복제 프로세스 가동]", font=("Malgun Gothic", 12, "bold"), 
                                 bg="#1B5E20", fg="white", height=2, command=self.controller.handle_start)
        self.run_btn.pack(fill="x", pady=15)
    
    def log(self, msg):
        """로그 메시지 추가"""
        self.log_txt.insert("end", f"[*] {msg}\n")
        self.log_txt.see("end")
    
    def clear_log(self):
        """로그 초기화"""
        self.log_txt.delete("1.0", "end")
    
    def update_name_entry(self, text):
        """명칭 소스 엔트리 업데이트"""
        self.name_ent.config(state="normal")
        self.name_ent.delete(0, "end")
        self.name_ent.insert(0, text)
        self.name_ent.config(state="readonly")
    
    def set_running_state(self, is_running):
        """실행 상태 UI 업데이트"""
        if is_running:
            self.run_btn.config(state="disabled", bg="#9E9E9E")
            if hasattr(self, "transfer_btn"):
                self.transfer_btn.config(state="disabled", bg="#9E9E9E")
        else:
            self.run_btn.config(state="normal", bg="#1B5E20")
            if hasattr(self, "transfer_btn"):
                self.transfer_btn.config(state="normal", bg="#0288D1")
    
    def update_transfer_entry(self, text):
        """명칭 소스 엔트리 업데이트 (Section 5)"""
        if hasattr(self, "transfer_ent"):
            self.transfer_ent.config(state="normal")
            self.transfer_ent.delete(0, "end")
            self.transfer_ent.insert(0, text)
            self.transfer_ent.config(state="readonly")


# ==========================================
# 3. APPLICATION LAYER (Controller - 로직 연결)
# ==========================================
class FileCopyController:
    """Controller: View와 Engine을 연결하고 사용자 이벤트 처리"""
    
    def __init__(self, root):
        self.root = root
        self.engine = FileCopyEngine()
        self.is_running = False
        
        # === 상태 변수 초기화 (기본값 설정 v1.7 리팩토링) ===
        default_src = "D:/02 기숙사 및 사택/02 견적작업/01 견적 갑지/02특기안전시방_(시방서 유첨용).pdf"
        self.source_file = tk.StringVar(value=default_src)
        self.target_dir = tk.StringVar()
        self.name_source_mode = tk.StringVar(value="폴더내파일")
        self.selected_name_files = []
        self.selected_name_folders = []
        self.base_folder_path = tk.StringVar(value=os.path.dirname(default_src))
        self.folder_filter_mode = tk.StringVar(value="필터")
        self.folder_filter_text = tk.StringVar(value="*시방서*.pp*")
        self.target_folder_mode = tk.StringVar(value="원본동일")
        
        # Section 5: Bulk Transfer 전용 상태 변수
        self.transfer_name_source_mode = tk.StringVar(value="폴더내파일")
        self.transfer_selected_name_files = []
        self.transfer_selected_name_folders = []
        self.transfer_base_folder_path = tk.StringVar(value=os.path.dirname(default_src))
        self.transfer_folder_filter_mode = tk.StringVar(value="필터")
        self.transfer_folder_filter_text = tk.StringVar(value="*작업요청*.pp*")
        self.transfer_target_folder_mode = tk.StringVar(value="원본동일")
        self.transfer_target_dir = tk.StringVar()
        self.transfer_mode = tk.StringVar(value="복사")
        
        # Edit Modes (기본 규칙 리팩토링)
        self.edit_mode = tk.StringVar(value="삽입")
        self.naming_mode = tk.StringVar(value="접두사")
        self.prefix_text = tk.StringVar(value="특기")
        self.prefix_pos = tk.IntVar(value=2)
        self.suffix_text = tk.StringVar(value="특기")
        self.suffix_pos = tk.IntVar(value=2)
        self.range_start = tk.IntVar(value=0)
        self.range_end = tk.IntVar(value=1)
        self.replace_text = tk.StringVar(value="특기")
        self.find_text = tk.StringVar(value="PR당사안")
        self.change_to_text = tk.StringVar(value="특기")
        self.sample_preview = tk.StringVar(value="대기 중...")
        
        # 실시간 미리보기 및 저장 경로 자동 업데이트 연결
        vars_to_watch = [self.edit_mode, self.naming_mode, self.prefix_text, self.prefix_pos,
                         self.suffix_text, self.suffix_pos, self.range_start, self.range_end,
                         self.replace_text, self.find_text, self.change_to_text,
                         self.base_folder_path, self.name_source_mode,
                         self.folder_filter_mode, self.folder_filter_text,
                         self.transfer_base_folder_path, self.transfer_name_source_mode,
                         self.transfer_folder_filter_mode, self.transfer_folder_filter_text,
                         self.transfer_target_folder_mode, self.target_folder_mode]
        for v in vars_to_watch:
            v.trace_add("write", lambda *args: self._update_sampling())
        
        # View 생성 (Controller 주입)
        self.view = FileCopyView(root, self)
        self.handle_mode_change()
        self.view.log("시스템이 초기 설정값(v1.7)으로 정밀 로드되었습니다.")
        self.view.log(f"- 원본: {os.path.basename(self.source_file.get())}")
        self.view.log(f"- 복사/이동 모드: {self.transfer_mode.get()} (기본)")
        self.view.log(f"- 모드: {self.name_source_mode.get()} (활성화)")
        self.view.log("- 저장: 대상 폴더 내 'Output_복제결과' 자동 생성 모드")
        
        # 초기 미리보기 및 분석 강제 갱신
        self._update_sampling()
    
    def _update_target_dir(self):
        """대상 명칭 추출 소스 폴더의 하위로 저장 폴더 자동 설정 및 입력칸 동기화"""
        if not hasattr(self, "view"):
            return
        
        # 2번 섹션 기준
        mode = self.name_source_mode.get()
        base_path = ""
        if mode == "폴더내파일":
            base_path = self.base_folder_path.get()
        elif mode == "파일" and self.selected_name_files:
            base_path = os.path.dirname(self.selected_name_files[0])
        elif mode == "폴더" and self.selected_name_folders:
            base_path = self.selected_name_folders[0]
            
        if base_path and os.path.isdir(base_path):
            if self.target_folder_mode.get() == "임시폴더":
                new_target = os.path.normpath(os.path.join(base_path, "Output_복제결과"))
            elif self.target_folder_mode.get() == "원본동일":
                new_target = os.path.normpath(base_path)
            else: # 수동
                new_target = self.target_dir.get()
                
            if self.target_folder_mode.get() != "수동":
                if self.target_dir.get() != new_target:
                    self.target_dir.set(new_target)
                
        # 5번 섹션 입력창 및 저장 폴더 동기화
        t_mode = self.transfer_name_source_mode.get()
        t_path = ""
        if t_mode == "폴더내파일":
            t_path = self.transfer_base_folder_path.get()
        elif t_mode == "파일" and self.transfer_selected_name_files:
            t_path = os.path.dirname(self.transfer_selected_name_files[0])
        elif t_mode == "폴더" and self.transfer_selected_name_folders:
            t_path = self.transfer_selected_name_folders[0]
            
        if t_path:
            self.view.update_transfer_entry(os.path.normpath(t_path))
            
        # 5번 섹션 대상 저장 폴더 자동 설정 (사용자 요청: 2번 섹션의 폴더 변경 시에만 자동 반영)
        if base_path and os.path.isdir(base_path):
            if self.transfer_target_folder_mode.get() == "임시폴더":
                new_t_target = os.path.normpath(os.path.join(base_path, "Output_복제결과"))
            elif self.transfer_target_folder_mode.get() == "원본동일":
                new_t_target = os.path.normpath(base_path)
            else: # 수동
                new_t_target = self.transfer_target_dir.get()
            
            # 수동 모드가 아닐 때만 2번 섹션의 경로를 따라감
            if self.transfer_target_folder_mode.get() != "수동":
                if self.transfer_target_dir.get() != new_t_target:
                    self.transfer_target_dir.set(new_t_target)
    
    def _get_options(self):
        """현재 설정 옵션을 dict로 반환"""
        return {
            "naming_mode": self.naming_mode.get(),
            "prefix_text": self.prefix_text.get(),
            "prefix_pos": self.prefix_pos.get(),
            "suffix_text": self.suffix_text.get(),
            "suffix_pos": self.suffix_pos.get(),
            "range_start": self.range_start.get(),
            "range_end": self.range_end.get(),
            "replace_text": self.replace_text.get(),
            "find_text": self.find_text.get(),
            "change_to_text": self.change_to_text.get()
        }
    
    def handle_browse_source(self):
        """원본 파일 선택"""
        f = filedialog.askopenfilename(filetypes=[("모든 파일", "*.*")])
        if f:
            self.source_file.set(f)
            self.view.log(f"원본 지정: {os.path.basename(f)}")
            self._update_sampling()
    
    def handle_browse_name_source(self):
        """명칭 추출 소스 선택"""
        self.view.clear_log()
        mode = self.name_source_mode.get()
        
        if mode == "파일":
            fs = filedialog.askopenfilenames(filetypes=[("모든 파일", "*.*")])
            if fs:
                self.selected_name_files.extend(list(fs))
                self.selected_name_files = list(set(self.selected_name_files))
                self.handle_mode_change()
        elif mode == "폴더":
            d = filedialog.askdirectory(title="폴더들을 선택하세요")
            if d:
                self.selected_name_folders.append(d)
                self.selected_name_folders = list(set(self.selected_name_folders))
                self.handle_mode_change()
        elif mode == "폴더내파일":
            d = filedialog.askdirectory(title="파일들을 추출할 폴더를 선택하세요")
            if d:
                self.base_folder_path.set(d)
                self.handle_mode_change()
    
    # _analyze_and_log 메서드 제거 (handle_mode_change 및 _update_sampling으로 통합)
    
    def handle_browse_target(self):
        """저장 폴더 선택 (Section 4)"""
        d = filedialog.askdirectory(title="저장할 폴더를 선택하세요")
        if d:
            self.target_folder_mode.set("수동")
            self.target_dir.set(d)
    
    def handle_browse_transfer_target(self):
        """복사/이동 저장 폴더 선택 (Section 5)"""
        d = filedialog.askdirectory(title="저장할 폴더를 선택하세요")
        if d:
            self.transfer_target_folder_mode.set("수동")
            self.transfer_target_dir.set(d)

    def handle_browse_transfer_source(self):
        """복사/이동 소스 선택 (Section 5)"""
        mode = self.transfer_name_source_mode.get()
        if mode == "파일":
            fs = filedialog.askopenfilenames(filetypes=[("모든 파일", "*.*")])
            if fs:
                self.transfer_selected_name_files = list(fs) # 신규 선택 시 리스트 교체
                self.handle_mode_change()
        elif mode == "폴더":
            d = filedialog.askdirectory(title="폴더를 선택하세요")
            if d:
                self.transfer_selected_name_folders = [d] # 신규 선택 시 리스트 교체
                self.handle_mode_change()
        elif mode == "폴더내파일":
            d = filedialog.askdirectory(title="파일들을 추출할 폴더를 선택하세요")
            if d:
                self.transfer_base_folder_path.set(d)
                self.handle_mode_change()

    def handle_mode_change(self):
        """명칭 소스 모드 변경 및 경로 즉시 표시"""
        mode = self.name_source_mode.get()
        if mode == "파일":
            self.view.update_name_entry(f"{len(self.selected_name_files)}개 파일 선택됨")
        elif mode == "폴더":
            if len(self.selected_name_folders) == 1:
                self.view.update_name_entry(os.path.normpath(self.selected_name_folders[0]))
            else:
                self.view.update_name_entry(f"{len(self.selected_name_folders)}개 폴더 선택됨")
        elif mode == "폴더내파일":
            self.view.update_name_entry(os.path.normpath(self.base_folder_path.get()))
        
        # 5번 섹션 동기화
        t_mode = self.transfer_name_source_mode.get()
        if t_mode == "파일":
            self.view.update_transfer_entry(f"{len(self.transfer_selected_name_files)}개 파일 선택됨")
        elif t_mode == "폴더":
            if len(self.transfer_selected_name_folders) == 1:
                self.view.update_transfer_entry(os.path.normpath(self.transfer_selected_name_folders[0]))
            else:
                self.view.update_transfer_entry(f"{len(self.transfer_selected_name_folders)}개 폴더 선택됨")
        elif t_mode == "폴더내파일":
            self.view.update_transfer_entry(os.path.normpath(self.transfer_base_folder_path.get()))
        
        self._update_sampling()
    
    def handle_tab_change(self, event):
        """탭 변경 처리"""
        idx = self.view.rule_tabs.index("current")
        modes = ["삽입", "삭제", "구간교체", "단어치환"]
        self.edit_mode.set(modes[idx])
        self._update_sampling()
    
    def _update_sampling(self, *args):
        """실시간 미리보기 및 저장 경로 업데이트"""
        if not hasattr(self, "view"):
            return
        self._update_target_dir()
        try:
            src_ext = os.path.splitext(self.source_file.get())[1] if self.source_file.get() else ".pdf"
            sample_base = "파일명예시"
            
            mode = self.name_source_mode.get()
            if mode == "파일" and self.selected_name_files:
                sample_base = os.path.splitext(os.path.basename(self.selected_name_files[0]))[0]
            elif mode == "폴더" and self.selected_name_folders:
                sample_base = os.path.basename(self.selected_name_folders[0])
            elif mode == "폴더내파일" and self.base_folder_path.get():
                # 필터링된 목록의 첫 번째 파일을 샘플로 사용
                names = self.engine.get_name_list(
                    mode, None, None, self.base_folder_path.get(),
                    self.folder_filter_text.get() if self.folder_filter_mode.get() == "필터" else None
                )
                if names:
                    sample_base = os.path.splitext(names[0])[0]
            
            res_base = self.engine.apply_edit(sample_base, self.edit_mode.get(), self._get_options())
            self.sample_preview.set(f"[{self.edit_mode.get()}] {sample_base}  ➔  {res_base}{src_ext}")
        except:
            self.sample_preview.set("설정 대기 중...")
        
        # 필터링 및 옵션 변경 시 로그 분석 결과도 함께 실시간 갱신 (섹션 2 & 5)
        if not self.is_running:
            # 1. 복제 로그 (Section 2 기준)
            name_list1 = self.engine.get_name_list(
                self.name_source_mode.get(),
                self.selected_name_files,
                self.selected_name_folders,
                self.base_folder_path.get(),
                self.folder_filter_text.get() if self.folder_filter_mode.get() == "필터" else None
            )
            # 2. 전송 로그 (Section 5 기준)
            name_list2 = self.engine.get_name_list(
                self.transfer_name_source_mode.get(),
                self.transfer_selected_name_files,
                self.transfer_selected_name_folders,
                self.transfer_base_folder_path.get(),
                self.transfer_folder_filter_text.get() if self.transfer_folder_filter_mode.get() == "필터" else None
            )
            
            self.view.clear_log()
            self.view.log(f"--- [복합 분석 결과] 복제대상:{len(name_list1)}건 / 전송대상:{len(name_list2)}건 ---")
            if name_list1:
                self.view.log("<복제 예정 목록>")
                for n in sorted(name_list1)[:10]: self.view.log(f" [복제] {n}")
            if name_list2:
                self.view.log("<복사/이동 예정 목록>")
                for n in sorted(name_list2)[:10]: self.view.log(f" [{self.transfer_mode.get()}] {n}")
    
    def handle_start(self):
        """복제 프로세스 시작"""
        if self.is_running:
            return
        
        src = self.source_file.get()
        out = self.target_dir.get()
        
        if not src or not os.path.exists(src):
            messagebox.showerror("오류", "원본 파일을 선택해 주세요.")
            return
        if not out:
            messagebox.showerror("오류", "저장 폴더를 지정해 주세요.")
            return
        
        self.is_running = True
        self.view.set_running_state(True)
        self.view.clear_log()
        
        def run():
            try:
                name_list = self.engine.get_name_list(
                    self.name_source_mode.get(),
                    self.selected_name_files,
                    self.selected_name_folders,
                    self.base_folder_path.get(),
                    self.folder_filter_text.get() if self.folder_filter_mode.get() == "필터" else None
                )
                
                if not name_list:
                    self.root.after(0, lambda: self._finalize(0))
                    return
                
                result = self.engine.copy_files(
                    src, out, name_list,
                    self.edit_mode.get(),
                    self._get_options(),
                    self._callback
                )
                self.root.after(0, lambda: self._finalize(result["success"]))
            except Exception as e:
                self.root.after(0, lambda: self.view.log(f"치명적 오류: {str(e)}"))
                self.root.after(0, lambda: self._finalize(0))
        
        threading.Thread(target=run, daemon=True).start()
    
    def _callback(self, msg_type, msg):
        """엔진 콜백 (스레드 안전)"""
        self.root.after(0, lambda: self.view.log(msg))
    
    def _finalize(self, success_count):
        """작업 완료 처리"""
        self.is_running = False
        self.view.set_running_state(False)
        messagebox.showinfo("완료", f"총 {success_count}건 작업 완료")

    def handle_transfer(self):
        """선택 파일 일괄 복사/이동 실행"""
        if self.is_running:
            return
        
        mode = self.transfer_name_source_mode.get()
        src_folder = ""
        
        if mode == "폴더내파일":
            src_folder = self.transfer_base_folder_path.get()
        elif mode == "파일" and self.transfer_selected_name_files:
            src_folder = os.path.dirname(self.transfer_selected_name_files[0])
        elif mode == "폴더" and self.transfer_selected_name_folders:
            src_folder = self.transfer_selected_name_folders[0]
        
        out_folder = self.transfer_target_dir.get()
        
        if not src_folder or not os.path.exists(src_folder):
            messagebox.showerror("오류", "소스 폴더가 유효하지 않습니다.")
            return
        if not out_folder:
            messagebox.showerror("오류", "저장 폴더가 지정되지 않았습니다.")
            return
            
        t_mode = self.transfer_mode.get()
        if not messagebox.askyesno("확인", f"선택된 파일들을 {out_folder} 폴더로 {t_mode} 하시겠습니까?"):
            return

        self.is_running = True
        self.view.set_running_state(True)
        self.view.clear_log()
        
        def run():
            try:
                name_list = self.engine.get_name_list(
                    self.transfer_name_source_mode.get(),
                    self.transfer_selected_name_files,
                    self.transfer_selected_name_folders,
                    self.transfer_base_folder_path.get(),
                    self.transfer_folder_filter_text.get() if self.transfer_folder_filter_mode.get() == "필터" else None
                )
                
                if not name_list:
                    self.root.after(0, lambda: self._finalize(0))
                    return
                
                result = self.engine.transfer_files(
                    src_folder, out_folder, name_list,
                    t_mode,
                    self._callback
                )
                self.root.after(0, lambda: self._finalize(result["success"]))
            except Exception as e:
                self.root.after(0, lambda: self.view.log(f"전송 오류: {str(e)}"))
                self.root.after(0, lambda: self._finalize(0))
        
        threading.Thread(target=run, daemon=True).start()


# ==========================================
# 4. ENTRY POINT
# ==========================================
if __name__ == "__main__":
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    FileCopyController(root)
    root.mainloop()
