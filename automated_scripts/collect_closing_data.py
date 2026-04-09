"""
================================================================================
 [FILE] 지능형 파일 수집기 (Intelligent File Collector) v34.1.16
================================================================================
 - 클린 레이어 아키텍처: Engine(Domain) / View(Presentation) / Controller(Application)
 - 사용자 친화적 GUI: 모든 옵션을 한 화면에서 설정 후 실행
 - 안전한 복사: 중복 파일 자동 보존, 미리보기 모드 지원
 - 무결성 보증: sys.stdout UTF-8 강제 설정 및 이모지 제거 (CP949 호환)
================================================================================
"""
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import os
import shutil
import threading
import time
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# ==========================================
# 1. DOMAIN LAYER (Core Engine - 비즈니스 로직)
# ==========================================
class FileCollectorEngine:
    """파일 수집 핵심 엔진: GUI와 분리된 순수 로직"""
    
    @staticmethod
    def scan_files(source_dir, pattern, extension_filter=".pdf", exclude_folders=None, callback=None):
        """
        모든 하위 폴더를 재귀적으로 탐색하여 패턴에 맞는 파일 목록 반환
        exclude_folders: 탐색에서 제외할 폴더명 리스트
        callback: 진행 상황 업데이트용 (type, data)
        """
        matched_files = []
        folder_count = 0
        skipped_count = 0
        exclude_set = set(f.strip().lower() for f in (exclude_folders or []) if f.strip())
        
        for root_path, dirs, files in os.walk(source_dir):
            # 제외 폴더 필터링 (dirs를 수정하면 os.walk가 해당 폴더를 건너뜀)
            if exclude_set:
                original_dirs = dirs[:]
                dirs[:] = [d for d in dirs if d.lower() not in exclude_set]
                skipped = len(original_dirs) - len(dirs)
                if skipped > 0:
                    skipped_count += skipped
                    if callback:
                        callback("skip", f"제외됨: {[d for d in original_dirs if d.lower() in exclude_set]}")
            
            folder_count += 1
            if callback:
                callback("status", f"탐색 중: {os.path.basename(root_path)} (폴더 #{folder_count})")
            
            for f_name in files:
                # 확장자 필터링
                if extension_filter and extension_filter.lower() not in f_name.lower():
                    continue
                # 패턴 필터링 (대소문자 무시)
                if pattern.lower() in f_name.lower():
                    file_path = os.path.join(root_path, f_name)
                    rel_path = os.path.relpath(file_path, source_dir)
                    matched_files.append({
                        "full_path": file_path,
                        "rel_path": rel_path,
                        "name": f_name,
                        "folder": os.path.basename(root_path)
                    })
                    if callback:
                        callback("found", f"발견: {rel_path}")
        
        return {"files": matched_files, "folder_count": folder_count, "skipped_folders": skipped_count}

    
    @staticmethod
    def apply_padding(text, enabled=True):
        """독립된 한 자릿수 숫자를 감지하여 0 삽입 (예: 1 -> 01)"""
        if not enabled:
            return text
        return re.sub(r'(?<!\d)(\d)(?!\d)', r'0\1', text)

    @staticmethod
    def copy_files(files, dest_dir, prefix="", suffix="", smart_padding=False, callback=None):
        """
        파일들을 대상 폴더로 복사 (중복명 자동 처리)
        """
        dest_path = Path(dest_dir)
        if not dest_path.exists():
            dest_path.mkdir(parents=True)
        
        success = 0
        errors = []
        
        for i, f_info in enumerate(files, 1):
            try:
                # 파일명 생성 로직 (접두사 + 원본명 + 접미사)
                orig_path = Path(f_info['name'])
                stem = orig_path.stem
                ext = orig_path.suffix
                
                # 지능형 패딩 적용 (원본 이름에 대해)
                stem = FileCollectorEngine.apply_padding(stem, smart_padding)
                
                # 접두사/접미사 적용
                res_name = f"{prefix}_{stem}" if prefix else stem
                if suffix:
                    res_name = f"{res_name}_{suffix}"
                
                new_name = f"{res_name}{ext}"
                target = dest_path / new_name
                
                # 중복 파일명 처리
                counter = 1
                while target.exists():
                    cnt_str = f"{counter:02d}" if smart_padding else str(counter)
                    new_name = f"{res_name}_{cnt_str}{ext}"
                    target = dest_path / new_name
                    counter += 1
                
                shutil.copy2(f_info['full_path'], target)
                success += 1
                
                if callback:
                    callback("copy", f"[{i}/{len(files)}] 복사 완료: {new_name}")
            except Exception as e:
                errors.append(f"{f_info['name']}: {str(e)}")
                if callback:
                    callback("error", f"오류: {f_info['name']}")
        
        return {"success": success, "errors": errors}


# ==========================================
# 2. PRESENTATION LAYER (View - GUI 정의)
# ==========================================
class CollectorView:
    """GUI 화면 구성: Controller와 분리된 순수 View"""
    
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[지능형 파일 수집기 v34.1.16]")
        self.root.geometry("900x750")
        self._build_ui()
    
    def _build_ui(self):
        # ===== 헤더 영역 =====
        header = tk.Frame(self.root, bg="#1565C0", pady=12)
        header.pack(fill="x")
        tk.Label(header, text="[지능형 파일 수집기 v34.1.16]", font=("Malgun Gothic", 16, "bold"),
                 bg="#1565C0", fg="white").pack()
        tk.Label(header, text="하위 폴더 전수 조사 → 패턴 매칭 → 접두사 부여 → 안전 복사",
                 font=("Malgun Gothic", 9), bg="#1565C0", fg="#BBDEFB").pack()
        
        # ===== 메인 설정 영역 =====
        main = tk.Frame(self.root, padx=20, pady=10)
        main.pack(fill="both", expand=True)
        
        # --- 1. 소스 폴더 설정 ---
        box1 = tk.LabelFrame(main, text="[1] 검색 대상 폴더", font=("Malgun Gothic", 10, "bold"), padx=10, pady=8)
        box1.pack(fill="x", pady=5)
        
        src_f = tk.Frame(box1)
        src_f.pack(fill="x")
        tk.Label(src_f, text="[경로]:", font=("Malgun Gothic", 9)).pack(side="left")
        self.src_entry = tk.Entry(src_f, textvariable=self.controller.source_dir_var, font=("Consolas", 9))
        self.src_entry.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(src_f, text="폴더 선택", command=self.controller.handle_select_source,
                  font=("Malgun Gothic", 8), bg="#E3F2FD").pack(side="left", padx=2)
        
        # --- 2. 검색 조건 설정 ---
        box2 = tk.LabelFrame(main, text="[2] 검색 조건 설정", font=("Malgun Gothic", 10, "bold"), padx=10, pady=8)
        box2.pack(fill="x", pady=5)
        
        cond_f = tk.Frame(box2)
        cond_f.pack(fill="x", pady=2)
        tk.Label(cond_f, text="[파일명 패턴]:", font=("Malgun Gothic", 9)).pack(side="left")
        tk.Entry(cond_f, textvariable=self.controller.pattern_var, font=("Malgun Gothic", 9), width=30).pack(side="left", padx=5)
        tk.Label(cond_f, text="[확장자]:", font=("Malgun Gothic", 9)).pack(side="left", padx=(10,0))
        ttk.Combobox(cond_f, textvariable=self.controller.ext_var, values=[".pdf", ".xlsx", ".docx", ".pptx", ".png", ".jpg", "(모든 파일)"],
                     font=("Malgun Gothic", 9), width=12, state="readonly").pack(side="left", padx=5)
        
        # 제외 폴더 설정
        excl_f = tk.Frame(box2)
        excl_f.pack(fill="x", pady=5)
        tk.Label(excl_f, text="[제외 폴더]:", font=("Malgun Gothic", 9)).pack(side="left")
        tk.Entry(excl_f, textvariable=self.controller.exclude_folders_var, font=("Malgun Gothic", 9), width=50).pack(side="left", padx=5, fill="x", expand=True)
        tk.Label(excl_f, text="(쉼표로 구분, 예: 000_Backup, temp, 임시)", font=("Malgun Gothic", 8), fg="#666").pack(side="left")

        
        # --- 3. 출력 설정 ---
        box3 = tk.LabelFrame(main, text="[3] 출력 설정", font=("Malgun Gothic", 10, "bold"), padx=10, pady=8)
        box3.pack(fill="x", pady=5)
        
        out_f = tk.Frame(box3)
        out_f.pack(fill="x", pady=2)
        tk.Label(out_f, text="[저장 폴더]:", font=("Malgun Gothic", 9)).pack(side="left")
        self.dest_entry = tk.Entry(out_f, textvariable=self.controller.dest_dir_var, font=("Consolas", 9))
        self.dest_entry.pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(out_f, text="폴더 선택", command=self.controller.handle_select_dest,
                  font=("Malgun Gothic", 8), bg="#E3F2FD").pack(side="left", padx=2)
        
        # 접두사 설정 라인
        prefix_f = tk.Frame(box3)
        prefix_f.pack(fill="x", pady=3)
        tk.Checkbutton(prefix_f, text="상위 폴더명을 접두사로 사용", variable=self.controller.use_prefix_var,
                       font=("Malgun Gothic", 9)).pack(side="left")
        tk.Label(prefix_f, text="  |  지정 접두사:", font=("Malgun Gothic", 9), width=12, anchor="e").pack(side="left", padx=(10,0))
        tk.Entry(prefix_f, textvariable=self.controller.custom_prefix_var, font=("Malgun Gothic", 9), width=15).pack(side="left", padx=5)
        
        # 접미사 설정 라인
        suffix_f = tk.Frame(box3)
        suffix_f.pack(fill="x", pady=3)
        tk.Checkbutton(suffix_f, text="상위 폴더명을 접미사로 사용", variable=self.controller.use_suffix_var,
                       font=("Malgun Gothic", 9)).pack(side="left")
        tk.Label(suffix_f, text="  |  지정 접미사:", font=("Malgun Gothic", 9), width=12, anchor="e").pack(side="left", padx=(10,0))
        tk.Entry(suffix_f, textvariable=self.controller.custom_suffix_var, font=("Malgun Gothic", 9), width=15).pack(side="left", padx=5)
        
        # 설정 통합 라인
        opt_f = tk.Frame(box3)
        opt_f.pack(fill="x", pady=3)
        tk.Checkbutton(opt_f, text="숫자 자릿수 보정 (1 → 01)", variable=self.controller.smart_padding_var,
                       font=("Malgun Gothic", 9, "bold"), fg="#D32F2F").pack(side="left")
        tk.Label(opt_f, text=" *파일명 내의 숫자를 지능형으로 분석하여 정렬을 최적화합니다.", 
                 font=("Malgun Gothic", 8), fg="#777").pack(side="left")
        
        # --- 4. 로그/결과 영역 ---
        box4 = tk.LabelFrame(main, text="[4] 진행 상황 및 결과", font=("Malgun Gothic", 10, "bold"), padx=10, pady=8)
        box4.pack(fill="both", expand=True, pady=5)
        
        self.log_txt = tk.Text(box4, font=("Consolas", 9), bg="#FAFAFA", height=12)
        scr = tk.Scrollbar(box4, command=self.log_txt.yview)
        self.log_txt.config(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y")
        self.log_txt.pack(side="left", fill="both", expand=True)
        
        # --- 5. 액션 버튼 영역 ---
        act_f = tk.Frame(main, pady=10)
        act_f.pack(fill="x")
        
        self.scan_btn = tk.Button(act_f, text="[검색 미리보기]", font=("Malgun Gothic", 10, "bold"),
                                   bg="#FFA726", fg="white", width=18, command=self.controller.handle_preview)
        self.scan_btn.pack(side="left", expand=True, padx=5)
        
        self.run_btn = tk.Button(act_f, text="[수집 실행]", font=("Malgun Gothic", 10, "bold"),
                                  bg="#2E7D32", fg="white", width=18, command=self.controller.handle_collect)
        self.run_btn.pack(side="left", expand=True, padx=5)
        
        tk.Button(act_f, text="[도움말]", font=("Malgun Gothic", 9), bg="#607D8B", fg="white",
                  command=self.controller.show_help).pack(side="left", padx=10)
        
        # --- 하단 상태바 ---
        self.status_bar = tk.Label(self.root, text="준비됨 | 설정을 확인하고 '검색 미리보기'를 눌러주세요.",
                                    bd=1, relief="sunken", anchor="w", font=("Malgun Gothic", 8), padx=10)
        self.status_bar.pack(side="bottom", fill="x")
    
    def log(self, msg):
        """로그 메시지 추가"""
        self.log_txt.insert("end", f"{msg}\n")
        self.log_txt.see("end")
    
    def clear_log(self):
        """로그 초기화"""
        self.log_txt.delete("1.0", "end")
    
    def set_status(self, msg):
        """상태바 업데이트"""
        self.status_bar.config(text=msg)


# ==========================================
# 3. APPLICATION LAYER (Controller - 로직 연결)
# ==========================================
class CollectorController:
    """Controller: View와 Engine을 연결하고 사용자 이벤트 처리"""
    
    def __init__(self, root):
        self.root = root
        self.engine = FileCollectorEngine()
        self.is_running = False
        self.preview_files = []
        
        # === 상태 변수 초기화 ===
        self.source_dir_var = tk.StringVar(value=r"D:\02 기숙사 및 사택")
        self.pattern_var = tk.StringVar(value="마감자료")
        self.ext_var = tk.StringVar(value=".pdf")
        self.dest_dir_var = tk.StringVar(value=str(Path(__file__).parent / "00_Temp_Collected_Files"))
        self.use_prefix_var = tk.BooleanVar(value=True)
        self.use_suffix_var = tk.BooleanVar(value=False)
        self.smart_padding_var = tk.BooleanVar(value=True)
        self.custom_prefix_var = tk.StringVar(value="")
        self.custom_suffix_var = tk.StringVar(value="")
        self.exclude_folders_var = tk.StringVar(value="화학, 01 개별, 개별")
        
        # View 생성 (Controller 주입)
        self.view = CollectorView(root, self)

    
    def handle_select_source(self):
        """소스 폴더 선택"""
        d = filedialog.askdirectory(initialdir=self.source_dir_var.get(), title="검색할 최상위 폴더 선택")
        if d:
            self.source_dir_var.set(d)
            self.view.log(f"[*] 소스 폴더 설정: {d}")
    
    def handle_select_dest(self):
        """저장 폴더 선택"""
        d = filedialog.askdirectory(title="파일을 저장할 폴더 선택")
        if d:
            self.dest_dir_var.set(d)
            self.view.log(f"[*] 저장 폴더 설정: {d}")
    
    def handle_preview(self):
        """검색 미리보기 (파일 목록만 확인, 복사 안함)"""
        if self.is_running:
            return
        
        source = self.source_dir_var.get().strip()
        pattern = self.pattern_var.get().strip()
        ext = self.ext_var.get() if self.ext_var.get() != "(모든 파일)" else ""
        
        if not source or not os.path.isdir(source):
            messagebox.showwarning("경고", "유효한 소스 폴더를 선택해주세요.")
            return
        if not pattern:
            messagebox.showwarning("경고", "검색할 파일명 패턴을 입력해주세요.")
            return
        
        self.is_running = True
        self.view.clear_log()
        self.view.log(f"[INFO] 미리보기 시작 ({time.strftime('%H:%M:%S')})")
        self.view.log(f"   패턴: '{pattern}' | 확장자: '{ext or '모든 파일'}'")
        
        # 제외 폴더 파싱
        exclude_list = [f.strip() for f in self.exclude_folders_var.get().split(',') if f.strip()]
        if exclude_list:
            self.view.log(f"   제외 폴더: {exclude_list}")
        
        self.view.log("-" * 60)
        self.view.set_status("검색 중...")
        self.view.scan_btn.config(state="disabled")
        
        def run():
            result = self.engine.scan_files(source, pattern, ext, exclude_list, self._update_callback)
            self.root.after(0, lambda: self._preview_done(result))
        
        threading.Thread(target=run, daemon=True).start()

    
    def _update_callback(self, msg_type, msg):
        """엔진에서 호출되는 콜백 (스레드 안전)"""
        if msg_type == "status":
            self.root.after(0, lambda: self.view.set_status(msg))
        elif msg_type in ["found", "copy", "error", "skip"]:
            self.root.after(0, lambda: self.view.log(f"  {msg}"))
    
    def _preview_done(self, result):
        """미리보기 완료 처리"""
        self.is_running = False
        self.preview_files = result["files"]
        self.view.scan_btn.config(state="normal")
        
        self.view.log("-" * 60)
        self.view.log(f"[OK] 미리보기 완료!")
        self.view.log(f"   [F] 탐색된 폴더: {result['folder_count']}개")
        self.view.log(f"   [X] 제외된 폴더: {result.get('skipped_folders', 0)}개")
        self.view.log(f"   [P] 발견된 파일: {len(self.preview_files)}개")
        self.view.set_status(f"미리보기 완료: {len(self.preview_files)}개 파일 발견 (탐색: {result['folder_count']}개, 제외: {result.get('skipped_folders', 0)}개)")
        
        if self.preview_files:
            self.view.log("\n[INFO] 파일명 변경 분석 (일대일 비교):")
            self.view.log(f"{' [원본 상대 경로]':<45} | {' [변경 예정 파일명]'}")
            self.view.log("-" * 90)
            
            # 현재 설정값 가져오기
            prefix = ""
            if self.custom_prefix_var.get().strip():
                prefix = self.custom_prefix_var.get().strip()
            elif self.use_prefix_var.get():
                prefix = Path(self.source_dir_var.get()).name
            
            suffix = ""
            if self.custom_suffix_var.get().strip():
                suffix = self.custom_suffix_var.get().strip()
            elif self.use_suffix_var.get():
                suffix = Path(self.source_dir_var.get()).name
            
            padding = self.smart_padding_var.get()
            
            for i, f in enumerate(self.preview_files[:50], 1): # 가독성을 위해 상위 50개만 상세 출력
                orig_stem = Path(f['name']).stem
                ext = Path(f['name']).suffix
                
                # 패딩 적용
                padded_stem = FileCollectorEngine.apply_padding(orig_stem, padding)
                
                # 최종 조합
                final_stem = f"{prefix}_{padded_stem}" if prefix else padded_stem
                if suffix:
                    final_stem = f"{final_stem}_{suffix}"
                
                final_name = f"{final_stem}{ext}"
                
                # 경로가 길 경우 생략 표시
                orig_p = f['rel_path']
                if len(orig_p) > 42:
                    orig_p = "..." + orig_p[-39:]
                
                self.view.log(f" {i:02d}. {orig_p:<42}  --->  {final_name}")
            
            if len(self.preview_files) > 50:
                self.view.log(f"\n   ... 외 {len(self.preview_files) - 50}개 파일이 더 있습니다.")
            self.view.log("-" * 90)
    
    def handle_collect(self):
        """실제 파일 수집 실행"""
        if self.is_running:
            return
        
        if not self.preview_files:
            messagebox.showinfo("알림", "먼저 '검색 미리보기'를 실행하여 대상 파일을 확인해주세요.")
            return
        
        dest = self.dest_dir_var.get().strip()
        if not dest:
            messagebox.showwarning("경고", "저장 폴더를 선택해주세요.")
            return
        
        # 접두사 결정 (사용자 지정 텍스트가 있으면 우선, 없으면 체크박스 확인)
        if self.custom_prefix_var.get().strip():
            prefix = self.custom_prefix_var.get().strip()
        elif self.use_prefix_var.get():
            prefix = Path(self.source_dir_var.get()).name
        else:
            prefix = ""
            
        # 접미사 결정 (사용자 지정 텍스트가 있으면 우선, 없으면 체크박스 확인)
        if self.custom_suffix_var.get().strip():
            suffix = self.custom_suffix_var.get().strip()
        elif self.use_suffix_var.get():
            suffix = Path(self.source_dir_var.get()).name
        else:
            suffix = ""
        
        # 확인 다이얼로그
        msg = f"다음 작업을 실행합니다:\n\n" \
              f"• 복사할 파일: {len(self.preview_files)}개\n" \
              f"• 저장 위치: {dest}\n" \
              f"• 접두사: {prefix or '(없음)'}\n" \
              f"• 접미사: {suffix or '(없음)'}\n\n진행하시겠습니까?"
        if not messagebox.askyesno("확인", msg):
            return
        
        self.is_running = True
        self.view.log("\n" + "=" * 60)
        self.view.log(f"[INFO] 파일 수집 시작 ({time.strftime('%H:%M:%S')})")
        self.view.set_status("복사 중...")
        self.view.run_btn.config(state="disabled")
        
        def run():
            result = self.engine.copy_files(
                self.preview_files, dest, prefix, suffix, 
                smart_padding=self.smart_padding_var.get(), 
                callback=self._update_callback
            )
            self.root.after(0, lambda: self._collect_done(result, dest))
        
        threading.Thread(target=run, daemon=True).start()
    
    def _collect_done(self, result, dest):
        """수집 완료 처리"""
        self.is_running = False
        self.view.run_btn.config(state="normal")
        
        self.view.log("=" * 60)
        self.view.log(f"[OK] 수집 완료!")
        self.view.log(f"   [S] 성공: {result['success']}개")
        self.view.log(f"   [E] 오류: {len(result['errors'])}개")
        self.view.log(f"   [D] 저장 위치: {dest}")
        self.view.set_status(f"완료: {result['success']}개 복사됨")
        
        messagebox.showinfo("완료", f"파일 수집이 완료되었습니다.\n\n"
                                      f"• 성공: {result['success']}개\n"
                                      f"• 저장 위치: {dest}")
    
    def show_help(self):
        """도움말 표시"""
        help_msg = (
            "📖 [지능형 파일 수집기] 사용 가이드\n\n"
            "1️⃣ 검색 대상 폴더 선택\n"
            "   - 파일을 찾을 최상위 폴더를 선택합니다.\n"
            "   - 모든 하위 폴더가 재귀적으로 탐색됩니다.\n\n"
            "2️⃣ 검색 조건 설정\n"
            "   - 파일명 패턴: 찾을 파일명에 포함된 키워드\n"
            "   - 확장자: 특정 확장자만 필터링 가능\n\n"
            "3️⃣ 출력 설정\n"
            "   - 저장 폴더: 수집된 파일을 저장할 위치\n"
            "   - 접두사: 파일명 앞에 붙일 텍스트\n\n"
            "4️⃣ 실행 순서\n"
            "   ① '검색 미리보기' → 대상 파일 목록 확인\n"
            "   ② '수집 실행' → 실제 파일 복사\n\n"
            "💡 Tip: 중복 파일명은 '_1, _2' 등의 숫자가 붙어 모두 보존됩니다."
        )
        messagebox.showinfo("도움말", help_msg)


# ==========================================
# 4. ENTRY POINT
# ==========================================
if __name__ == "__main__":
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    CollectorController(root)
    root.mainloop()
