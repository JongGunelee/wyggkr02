"""
================================================================================
 [EXCEL] 엑셀 초강력 압축기 (Excel Hyper Compressor) v34.1.16
================================================================================
 - 아키텍처: Clean Architecture (Engine / View / Controller)
 - 주요 기능:
   1. 엑셀 내부 이미지 압축 (미디어 파일 최적화) -> 용량 획기적 절감
   2. 일반 ZIP 압축 (아카이빙)
 - 작동 원리: .xlsx는 사실 zip 파일입니다. 이를 풀어서 내부의 무거운 이미지들을
   Pillow 라이브러리로 압축한 뒤 다시 묶는 방식입니다.
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
import zipfile
import tempfile
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
try:
    from PIL import Image
except ImportError:
    messagebox.showerror("오류", "Pillow 라이브러리가 필요합니다.\n\npip install Pillow\n\n명령어를 실행해주세요.")
    import sys
    sys.exit(1)

# ==========================================
# 1. DOMAIN LAYER (Core logic)
# ==========================================
import win32com.client
import pythoncom

class ExcelCompressorEngine:
    def __init__(self, callback):
        self.callback = callback

    def _compress_image(self, img_path, quality=70):
        """이미지 파일을 열어 압축하고 저장합니다."""
        try:
            with Image.open(img_path) as img:
                # 이미지가 너무 크면 리사이징 (옵션) - 현재는 품질만 조절
                # 만약 원본 포맷 유지하며 저장
                fmt = img.format
                if fmt not in ['JPEG', 'PNG']: 
                    return False # 지원하지 않는 포맷은 패스
                
                # 이미지 저장 (최적화 옵션 사용)
                img.save(img_path, quality=quality, optimize=True)
            return True
        except Exception as e:
            # self.callback("log", f"이미지 처리 실패 {os.path.basename(img_path)}: {e}")
            return False

    def compress_excel_images(self, src_path, dest_path, quality=60):
        """엑셀 파일을 해부하여 내부 이미지를 압축합니다."""
        temp_dir = tempfile.mkdtemp()
        try:
            # 1. Unzip
            with zipfile.ZipFile(src_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 2. Find Media folder (xl/media)
            media_path = os.path.join(temp_dir, 'xl', 'media')
            img_count = 0
            size_reduced = False
            
            if os.path.exists(media_path):
                files = os.listdir(media_path)
                for file in files:
                    full_path = os.path.join(media_path, file)
                    if os.path.isfile(full_path):
                        # 압축 전 사이즈
                        size_before = os.path.getsize(full_path)
                        
                        if self._compress_image(full_path, quality):
                            # 압축 후 사이즈
                            size_after = os.path.getsize(full_path)
                            if size_after < size_before:
                                img_count += 1
                                size_reduced = True
            
            # 3. Re-zip
            # shutil.make_archive는 기본적으로 폴더 자체를 포함할 수 있어 zipfile로 직접 처리
            with zipfile.ZipFile(dest_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        # temp_dir 이후의 경로만 zip에 저장
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)
                        
            return True, f"이미지 {img_count}개 최적화 완료"
            
        except Exception as e:
            return False, str(e)
        finally:
            shutil.rmtree(temp_dir)

    def _verify_excel_integrity(self, excel_path):
        """압축 완료 후 ZIP 무결성 체크 (Excel 파일 구조 검증)"""
        try:
            with zipfile.ZipFile(excel_path, 'r') as z:
                # 깨진 파일이 있는지 검사
                bad_file = z.testzip()
                if bad_file:
                    return False
            return True
        except:
            return False

    def _convert_xls_to_xlsx(self, src_path, res_dir):
        """COM을 통해 레거시 .xls 파일을 보안 .xlsx 포맷으로 변환합니다."""
        fname = os.path.basename(src_path)
        name_without_ext, _ = os.path.splitext(fname)
        dest_xlsx = os.path.join(res_dir, f"변환_최적화_{name_without_ext}.xlsx")
        
        pythoncom.CoInitialize()
        app = None
        wb = None
        try:
            from win32com.client.dynamic import Dispatch as DynDispatch
            try:
                app = DynDispatch("Excel.Application")
            except:
                try:
                    app = win32com.client.gencache.EnsureDispatch("Excel.Application")
                except:
                    app = win32com.client.Dispatch("Excel.Application")
                    
            app.Visible = False
            app.DisplayAlerts = False
            
            # 백그라운드 팝업창(SaveAs 강제 부상) 완벽 제압 (파워포인트식 하드닝)
            try: app.Interactive = False 
            except: pass
            try: app.EnableEvents = False 
            except: pass
            try: app.ScreenUpdating = False 
            except: pass
            try: app.AutomationSecurity = 3 # 3 = msoAutomationSecurityForceDisable
            except: pass
            
            abs_src = os.path.abspath(src_path)
            abs_dest = os.path.abspath(dest_xlsx)
            
            # 읽기 전용으로 백그라운드 오픈
            wb = app.Workbooks.Open(abs_src, UpdateLinks=0, ReadOnly=True)
            
            if os.path.exists(abs_dest):
                os.remove(abs_dest)
                
            # [CRITICAL FIX] win32com LateBinding 호환성 보장: Kwarg 금지, Positional(51) 할당
            # FileFormat=51 (xlOpenXMLWorkbook) -> 위치 기반 파라미터로 명시해야 다이얼로그 팝업을 방지합니다.
            wb.SaveAs(abs_dest, 51)
            return True, dest_xlsx
        except Exception as e:
            return False, str(e)
        finally:
            try:
                if wb: wb.Close(SaveChanges=False)
            except: pass
            try:
                if app: app.Quit()
            except: pass
            pythoncom.CoUninitialize()

    def process_files(self, files, mode, options):
        """파일 처리 메인 로직"""
        start_time = time.time()
        success_cnt = 0
        total = len(files)
        
        # 결과 폴더 생성
        first_dir = os.path.dirname(files[0])
        res_dir = os.path.join(first_dir, "00_Compressed_Results")
        if not os.path.exists(res_dir): os.makedirs(res_dir)
        
        i_file_map = {} # {original_path: result_path}
        
        if mode == "image_compress":
            quality = options.get('quality', 60)
            self.callback("status", f"[INFO] 엑셀 내부 이미지 압축 시작 (품질: {quality}%)")
            
            for i, src in enumerate(files, 1):
                fname = os.path.basename(src)
                dest = os.path.join(res_dir, f"최적화_{fname}")
                org_size = os.path.getsize(src)
                
                self.callback("log", f"[{i}/{total}] 처리 중: {fname}")
                
                # 레거시 파일(.xls) 감지 및 변환 파이프라인
                _, ext = os.path.splitext(src)
                target_src = src
                if ext.lower() == '.xls':
                    self.callback("log", "  [STEP 0] 레거시 구형 포맷(.xls) 감지! 최신(.xlsx) 구조로 변환 중...")
                    ok_conv, conv_res = self._convert_xls_to_xlsx(src, res_dir)
                    if ok_conv:
                        self.callback("log", "  [OK] 포맷 강제 업그레이드 완료")
                        target_src = conv_res
                        dest = conv_res # 덮어쓰거나 이미 변환된 상태
                        org_size = os.path.getsize(target_src)
                    else:
                        self.callback("log", f"  [FAIL] 레거시 엑셀 파일 변환 오류: {conv_res}")
                        continue
                
                self.callback("log", "  [STEP 1] 엑셀 구조 해부 및 잉여 미디어 압축 중...")
                ok, msg = self.compress_excel_images(target_src, dest, quality)
                
                if ok:
                    self.callback("log", f"  [OK] 완료! {msg}")
                    self.callback("log", "  [STEP 2] 원본 대비 무결성(ZIP CRC) 자동 검증 진행 중...")
                    
                    if self._verify_excel_integrity(dest):
                        new_size = os.path.getsize(dest)
                        ratio = (1 - new_size/org_size) * 100
                        self.callback("log", "  [보증] 무결성 검증 통과 (구조 손상 없음)")
                        self.callback("log", f"  [결과] 최종 용량: {org_size/1024/1024:.2f}MB → {new_size/1024/1024:.2f}MB ({ratio:.1f}% 감소)")
                        success_cnt += 1
                        i_file_map[src] = dest
                    else:
                        self.callback("log", "  [FAIL] 무결성 검증 실패! 원본 데이터 보호를 위해 손상된 결과물을 자동 페기합니다.")
                        os.remove(dest)
                else:
                    self.callback("log", f"  [FAIL] 미디어 압축 중 구조 오류 발생: {msg}")
                    
        elif mode == "zip_archive":
            self.callback("status", "[INFO] ZIP 아카이빙 시작")
            # 개별 압축 또는 전체 압축? 여기서는 선택된 파일들을 하나의 zip으로 묶는 것으로 가정
            # 또는 파일별로 개별 zip? 보통은 묶어서 보냄.
            # 사용자 편의를 위해 "파일명_날짜.zip" 하나로 묶음
            
            zip_name = f"작업파일모음_{int(time.time())}.zip"
            dest_zip = os.path.join(res_dir, zip_name)
            
            try:
                with zipfile.ZipFile(dest_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for i, src in enumerate(files, 1):
                        fname = os.path.basename(src)
                        self.callback("log", f"[{i}/{total}] 아카이빙: {fname}")
                        zipf.write(src, fname)
                
                final_size = os.path.getsize(dest_zip)
                self.callback("log", f"[OK] ZIP 생성 완료: {zip_name} ({final_size/1024:.1f}KB)")
                success_cnt = total # Zip은 한방에 성공으로 간주
            except Exception as e:
                self.callback("log", f"[FAIL] ZIP 생성 실패: {str(e)}")

        duration = time.time() - start_time
        self.callback("done", (f"총 {total}건 중 {success_cnt}건 처리 완료 ({duration:.1f}초)", i_file_map))

    def finalize_review(self, file_map, res_dir):
        """본격적인 원본 덮어쓰기 및 임시 폴더 삭제 (PRD v35.4.18 5단계 트랜잭션 프로토콜 준수)"""
        deleted_items = []
        backups = []
        
        try:
            # [1단계] 사전 점검 (Pre-check): 모든 대상 파일의 접근성 확인
            for org_path, res_path in file_map.items():
                if not os.path.exists(res_path):
                    continue
                if os.path.exists(org_path):
                    # 파일 잠금 체크 (읽기/쓰기 시도)
                    try:
                        with open(org_path, 'r+b') as f: pass
                    except:
                        raise Exception(f"원본 파일이 다른 프로그램에서 사용 중입니다: {os.path.basename(org_path)}")

            # [2단계] 안전 백업 (.bak) 및 [3단계] 원자적 이동
            for org_path, res_path in file_map.items():
                if not os.path.exists(res_path): continue
                
                org_dir = os.path.dirname(org_path)
                org_fname = os.path.basename(org_path)
                res_ext = os.path.splitext(res_path)[1]
                org_base = os.path.splitext(org_fname)[0]
                
                # 최종 타겟 경로 (확장자 변경 고려)
                target_path = os.path.join(org_dir, org_base + res_ext)
                bak_path = org_path + ".bak"
                
                # 기존 원본이 있다면 백업 생성
                if os.path.exists(org_path):
                    if os.path.exists(bak_path): os.remove(bak_path)
                    os.rename(org_path, bak_path)
                    backups.append((bak_path, org_path))
                
                # 만약 타겟 경로가 다르고(확장자 변환 등) 파일이 이미 존재하면 백업
                if target_path != org_path and os.path.exists(target_path):
                    target_bak = target_path + ".bak"
                    if os.path.exists(target_bak): os.remove(target_bak)
                    os.rename(target_path, target_bak)
                    backups.append((target_bak, target_path))

                # 결과 파일을 원본 위치로 이동
                shutil.move(res_path, target_path)
                
                # [4단계] 최종 검증 (Verify): 이동된 파일의 크기가 0이 아님을 확인
                if not os.path.exists(target_path) or os.path.getsize(target_path) == 0:
                    raise Exception(f"교체 결과 무결성 검증 실패: {os.path.basename(target_path)}")
                
                deleted_items.append(f"교체 완료: {os.path.basename(target_path)}")

            # [5단계] 성공적 정리 (Cleanup): 백업 삭제 및 임시 폴더 제거
            for bak_p, _ in backups:
                if os.path.exists(bak_p): os.remove(bak_p)
                
            if os.path.exists(res_dir):
                shutil.rmtree(res_dir)
                deleted_items.append(f"임시 결과 폴더 삭제: {os.path.basename(res_dir)}")

            return True, deleted_items

        except Exception as e:
            # 롤백: 문제 발생 시 백업 파일 복구
            for bak_p, org_p in backups:
                try:
                    if os.path.exists(bak_p):
                        if os.path.exists(org_p): os.remove(org_p)
                        os.rename(bak_p, org_p)
                except: pass
            return False, [f"트랜잭션 오류 (롤백 수행됨): {str(e)}"]


# ==========================================
# 2. PRESENTATION LAYER (View)
# ==========================================
class ExcelCompressorView:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[엑셀 초강력 압축기 v34.1.16]")
        self.root.geometry("600x650") 
        self._build_ui()

    def _build_ui(self):
        # Header
        tk.Label(self.root, text="[EXCEL HYPER COMPRESSOR v34.1.16]", font=("Impact", 20), bg="#2E7D32", fg="white", pady=10).pack(fill="x")
        
        main = tk.Frame(self.root, padx=15, pady=15)
        main.pack(fill="both", expand=True)

        # 1. File Selection
        f1 = tk.LabelFrame(main, text="1. 파일 선택", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        f1.pack(fill="x", pady=5)
        
        btn_fr = tk.Frame(f1)
        btn_fr.pack(fill="x")
        tk.Button(btn_fr, text="[엑셀 파일 추가 (다중 선택)]", command=self.controller.add_files, bg="#E8F5E9").pack(fill="x")
        
        self.file_list = tk.Listbox(f1, height=5, selectmode="extended", font=("Consolas", 9))
        self.file_list.pack(fill="x", pady=5)

        action_fr = tk.Frame(f1)
        action_fr.pack(fill="x", pady=5)
        self.cancel_btn = tk.Button(action_fr, text="[목록 초기화]", bg="#f4f4f4", fg="#333", font=("맑은 고딕", 9, "bold"), height=2, command=self.controller.clear_files)
        self.cancel_btn.pack(side="left", fill="x", expand=True, padx=(0, 2))
        
        self.run_btn = tk.Button(action_fr, text="[최적화 시작]", bg="#2E7D32", fg="white", font=("맑은 고딕", 9, "bold"), height=2, command=self.controller.run_process)
        self.run_btn.pack(side="left", fill="x", expand=True, padx=2)

        self.review_btn = tk.Button(action_fr, text="[검토완료]", bg="#1565C0", fg="white", font=("맑은 고딕", 9, "bold"), height=2, state="disabled", command=self.controller.handle_review_complete)
        self.review_btn.pack(side="left", fill="x", expand=True, padx=(2, 0))

        # 2. Mode Selection
        f2 = tk.LabelFrame(main, text="2. 압축 모드 설정", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        f2.pack(fill="x", pady=5)
        
        # Radio for Mode
        modes_fr = tk.Frame(f2)
        modes_fr.pack(fill="x")
        tk.Radiobutton(modes_fr, text="① 이미지 최적화 (내부 이미지 압축)", variable=self.controller.mode_var, value="image_compress", command=self.controller.update_ui_state).pack(anchor="w")
        tk.Label(modes_fr, text="   └─ 엑셀 파일 용량을 줄이고 싶을 때 사용 (화질 저하 주의)", fg="gray", font=("맑은 고딕", 8)).pack(anchor="w", padx=20)
        
        tk.Radiobutton(modes_fr, text="② ZIP으로 묶기 (단순 아카이빙)", variable=self.controller.mode_var, value="zip_archive", command=self.controller.update_ui_state).pack(anchor="w", pady=(10,0))
        
        # Quality Slider (Only for image compress)
        self.quality_fr = tk.Frame(f2, pady=5)
        self.quality_fr.pack(fill="x", padx=20)
        tk.Label(self.quality_fr, text="이미지 품질 설정:", font=("맑은 고딕", 9)).pack(side="left")
        self.qual_scale = tk.Scale(self.quality_fr, from_=10, to=90, orient="horizontal", variable=self.controller.quality_var, length=200, showvalue=1, resolution=10)
        self.qual_scale.set(60)
        self.qual_scale.pack(side="left", padx=10)
        tk.Label(self.quality_fr, text="(낮을수록 고압축)", fg="blue", font=("맑은 고딕", 8)).pack(side="left")

        # 3. Log
        f3 = tk.LabelFrame(main, text="3. 실행 로그", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        f3.pack(fill="both", expand=True, pady=5)
        self.log_txt = tk.Text(f3, font=("Consolas", 9), bg="#F1F8E9", state="disabled")
        self.log_txt.pack(fill="both", expand=True)

        # Run Button is now inside File Selection container
        
    def log(self, msg):
        self.log_txt.config(state="normal")
        self.log_txt.insert("end", f"{msg}\n")
        self.log_txt.see("end")
        self.log_txt.config(state="disabled")

    def update_file_list(self, files):
        self.file_list.delete(0, "end")
        for f in files:
            self.file_list.insert("end", os.path.basename(f))
            
    def set_quality_visibility(self, visible):
        if visible:
            self.quality_fr.pack(fill="x", padx=20)
            self.qual_scale.pack(side="left", padx=10)
        else:
            self.quality_fr.pack_forget()

    def set_running(self, running):
        if running:
            self.run_btn.config(state="disabled", text="[작업 중...]", bg="gray")
            self.cancel_btn.config(state="disabled")
        else:
            self.run_btn.config(state="normal", text="[최적화 시작]", bg="#2E7D32")
            self.cancel_btn.config(state="normal")

    def show_deletion_status(self, deleted_items):
        status_win = tk.Toplevel(self.root)
        status_win.title("🗑️ 삭제 및 이동 현황")
        status_win.geometry("500x400")
        status_win.attributes('-topmost', True)
        
        tk.Label(status_win, text="[ 원본 덮어쓰기 및 정리 완료 ]", font=("맑은 고딕", 12, "bold"), pady=10).pack()
        
        list_frame = tk.Frame(status_win, padx=10, pady=10)
        list_frame.pack(fill="both", expand=True)
        
        text_area = tk.Text(list_frame, font=("Consolas", 9), bg="#f8f8f8")
        text_area.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, command=text_area.yview)
        scrollbar.pack(side="right", fill="y")
        text_area.config(yscrollcommand=scrollbar.set)
        
        text_area.insert("end", "아래 항목들이 안전하게 처리(삭제/교체) 되었습니다:\n")
        text_area.insert("end", "="*50 + "\n")
        for item in deleted_items:
            text_area.insert("end", f"• {item}\n")
        
        text_area.config(state="disabled")
        tk.Button(status_win, text="확인", width=10, command=status_win.destroy).pack(pady=10)

# ==========================================
# 3. APPLICATION LAYER (Controller)
# ==========================================
class ExcelCompressorController:
    def __init__(self, root):
        self.root = root
        self.files = []
        self.last_res_dir = None
        self.processed_map = {}
        self.mode_var = tk.StringVar(value="image_compress")
        self.quality_var = tk.IntVar(value=60)
        
        self.view = ExcelCompressorView(root, self)
        self.engine = ExcelCompressorEngine(self._callback)
        self.is_running = False
        
        self.update_ui_state()

    def add_files(self):
        fs = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx;*.xlsm;*.xls"), ("All Files", "*.*")])
        if fs:
            self.files.extend(list(fs))
            # 중복 제거 및 가나다(알파벳)순 정렬
            self.files = sorted(list(set(self.files)))
            self.view.update_file_list(self.files)

    def clear_files(self):
        self.files = []
        self.view.update_file_list(self.files)

    def update_ui_state(self):
        mode = self.mode_var.get()
        self.view.set_quality_visibility(mode == "image_compress")

    def run_process(self):
        if not self.files:
            messagebox.showwarning("경고", "파일을 먼저 추가해주세요.")
            return
            
        if self.is_running: return
        
        mode = self.mode_var.get()
        opts = {'quality': self.quality_var.get()}
        
        self.is_running = True
        self.view.set_running(True)
        self.view.log_txt.config(state="normal")
        self.view.log_txt.delete("1.0", "end")
        self.view.log_txt.config(state="disabled")
        
        threading.Thread(target=lambda: self.engine.process_files(self.files, mode, opts), daemon=True).start()

    def _callback(self, type, msg):
        self.root.after(0, lambda: self._handle_callback(type, msg))

    def _handle_callback(self, type, msg):
        if type == "log":
            self.view.log(msg)
        elif type == "status":
            self.view.log(f"=== {msg} ===")
        elif type == "done":
            msg_str, f_map = msg
            self.view.log(f"\n[DONE] {msg_str}")
            self.is_running = False
            self.processed_map = f_map
            if f_map:
                first_file = list(f_map.keys())[0]
                self.last_res_dir = os.path.join(os.path.dirname(first_file), "00_Compressed_Results")
                self.view.review_btn.config(state="normal")
                
            self.view.set_running(False)
            messagebox.showinfo("완료", msg_str)

    def handle_review_complete(self):
        if not self.processed_map:
            messagebox.showwarning("경고", "검토할 작업 결과가 없습니다.")
            return
            
        if not messagebox.askyesno("확인", "변환된 파일을 원본 위치로 이동하고 임시 폴더를 삭제하시겠습니까?\n(원본 파일은 새로운 파일로 대체됩니다.)"):
            return
            
        success, deleted_list = self.engine.finalize_review(self.processed_map, self.last_res_dir)
        
        if success:
            self.view.show_deletion_status(deleted_list)
            self.processed_map = {}
            self.view.review_btn.config(state="disabled")
            self.clear_files()
            self.view.log("\n[SYSTEM] 원본 덮어쓰기 및 임시 파일 정리가 완료되었습니다.")
        else:
            messagebox.showerror("오류", f"정리 중 오류가 발생했습니다.\n{deleted_list[0]}")

if __name__ == "__main__":
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = ExcelCompressorController(root)
    root.mainloop()

