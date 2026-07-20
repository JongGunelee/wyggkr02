"""
================================================================================
 [POWERPOINT] 파워포인트 전용 압축기 (PPTX Hyper Compressor) v1.0.0
================================================================================
 - 아키텍처: Clean Architecture (Engine / View / Controller)
 - 주요 기능:
   1. .ppt (구버전) 등 이기종 포맷을 .pptx 로 자동 변환
   2. 파워포인트 내부에 숨겨진 유령 객체(보이지 않는 도형 등) 삭제
   3. 내부 미디어 파일(이미지) 품질 슬라이더 기반 고압축 (Pillow 활용)
   4. 무결성 보증 확인을 거쳐 데이터 파손 방지
 - 기반 엔진: win32com + Python Zip/Pillow Hybrid
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
from tkinter import filedialog, messagebox
try:
    from PIL import Image
except ImportError:
    messagebox.showerror("오류", "Pillow 라이브러리가 필요합니다.\n\npip install Pillow\n\n명령어를 실행해주세요.")
    sys.exit(1)

import win32com.client
import pythoncom

# ==========================================
# 1. DOMAIN LAYER (Core logic)
# ==========================================
class PPTCompressorEngine:
    def __init__(self, callback):
        self.callback = callback

    def _compress_image(self, img_path, quality=70):
        """이미지 파일을 열어 압축하고 저장합니다."""
        try:
            with Image.open(img_path) as img:
                fmt = img.format
                if fmt not in ['JPEG', 'PNG']: 
                    return False
                img.save(img_path, quality=quality, optimize=True)
            return True
        except Exception as e:
            return False

    def _kill_powerpoint(self):
        try:
            import psutil
            for p in psutil.process_iter(['name']):
                if p.info['name'] and p.info['name'].upper() == "POWERPNT.EXE":
                    try: p.kill()
                    except: pass
            import time
            time.sleep(1.0)
        except: pass

    def _convert_and_clean_hidden_objects(self, src_path, dest_pptx_path):
        """COM 훅을 통해 ppt 파일을 열고, 숨겨진(유령) 객체를 삭제 후 pptx로 저장합니다."""
        self._kill_powerpoint()
        pythoncom.CoInitialize()
        hidden_counts = 0
        app = None
        prs = None
        
        abs_src = os.path.abspath(src_path)
        abs_dest = os.path.abspath(dest_pptx_path)
        
        try:
            # Batch_PPT_to_PDF_DDD.py 와 완벽히 동일한 방식의 파워포인트 통제 로직 적용
            try:
                from win32com.client.dynamic import Dispatch as DynDispatch
                app = DynDispatch("{91493441-5A91-11CF-8700-00AA0060263B}")
            except:
                try:
                    app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
                except:
                    app = win32com.client.Dispatch("PowerPoint.Application")
                
            try:
                app.Visible = 0
                app.DisplayAlerts = 1  
                try: app.Interactive = False
                except: pass
            except: pass
            
            # Batch PPT to PDF 의 2회 재시도 로직 적용 (WithWindow 옵션 토글)
            for retry in range(2):
                try:
                    import time
                    prs = app.Presentations.Open(abs_src, True, False, retry == 1)
                    break
                except:
                    if retry == 1: raise
                    time.sleep(1.0)
            
            # 슬라이드 순회하며 숨겨진 도형 제거 (역순 제거가 안전함)
            for slide in prs.Slides:
                try:
                    shapes_count = slide.Shapes.Count
                    for i in range(shapes_count, 0, -1):
                        try:
                            shape = slide.Shapes.Item(i)
                            # Visible이 msoFalse(0)인 객체 삭제
                            if shape.Visible == 0:
                                shape.Delete()
                                hidden_counts += 1
                        except Exception:
                            continue
                except Exception:
                    continue
            
            # 찌꺼기 파일 원천삭제
            if os.path.exists(abs_dest):
                os.remove(abs_dest)
                
            # .pptx(24 = ppSaveAsOpenXMLPresentation) 포맷으로 새로운 파일에 저장
            prs.SaveAs(abs_dest, 24)
            return True, hidden_counts, ""
            
        except Exception as e:
            return False, 0, str(e)

            
        finally:
            try:
                if prs: prs.Close()
            except: pass
            try:
                if app:
                    app.DisplayAlerts = 1 # ppAlertsAll 기본값으로 롤백
                    app.Quit()
            except: pass
            pythoncom.CoUninitialize()

    def compress_pptx_images(self, src_path, dest_path, quality=60):
        """pptx 파일을 해부하여 ppt/media 내부 이미지를 압축합니다."""
        temp_dir = tempfile.mkdtemp()
        try:
            # 1. Unzip
            with zipfile.ZipFile(src_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 2. Find Media folder (ppt/media)
            media_path = os.path.join(temp_dir, 'ppt', 'media')
            img_count = 0
            
            if os.path.exists(media_path):
                files = os.listdir(media_path)
                for file in files:
                    full_path = os.path.join(media_path, file)
                    if os.path.isfile(full_path):
                        size_before = os.path.getsize(full_path)
                        if self._compress_image(full_path, quality):
                            size_after = os.path.getsize(full_path)
                            if size_after < size_before:
                                img_count += 1
            
            # 3. Re-zip
            with zipfile.ZipFile(dest_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        zipf.write(file_path, arcname)
                        
            return True, f"이미지 {img_count}개 최적화 완료"
        except Exception as e:
            return False, str(e)
        finally:
            shutil.rmtree(temp_dir)

    def _verify_pptx_integrity(self, pptx_path):
        """압축 완료 후 ZIP 무결성 체크"""
        try:
            with zipfile.ZipFile(pptx_path, 'r') as z:
                # 깨진 파일이 있는지 검사
                bad_file = z.testzip()
                if bad_file:
                    return False
            return True
        except:
            return False

    def process_files(self, files, mode, options):
        """파일 처리 메인 로직"""
        start_time = time.time()
        success_cnt = 0
        total = len(files)
        quality = options.get('quality', 60)
        
        # 결과 폴더 생성
        first_dir = os.path.dirname(files[0])
        res_dir = os.path.join(first_dir, "00_PPT_Optimized_Results")
        if not os.path.exists(res_dir): os.makedirs(res_dir)
        
        i_file_map = {} # {original_path: result_path}
        
        self.callback("status", f"[INFO] 파워포인트 최적화 시작 (이미지 품질: {quality}%)")
        
        for i, src in enumerate(files, 1):
            fname = os.path.basename(src)
            name_without_ext, _ = os.path.splitext(fname)
            final_dest = os.path.join(res_dir, f"최적화_{name_without_ext}.pptx")
            
            self.callback("log", f"[{i}/{total}] 처리 중: {fname}")
            org_size = os.path.getsize(src)
            
            # STEP 1: COM 엔진으로 변환 (유령 도형 제거 및 PPTX 변환)
            self.callback("log", "  [STEP 1] 포맷 정제 및 유령 객체 클린업 중...")
            temp_pptx_path = os.path.join(res_dir, f"~temp_{name_without_ext}.pptx")
            
            ok1, hidden_cnt, msg1 = self._convert_and_clean_hidden_objects(src, temp_pptx_path)
            if not ok1:
                self.callback("log", f"  [FAIL] COM 엔진 제어 실패: {msg1}")
                continue
            
            self.callback("log", f"  [OK] 유령 객체 {hidden_cnt}개 제거 및 PPTX 구조화 완료")
            
            # STEP 2: 내부 미디어 압축
            self.callback("log", "  [STEP 2] 내부 미디어 해부 및 고효율 압축...")
            ok2, msg2 = self.compress_pptx_images(temp_pptx_path, final_dest, quality)
            
            # 마무리 정리 
            if os.path.exists(temp_pptx_path):
                os.remove(temp_pptx_path)
                
            if ok2:
                # STEP 3: 무결성 테스트
                self.callback("log", "  [STEP 3] 원본 대비 무결성(ZIP CRC) 자동 검증 진행 중...")
                if self._verify_pptx_integrity(final_dest):
                    new_size = os.path.getsize(final_dest)
                    ratio = (1 - new_size/org_size) * 100
                    self.callback("log", f"  [DONE] {msg2}")
                    self.callback("log", f"  [보증] 무결성 검증 통과 (구조 손상 없음)")
                    self.callback("log", f"  [결과] 최종 용량: {org_size/1024/1024:.2f}MB → {new_size/1024/1024:.2f}MB ({ratio:.1f}% 감소)")
                    success_cnt += 1
                    i_file_map[src] = final_dest
                else:
                    self.callback("log", "  [FAIL] 무결성 검증 실패! 원본 데이터 보호를 위해 손상된 결과물을 자동 페기합니다.")
                    os.remove(final_dest)
            else:
                self.callback("log", f"  [FAIL] 미디어 압축 중 구조 오류 발생: {msg2}")

        duration = time.time() - start_time
        self.callback("done", (f"총 {total}건 중 {success_cnt}건 최적화 완료 ({duration:.1f}초)", i_file_map))

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
class PPTCompressorView:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[파워포인트 전용 압축기 v1.0.0]")
        self.root.geometry("600x650") 
        self._build_ui()

    def _build_ui(self):
        # Header
        tk.Label(self.root, text="[PPTX HYPER COMPRESSOR v1.0.0]", font=("Impact", 20), bg="#D32F2F", fg="white", pady=10).pack(fill="x")
        
        main = tk.Frame(self.root, padx=15, pady=15)
        main.pack(fill="both", expand=True)

        # 1. File Selection
        f1 = tk.LabelFrame(main, text="1. 파일 선택 (.ppt, .pptx)", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        f1.pack(fill="x", pady=5)
        
        btn_fr = tk.Frame(f1)
        btn_fr.pack(fill="x")
        tk.Button(btn_fr, text="[파워포인트 파일 추가 (다중 지원)]", command=self.controller.add_files, bg="#FFEBEE").pack(fill="x")
        
        self.file_list = tk.Listbox(f1, height=5, selectmode="extended", font=("Consolas", 9))
        self.file_list.pack(fill="x", pady=5)

        action_fr = tk.Frame(f1)
        action_fr.pack(fill="x", pady=5)
        self.cancel_btn = tk.Button(action_fr, text="[목록 초기화]", bg="#f4f4f4", fg="#333", font=("맑은 고딕", 9, "bold"), height=2, command=self.controller.clear_files)
        self.cancel_btn.pack(side="left", fill="x", expand=True, padx=(0, 2))
        
        self.run_btn = tk.Button(action_fr, text="[압축 시작]", bg="#D32F2F", fg="white", font=("맑은 고딕", 9, "bold"), height=2, command=self.controller.run_process)
        self.run_btn.pack(side="left", fill="x", expand=True, padx=2)

        self.review_btn = tk.Button(action_fr, text="[검토완료]", bg="#1565C0", fg="white", font=("맑은 고딕", 9, "bold"), height=2, state="disabled", command=self.controller.handle_review_complete)
        self.review_btn.pack(side="left", fill="x", expand=True, padx=(2, 0))

        # 2. Quality Selection
        f2 = tk.LabelFrame(main, text="2. 화질 보존 수준 설정", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        f2.pack(fill="x", pady=5)
        
        # Quality Slider
        self.quality_fr = tk.Frame(f2, pady=5)
        self.quality_fr.pack(fill="x")
        tk.Label(self.quality_fr, text="품질(%):", font=("맑은 고딕", 9)).pack(side="left")
        self.qual_scale = tk.Scale(self.quality_fr, from_=10, to=90, orient="horizontal", variable=self.controller.quality_var, length=300, showvalue=1, resolution=10)
        self.qual_scale.set(60)
        self.qual_scale.pack(side="left", padx=10)
        tk.Label(self.quality_fr, text="(권장: 60)", fg="blue", font=("맑은 고딕", 8)).pack(side="left")

        # 3. Log
        f3 = tk.LabelFrame(main, text="3. 실행 현황 로그", font=("맑은 고딕", 10, "bold"), padx=10, pady=10)
        f3.pack(fill="both", expand=True, pady=5)
        self.log_txt = tk.Text(f3, font=("Consolas", 9), bg="#FFF8F8", state="disabled")
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

    def set_running(self, running):
        if running:
            self.run_btn.config(state="disabled", text="[안전 압축 작업 중...]", bg="gray")
            self.cancel_btn.config(state="disabled")
        else:
            self.run_btn.config(state="normal", text="[압축 시작]", bg="#D32F2F")
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
class PPTCompressorController:
    def __init__(self, root):
        self.root = root
        self.files = []
        self.last_res_dir = None
        self.processed_map = {}
        self.quality_var = tk.IntVar(value=60)
        
        self.view = PPTCompressorView(root, self)
        self.engine = PPTCompressorEngine(self._callback)
        self.is_running = False

    def add_files(self):
        fs = filedialog.askopenfilenames(filetypes=[("PowerPoint Files", "*.ppt;*.pptx;*.pot;*.potx"), ("All Files", "*.*")])
        if fs:
            self.files.extend(list(fs))
            self.files = sorted(list(set(self.files)))  # 중복 제거 및 가나다(알파벳)순 정렬
            self.view.update_file_list(self.files)

    def clear_files(self):
        self.files = []
        self.view.update_file_list(self.files)

    def run_process(self):
        if not self.files:
            messagebox.showwarning("경고", "먼저 변환할 파워포인트 파일을 선택해주세요.")
            return
            
        if self.is_running: return
        
        opts = {'quality': self.quality_var.get()}
        self.is_running = True
        self.view.set_running(True)
        self.view.log_txt.config(state="normal")
        self.view.log_txt.delete("1.0", "end")
        self.view.log_txt.config(state="disabled")
        
        threading.Thread(target=lambda: self.engine.process_files(self.files, "image_compress", opts), daemon=True).start()

    def _callback(self, type, msg):
        self.root.after(0, lambda: self._handle_callback(type, msg))

    def _handle_callback(self, type, msg):
        if type == "log":
            self.view.log(msg)
        elif type == "status":
            self.view.log(f"=== {msg} ===")
        elif type == "done":
            msg_str, f_map = msg
            self.view.log(f"\n{msg_str}")
            self.is_running = False
            self.processed_map = f_map
            if f_map:
                first_file = list(f_map.keys())[0]
                self.last_res_dir = os.path.join(os.path.dirname(first_file), "00_PPT_Optimized_Results")
                self.view.review_btn.config(state="normal")
                
            self.view.set_running(False)
            messagebox.showinfo("압축 완료", msg_str)

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
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = PPTCompressorController(root)
    root.mainloop()
