"""
================================================================================
[PPT] PPT -> PDF 일괄 변환기 (Batch PPT to PDF) v34.1.16
================================================================================
- 아키텍처: Clean Layer Architecture (Domain / Presentation / Application)
- 주요 기능: PPT/PPTX 일괄 PDF 변환, 북마크 자동 생성, 품질 최적화
- 가이드라인 준수: 00 PRD 가이드.md | AI_CODING_GUIDELINES_2026.md
- 무결성 보증: Force Process Kill (좀비 소거), UTF-8 Enforcement
- 업데이트: [v34.1.16] 관리자 권한 강제 승격 제거 (UAC 격리 방지 및 COM 안정화)
================================================================================
"""
import os
import glob
import dataclasses
import time
import subprocess
import datetime
import sys
import abc
import threading
import uuid
from typing import List, Tuple, Optional, Any
import psutil
import pythoncom
import win32com.client
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

# Enforce UTF-8 for process communication
if hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8')
    except: pass
else:
    import io
    try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except: pass

# ═══════════════════════════════════════════════════════════
# LAYER 1: DOMAIN (Entities & Value Objects)
# ═══════════════════════════════════════════════════════════

@dataclasses.dataclass
class ConversionResult:
    original_path: str
    pdf_path: str
    original_size: int
    pdf_size: int
    success: bool
    error_message: str = ""

    @property
    def savings_bytes(self) -> int:
        return max(0, self.original_size - self.pdf_size)

    @property
    def savings_percentage(self) -> float:
        if self.original_size <= 0: return 0.0
        return float(round((1 - (self.pdf_size / self.original_size)) * 100, 1))

@dataclasses.dataclass
class ConversionStats:
    total_files: int = 0
    successful_files: int = 0
    failed_files: int = 0
    total_original_bytes: int = 0
    total_pdf_bytes: int = 0
    
    def add_result(self, result: ConversionResult):
        self.total_files += 1
        if result.success:
            self.successful_files += 1
            self.total_original_bytes += result.original_size
            self.total_pdf_bytes += result.pdf_size
        else:
            self.failed_files += 1
    
    @property
    def total_savings_bytes(self) -> int:
        return max(0, self.total_original_bytes - self.total_pdf_bytes)

    @property
    def total_savings_percentage(self) -> float:
        if self.total_original_bytes <= 0: return 0.0
        return float(round((1 - (self.total_pdf_bytes / self.total_original_bytes)) * 100, 1))

    @property
    def original_mb(self) -> float:
        return float(round(self.total_original_bytes / (1024 * 1024), 2))
        
    @property
    def pdf_mb(self) -> float:
        return float(round(self.total_pdf_bytes / (1024 * 1024), 2))
        
    @property
    def savings_mb(self) -> float:
        return float(round(self.total_savings_bytes / (1024 * 1024), 2))

# ═══════════════════════════════════════════════════════════
# LAYER 2: APPLICATION (Interfaces & Use Cases)
# ═══════════════════════════════════════════════════════════

class IPowerPointService(abc.ABC):
    @abc.abstractmethod
    def open_session(self) -> None: pass
    
    @abc.abstractmethod
    def close_session(self) -> None: pass

    @abc.abstractmethod
    def convert_to_pdf(self, ppt_path: str, pdf_path: str, base_name: str) -> Tuple[bool, str]:
        return False, ""

class IFileService(abc.ABC):
    @abc.abstractmethod
    def get_ppt_files(self, folder_path: str) -> List[str]: return []
    @abc.abstractmethod
    def get_file_size(self, file_path: str) -> int: return 0
    @abc.abstractmethod
    def create_pdf_path(self, ppt_path: str) -> str: return ""
    @abc.abstractmethod
    def get_base_name(self, file_path: str) -> str: return ""
    @abc.abstractmethod
    def file_exists(self, file_path: str) -> bool: return False

class IDialogService(abc.ABC):
    @abc.abstractmethod
    def ask_for_folder(self) -> str: return ""

class ILogger(abc.ABC):
    @abc.abstractmethod
    def log(self, level_or_msg: str, message: Optional[str] = None) -> None: pass

class BatchConverterUseCase:
    def __init__(self, ppt_service: IPowerPointService, file_service: IFileService, logger: ILogger):
        self.ppt_service = ppt_service
        self.file_service = file_service
        self.logger = logger

    def execute(self, target_folder: str, progress_callback=None) -> ConversionStats:
        stats = ConversionStats()
        ppt_files = self.file_service.get_ppt_files(target_folder)
        
        if not ppt_files:
            self.logger.log("error", f"No PPT files found in: {target_folder}")
            return stats
            
        self.logger.log("============================================")
        self.logger.log(" PPT to PDF Batch Converter (Optimized v34.1.16)")
        self.logger.log("============================================")
        self.logger.log("info", f"Target: {target_folder}")
        self.logger.log("info", f"Files:  {len(ppt_files)} PPT files")
        self.logger.log("--------------------------------------------")

        try:
            self.ppt_service.open_session()
            for index, ppt_path in enumerate(ppt_files, start=1):
                base_name = self.file_service.get_base_name(ppt_path)
                pdf_path = self.file_service.create_pdf_path(ppt_path)
                abs_ppt_path = os.path.abspath(ppt_path)
                abs_pdf_path = os.path.abspath(pdf_path)
                
                if progress_callback:
                    progress_callback(index, len(ppt_files), base_name)
                    
                original_size = self.file_service.get_file_size(ppt_path)
                try:
                    success, err_msg = self.ppt_service.convert_to_pdf(abs_ppt_path, abs_pdf_path, base_name)
                    if success and self.file_service.file_exists(abs_pdf_path):
                        pdf_size = self.file_service.get_file_size(abs_pdf_path)
                        result = ConversionResult(ppt_path, pdf_path, original_size, pdf_size, True)
                        self.logger.log("info", f"[{index}/{len(ppt_files)}] {base_name} -> [OK]")
                        stats.add_result(result)
                    else:
                        self.logger.log("error", f"[{index}/{len(ppt_files)}] {base_name} -> [FAIL]: {err_msg}")
                        stats.add_result(ConversionResult(ppt_path, pdf_path, original_size, 0, False, err_msg))
                except Exception as e:
                    self.logger.log("error", f"[{index}/{len(ppt_files)}] {base_name} -> [ERROR]: {str(e)}")
                    stats.add_result(ConversionResult(ppt_path, pdf_path, original_size, 0, False, str(e)))
        finally:
            self.ppt_service.close_session()

        return stats

# ═══════════════════════════════════════════════════════════
# LAYER 3: INFRASTRUCTURE (Concrete Implementations)
# ═══════════════════════════════════════════════════════════

class PyWin32PowerPointService(IPowerPointService):
    def __init__(self, logger: ILogger):
        self.powerpoint: Any = None
        self.logger = logger

    def _kill_office_processes(self) -> None:
        targets = ["POWERPNT.EXE"]
        try:
            for p in psutil.process_iter(['name']):
                if p.info['name'] and p.info['name'].upper() in targets:
                    try: p.kill()
                    except: pass
            time.sleep(1.0)
        except: pass

    def open_session(self) -> None:
        self._kill_office_processes()
        pythoncom.CoInitialize()
        self.powerpoint = self._start_native_engine("PowerPoint.Application")

    def _start_native_engine(self, prog_id):
        app = None
        clsid = "{91493441-5A91-11CF-8700-00AA0060263B}"
        try:
            from win32com.client.dynamic import Dispatch as DynDispatch
            app = DynDispatch(clsid)
            if app: return self._finalize_app(app, "P0:CLSID")
        except: pass

        try:
            from win32com.client import gencache
            app = gencache.EnsureDispatch(prog_id)
            if app: return self._finalize_app(app, "P1:Ensure")
        except: pass

        try:
            app = win32com.client.Dispatch(prog_id)
            if app: return self._finalize_app(app, "P3:Std")
        except: pass

        raise Exception("[FAIL] PowerPoint 엔진을 기동할 수 없습니다.")

    def _finalize_app(self, app, mode_name):
        try:
            app.Visible = 0
            app.DisplayAlerts = 1
            try: app.Interactive = False
            except: pass
        except: pass
        self.logger.log("info", f"[OK] PowerPoint 엔진 기동 완료 ({mode_name})")
        return app

    def close_session(self) -> None:
        if self.powerpoint:
            try: self.powerpoint.Quit()
            except: pass
            self.powerpoint = None
        pythoncom.CoUninitialize()

    def convert_to_pdf(self, ppt_path: str, pdf_path: str, base_name: str) -> Tuple[bool, str]:
        presentation = None
        temp_pdf = f"{pdf_path}.{uuid.uuid4().hex[:8]}.tmp"
        try:
            for retry in range(2):
                try:
                    presentation = self.powerpoint.Presentations.Open(ppt_path, True, False, retry == 1)
                    break
                except:
                    if retry == 1: raise
                    time.sleep(1.0)
            
            if not presentation: return False, "Failed to open presentation"

            slide_index = 1
            for slide in presentation.Slides:
                bookmark_text = f"{base_name}-{slide_index}"
                if not slide.Shapes.HasTitle:
                    try: slide.Shapes.AddTitle()
                    except: pass
                try:
                    title_shape = slide.Shapes.Title
                    if title_shape:
                        title_shape.TextFrame.TextRange.Text = bookmark_text
                        title_shape.Top = -2000 
                except: pass
                slide_index += 1

            presentation.SaveAs(temp_pdf, 32)
            if os.path.exists(temp_pdf) and os.path.getsize(temp_pdf) > 0:
                if os.path.exists(pdf_path): os.remove(pdf_path)
                os.rename(temp_pdf, pdf_path)
                return True, ""
            return False, "SaveAs failed"
        except Exception as e:
            if os.path.exists(temp_pdf):
                try: os.remove(temp_pdf)
                except: pass
            return False, str(e)
        finally:
            if presentation:
                try: presentation.Close()
                except: pass

class LocalFileService(IFileService):
    def get_ppt_files(self, folder_path: str) -> List[str]:
        files = []
        for ext in ("*.ppt", "*.pptx", "*.pps", "*.ppsx", "*.pot", "*.potx"):
            files.extend(glob.glob(os.path.join(folder_path, ext)))
        return [f for f in files if not os.path.basename(f).startswith('~$')]

    def get_file_size(self, file_path: str) -> int:
        return os.path.getsize(file_path)

    def create_pdf_path(self, ppt_path: str) -> str:
        return os.path.splitext(ppt_path)[0] + ".pdf"
        
    def get_base_name(self, file_path: str) -> str:
        return os.path.splitext(os.path.basename(file_path))[0]
        
    def file_exists(self, file_path: str) -> bool:
        return os.path.exists(file_path)

class TkinterDialogService(IDialogService):
    def ask_for_folder(self) -> str:
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
        d = filedialog.askdirectory(title="Select Folder"); root.destroy()
        return d

class GUILogger(ILogger):
    def __init__(self, text_widget: tk.Text):
        self.text_widget = text_widget
    def log(self, level_or_msg: str, message: Optional[str] = None) -> None:
        msg = f"[{level_or_msg.upper()}] {message}" if message else level_or_msg
        self.text_widget.after(0, self._append_text, msg)
    def _append_text(self, message: str):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message + "\n")
        self.text_widget.see(tk.END)
        self.text_widget.config(state=tk.DISABLED)

class BatchConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("[PPT to PDF v34.1.16]")
        self.root.geometry("800x600")
        self.root.configure(padx=20, pady=20)
        
        ttk.Label(root, text="[STATUS]: PPT 일괄 고속 변환기", font=("Malgun Gothic", 12, "bold")).pack(anchor="w")
        self.lbl_status = ttk.Label(root, text="준비됨")
        self.lbl_status.pack(anchor="w", pady=5)
        
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(root, variable=self.progress_var, maximum=100).pack(fill=tk.X, pady=10)
        
        self.txt_log = tk.Text(root, height=15, state=tk.DISABLED, bg="#f5f5f5", font=("Consolas", 10))
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.btn_run = ttk.Button(root, text="변환 시작 (폴더 선택)", command=self.start_conversion)
        self.btn_run.pack(pady=10)
        
        self.file_service = LocalFileService()
        self.logger = GUILogger(self.txt_log)
        self.ppt_service = PyWin32PowerPointService(self.logger)
        self.dialog_service = TkinterDialogService()

    def update_progress(self, current: int, total: int, name: str):
        def _upd():
            self.progress_var.set((current/total)*100)
            self.lbl_status.config(text=f"진행 중 ({current}/{total}): {name}")
        self.root.after(0, _upd)

    def start_conversion(self):
        d = self.dialog_service.ask_for_folder()
        if not d: return
        self.btn_run.config(state=tk.DISABLED)
        def run():
            start = datetime.datetime.now()
            use_case = BatchConverterUseCase(self.ppt_service, self.file_service, self.logger)
            stats = use_case.execute(d, progress_callback=self.update_progress)
            dur = datetime.datetime.now() - start
            def final():
                self.lbl_status.config(text="완료")
                self.btn_run.config(state=tk.NORMAL)
                messagebox.showinfo("완료", f"성공: {stats.successful_files}/{stats.total_files}\n시간: {dur.seconds}초")
            self.root.after(0, final)
        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = BatchConverterApp(root)
    root.mainloop()
