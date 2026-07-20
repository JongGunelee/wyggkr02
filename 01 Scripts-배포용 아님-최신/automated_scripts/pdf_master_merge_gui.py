import fitz
import os
import math
import re
import copy
from io import BytesIO
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import tkinter.font as tkfont
from PIL import Image
import threading
import queue
import argparse
import sys

# UTF-8 stdout reconfiguration standard (v34.1.16 standard) for CP949 environment stability
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    else:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
except Exception:
    pass

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import List, Dict, Any, Callable, Optional, Tuple

# =====================================================================
# 1. DOMAIN LAYER (Entities & App State)
# =====================================================================

def get_backup_dir() -> str:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(script_dir, "..", ".temp_backups"))

def check_disk_space_warning() -> str:
    try:
        import shutil
        total, used, free = shutil.disk_usage('C:\\')
        free_mb = free / 1024 / 1024
        if free_mb < 200:
            return f"\n\n[디스크 공간 경고] 현재 C 드라이브의 여유 공간이 {free_mb:.1f} MB로 매우 부족합니다. MS Office 변환기는 작업 시 임시 파일 쓰기 공간이 필요하므로, C 드라이브의 디스크 여유 공간을 확보해 주세요."
    except Exception:
        pass
    return ""

def cleanup_temp_files():
    backup_dir = get_backup_dir()
    if os.path.exists(backup_dir):
        for f in os.listdir(backup_dir):
            if f.startswith("temp_img_") or f.startswith("temp_office_"):
                try:
                    os.remove(os.path.join(backup_dir, f))
                except Exception:
                    pass

@dataclass
class PDFItem:
    file_idx: int
    filepath: str
    filename: str
    page_count: int
    original_filepath: str = ""
    preprocessed_office_path: str = ""

@dataclass
class TOCItem:
    file_idx: int
    file_name: str
    level: int
    title: str
    dest_page: int  # 1-based page index in the individual PDF
    is_file_header: bool = False

@dataclass
class MergeConfig:
    compression_mode: str = "strong"   # "strong", "balanced", "basic"
    filename_mode: str = "toc_bookmark"  # "toc_bookmark", "bookmark_only", "separate_page", "none"
    toc_paper_size: str = "A4"           # "A4", "A3", "Letter", "Legal"
    toc_orientation: str = "portrait"    # "portrait", "landscape"
    save_preprocessed: bool = False

class AppState:
    def __init__(self):
        self.pdf_items: List[PDFItem] = []
        self.toc_data: List[TOCItem] = []
        self.page_offsets: List[int] = []  # Cumulative page count offsets
        self.page_rotations: Dict[Tuple[int, int], int] = {}  # Store page rotations as (file_idx, page_idx) -> degrees
        self.manually_rotated: List[Tuple[int, int]] = []  # Track pages manually rotated by user
        self.undo_stack: List[Tuple[List[TOCItem], Dict[Tuple[int, int], int], List[Tuple[int, int]], List[str]]] = []  # Copies of (toc_data, page_rotations, manually_rotated, selected_iids)
        self.listeners: List[Callable[[], None]] = []
        self.level_prefixes: Dict[int, str] = {
            1: "■ ",
            2: "● ",
            3: "▲ ",
            4: "▼ ",
            5: "◆ ",
            6: "○ ",
            7: "□ ",
            8: "△ ",
            9: "▽ ",
            10: "◇ "
        }
        self.load_level_prefixes()
        
        # Clean up any leftover temp files from previous runs
        cleanup_temp_files()

    def get_config_path(self) -> str:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, "pdf_master_prefixes_config.json")

    def load_level_prefixes(self):
        config_path = self.get_config_path()
        if os.path.exists(config_path):
            try:
                import json
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    for k, v in data.items():
                        try:
                            lvl = int(k)
                            if 1 <= lvl <= 10:
                                self.level_prefixes[lvl] = str(v)
                        except ValueError:
                            pass
            except Exception as e:
                print(f"접두기호 설정 로드 실패: {e}")

    def save_level_prefixes(self):
        config_path = self.get_config_path()
        try:
            import json
            # Convert keys to strings for JSON serialization
            data = {str(k): v for k, v in self.level_prefixes.items()}
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"접두기호 설정 저장 실패: {e}")

    def register_listener(self, listener: Callable[[], None]):
        self.listeners.append(listener)

    def notify_listeners(self):
        for listener in self.listeners:
            listener()

    def set_pdfs(self, filepaths: List[str], save_preprocessed: bool = False, compression_mode: str = "balanced"):
        loader = PDFLoadUseCase(self, filepaths, save_preprocessed, compression_mode)
        loader._execute_sync()

    def push_undo(self, current_selection_iids: List[str]):
        self.undo_stack.append((
            copy.deepcopy(self.pdf_items),
            copy.deepcopy(self.toc_data),
            copy.deepcopy(self.page_offsets),
            copy.deepcopy(self.page_rotations),
            list(self.manually_rotated),
            list(current_selection_iids)
        ))
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)

    def pop_undo(self) -> Optional[Tuple[List[PDFItem], List[TOCItem], List[int], Dict[Tuple[int, int], int], List[Tuple[int, int]], List[str]]]:
        if not self.undo_stack:
            return None
        pdf_items, toc, offsets, rotations, manual, sel = self.undo_stack.pop()
        self.pdf_items = pdf_items
        self.toc_data = toc
        self.page_offsets = offsets
        self.page_rotations = rotations
        self.manually_rotated = manual
        self.notify_listeners()
        return pdf_items, toc, offsets, rotations, manual, sel

    def delete_items(self, indices: List[int]):
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self.toc_data):
                self.toc_data.pop(idx)
        self.notify_listeners()

    def strip_prefix(self, title: str) -> str:
        if not title:
            return ""
        pattern = r'^[\s■□●○▲▼◆◇★☆▶▷\-*•▪▫♦♥♣♠]+\s*'
        import re
        return re.sub(pattern, '', title)

    def update_item_level(self, idx: int, new_level: int):
        if 0 <= idx < len(self.toc_data):
            item = self.toc_data[idx]
            item.level = new_level
            if new_level == 0:
                prefix = ""
            else:
                prefix = self.level_prefixes.get(new_level, self.level_prefixes.get(10, ""))
            clean_title = self.strip_prefix(item.title)
            item.title = f"{prefix}{clean_title}".strip()
            self.notify_listeners()

    def apply_level_prefixes(self):
        for item in self.toc_data:
            if item.level == 0:
                prefix = ""
            else:
                prefix = self.level_prefixes.get(item.level, self.level_prefixes.get(10, ""))
            clean_title = self.strip_prefix(item.title)
            item.title = f"{prefix}{clean_title}".strip()
        self.notify_listeners()

    def apply_level_prefixes_to_indices(self, indices: List[int]):
        for idx in indices:
            if 0 <= idx < len(self.toc_data):
                item = self.toc_data[idx]
                if item.level == 0:
                    prefix = ""
                else:
                    prefix = self.level_prefixes.get(item.level, self.level_prefixes.get(10, ""))
                clean_title = self.strip_prefix(item.title)
                item.title = f"{prefix}{clean_title}".strip()
        self.notify_listeners()

    def update_item_title(self, idx: int, new_title: str):
        if 0 <= idx < len(self.toc_data):
            self.toc_data[idx].title = new_title
            self.notify_listeners()

    def move_item_up(self, idx: int) -> bool:
        if 0 < idx < len(self.toc_data):
            self.toc_data[idx], self.toc_data[idx-1] = self.toc_data[idx-1], self.toc_data[idx]
            self.notify_listeners()
            return True
        return False

    def move_item_down(self, idx: int) -> bool:
        if 0 <= idx < len(self.toc_data) - 1:
            self.toc_data[idx], self.toc_data[idx+1] = self.toc_data[idx+1], self.toc_data[idx]
            self.notify_listeners()
            return True
        return False

    def batch_edit_titles(
        self,
        search: str,
        value: str,
        action: str,
        scope: str,
        case_sensitive: bool,
        mode: str,
        position: str,
        length: str,
        selected_iids: List[str]
    ) -> Optional[int]:
        search = search.strip()

        if mode == "search" and not search:
            return None

        if mode != "search":
            try:
                pos = int(position)
                lg = int(length)
            except ValueError:
                return None
            if pos < 1 or lg < 1:
                return None
        else:
            pos = 1
            lg = 1

        if action != "delete" and value == "":
            return None

        if scope == "selected":
            if not selected_iids:
                return None
            indices = [int(iid) for iid in selected_iids]
        else:
            indices = list(range(len(self.toc_data)))

        changed = 0
        self.push_undo(selected_iids)
        
        for idx in indices:
            if idx < 0 or idx >= len(self.toc_data):
                continue
            old_title = self.toc_data[idx].title
            if mode == "search":
                new_title = self._edit_title_text(old_title, search, value, action, case_sensitive).strip()
            else:
                new_title = self._edit_title_by_position(old_title, value, action, mode, pos, lg).strip()
            if new_title and new_title != old_title:
                self.toc_data[idx].title = new_title
                changed += 1

        if changed == 0:
            self.undo_stack.pop()  # Discard empty state from history
            return 0

        self.notify_listeners()
        return changed

    def _edit_title_text(self, title: str, search: str, value: str, action: str, case_sensitive: bool) -> str:
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(re.escape(search), flags)
        if not pattern.search(title):
            return title

        if action == "replace":
            return pattern.sub(value, title)
        if action == "delete":
            return pattern.sub("", title)
        if action == "insert_before":
            return pattern.sub(lambda match: value + match.group(0), title)
        if action == "insert_after":
            return pattern.sub(lambda match: match.group(0) + value, title)
        return title

    def _edit_title_by_position(self, title: str, value: str, action: str, mode: str, position: int, length: int) -> str:
        if not title:
            return title

        if action in ("insert_at", "insert_before", "insert_after"):
            if mode == "from_start":
                index = min(max(position - 1, 0), len(title))
            else:
                index = max(len(title) - position + 1, 0)
            return title[:index] + value + title[index:]

        if len(title) <= 2:
            return title

        if mode == "from_start":
            index = min(max(position, 1), len(title) - 1)
        else:
            index = max(len(title) - position - length, 1)

        if action == "delete":
            end = min(index + length, len(title) - 1)
            return title[:index] + title[end:]
        if action == "replace":
            end = min(index + length, len(title) - 1)
            return title[:index] + value + title[end:]
        return title


# =====================================================================
# 2. INFRASTRUCTURE SERVICES (PDF & Image Engines)
# =====================================================================

def safe_path(path):
    r"""
    Windows MAX_PATH (260 chars) limitation handling.
    Normalizes path and applies \\?\ prefix if path is long or a UNC path.
    """
    path = os.path.abspath(os.path.normpath(path))
    if len(path) > 240 or path.startswith('\\\\'):
        if not path.startswith('\\\\?\\'):
            if path.startswith('\\\\'):
                return '\\\\?\\UNC\\' + path[2:]
            return '\\\\?\\' + path
    return path

def get_com_path(safe_path):
    r"""
    COM Safe Path Protocol.
    Office COM cannot resolve \\?\ prefix on paths < 260 chars.
    Removes it for short paths.
    """
    if len(safe_path) < 260 and safe_path.startswith('\\\\?\\'):
        return safe_path[4:]
    return safe_path

def get_unique_output_path(base_path: str) -> str:
    if not os.path.exists(base_path):
        return base_path
    dir_name, file_name = os.path.split(base_path)
    name, ext = os.path.splitext(file_name)
    counter = 1
    candidate = os.path.join(dir_name, f"{name}_preprocessed{ext}")
    while os.path.exists(candidate):
        candidate = os.path.join(dir_name, f"{name}_preprocessed ({counter}){ext}")
        counter += 1
    return candidate

def _compress_media_bytes(data: bytes, filename: str, quality: int, max_side: int) -> Optional[bytes]:
    try:
        from PIL import Image
        from io import BytesIO
        with Image.open(BytesIO(data)) as img:
            if img.width < 150 and img.height < 150:
                return None
                
            fmt = img.format
            if not fmt:
                if filename.endswith(".png"): fmt = "PNG"
                elif filename.endswith((".jpg", ".jpeg")): fmt = "JPEG"
                elif filename.endswith(".bmp"): fmt = "BMP"
                else: return None
                
            is_screenshot_or_diagram = False
            if fmt in ('PNG', 'GIF') or img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                is_screenshot_or_diagram = True
            else:
                try:
                    thumb = img.resize((100, 100))
                    colors = thumb.getcolors(maxcolors=256)
                    if colors is not None:
                        is_screenshot_or_diagram = True
                except Exception:
                    pass

            orig_w, orig_h = img.size
            work = img
            if max(orig_w, orig_h) > max_side:
                scale = max_side / max(orig_w, orig_h)
                new_size = (max(1, int(orig_w * scale)), max(1, int(orig_h * scale)))
                work = img.resize(new_size, Image.Resampling.LANCZOS)

            output = BytesIO()
            if is_screenshot_or_diagram:
                # For diagrams/screenshots, preserve quality but quantize PNG if RGB/RGBA to save space
                if work.mode in ("RGB", "RGBA"):
                    try:
                        palette_img = work.quantize(colors=256, dither=Image.Dither.NONE)
                        test_out = BytesIO()
                        palette_img.save(test_out, format="PNG", optimize=True)
                        if len(test_out.getvalue()) < len(data):
                            work = palette_img
                    except Exception:
                        pass
                work.save(output, format="PNG", optimize=True)
            else:
                if work.mode != "RGB":
                    work = work.convert("RGB")
                work.save(output, format="JPEG", quality=quality, optimize=True)
                
            compressed = output.getvalue()
            if len(compressed) < len(data) * 0.98:
                return compressed
    except Exception:
        pass
    return None

def compress_openxml_media(file_path: str, compression_mode: str = "balanced") -> bool:
    import zipfile
    import shutil
    import os
    
    if compression_mode == "basic":
        return False
        
    if not zipfile.is_zipfile(file_path):
        return False
        
    if compression_mode == "strong":
        quality = 70
        max_side = 800
    else:  # balanced
        quality = 78
        max_side = 1200
        
    temp_zip_path = file_path + ".tmp_zip"
    try:
        modified = False
        with zipfile.ZipFile(file_path, 'r') as zin:
            with zipfile.ZipFile(temp_zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    is_media_image = False
                    lower_name = item.filename.lower()
                    if "/media/" in lower_name and any(lower_name.endswith(ext) for ext in [".png", ".jpg", ".jpeg", ".bmp"]):
                        is_media_image = True
                        
                    if is_media_image:
                        compressed_data = _compress_media_bytes(data, lower_name, quality, max_side)
                        if compressed_data and len(compressed_data) < len(data):
                            zout.writestr(item, compressed_data)
                            modified = True
                            continue
                    
                    zout.writestr(item, data)
                    
        if modified:
            shutil.move(temp_zip_path, file_path)
            return True
        else:
            if os.path.exists(temp_zip_path):
                os.remove(temp_zip_path)
            return False
    except Exception as e:
        print(f"OpenXML 미디어 압축 중 오류: {e}")
        if os.path.exists(temp_zip_path):
            os.remove(temp_zip_path)
        return False

class COMOfficeManager:
    def __init__(self):
        self.word = None
        self.excel = None
        self.ppt = None
        
    def _kill_office_processes(self):
        """
        WinError 32 prevention: force clean all office processes before starting work
        and wait 1.0 second. Done silently without flashing CMD shell window.
        """
        import subprocess
        import time
        for process in ["winword.exe", "excel.exe", "powerpnt.exe"]:
            try:
                subprocess.run(
                    ["taskkill", "/f", "/im", process],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=0x08000000  # CREATE_NO_WINDOW
                )
            except Exception:
                pass
        time.sleep(1.0)

    def _robust_dispatch(self, prog_id, clsid):
        """
        3-Stage Recovery and 5-Fold COM Binding Filter:
        DynDispatch -> DispatchEx -> Standard Dispatch -> GetActiveObject -> GetObject
        """
        import win32com.client
        from win32com.client import Dispatch, DispatchEx
        from win32com.client.dynamic import Dispatch as DynDispatch
        
        # Phase 0: CLSID Direct Dispatch (가장 강력한 직접 바인딩)
        try:
            app = DynDispatch(clsid)
            if app:
                return app
        except Exception:
            pass

        # Phase 1: DispatchEx (독립된 신규 프로세스로 기동)
        try:
            app = DispatchEx(prog_id)
            if app:
                return app
        except Exception:
            pass

        # Phase 2: Standard Dispatch
        try:
            app = Dispatch(prog_id)
            if app:
                return app
        except Exception:
            pass

        # Phase 3: GetActiveObject (이미 실행 중인 인스턴스에 바인딩)
        try:
            app = win32com.client.GetActiveObject(prog_id)
            if app:
                return app
        except Exception:
            pass

        # Phase 4: GetObject (모니커 바인딩을 통한 우회)
        try:
            app = win32com.client.GetObject(None, prog_id)
            if app:
                return app
        except Exception:
            pass

        raise RuntimeError(f"{prog_id} COM 엔진을 가동할 수 없습니다. MS Office 설치와 UAC 권한 상태를 확인해주세요.")

    def get_word(self):
        import pythoncom
        pythoncom.CoInitialize()
        if not self.word:
            self._kill_office_processes()
            app = self._robust_dispatch("Word.Application", "{000209FF-0000-0000-C000-000000000046}")
            try:
                app.Visible = False
                app.DisplayAlerts = 0  # wdAlertsNone
            except Exception:
                pass
            self.word = app
        return self.word

    def get_excel(self):
        import pythoncom
        pythoncom.CoInitialize()
        if not self.excel:
            self._kill_office_processes()
            app = self._robust_dispatch("Excel.Application", "{00024500-0000-0000-C000-000000000046}")
            try:
                app.Visible = False
                app.DisplayAlerts = False
                app.AskToUpdateLinks = False
                app.Interactive = False
                app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            self.excel = app
        return self.excel

    def get_ppt(self):
        import pythoncom
        pythoncom.CoInitialize()
        if not self.ppt:
            self._kill_office_processes()
            app = self._robust_dispatch("Powerpoint.Application", "{91493441-5A91-11CF-8700-0020AFF12257}")
            try:
                app.DisplayAlerts = 1  # ppAlertsNone
            except Exception:
                pass
            self.ppt = app
        return self.ppt

    def close_all(self):
        if self.word:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None
        if self.excel:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None
        if self.ppt:
            try:
                self.ppt.Quit()
            except Exception:
                pass
            self.ppt = None
        import pythoncom
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

class FontService:
    @staticmethod
    def get_korean_font_path() -> Optional[str]:
        # Cross-platform common Korean font locations
        paths = [
            r"C:\Windows\Fonts\malgun.ttf",
            r"C:\Windows\Fonts\malgunbd.ttf",
            r"C:\Windows\Fonts\batang.ttc",
            r"C:\Windows\Fonts\gulim.ttc",
            "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
            "/Library/Fonts/NanumGothic.ttf",
            "/Library/Fonts/NanumBarunGothic.ttf",
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/usr/share/fonts/nanum/NanumGothic.ttf",
            "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
        ]
        for p in paths:
            if os.path.exists(p):
                return p
        return None

class ImageCompressionService:
    @staticmethod
    def _compress_single_image(original_bytes: bytes, settings: Dict[str, Any]) -> Optional[bytes]:
        try:
            with Image.open(BytesIO(original_bytes)) as img:
                # Skip very small icons/decorations to save computation
                if img.width < 150 and img.height < 150:
                    return None

                # Screenshot / diagram / text readability safeguard heuristic:
                is_screenshot_or_diagram = False
                if img.format in ('PNG', 'GIF') or img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                    is_screenshot_or_diagram = True
                else:
                    try:
                        thumb = img.resize((100, 100))
                        colors = thumb.getcolors(maxcolors=256)
                        if colors is not None:
                            is_screenshot_or_diagram = True
                    except Exception:
                        pass

                if is_screenshot_or_diagram:
                    # Lossless PNG compression to preserve readability of text and thin lines
                    work = img
                    max_side = max(work.size)
                    if max_side > settings["max_side"]:
                        scale = settings["max_side"] / max_side
                        new_size = (max(1, int(work.width * scale)), max(1, int(work.height * scale)))
                        work = work.resize(new_size, Image.Resampling.LANCZOS)
                    
                    output = BytesIO()
                    work.save(output, format="PNG", optimize=True)
                    compressed = output.getvalue()
                    if len(compressed) < len(original_bytes) * 0.95:
                        return compressed
                else:
                    # Photographic image: Use JPEG lossy compression
                    work = img
                    if work.mode in ("RGBA", "LA") or (work.mode == "P" and "transparency" in work.info):
                        background = Image.new("RGB", work.size, (255, 255, 255))
                        alpha = work.convert("RGBA").split()[-1]
                        background.paste(work.convert("RGB"), mask=alpha)
                        work = background
                    else:
                        work = work.convert("RGB")

                    max_side = max(work.size)
                    if max_side > settings["max_side"]:
                        scale = settings["max_side"] / max_side
                        new_size = (max(1, int(work.width * scale)), max(1, int(work.height * scale)))
                        work = work.resize(new_size, Image.Resampling.LANCZOS)

                    output = BytesIO()
                    work.save(
                        output,
                        format="JPEG",
                        quality=settings["quality"],
                        optimize=True,
                        progressive=True,
                    )
                    compressed = output.getvalue()
                    if len(compressed) < len(original_bytes) * 0.92:
                        return compressed
        except Exception:
            pass
        return None

    @classmethod
    def compress_pdf_images(cls, doc: fitz.Document, mode: str, progress_callback: Callable[[str, Optional[int]], None]) -> Dict[str, Any]:
        if mode == "basic":
            return {"checked": 0, "replaced": 0, "saved_bytes": 0}

        settings = {
            "balanced": {"quality": 78, "max_side": 2200, "min_bytes": 35 * 1024},
            "strong": {"quality": 62, "max_side": 1600, "min_bytes": 20 * 1024},
        }[mode]

        stats = {"checked": 0, "replaced": 0, "saved_bytes": 0}
        seen_xrefs = {}
        
        # 1. Scan pages and extract unique image XREFs and map to page indices
        for page_idx in range(len(doc)):
            for img_info in doc[page_idx].get_images(full=True):
                xref = img_info[0]
                if xref not in seen_xrefs:
                    seen_xrefs[xref] = page_idx
                    
        total_images = len(seen_xrefs)
        if total_images == 0:
            return stats

        progress_callback(f"이미지 추출 완료: 총 {total_images}개 발견. 압축 작업 시작...", 60)
        
        # 2. Extract original bytes for processing tasks
        tasks = []
        for xref, page_idx in seen_xrefs.items():
            try:
                extracted = doc.extract_image(xref)
                original_bytes = extracted.get("image")
                if original_bytes and len(original_bytes) >= settings["min_bytes"]:
                    tasks.append((xref, page_idx, original_bytes))
            except Exception:
                continue

        stats["checked"] = len(tasks)
        if not tasks:
            return stats

        # 3. Compress images in parallel using ThreadPoolExecutor
        compressed_results = {}
        progress_callback("이미지 압축 중 (병렬 처리)...", 60)
        
        max_workers = min(os.cpu_count() or 4, 8)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_xref = {
                executor.submit(cls._compress_single_image, task[2], settings): (task[0], task[1], len(task[2]))
                for task in tasks
            }
            
            completed_count = 0
            for future in future_to_xref:
                xref, page_idx, original_len = future_to_xref[future]
                try:
                    compressed_bytes = future.result()
                    if compressed_bytes:
                        compressed_results[xref] = (page_idx, compressed_bytes, original_len)
                except Exception:
                    pass
                completed_count += 1
                if completed_count % 5 == 0 or completed_count == len(future_to_xref):
                    pct = 60 + int((completed_count / len(future_to_xref)) * 30)
                    progress_callback(f"이미지 압축 진행 중... {completed_count}/{len(future_to_xref)}", pct)

        # 4. Sequentially replace compressed image objects in the PDF
        replaced_count = 0
        for xref, (page_idx, compressed_bytes, original_len) in compressed_results.items():
            try:
                doc[page_idx].replace_image(xref, stream=compressed_bytes)
                replaced_count += 1
                stats["saved_bytes"] += (original_len - len(compressed_bytes))
            except Exception:
                pass
                
        stats["replaced"] = replaced_count
        return stats

class PDFService:
    @staticmethod
    def get_toc_layout_settings(config: MergeConfig) -> Dict[str, Any]:
        paper_sizes = {
            "A4": (595.0, 842.0),
            "A3": (842.0, 1191.0),
            "Letter": (612.0, 792.0),
            "Legal": (612.0, 1008.0),
        }
        base_width, base_height = paper_sizes.get(config.toc_paper_size, paper_sizes["A4"])
        if config.toc_orientation == "landscape":
            page_w, page_h = max(base_width, base_height), min(base_width, base_height)
            orientation_label = "가로"
        else:
            page_w, page_h = min(base_width, base_height), max(base_width, base_height)
            orientation_label = "세로"

        scale_factor = max(0.86, min(1.18, page_w / 595.0))
        if config.toc_orientation == "landscape":
            margin_top = 62 * scale_factor
            margin_bottom = 38 * scale_factor
            margin_left = 42 * scale_factor
            line_height = 24 * scale_factor
            title_font_size = 16 * scale_factor
            category_font_size = 12 * scale_factor
            item_font_size = 10 * scale_factor
        else:
            margin_top = 80 * scale_factor
            margin_bottom = 50 * scale_factor
            margin_left = 50 * scale_factor
            line_height = 32 * scale_factor
            title_font_size = 18 * scale_factor
            category_font_size = 13 * scale_factor
            item_font_size = 11 * scale_factor

        return {
            "paper_name": config.toc_paper_size,
            "orientation": config.toc_orientation,
            "orientation_label": orientation_label,
            "page_width": page_w,
            "page_height": page_h,
            "margin_top": margin_top,
            "margin_bottom": margin_bottom,
            "margin_left": margin_left,
            "line_height": line_height,
            "title_font_size": title_font_size,
            "category_font_size": category_font_size,
            "item_font_size": item_font_size,
            "move_button_width": 58 * scale_factor,
            "move_button_gap": 10 * scale_factor,
        }

    @staticmethod
    def fit_text_to_width(text: str, available_width: float, fontsize: float) -> str:
        if not text:
            return ""
        avg_char_width = max(fontsize * 0.88, 1)
        max_chars = max(8, int(available_width / avg_char_width))
        if len(text) <= max_chars:
            return text
        return text[:max(1, max_chars - 3)] + "..."


# =====================================================================
# 3. APPLICATION LAYER (Use Cases & Threading)
# =====================================================================

class PDFLoadUseCase:
    def __init__(self, state: AppState, filepaths: List[str], save_preprocessed: bool = False, compression_mode: str = "balanced", custom_filenames: Dict[str, str] = None):
        self.state = state
        self.filepaths = filepaths
        self.save_preprocessed = save_preprocessed
        self.compression_mode = compression_mode
        self.custom_filenames = custom_filenames or {}
        self.progress_queue = queue.Queue()
        self.response_queue = queue.Queue()
        self.bookmark_choice_override = None

    def run_in_background(self):
        thread = threading.Thread(target=self._execute, daemon=True)
        thread.start()

    def _preprocess_office_file(self, filepath: str, ext: str, manager, backup_dir: str, h: str) -> Tuple[str, bool]:
        """
        구버전 오피스 문서 (.doc, .xls, .ppt)를 최신 포맷 XML 문서 (.docx, .xlsx, .pptx)로
        변환하여 최신 포맷 그래픽 압축 기술을 적용합니다.
        """
        abs_in = get_com_path(safe_path(filepath))

        if ext == ".doc":
            temp_docx = os.path.join(backup_dir, f"temp_conv_{h}.docx")
            abs_out = get_com_path(safe_path(temp_docx))
            word = manager.get_word()
            doc_obj = word.Documents.Open(abs_in, False, True)
            doc_obj.SaveAs(abs_out, 16)  # wdFormatXMLDocument = 16
            doc_obj.Close(0)
            compress_openxml_media(temp_docx, self.compression_mode)
            return temp_docx, True
            
        elif ext == ".xls":
            temp_xlsx = os.path.join(backup_dir, f"temp_conv_{h}.xlsx")
            abs_out = get_com_path(safe_path(temp_xlsx))
            excel = manager.get_excel()
            wb = excel.Workbooks.Open(abs_in, 0, True)
            wb.SaveAs(abs_out, 51)  # xlOpenXMLWorkbook = 51
            wb.Close(False)
            compress_openxml_media(temp_xlsx, self.compression_mode)
            return temp_xlsx, True
            
        elif ext == ".ppt":
            temp_pptx = os.path.join(backup_dir, f"temp_conv_{h}.pptx")
            abs_out = get_com_path(safe_path(temp_pptx))
            ppt = manager.get_ppt()
            deck = ppt.Presentations.Open(abs_in, True, False, False)
            deck.SaveAs(abs_out, 24)  # ppSaveAsOpenXMLPresentation = 24
            deck.Close()
            compress_openxml_media(temp_pptx, self.compression_mode)
            return temp_pptx, True
            
        return filepath, False

    def _optimize_temp_pdf(self, pdf_path: str):
        """
        임시 변환된 PDF 파일 내부의 이미지와 객체를 로드 시점에 즉시 압축하고
        가비지 컬렉션을 수행하여 불필요한 GDI 비트맵 스트림 찌꺼기를 해제합니다.
        """
        try:
            doc = fitz.open(pdf_path)
            # Use the selected compression mode instead of hardcoded balanced mode
            ImageCompressionService.compress_pdf_images(
                doc, self.compression_mode, lambda msg, pct: None
            )
            temp_save = pdf_path + ".tmp"
            doc.save(
                temp_save,
                garbage=4,
                deflate=True,
                deflate_images=True,
                deflate_fonts=True,
                clean=True,
                use_objstms=1,
                compression_effort=100,
            )
            doc.close()
            if os.path.exists(temp_save) and os.path.getsize(temp_save) > 0:
                os.replace(temp_save, pdf_path)
            else:
                if os.path.exists(temp_save):
                    os.remove(temp_save)
        except Exception:
            pass

    def _execute(self):
        try:
            self.progress_queue.put({"type": "start", "message": "파일 분석 및 로딩을 준비 중입니다...", "percent": 0})
            
            # Ensure temp backups directory exists
            backup_dir = get_backup_dir()
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir, exist_ok=True)
                
            manager = COMOfficeManager()
            
            pdf_items = []
            toc_data = []
            page_offsets = []
            page_rotations = {}
            
            cumulative_pages = 0
            total_files = len(self.filepaths)
            
            for i, filepath in enumerate(self.filepaths):
                filename = self.custom_filenames.get(filepath, os.path.basename(filepath))
                ext = os.path.splitext(filename)[1].lower()
                
                pct = int((i / total_files) * 90)
                self.progress_queue.put({"type": "progress", "message": f"[{i+1}/{total_files}] {filename} 처리 중...", "percent": pct})
                
                temp_pdf_path = ""
                original_filepath = filepath
                temp_office_path = ""
                is_temp_office = False
                
                if ext == ".pdf":
                    target_pdf_path = filepath
                elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff"):
                    import hashlib
                    h = hashlib.md5(filepath.encode('utf-8')).hexdigest()[:8]
                    temp_pdf_path = os.path.join(backup_dir, f"temp_img_{h}.pdf")
                    
                    img_doc = fitz.open(filepath)
                    pdf_bytes = img_doc.convert_to_pdf()
                    img_doc.close()
                    
                    with open(temp_pdf_path, "wb") as f:
                        f.write(pdf_bytes)
                    self._optimize_temp_pdf(temp_pdf_path)
                    target_pdf_path = temp_pdf_path
                elif ext in (".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"):
                    import hashlib
                    h = hashlib.md5(filepath.encode('utf-8')).hexdigest()[:8]
                    temp_pdf_path = os.path.join(backup_dir, f"temp_office_{h}.pdf")
                    
                    # Preprocess legacy formats to modern formats for image/graphics compression
                    temp_office_path = filepath
                    is_temp_office = False
                    if ext in (".doc", ".xls", ".ppt"):
                        try:
                            temp_office_path, is_temp_office = self._preprocess_office_file(filepath, ext, manager, backup_dir, h)
                            conv_ext = os.path.splitext(temp_office_path)[1].lower()
                        except Exception:
                            # Graceful fallback to original path if pre-conversion fails
                            temp_office_path = filepath
                            is_temp_office = False
                            conv_ext = ext
                    elif ext in (".docx", ".xlsx", ".pptx"):
                        try:
                            import shutil
                            temp_office_path = os.path.join(backup_dir, f"temp_conv_{h}{ext}")
                            shutil.copy2(filepath, temp_office_path)
                            is_temp_office = True
                            compress_openxml_media(temp_office_path, self.compression_mode)
                            conv_ext = ext
                        except Exception:
                            temp_office_path = filepath
                            is_temp_office = False
                            conv_ext = ext
                    else:
                        conv_ext = ext
                        
                    abs_in = get_com_path(safe_path(temp_office_path))
                    abs_out = get_com_path(safe_path(temp_pdf_path))
                    
                    if conv_ext in (".doc", ".docx"):
                        try:
                            word = manager.get_word()
                            doc = word.Documents.Open(abs_in, False, True)  # FileName, ConfirmConversions=False, ReadOnly=True
                            doc.SaveAs(abs_out, 17) # FileName, FileFormat=17
                            doc.Close(0) # SaveChanges=0 (wdDoNotSaveChanges)
                        except Exception as e:
                            disk_warn = check_disk_space_warning()
                            raise RuntimeError(f"Word 파일 변환 오류 ({filename}):\n{str(e)}{disk_warn}\n\nMS Office 설치 상태와 라이선스를 확인해주세요.")
                    elif conv_ext in (".xls", ".xlsx"):
                        try:
                            excel = manager.get_excel()
                            wb = excel.Workbooks.Open(abs_in, 0, True)  # Filename, UpdateLinks=0, ReadOnly=True
                            
                            # Set print scaling options for each worksheet so columns fit page width
                            try:
                                for ws in wb.Worksheets:
                                    ws.PageSetup.Zoom = False
                                    ws.PageSetup.FitToPagesWide = 1
                                    ws.PageSetup.FitToPagesTall = 0
                            except Exception:
                                pass
                                
                            wb.ExportAsFixedFormat(0, abs_out) # Type=0 (xlTypePDF), Filename
                            wb.Close(False)
                        except Exception as e:
                            disk_warn = check_disk_space_warning()
                            raise RuntimeError(f"Excel 파일 변환 오류 ({filename}):\n{str(e)}{disk_warn}\n\nMS Office 설치 상태와 라이선스를 확인해주세요.")
                    elif conv_ext in (".ppt", ".pptx"):
                        try:
                            ppt = manager.get_ppt()
                            deck = ppt.Presentations.Open(abs_in, True, False, False)  # FileName, ReadOnly=True, Untitled=False, WithWindow=False
                            deck.SaveAs(abs_out, 32) # FileName, FileFormat=32
                            deck.Close()
                        except Exception as e:
                            disk_warn = check_disk_space_warning()
                            raise RuntimeError(f"PowerPoint 파일 변환 오류 ({filename}):\n{str(e)}{disk_warn}\n\nMS Office 설치 상태와 라이선스를 확인해주세요.")
                    self._optimize_temp_pdf(temp_pdf_path)
                    target_pdf_path = temp_pdf_path
                else:
                    raise ValueError(f"지원하지 않는 파일 형식입니다: {ext}")
                
                page_offsets.append(cumulative_pages)
                doc = fitz.open(target_pdf_path)
                page_count = len(doc)
                
                preprocessed_office_path = temp_office_path if (is_temp_office and os.path.exists(temp_office_path)) else ""
                
                pdf_items.append(PDFItem(
                    file_idx=i,
                    filepath=target_pdf_path,
                    filename=filename,
                    page_count=page_count,
                    original_filepath=original_filepath,
                    preprocessed_office_path=preprocessed_office_path
                ))
                
                for p_idx in range(page_count):
                    page = doc[p_idx]
                    page_rotations[(i, p_idx)] = page.rotation
                
                base_page = 1 if page_count > 0 else 0
                cat_name = filename
                for drop_ext in (".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"):
                    if cat_name.lower().endswith(drop_ext):
                        cat_name = cat_name[:-len(drop_ext)]
                        break
                        
                prefix_lvl1 = self.state.level_prefixes.get(1, "")
                clean_cat = self.state.strip_prefix(cat_name)
                toc_data.append(TOCItem(
                    file_idx=i,
                    file_name=filename,
                    level=1,
                    title=f"{prefix_lvl1}{clean_cat}".strip(),
                    dest_page=base_page,
                    is_file_header=True
                ))
                
                has_bookmarks = False
                toc = []
                if ext == ".pdf":
                    try:
                        toc = doc.get_toc()
                        if toc:
                            has_bookmarks = True
                    except Exception:
                        pass
                
                choice = "bookmark"
                if has_bookmarks:
                    if self.bookmark_choice_override is not None:
                        choice = self.bookmark_choice_override
                    else:
                        self.progress_queue.put({"type": "ask_bookmarks", "filename": filename})
                        try:
                            choice = self.response_queue.get()  # Block until user responds
                            if choice in ("filename_all", "bookmark_all"):
                                override_val = choice.replace("_all", "")
                                self.bookmark_choice_override = override_val
                                choice = override_val
                        except Exception:
                            choice = "bookmark"
                
                if choice == "bookmark" and ext == ".pdf":
                    try:
                        for item in toc:
                            level, title, page = item
                            if "슬라이드" in title:
                                continue
                            lvl = level + 1
                            lvl = max(0, min(10, lvl))
                            prefix = "" if lvl == 0 else self.state.level_prefixes.get(lvl, self.state.level_prefixes.get(10, ""))
                            clean_title = self.state.strip_prefix(title)
                            toc_data.append(TOCItem(
                                file_idx=i,
                                file_name=filename,
                                level=lvl,
                                title=f"{prefix}{clean_title}".strip(),
                                dest_page=page,
                                is_file_header=False
                            ))
                    except Exception:
                        pass
                
                cumulative_pages += page_count
                doc.close()
                
            manager.close_all()
            
            self.progress_queue.put({
                "type": "done",
                "message": "파일을 성공적으로 읽어왔습니다.",
                "pdf_items": pdf_items,
                "toc_data": toc_data,
                "page_offsets": page_offsets,
                "page_rotations": page_rotations,
                "percent": 100
            })
            
        except Exception as e:
            try:
                manager.close_all()
            except Exception:
                pass
            self.progress_queue.put({"type": "error", "message": f"파일 로드 실패:\n{str(e)}", "percent": 0})
            
    def _execute_sync(self):
        backup_dir = get_backup_dir()
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir, exist_ok=True)
            
        manager = COMOfficeManager()
        pdf_items = []
        toc_data = []
        page_offsets = []
        page_rotations = {}
        cumulative_pages = 0
        
        try:
            for i, filepath in enumerate(self.filepaths):
                filename = self.custom_filenames.get(filepath, os.path.basename(filepath))
                ext = os.path.splitext(filename)[1].lower()
                
                temp_pdf_path = ""
                original_filepath = filepath
                temp_office_path = ""
                is_temp_office = False

                if ext == ".pdf":
                    target_pdf_path = filepath
                elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff"):
                    import hashlib
                    h = hashlib.md5(filepath.encode('utf-8')).hexdigest()[:8]
                    temp_pdf_path = os.path.join(backup_dir, f"temp_img_{h}.pdf")
                    img_doc = fitz.open(filepath)
                    pdf_bytes = img_doc.convert_to_pdf()
                    img_doc.close()
                    with open(temp_pdf_path, "wb") as f:
                        f.write(pdf_bytes)
                    self._optimize_temp_pdf(temp_pdf_path)
                    target_pdf_path = temp_pdf_path
                elif ext in (".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"):
                    import hashlib
                    h = hashlib.md5(filepath.encode('utf-8')).hexdigest()[:8]
                    temp_pdf_path = os.path.join(backup_dir, f"temp_office_{h}.pdf")
                    
                    # Preprocess legacy formats to modern formats for image/graphics compression
                    temp_office_path = filepath
                    is_temp_office = False
                    if ext in (".doc", ".xls", ".ppt"):
                        try:
                            temp_office_path, is_temp_office = self._preprocess_office_file(filepath, ext, manager, backup_dir, h)
                            conv_ext = os.path.splitext(temp_office_path)[1].lower()
                        except Exception:
                            # Graceful fallback to original path if pre-conversion fails
                            temp_office_path = filepath
                            is_temp_office = False
                            conv_ext = ext
                    elif ext in (".docx", ".xlsx", ".pptx"):
                        try:
                            import shutil
                            temp_office_path = os.path.join(backup_dir, f"temp_conv_{h}{ext}")
                            shutil.copy2(filepath, temp_office_path)
                            is_temp_office = True
                            compress_openxml_media(temp_office_path, self.compression_mode)
                            conv_ext = ext
                        except Exception:
                            temp_office_path = filepath
                            is_temp_office = False
                            conv_ext = ext
                    else:
                        conv_ext = ext
                        
                    abs_in = get_com_path(safe_path(temp_office_path))
                    abs_out = get_com_path(safe_path(temp_pdf_path))
                    
                    if conv_ext in (".doc", ".docx"):
                        try:
                            word = manager.get_word()
                            doc = word.Documents.Open(abs_in, False, True)  # FileName, ConfirmConversions=False, ReadOnly=True
                            doc.SaveAs(abs_out, 17) # FileName, FileFormat=17
                            doc.Close(0) # SaveChanges=0 (wdDoNotSaveChanges)
                        except Exception as e:
                            disk_warn = check_disk_space_warning()
                            raise RuntimeError(f"Word 파일 변환 오류 ({filename}):\n{str(e)}{disk_warn}\n\nMS Office 설치 상태와 라이선스를 확인해주세요.")
                    elif conv_ext in (".xls", ".xlsx"):
                        try:
                            excel = manager.get_excel()
                            wb = excel.Workbooks.Open(abs_in, 0, True)  # Filename, UpdateLinks=0, ReadOnly=True
                            try:
                                for ws in wb.Worksheets:
                                    ws.PageSetup.Zoom = False
                                    ws.PageSetup.FitToPagesWide = 1
                                    ws.PageSetup.FitToPagesTall = 0
                            except Exception:
                                pass
                            wb.ExportAsFixedFormat(0, abs_out) # Type=0 (xlTypePDF), Filename
                            wb.Close(False)
                        except Exception as e:
                            disk_warn = check_disk_space_warning()
                            raise RuntimeError(f"Excel 파일 변환 오류 ({filename}):\n{str(e)}{disk_warn}\n\nMS Office 설치 상태와 라이선스를 확인해주세요.")
                    elif conv_ext in (".ppt", ".pptx"):
                        try:
                            ppt = manager.get_ppt()
                            deck = ppt.Presentations.Open(abs_in, True, False, False)  # FileName, ReadOnly=True, Untitled=False, WithWindow=False
                            deck.SaveAs(abs_out, 32) # FileName, FileFormat=32
                            deck.Close()
                        except Exception as e:
                            disk_warn = check_disk_space_warning()
                            raise RuntimeError(f"PowerPoint 파일 변환 오류 ({filename}):\n{str(e)}{disk_warn}\n\nMS Office 설치 상태와 라이선스를 확인해주세요.")
                    self._optimize_temp_pdf(temp_pdf_path)
                    target_pdf_path = temp_pdf_path
                else:
                    raise ValueError(f"Unsupported format: {ext}")
                
                page_offsets.append(cumulative_pages)
                doc = fitz.open(target_pdf_path)
                page_count = len(doc)
                
                preprocessed_office_path = temp_office_path if (is_temp_office and os.path.exists(temp_office_path)) else ""
                
                pdf_items.append(PDFItem(
                    file_idx=i,
                    filepath=target_pdf_path,
                    filename=filename,
                    page_count=page_count,
                    original_filepath=original_filepath,
                    preprocessed_office_path=preprocessed_office_path
                ))
                
                for p_idx in range(page_count):
                    page = doc[p_idx]
                    page_rotations[(i, p_idx)] = page.rotation
                
                base_page = 1 if page_count > 0 else 0
                cat_name = filename
                for drop_ext in (".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"):
                    if cat_name.lower().endswith(drop_ext):
                        cat_name = cat_name[:-len(drop_ext)]
                        break
                        
                prefix_lvl1 = self.state.level_prefixes.get(1, "")
                clean_cat = self.state.strip_prefix(cat_name)
                toc_data.append(TOCItem(
                    file_idx=i,
                    file_name=filename,
                    level=1,
                    title=f"{prefix_lvl1}{clean_cat}".strip(),
                    dest_page=base_page,
                    is_file_header=True
                ))
                
                if ext == ".pdf":
                    try:
                        toc = doc.get_toc()
                        for item in toc:
                            level, title, page = item
                            if "슬라이드" in title:
                                continue
                            lvl = level + 1
                            lvl = max(0, min(10, lvl))
                            prefix = "" if lvl == 0 else self.state.level_prefixes.get(lvl, self.state.level_prefixes.get(10, ""))
                            clean_title = self.state.strip_prefix(title)
                            toc_data.append(TOCItem(
                                file_idx=i,
                                file_name=filename,
                                level=lvl,
                                title=f"{prefix}{clean_title}".strip(),
                                dest_page=page,
                                is_file_header=False
                            ))
                    except Exception:
                        pass
                
                cumulative_pages += page_count
                doc.close()
                
            manager.close_all()
            self.state.pdf_items = pdf_items
            self.state.toc_data = toc_data
            self.state.page_offsets = page_offsets
            self.state.page_rotations = page_rotations
            self.state.manually_rotated = []
            self.state.undo_stack = []
            self.state.notify_listeners()
        except Exception as e:
            manager.close_all()
            raise e

class PDFMergeUseCase:
    def __init__(self, state: AppState, config: MergeConfig):
        self.state = state
        self.config = config
        self.progress_queue = queue.Queue()
        self._response_queue = queue.Queue()

    def run_in_background(self, save_path: str):
        thread = threading.Thread(target=self._execute, args=(save_path,), daemon=True)
        thread.start()

    def _execute(self, save_path: str):
        try:
            self.progress_queue.put({"type": "start", "message": "PDF 병합을 준비하고 있습니다...", "percent": 0})
            if not self.state.pdf_items:
                raise ValueError("병합할 PDF 파일이 없습니다.")
                
            # 1. Determine active files from toc_data file headers
            active_headers = [item for item in self.state.toc_data if item.is_file_header]
            if not active_headers:
                raise ValueError("병합할 활성화된 PDF 파일이 없습니다.")
                
            # Reconstruct the merge sequence (list of original file indices to merge)
            merge_sequence = [header.file_idx for header in active_headers]
            
            # Recalculate page offsets for active files based on the merge order
            new_page_offsets = {}
            current_offset = 0
            for file_idx in merge_sequence:
                new_page_offsets[file_idx] = current_offset
                current_offset += self.state.pdf_items[file_idx].page_count
                
            expected_file_count = len(merge_sequence)
            expected_page_count = current_offset
            
            self.progress_queue.put({"type": "progress", "message": "1/5. 원본 PDF 파일을 병합하고 크기를 조절하는 중...", "percent": 5})
            merged_doc = fitz.open()
            
            paper_sizes = {
                "A4": (595.0, 842.0),
                "A3": (842.0, 1191.0),
                "Letter": (612.0, 792.0),
                "Legal": (612.0, 1008.0),
            }
            base_w, base_h = paper_sizes.get(self.config.toc_paper_size, (595.0, 842.0))
            
            # Merge pages in the order of merge_sequence
            successful_files = 0
            for new_idx, file_idx in enumerate(merge_sequence):
                item = self.state.pdf_items[file_idx]
                doc = fitz.open(item.filepath)
                for p_idx in range(item.page_count):
                    src_page = doc[p_idx]
                    
                    orig_rot = src_page.rotation
                    orig_rect = src_page.rect
                    
                    # Determine physical aspect ratio based on displayed dimensions
                    phys_portrait = orig_rect.width <= orig_rect.height
                    
                    # Determine target rotation (either manual or native)
                    is_manual = (file_idx, p_idx) in self.state.manually_rotated
                    if is_manual:
                        rot = self.state.page_rotations.get((file_idx, p_idx), 0)
                    else:
                        rot = orig_rot
                        
                    # Determine target paper dimensions
                    if phys_portrait:
                        target_w, target_h = base_w, base_h
                    else:
                        target_w, target_h = base_h, base_w
                        
                    # Create a new empty page of target size in merged_doc
                    new_page = merged_doc.new_page(width=target_w, height=target_h)
                    
                    # Draw source page content scaled onto the new page (filling the page)
                    target_rect = fitz.Rect(0, 0, target_w, target_h)
                    new_page.show_pdf_page(target_rect, doc, p_idx, keep_proportion=False)
                    
                    # Compute relative rotation to avoid double-rotation bug
                    relative_rot = (rot - orig_rot) % 360
                    new_page.set_rotation(relative_rot)
                doc.close()
                successful_files += 1
                
            # 2. Extract and layout row variables
            toc_rows = []
            outline_rows = []
            file_list_rows = []
            
            for item in self.state.toc_data:
                # Skip bookmarks if their parent file header was deleted
                if item.file_idx not in new_page_offsets:
                    continue
                    
                global_page = new_page_offsets[item.file_idx] + item.dest_page
                row = [item.level, item.title, global_page]
                
                if item.is_file_header:
                    if self.config.filename_mode == "toc_bookmark":
                        toc_rows.append(row[:])
                    elif self.config.filename_mode == "separate_page":
                        file_list_rows.append(row[:])
                    
                    # Always include file headers in PDF bookmarks outline if not 'none'
                    if self.config.filename_mode in ("toc_bookmark", "bookmark_only", "separate_page"):
                        outline_rows.append(row[:])
                else:
                    if self.config.filename_mode == "toc_bookmark":
                        toc_rows.append(row[:])
                    
                    if self.config.filename_mode in ("toc_bookmark", "bookmark_only", "separate_page"):
                        outline_rows.append(row[:])
                        
            # Calculate metrics
            layout = PDFService.get_toc_layout_settings(self.config)
            page_w = layout["page_width"]
            page_h = layout["page_height"]
            margin_top = layout["margin_top"]
            margin_bottom = layout["margin_bottom"]
            margin_left = layout["margin_left"]
            line_height = layout["line_height"]
            title_font_size = layout["title_font_size"]
            category_font_size = layout["category_font_size"]
            item_font_size = layout["item_font_size"]
            move_button_width = layout["move_button_width"]
            move_button_gap = layout["move_button_gap"]
            
            items_per_page = int((page_h - margin_top - margin_bottom) / line_height)
            
            # File list page is generated only in separate_page mode
            if self.config.filename_mode == "separate_page":
                num_file_list_pages = math.ceil(len(file_list_rows) / items_per_page) if file_list_rows else 1
            else:
                num_file_list_pages = 0
                
            # TOC page is generated only in toc_bookmark mode
            if self.config.filename_mode == "toc_bookmark":
                num_toc_pages = math.ceil(len(toc_rows) / items_per_page) if toc_rows else 1
            else:
                num_toc_pages = 0
                
            # Double safety check: Force 0 front pages if mode is bookmark_only or none
            if self.config.filename_mode in ("bookmark_only", "none"):
                num_file_list_pages = 0
                num_toc_pages = 0
                
            total_front_pages = num_file_list_pages + num_toc_pages
            
            # Recalculate dest pages to account for newly inserted front pages
            for rows in (toc_rows, outline_rows, file_list_rows):
                for item in rows:
                    item[2] += total_front_pages
                    
            # 3. Create TOC pages
            self.progress_queue.put({"type": "progress", "message": "2/5. 목차 페이지 생성 및 그리기...", "percent": 20})
            toc_doc = fitz.open()
            toc_link_records = []
            if total_front_pages > 0:
                for _ in range(total_front_pages):
                    toc_doc.new_page(width=page_w, height=page_h)
                    
                font_path = FontService.get_korean_font_path()
                has_font = font_path is not None
                
                for i in range(total_front_pages):
                    if has_font:
                        toc_doc[i].insert_font(fontname="Malgun", fontfile=font_path)
            
            # Draw File List pages if applicable
            if file_list_rows:
                item_idx = 0
                file_row_count = 0
                for p_idx in range(num_file_list_pages):
                    page = toc_doc[p_idx]
                    page.insert_text((margin_left, margin_top - 30), f"파일 목록 - {p_idx + 1}/{num_file_list_pages}", fontsize=title_font_size, fontname="Malgun" if has_font else "helv", color=(0, 0, 0))
                    
                    y = margin_top
                    items_printed = 0
                    indent = margin_left
                    while items_printed < items_per_page and item_idx < len(file_list_rows):
                        _level, title, dest_page = file_list_rows[item_idx]
                        clean_title = title.replace("■", "").strip()
                        text_x = indent + move_button_width + move_button_gap + 18
                        rect = fitz.Rect(indent, y - item_font_size - 3, indent + move_button_width, y + 4)
                        suffix = f" (Page {dest_page})"
                        avail_w = page_w - margin_left - text_x
                        display_title = PDFService.fit_text_to_width(clean_title, max(20, avail_w - len(suffix) * item_font_size * 0.45), item_font_size)
                        
                        row_rect = fitz.Rect(indent - 4, y - item_font_size - 3, page_w - margin_left, y + 4)
                        
                        # 심리학적 가독성 교차 색상 (Blue-Gray / Sage-Green) - 대비 강화 및 옅고 투명한 투과 효과 교차
                        if file_row_count % 2 == 0:
                            btn_bg = (0.88, 0.93, 0.99)
                            btn_border = (0.55, 0.70, 0.90)
                            title_bg = (0.93, 0.95, 0.98)
                            title_border = (0.79, 0.83, 0.89)
                            fill_op = 1.0
                            stroke_op = 1.0
                        else:
                            btn_bg = (0.87, 0.93, 0.87)
                            btn_border = (0.58, 0.76, 0.58)
                            title_bg = (0.91, 0.94, 0.91)
                            title_border = (0.77, 0.83, 0.77)
                            fill_op = 0.35
                            stroke_op = 0.45

                        shape = page.new_shape()
                        shape.draw_rect(row_rect)
                        shape.finish(color=title_border, fill=title_bg, fill_opacity=fill_op, stroke_opacity=stroke_op)
                        shape.draw_rect(rect)
                        shape.finish(color=btn_border, fill=btn_bg, fill_opacity=fill_op, stroke_opacity=stroke_op)
                        shape.commit()
                        
                        # PDF 자체에 검사용 대화형 체크박스 위젯 추가
                        chk_rect = fitz.Rect(indent + move_button_width + 4, y - item_font_size + 1, indent + move_button_width + 16, y - item_font_size + 13)
                        widget = fitz.Widget()
                        widget.rect = chk_rect
                        widget.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
                        widget.field_name = f"check_list_{p_idx}_{item_idx}"
                        widget.field_value = "Off"
                        widget.border_color = (0.2, 0.4, 0.8)
                        widget.fill_color = (1.0, 1.0, 1.0)
                        page.add_widget(widget)
                        
                        # 대화형 투명 버튼 위젯 추가 (클릭 시 체크박스 자동 선택 및 해당 페이지로 이동 처리)
                        btn_widget = fitz.Widget()
                        btn_widget.rect = rect
                        btn_widget.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
                        btn_widget.field_name = f"btn_list_{p_idx}_{item_idx}"
                        btn_widget.button_caption = ""
                        btn_widget.border_width = 0
                        btn_widget.script = f'this.getField("check_list_{p_idx}_{item_idx}").value = "Yes"; this.pageNum = {dest_page - 1};'
                        page.add_widget(btn_widget)
                        
                        page.insert_text((indent, y), "[이동]", fontsize=item_font_size, fontname="Malgun" if has_font else "helv", color=(0.15, 0.35, 0.65))
                        page.insert_text((text_x, y), f" {display_title}{suffix}", fontsize=item_font_size, fontname="Malgun" if has_font else "helv", color=(0.15, 0.19, 0.25))
                        y += line_height
                        item_idx += 1
                        items_printed += 1
                        file_row_count += 1
                        
            # Draw TOC rows
            item_idx = 0
            sub_item_count = 0
            if num_toc_pages > 0 and not toc_rows:
                 # If we requested 1 TOC page but have no rows, we can skip or just draw empty title.
                 page_index = num_file_list_pages
                 page = toc_doc[page_index]
                 page.insert_text((margin_left, margin_top - 30), f"목차 (Table of Contents) - 1/1", fontsize=title_font_size, fontname="Malgun" if has_font else "helv", color=(0, 0, 0))

            for p_idx in range(num_toc_pages):
                if not toc_rows:
                    break
                page_index = num_file_list_pages + p_idx
                page = toc_doc[page_index]
                page.insert_text((margin_left, margin_top - 30), f"목차 (Table of Contents) - {p_idx + 1}/{num_toc_pages}", fontsize=title_font_size, fontname="Malgun" if has_font else "helv", color=(0, 0, 0))
                
                y = margin_top
                items_printed = 0
                while items_printed < items_per_page and item_idx < len(toc_rows):
                    level, title, dest_page = toc_rows[item_idx]
                    indent = margin_left + max(0, level - 1) * 20
                    
                    if level == 1 and title.startswith("■"):
                        avail_w = page_w - margin_left - indent
                        display_title = PDFService.fit_text_to_width(title, avail_w, category_font_size)
                        page.insert_text((indent, y), display_title, fontsize=category_font_size, fontname="Malgun" if has_font else "helv", color=(0.15, 0.19, 0.25))
                        rect = fitz.Rect(indent, y - category_font_size - 3, page_w - margin_left, y + 4)
                        toc_link_records.append((page_index, rect, dest_page))
                    else:
                         btn_text = "[이동]"
                         text_x = indent + move_button_width + move_button_gap + 18
                         rect = fitz.Rect(indent, y - item_font_size - 3, indent + move_button_width, y + 4)
                         row_rect = fitz.Rect(indent - 4, y - item_font_size - 3, page_w - margin_left, y + 4)
                         
                         suffix = f" (Page {dest_page})"
                         avail_w = page_w - margin_left - text_x
                         display_title = PDFService.fit_text_to_width(title, max(20, avail_w - len(suffix) * item_font_size * 0.45), item_font_size)
                         text = f" {display_title}{suffix}"
                         
                         # 심리학적 가독성 교차 색상 (Blue-Gray / Sage-Green) - 대비 강화 및 옅고 투명한 투과 효과 교차
                         if sub_item_count % 2 == 0:
                             btn_bg = (0.88, 0.93, 0.99)
                             btn_border = (0.55, 0.70, 0.90)
                             title_bg = (0.93, 0.95, 0.98)
                             title_border = (0.79, 0.83, 0.89)
                             fill_op = 1.0
                             stroke_op = 1.0
                         else:
                             btn_bg = (0.87, 0.93, 0.87)
                             btn_border = (0.58, 0.76, 0.58)
                             title_bg = (0.91, 0.94, 0.91)
                             title_border = (0.77, 0.83, 0.77)
                             fill_op = 0.35
                             stroke_op = 0.45

                         shape = page.new_shape()
                         shape.draw_rect(row_rect)
                         shape.finish(color=title_border, fill=title_bg, fill_opacity=fill_op, stroke_opacity=stroke_op)
                         shape.draw_rect(rect)
                         shape.finish(color=btn_border, fill=btn_bg, fill_opacity=fill_op, stroke_opacity=stroke_op)
                         shape.commit()
                         
                         # PDF 자체에 검사용 대화형 체크박스 위젯 추가
                         chk_rect = fitz.Rect(indent + move_button_width + 4, y - item_font_size + 1, indent + move_button_width + 16, y - item_font_size + 13)
                         widget = fitz.Widget()
                         widget.rect = chk_rect
                         widget.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
                         widget.field_name = f"check_toc_{p_idx}_{item_idx}"
                         widget.field_value = "Off"
                         widget.border_color = (0.2, 0.4, 0.8)
                         widget.fill_color = (1.0, 1.0, 1.0)
                         page.add_widget(widget)
                         
                         # 대화형 투명 버튼 위젯 추가 (클릭 시 체크박스 자동 선택 및 해당 페이지로 이동 처리)
                         btn_widget = fitz.Widget()
                         btn_widget.rect = rect
                         btn_widget.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
                         btn_widget.field_name = f"btn_toc_{p_idx}_{item_idx}"
                         btn_widget.button_caption = ""
                         btn_widget.border_width = 0
                         btn_widget.script = f'this.getField("check_toc_{p_idx}_{item_idx}").value = "Yes"; this.pageNum = {dest_page - 1};'
                         page.add_widget(btn_widget)
                         
                         page.insert_text((indent, y), btn_text, fontsize=item_font_size, fontname="Malgun" if has_font else "helv", color=(0.15, 0.35, 0.65))
                         page.insert_text((text_x, y), text, fontsize=item_font_size, fontname="Malgun" if has_font else "helv", color=(0.15, 0.19, 0.25))
                         sub_item_count += 1
                        
                    y += line_height
                    item_idx += 1
                    items_printed += 1
                    
            # 4. Compile final PDF document
            final_doc = fitz.open()
            if total_front_pages > 0:
                final_doc.insert_pdf(toc_doc)
            final_doc.insert_pdf(merged_doc)
            final_page_count = len(final_doc)
            
            # Apply TOC Links
            if total_front_pages > 0:
                for toc_page_index, rect, dest_page in toc_link_records:
                    target_page = max(0, min(final_page_count - 1, dest_page - 1))
                    final_doc[toc_page_index].insert_link({"kind": fitz.LINK_GOTO, "page": target_page, "from": rect})
                
            # Set PDF Bookmarks Outline
            safe_tocs = []
            if self.config.filename_mode != "none":
                if file_list_rows:
                    safe_tocs.append([1, "파일 목록", 1])
                for level, title, dest_page in outline_rows:
                    safe_level = max(1, level)
                    safe_page = max(1, min(final_page_count, int(dest_page)))
                    safe_tocs.append([safe_level, title, safe_page])
                if safe_tocs:
                    final_doc.set_toc(safe_tocs)
                
            # 5. Insert Navigation Links on ALL pages (HOME, PREV, NEXT)
            if self.config.filename_mode != "none":
                self.progress_queue.put({"type": "progress", "message": "3/5. 내비게이션 버튼(HOME/PREV/NEXT) 그리기...", "percent": 40})
                nav_bg = (0.93, 0.96, 0.99)
                nav_border = (0.67, 0.78, 0.92)
                nav_fg = (0.15, 0.35, 0.65)
                
                for i in range(len(final_doc)):
                    page = final_doc[i]
                    y0 = 20
                    x_home = page.rect.width - 70
                    rect_home = fitz.Rect(x_home, y0, x_home + 50, y0 + 15)
                    
                    x_next = x_home - 60
                    rect_next = fitz.Rect(x_next, y0, x_next + 50, y0 + 15)
                    
                    x_prev = x_next - 60
                    rect_prev = fitz.Rect(x_prev, y0, x_prev + 50, y0 + 15)
                    
                    # Apply derotation matrix to map display coords back to unrotated user space
                    mat = page.derotation_matrix
                    rot = page.rotation
                    
                    rect_home_tr = rect_home * mat
                    rect_next_tr = rect_next * mat
                    rect_prev_tr = rect_prev * mat
                    
                    shape = page.new_shape()
                    shape.draw_rect(rect_home_tr)
                    shape.finish(color=nav_border, fill=nav_bg)
                    
                    if i < len(final_doc) - 1:
                        shape.draw_rect(rect_next_tr)
                        shape.finish(color=nav_border, fill=nav_bg)
                        
                    if i > 0:
                        shape.draw_rect(rect_prev_tr)
                        shape.finish(color=nav_border, fill=nav_bg)
                    shape.commit()
                    
                    page.insert_text(fitz.Point(x_home + 10, y0 + 11) * mat, "HOME", fontsize=9, color=nav_fg, rotate=rot)
                    
                    # HOME 투명 위젯 추가 (손가락 커서 강제 및 JS 이동)
                    btn_home = fitz.Widget()
                    btn_home.rect = rect_home_tr
                    btn_home.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
                    btn_home.field_name = f"btn_nav_home_{i}"
                    btn_home.button_caption = ""
                    btn_home.border_width = 0
                    btn_home.script = "this.pageNum = 0;"
                    page.add_widget(btn_home)
                    
                    if i < len(final_doc) - 1:
                        page.insert_text(fitz.Point(x_next + 12, y0 + 11) * mat, "NEXT", fontsize=9, color=nav_fg, rotate=rot)
                        
                        btn_next = fitz.Widget()
                        btn_next.rect = rect_next_tr
                        btn_next.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
                        btn_next.field_name = f"btn_nav_next_{i}"
                        btn_next.button_caption = ""
                        btn_next.border_width = 0
                        btn_next.script = f"this.pageNum = {i + 1};"
                        page.add_widget(btn_next)
                        
                    if i > 0:
                        page.insert_text(fitz.Point(x_prev + 12, y0 + 11) * mat, "PREV", fontsize=9, color=nav_fg, rotate=rot)
                        
                        btn_prev = fitz.Widget()
                        btn_prev.rect = rect_prev_tr
                        btn_prev.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
                        btn_prev.field_name = f"btn_nav_prev_{i}"
                        btn_prev.button_caption = ""
                        btn_prev.border_width = 0
                        btn_prev.script = f"this.pageNum = {i - 1};"
                        page.add_widget(btn_prev)
            else:
                self.progress_queue.put({"type": "progress", "message": "3/5. 내비게이션 비활성화...", "percent": 40})
                    
            # 6. Compress images with parallel executor
            self.progress_queue.put({"type": "progress", "message": "4/5. 이미지 최적화 및 용량 압축 중...", "percent": 60})
            
            def notify_img_progress(msg: str, percent: int = None):
                self.progress_queue.put({"type": "progress", "message": f"4/5. {msg}", "percent": percent or 60})
                
            compression_stats = ImageCompressionService.compress_pdf_images(
                final_doc, self.config.compression_mode, notify_img_progress
            )
            
            # Original total size
            original_total_size = 0
            for item in self.state.pdf_items:
                path_to_check = item.original_filepath if (item.original_filepath and os.path.exists(item.original_filepath)) else item.filepath
                if os.path.exists(path_to_check):
                    original_total_size += os.path.getsize(path_to_check)
            
            # 7. Compress and save PDF
            self.progress_queue.put({"type": "progress", "message": "5/5. 최종 PDF 최적화 저장 중...", "percent": 92})
            
            # Atomic Save transaction (5-Stage Transaction Sequence)
            temp_out = save_path + ".tmp"
            bak_file = save_path + ".bak"
            
            # Stage 1: Pre-lock occupancy check
            if os.path.exists(save_path):
                try:
                    with open(save_path, 'a'):
                        pass
                except Exception as e:
                    raise RuntimeError(f"출력 파일이 다른 프로그램에 의해 열려 있거나 권한이 없어 쓸 수 없습니다: {e}")
                    
            # Subsetting embedded fonts to drastically reduce file size (stripping unused glyphs)
            try:
                if hasattr(final_doc, "subset_fonts"):
                    final_doc.subset_fonts()
            except Exception:
                pass

            # Stage 2: Save to temp_out
            final_doc.save(
                temp_out,
                garbage=4,
                deflate=True,
                deflate_images=True,
                deflate_fonts=True,
                clean=True,
                use_objstms=1,
                compression_effort=100,
            )
            
            toc_doc.close()
            merged_doc.close()
            final_doc.close()
            
            # Stage 3: Atomic Swap with Backup-First
            existed = os.path.exists(save_path)
            if existed:
                if os.path.exists(bak_file):
                    try: os.remove(bak_file)
                    except: pass
                try:
                    os.rename(save_path, bak_file)
                except Exception as e:
                    raise RuntimeError(f"원본 파일을 백업 파일로 변경하는 데 실패했습니다: {e}")
                    
            # Rename temp to target
            try:
                os.rename(temp_out, save_path)
            except Exception as swap_err:
                # Rollback from backup
                if existed and os.path.exists(bak_file):
                    try: os.rename(bak_file, save_path)
                    except: pass
                raise RuntimeError(f"임시 파일을 최종 경로로 변경하는 데 실패했습니다: {swap_err}")
                
            # Stage 4: Verification of physical existence and size > 0
            if not os.path.exists(save_path) or os.path.getsize(save_path) == 0:
                # Rollback
                if existed and os.path.exists(bak_file):
                    try: os.rename(bak_file, save_path)
                    except: pass
                raise RuntimeError("생성된 PDF 파일의 크기가 0이거나 정상적으로 파일이 기록되지 않았습니다.")
                
            # Stage 5: Success Commit (Clean up backup file)
            if existed and os.path.exists(bak_file):
                try: os.remove(bak_file)
                except: pass
                
            # Perform logical verification of target file count and page count
            actual_page_count = 0
            verification_status = "실패"
            actual_file_count = 0
            try:
                verify_doc = fitz.open(save_path)
                actual_page_count = len(verify_doc)
                
                # Check page count
                page_count_ok = (actual_page_count == expected_page_count + total_front_pages)
                
                # Check outline bookmark count if outline was expected
                outline_ok = True
                if self.config.filename_mode != "none" and safe_tocs:
                    actual_toc = verify_doc.get_toc()
                    if len(actual_toc) != len(safe_tocs):
                        outline_ok = False
                
                verify_doc.close()
                
                if page_count_ok and outline_ok:
                    verification_status = "검증 완료 (일치)"
                    actual_file_count = successful_files
                elif not page_count_ok:
                    verification_status = f"검증 실패 (페이지 불일치: 예상 {expected_page_count + total_front_pages}장, 실제 {actual_page_count}장)"
                    actual_file_count = 0
                else:
                    verification_status = f"검증 실패 (북마크/목차 불일치: 예상 {len(safe_tocs)}개, 실제 {len(actual_toc)}개)"
                    actual_file_count = 0
            except Exception as e:
                verification_status = f"검증 실패 (오류: {e})"
                actual_file_count = 0
            
            # Stage 6: Copy preprocessed files to original folder if option is enabled
            copied_files = []
            copy_errors = []
            if self.config.save_preprocessed:
                self.progress_queue.put({"type": "progress", "message": "6/6. 전처리 파일 원본 폴더에 저장 중...", "percent": 97})
                import shutil
                for item in self.state.pdf_items:
                    src = item.preprocessed_office_path
                    if not src or not os.path.exists(src):
                        continue
                    orig_dir = os.path.dirname(item.original_filepath) if item.original_filepath else ""
                    if not orig_dir:
                        continue
                    orig_name = os.path.basename(item.original_filepath)
                    base_name, _ = os.path.splitext(orig_name)
                    dest_ext = os.path.splitext(src)[1]
                    dest_path = get_unique_output_path(os.path.join(orig_dir, base_name + dest_ext))
                    try:
                        shutil.copy2(src, dest_path)
                        copied_files.append(os.path.basename(dest_path))
                    except Exception as copy_err:
                        err_msg = (
                            f"전처리 파일 저장 실패:\n{os.path.basename(dest_path)}\n\n"
                            f"오류: {copy_err}\n\n"
                            f"나머지 전처리 파일 저장을 계속하시겠습니까?"
                        )
                        copy_errors.append(os.path.basename(dest_path))
                        self.progress_queue.put({"type": "ask_continue", "message": err_msg})
                        # Wait for GUI response
                        try:
                            user_continue = self._response_queue.get(timeout=60)
                            if not user_continue:
                                break
                        except Exception:
                            break

            output_size = os.path.getsize(save_path) if os.path.exists(save_path) else 0
            saved_total = max(0, original_total_size - output_size)
            saved_image = compression_stats["saved_bytes"]
            mode_label = {"strong": "강력 압축", "balanced": "균형 압축", "basic": "기본 최적화"}[self.config.compression_mode]

            copy_summary = ""
            if self.config.save_preprocessed:
                if copied_files:
                    copy_summary = f"\n전처리 파일 저장: {len(copied_files)}개 완료 ({', '.join(copied_files[:3])}{'...' if len(copied_files) > 3 else ''})"
                if copy_errors:
                    copy_summary += f"\n전처리 파일 저장 실패: {len(copy_errors)}개 ({', '.join(copy_errors[:3])}{'...' if len(copy_errors) > 3 else ''})"

            result_message = (
                f"성공적으로 병합 및 압축이 완료되었습니다!\n\n"
                f"=== [병합 및 무결성 검증 결과] ===\n"
                f"• 검증 상태: {verification_status}\n"
                f"• 대상 파일 수: {expected_file_count}개 / 실제 병합된 파일 수: {actual_file_count}개\n"
                f"• 대상 페이지 수: {expected_page_count + total_front_pages}장 (목차 {total_front_pages}장 포함) / 실제 병합된 페이지 수: {actual_page_count}장\n\n"
                f"=== [용량 압축 결과] ===\n"
                f"• 압축 모드: {mode_label}\n"
                f"• 이미지 최적화: {compression_stats['replaced']}개 적용"
                f" (예상 절감 {saved_image / 1024 / 1024:.1f} MB)\n"
                f"• 원본 합계: {original_total_size / 1024 / 1024:.1f} MB\n"
                f"• 최종 파일: {output_size / 1024 / 1024:.1f} MB\n"
                f"• 전체 절감: {saved_total / 1024 / 1024:.1f} MB{copy_summary}\n\n"
                f"저장 위치: {save_path}"
            )
            self.progress_queue.put({"type": "done", "message": result_message, "save_path": save_path, "percent": 100})
            
        except Exception as e:
            self.progress_queue.put({"type": "error", "message": f"병합 중 오류가 발생했습니다:\n{str(e)}", "percent": 0})


# =====================================================================
# 4. PRESENTATION LAYER (Tkinter UI)
# =====================================================================

class BookmarkPromptDialog(tk.Toplevel):
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.parent = parent
        self.filename = filename
        self.result = "filename"  # default fallback
        
        self.title("북마크 가져오기 선택")
        self.geometry("520x220")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # Center dialog relative to parent window
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 150, parent.winfo_rooty() + 150))
        
        main_frame = tk.Frame(self, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        lbl_msg = tk.Label(
            main_frame,
            text=f"불러온 PDF 파일 [{filename}] 내부에 북마크(목차) 정보가 있습니다.\n"
                 f"목차 제목으로 무엇을 사용하시겠습니까?",
            font=("Malgun Gothic", 10, "bold"),
            justify=tk.LEFT,
            anchor=tk.W
        )
        lbl_msg.pack(fill=tk.X, pady=(0, 15))
        
        btn_frame1 = tk.Frame(main_frame)
        btn_frame1.pack(fill=tk.X, pady=5)
        
        btn_frame2 = tk.Frame(main_frame)
        btn_frame2.pack(fill=tk.X, pady=5)
        
        btn_filename = tk.Button(
            btn_frame1, 
            text="파일 이름 사용", 
            command=lambda: self.set_result("filename"),
            width=24,
            font=("Malgun Gothic", 9)
        )
        btn_filename.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        btn_bookmark = tk.Button(
            btn_frame1, 
            text="파일 내부 북마크 사용", 
            command=lambda: self.set_result("bookmark"),
            width=24,
            bg="#3b82f6", 
            fg="white",
            font=("Malgun Gothic", 9, "bold")
        )
        btn_bookmark.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        btn_filename_all = tk.Button(
            btn_frame2, 
            text="이후 모두 파일 이름 사용", 
            command=lambda: self.set_result("filename_all"),
            width=24,
            font=("Malgun Gothic", 9)
        )
        btn_filename_all.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        btn_bookmark_all = tk.Button(
            btn_frame2, 
            text="이후 모두 내부 북마크 사용", 
            command=lambda: self.set_result("bookmark_all"),
            width=24,
            font=("Malgun Gothic", 9)
        )
        btn_bookmark_all.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        self.protocol("WM_DELETE_WINDOW", lambda: self.set_result("filename"))
        self.wait_window()

    def set_result(self, val):
        self.result = val
        self.destroy()


class PrefixSettingsDialog(tk.Toplevel):
    def __init__(self, parent_widget, app, state):
        super().__init__(parent_widget)
        self.parent_widget = parent_widget
        self.app = app
        self.state = state
        
        self.title("레벨별 접두기호 설정")
        self.geometry("450x480")
        self.resizable(False, False)
        self.transient(parent_widget)
        self.grab_set()
        
        # Center dialog relative to parent window
        self.geometry("+%d+%d" % (parent_widget.winfo_rootx() + 150, parent_widget.winfo_rooty() + 80))
        
        main_frame = tk.Frame(self, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="레벨별 접두기호 설정 (공백 포함 입력 가능)", font=("Malgun Gothic", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        grid_frame = tk.Frame(main_frame)
        grid_frame.pack(fill=tk.X, pady=5)
        grid_frame.columnconfigure(1, weight=1)
        
        self.entries = {}
        for lvl in range(1, 11):
            lbl_text = f"레벨 {lvl} 접두기호:"
            tk.Label(grid_frame, text=lbl_text, font=("Malgun Gothic", 9, "bold")).grid(row=lvl-1, column=0, sticky=tk.W, pady=4, padx=(0, 10))
            
            val = self.state.level_prefixes.get(lvl, "")
            entry_var = tk.StringVar(value=val)
            entry = tk.Entry(grid_frame, textvariable=entry_var, font=("Malgun Gothic", 10))
            entry.grid(row=lvl-1, column=1, sticky=tk.EW, pady=4)
            self.entries[lvl] = entry_var
            
        btn_frame = tk.Frame(main_frame, pady=10)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        btn_apply_sel = tk.Button(
            btn_frame, 
            text="선택 항목 적용", 
            command=self.apply_to_selected, 
            width=13, 
            bg="#4CAF50", 
            fg="white", 
            font=("Malgun Gothic", 9, "bold")
        )
        btn_apply_sel.pack(side=tk.LEFT, padx=3)
        
        btn_apply_all = tk.Button(
            btn_frame, 
            text="전체 항목 적용", 
            command=self.apply_to_all, 
            width=13, 
            bg="#FF9800", 
            fg="white", 
            font=("Malgun Gothic", 9, "bold")
        )
        btn_apply_all.pack(side=tk.LEFT, padx=3)
        
        btn_save_close = tk.Button(
            btn_frame, 
            text="설정 저장 및 닫기", 
            command=self.save_and_close, 
            width=14, 
            bg="#3b82f6", 
            fg="white", 
            font=("Malgun Gothic", 9, "bold")
        )
        btn_save_close.pack(side=tk.RIGHT, padx=3)
        
    def apply_to_selected(self):
        # Update state prefixes first
        for lvl, var in self.entries.items():
            self.state.level_prefixes[lvl] = var.get()
        # Save to disk
        self.state.save_level_prefixes()
        
        # Get selected items
        selected_iids = list(self.app.tree.selection())
        if not selected_iids:
            messagebox.showwarning("경고", "선택된 목차 항목이 없습니다. 항목을 먼저 선택해 주세요.")
            return
            
        # Save undo state
        self.app.state.push_undo(selected_iids)
        
        # Apply to selected items in AppState
        indices = [int(iid) for iid in selected_iids]
        self.state.apply_level_prefixes_to_indices(indices)
        
        messagebox.showinfo("완료", f"선택한 {len(indices)}개 목차 항목에 새로운 접두기호를 적용하고 설정을 저장했습니다.")
        self.destroy()

    def apply_to_all(self):
        # Update state prefixes first
        for lvl, var in self.entries.items():
            self.state.level_prefixes[lvl] = var.get()
        # Save to disk
        self.state.save_level_prefixes()
        
        # Save undo state
        self.app.state.push_undo(list(self.app.tree.selection()))
        
        # Apply to all items
        self.state.apply_level_prefixes()
        messagebox.showinfo("완료", "모든 목차 항목에 새로운 접두기호를 일괄 적용하고 설정을 저장했습니다.")
        self.destroy()
        
    def save_and_close(self):
        for lvl, var in self.entries.items():
            self.state.level_prefixes[lvl] = var.get()
        # Save to disk
        self.state.save_level_prefixes()
        self.destroy()


class DuplicateResolveDialog(tk.Toplevel):
    def __init__(self, parent, filename):
        super().__init__(parent)
        self.title("중복 파일 추가 선택")
        self.geometry("450x180")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # Center relative to parent
        self.geometry("+%d+%d" % (parent.winfo_rootx() + 200, parent.winfo_rooty() + 150))
        
        self.result = "cancel" # Default choice
        
        label_msg = f"목록에 이미 존재하는 파일 이름이 감지되었습니다.\n\n파일명: {filename}\n\n어떻게 처리할까요?"
        tk.Label(self, text=label_msg, font=("Malgun Gothic", 10), justify=tk.LEFT, padx=20, pady=20).pack(fill=tk.BOTH, expand=True)
        
        btn_frame = tk.Frame(self, pady=10, bg="#f1f5f9")
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        btn_add = tk.Button(btn_frame, text="그대로 추가", command=lambda: self.set_result("add"), width=12, bg="#4CAF50", fg="white", font=("Malgun Gothic", 9, "bold"))
        btn_add.pack(side=tk.LEFT, padx=10, expand=True)
        
        btn_rename = tk.Button(btn_frame, text="이름 변경 후 추가", command=lambda: self.set_result("rename"), width=16, bg="#3b82f6", fg="white", font=("Malgun Gothic", 9, "bold"))
        btn_rename.pack(side=tk.LEFT, padx=10, expand=True)
        
        btn_cancel = tk.Button(btn_frame, text="추가 취소", command=lambda: self.set_result("cancel"), width=12, font=("Malgun Gothic", 9))
        btn_cancel.pack(side=tk.LEFT, padx=10, expand=True)
        
    def set_result(self, val):
        self.result = val
        self.destroy()


class PDFMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 통합 및 목차 마스터 (PDF Master Merge)")
        self.root.geometry("1180x700")
        self.root.minsize(900, 600)
        
        # State and Config Initialization
        self.state = AppState()
        self.config = MergeConfig()
        
        # Link state change to UI redraw
        self.state.register_listener(self.refresh_tree)
        
        # GUI reactive variables
        self.compression_mode = tk.StringVar(value=self.config.compression_mode)
        self.file_name_mode = tk.StringVar(value=self.config.filename_mode)
        self.toc_paper_size = tk.StringVar(value=self.config.toc_paper_size)
        self.toc_orientation = tk.StringVar(value=self.config.toc_orientation)
        
        # Bind traces to sync config and UI views
        self.compression_mode.trace_add("write", lambda *_: self._sync_config())
        self.file_name_mode.trace_add("write", lambda *_: {self._sync_config(), self.update_toc_preview()})
        self.toc_paper_size.trace_add("write", lambda *_: {self._sync_config(), self.update_toc_preview()})
        self.toc_orientation.trace_add("write", lambda *_: {self._sync_config(), self.update_toc_preview()})
        
        self.status_text = tk.StringVar(value="PDF 파일을 선택하세요.")
        self.preview_click_counts = {}
        self.buttons_to_disable = []
        self.usecase: Optional[PDFMergeUseCase] = None
        self.load_usecase: Optional[PDFLoadUseCase] = None

        # 이기종 전처리 저장 옵션 변수 (create_widgets 이전 초기화 필요)
        self.save_preprocessed = tk.BooleanVar(value=False)
        self.save_preprocessed.trace_add("write", lambda *_: self._sync_config())

        self.create_widgets()

    def _sync_config(self):
        self.config.compression_mode = self.compression_mode.get()
        self.config.filename_mode = self.file_name_mode.get()
        self.config.toc_paper_size = self.toc_paper_size.get()
        self.config.toc_orientation = self.toc_orientation.get()
        if hasattr(self, "save_preprocessed"):
            self.config.save_preprocessed = self.save_preprocessed.get()

    def create_widgets(self):
        # Custom Progressbar Style for thickness
        style = ttk.Style(self.root)
        style.configure("Big.Horizontal.TProgressbar", thickness=20)
        
        # Fix for Treeview tag colors on Windows under default themes (vista/xpnative)
        def fixed_map(option):
            return [elm for elm in style.map("Treeview", query_opt=option)
                    if elm[:2] != ("!disabled", "!selected")]
        style.map("Treeview", foreground=fixed_map("foreground"), background=fixed_map("background"))
        
        # Status and Progress Bar (Moved to Top & Visual Enhancement)
        status_frame = tk.Frame(self.root, padx=15, pady=10, bg="#eef2f7")
        status_frame.pack(fill=tk.X, side=tk.TOP)
        
        self.progress = ttk.Progressbar(status_frame, style="Big.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 15))
        
        self.lbl_percent = tk.Label(status_frame, text="0%", font=("Malgun Gothic", 11, "bold"), width=6, bg="#eef2f7", fg="#2f65d9")
        self.lbl_percent.pack(side=tk.LEFT, padx=(0, 15))
        
        tk.Label(status_frame, textvariable=self.status_text, font=("Malgun Gothic", 10, "bold"), bg="#eef2f7", fg="#333333").pack(side=tk.RIGHT)

        # Top Frame for files
        top_frame = tk.Frame(self.root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        
        btn_select = tk.Button(top_frame, text="1. 병합할 PDF 다중 선택", command=self.select_files, font=("Malgun Gothic", 10, "bold"))
        btn_select.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_select)
        
        self.lbl_files = tk.Label(top_frame, text="선택된 파일이 없습니다.", font=("Malgun Gothic", 10))
        self.lbl_files.pack(side=tk.LEFT, padx=10)
        
        # Main Frame for Treeview
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        
        # Instructions
        tk.Label(main_frame, text="2. 목차 사전 검수 및 편집 (더블클릭하여 이름 수정, 항목 선택 후 삭제/순서변경 가능)", font=("Malgun Gothic", 10)).pack(anchor=tk.W, pady=5)
        
        # Treeview
        columns = ("file", "level", "title", "rotation", "page")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("file", text="원본 파일", command=lambda: self.sort_by_column("file"))
        self.tree.heading("level", text="레벨", command=lambda: self.sort_by_column("level"))
        self.tree.heading("title", text="목차 제목 (더블클릭 편집)", command=lambda: self.sort_by_column("title"))
        self.tree.heading("rotation", text="회전", command=lambda: self.sort_by_column("rotation"))
        self.tree.heading("page", text="목적 페이지", command=lambda: self.sort_by_column("page"))
        
        self.tree.column("file", width=200, stretch=False)
        self.tree.column("level", width=50, stretch=False, anchor=tk.CENTER)
        self.tree.column("title", width=420, stretch=True)
        self.tree.column("rotation", width=80, stretch=False, anchor=tk.CENTER)
        self.tree.column("page", width=80, stretch=False, anchor=tk.CENTER)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Zebra Striping tag configure for table view
        self.tree.tag_configure("evenrow", background="#f4f7fc")
        self.tree.tag_configure("oddrow", background="#ffffff")
        
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # Right Frame for Action Buttons
        right_frame = tk.Frame(self.root, padx=10, pady=10)
        
        btn_del = tk.Button(right_frame, text="선택 항목 삭제", command=self.delete_item, width=13)
        btn_del.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_del)
        
        btn_select_all = tk.Button(right_frame, text="전체 선택", command=self.select_all_toc_items, width=9)
        btn_select_all.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_select_all)
        
        btn_clear_sel = tk.Button(right_frame, text="전체 해제", command=self.clear_toc_selection, width=9)
        btn_clear_sel.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_clear_sel)
        
        btn_batch = tk.Button(right_frame, text="목차명 일괄 변경", command=self.open_batch_edit_dialog, width=14)
        btn_batch.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_batch)
        
        btn_prefix = tk.Button(right_frame, text="접두기호 설정/적용", command=self.open_prefix_settings_dialog, width=15)
        btn_prefix.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_prefix)
        
        btn_up = tk.Button(right_frame, text="▲ 위로 이동", command=self.move_up, width=11)
        btn_up.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_up)
        
        btn_down = tk.Button(right_frame, text="▼ 아래로 이동", command=self.move_down, width=11)
        btn_down.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.append(btn_down)
        
        btn_merge = tk.Button(right_frame, text="3. 최종 병합 및 압축 실행", command=self.merge_pdfs, font=("Malgun Gothic", 11, "bold"), bg="#4CAF50", fg="white", width=25)
        btn_merge.pack(side=tk.RIGHT, padx=5)
        self.buttons_to_disable.append(btn_merge)

        # ⚙ 전처리 저장 옵션 — 병합 버튼 왼쪽에 항상 보이게 배치
        chk_save_prep = tk.Checkbutton(
            right_frame,
            text=" ⚙ 이기종 전처리 파일(docx/xlsx/pptx) 원본 폴더에 저장",
            variable=self.save_preprocessed,
            font=("Malgun Gothic", 9, "bold"),
            fg="#1a3a6a",
            activeforeground="#0d47a1",
            selectcolor="#dce8ff",
            bg=right_frame.cget("bg"),
            activebackground="#dce8ff",
            padx=8, pady=4,
            cursor="hand2",
            relief="flat",
            bd=0,
            anchor="w",
        )
        chk_save_prep.pack(side=tk.RIGHT, padx=(0, 10))
        self.buttons_to_disable.append(chk_save_prep)

        options_area = tk.Frame(self.root, padx=10)
        options_area.columnconfigure(0, weight=1)
        options_area.columnconfigure(1, weight=1)

        # Compression options
        option_frame = tk.LabelFrame(options_area, text="용량 압축 옵션", padx=10, pady=8)
        option_frame.grid(row=0, column=0, sticky="ew", padx=(0, 5), pady=(0, 5))

        r_comp1 = tk.Radiobutton(option_frame, text="강력 압축 (강력 추천: 스캔/사진 PDF 용량 대폭 감소)", variable=self.compression_mode, value="strong")
        r_comp1.pack(side=tk.LEFT, padx=5)
        r_comp2 = tk.Radiobutton(option_frame, text="균형 압축 (화질 우선)", variable=self.compression_mode, value="balanced")
        r_comp2.pack(side=tk.LEFT, padx=5)
        r_comp3 = tk.Radiobutton(option_frame, text="기본 최적화만", variable=self.compression_mode, value="basic")
        r_comp3.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.extend([r_comp1, r_comp2, r_comp3])

        # File bookmark modes
        file_name_frame = tk.LabelFrame(options_area, text="개별 PDF 파일명 목록/북마크", padx=10, pady=8)
        file_name_frame.grid(row=0, column=1, sticky="ew", padx=(5, 0), pady=(0, 5))
        
        r_fn1 = tk.Radiobutton(file_name_frame, text="목차+북마크", variable=self.file_name_mode, value="toc_bookmark")
        r_fn1.pack(side=tk.LEFT, padx=5)
        r_fn2 = tk.Radiobutton(file_name_frame, text="북마크만", variable=self.file_name_mode, value="bookmark_only")
        r_fn2.pack(side=tk.LEFT, padx=5)
        r_fn3 = tk.Radiobutton(file_name_frame, text="별도 파일목록 페이지", variable=self.file_name_mode, value="separate_page")
        r_fn3.pack(side=tk.LEFT, padx=5)
        r_fn4 = tk.Radiobutton(file_name_frame, text="사용 안 함", variable=self.file_name_mode, value="none")
        r_fn4.pack(side=tk.LEFT, padx=5)
        self.buttons_to_disable.extend([r_fn1, r_fn2, r_fn3, r_fn4])

        # TOC layout orientation options
        toc_option_frame = tk.LabelFrame(options_area, text="목차 페이지 방향", padx=10, pady=8)
        toc_option_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        toc_option_frame.columnconfigure(1, weight=1)

        toc_controls = tk.Frame(toc_option_frame)
        toc_controls.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        tk.Label(toc_controls, text="용지 크기", font=("Malgun Gothic", 9, "bold")).pack(anchor=tk.W)
        
        paper_combo = ttk.Combobox(toc_controls, textvariable=self.toc_paper_size, values=("A4", "A3", "Letter", "Legal"), state="readonly", width=10)
        paper_combo.pack(anchor=tk.W, pady=(2, 8))
        self.buttons_to_disable.append(paper_combo)
        
        tk.Label(toc_controls, text="방향", font=("Malgun Gothic", 9, "bold")).pack(anchor=tk.W)
        r_or1 = tk.Radiobutton(toc_controls, text="세로", variable=self.toc_orientation, value="portrait")
        r_or1.pack(anchor=tk.W, pady=2)
        r_or2 = tk.Radiobutton(toc_controls, text="가로", variable=self.toc_orientation, value="landscape")
        r_or2.pack(anchor=tk.W, pady=2)
        self.buttons_to_disable.extend([r_or1, r_or2])
        
        btn_zoom = tk.Button(toc_controls, text="크게 보기", command=self.open_toc_preview_window, width=12)
        btn_zoom.pack(anchor=tk.W, pady=(8, 0))
        self.buttons_to_disable.append(btn_zoom)
        
        self.lbl_toc_preview = tk.Label(toc_controls, text="", font=("Malgun Gothic", 9), fg="#333333")
        self.lbl_toc_preview.pack(anchor=tk.W, pady=(8, 0))

        self.toc_preview_canvas = tk.Canvas(toc_option_frame, width=520, height=110, bg="#f5f5f5", highlightthickness=1, highlightbackground="#cccccc")
        self.toc_preview_canvas.grid(row=0, column=1, sticky="ew")
        self.toc_preview_canvas.bind("<Configure>", lambda _event: self.update_toc_preview())
        self.update_toc_preview()

        # Pack frames in bottom-up order to prevent options panel clipping
        options_area.pack(fill=tk.X, side=tk.BOTTOM, pady=(0, 5))
        right_frame.pack(fill=tk.X, side=tk.BOTTOM)
        main_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

        # (전처리 저장 옵션 체크박스는 right_frame 병합 버튼 행에 통합됨)

        # (Status and progress components were moved to the top of create_widgets)
        pass

    def set_ui_state(self, state: str):
        for widget in self.buttons_to_disable:
            try:
                widget.configure(state=state)
            except Exception:
                pass

    def update_toc_preview(self):
        if not hasattr(self, "toc_preview_canvas"):
            return
        self.render_toc_preview(self.toc_preview_canvas, compact=True)
        settings = PDFService.get_toc_layout_settings(self.config)
        self.lbl_toc_preview.config(text=f"{settings['paper_name']} / {settings['orientation_label']}")

    def render_toc_preview(self, canvas, compact=True, zoom_scale=None):
        canvas.delete("all")
        canvas._toc_hitboxes = []
        canvas._toc_compact = compact
        canvas._toc_zoom_scale = zoom_scale
        canvas.bind("<Button-1>", lambda event, target=canvas: self.handle_toc_preview_click(event, target))
        
        settings = PDFService.get_toc_layout_settings(self.config)
        page_w = settings["page_width"]
        page_h = settings["page_height"]
        margin_top = settings["margin_top"]
        margin_bottom = settings["margin_bottom"]
        margin_left = settings["margin_left"]
        line_height = settings["line_height"]
        move_button_width = settings["move_button_width"]
        move_button_gap = settings["move_button_gap"]

        canvas.update_idletasks()
        canvas_w = max(canvas.winfo_width(), int(canvas.cget("width")), 320)
        canvas_h = max(canvas.winfo_height(), int(canvas.cget("height")), 150)
        if zoom_scale is not None:
            scale = zoom_scale
        elif compact:
            scale = min((canvas_w - 34) / page_w, (canvas_h - 24) / page_h)
        else:
            scale = max(0.45, (canvas_w - 80) / page_w)
            
        draw_w = page_w * scale
        draw_h = page_h * scale
        x0 = max(18, (canvas_w - draw_w) / 2) if compact else 34
        y0 = max(10, (canvas_h - draw_h) / 2) if compact else 24
        x1 = x0 + draw_w
        y1 = y0 + draw_h

        # Page Shadow and Border
        canvas.create_rectangle(x0 + 3, y0 + 3, x1 + 3, y1 + 3, fill="#d9d9d9", outline="")
        canvas.create_rectangle(x0, y0, x1, y1, fill="white", outline="#555555")
        
        preview_font_scale = 0.70 if compact else 1.15
        min_font = 4 if compact else 9
        title_font = max(min_font + 1, int(settings["title_font_size"] * scale * preview_font_scale))
        category_font = max(min_font, int(settings["category_font_size"] * scale * preview_font_scale))
        item_font = max(min_font, int(settings["item_font_size"] * scale * preview_font_scale))

        mode = self.file_name_mode.get()
        if mode == "none":
            canvas.create_rectangle(x0, y0, x1, y1, fill="#f8fafc", outline="#cbd5e1")
            canvas.create_text((x0 + x1)/2, (y0 + y1)/2 - 15, text="목차 / 북마크 미사용", font=("Malgun Gothic", title_font, "bold"), fill="#64748b", justify=tk.CENTER)
            canvas.create_text((x0 + x1)/2, (y0 + y1)/2 + 15, text="최종 PDF에 목차 페이지와 북마크를 생성하지 않습니다.", font=("Malgun Gothic", item_font), fill="#94a3b8", justify=tk.CENTER)
            return
            
        if mode == "bookmark_only":
            canvas.create_rectangle(x0, y0, x1, y1, fill="#f8fafc", outline="#cbd5e1")
            canvas.create_text((x0 + x1)/2, (y0 + y1)/2 - 20, text="북마크 전용 모드 (TOC 페이지 없음)", font=("Malgun Gothic", title_font, "bold"), fill="#3b82f6", justify=tk.CENTER)
            msg = "최종 PDF 파일 내부에 시각적인 목차 페이지가 삽입되지 않으며,\n아도비 리더 등 PDF 뷰어의 좌측 북마크 패널로만 목차 구조가 생성됩니다."
            canvas.create_text((x0 + x1)/2, (y0 + y1)/2 + 25, text=msg, font=("Malgun Gothic", item_font), fill="#64748b", justify=tk.CENTER)
            return

        if mode == "separate_page":
            title_text = "파일 목록 (File List)"
            filtered_toc = [item for item in self.state.toc_data if item.is_file_header]
            placeholder_items = [
                (None, TOCItem(0, "mock1.pdf", 1, f"{self.state.level_prefixes.get(1, '■ ')}PDF 파일 제목 1", 1, True)),
                (None, TOCItem(0, "mock2.pdf", 1, f"{self.state.level_prefixes.get(1, '■ ')}PDF 파일 제목 2", 3, True)),
            ]
        else: # "toc_bookmark"
            title_text = "목차 (Table of Contents)"
            filtered_toc = self.state.toc_data
            placeholder_items = [
                (None, TOCItem(0, "mock1.pdf", 1, f"{self.state.level_prefixes.get(1, '■ ')}PDF 파일 제목", 1, True)),
                (None, TOCItem(0, "mock1.pdf", 2, f"{self.state.level_prefixes.get(2, '● ')}첫 번째 목차 항목", 2, False)),
                (None, TOCItem(0, "mock1.pdf", 2, f"{self.state.level_prefixes.get(2, '● ')}긴 목차 항목은 방향에 맞게 자동 축약 표시", 3, False)),
                (None, TOCItem(0, "mock1.pdf", 2, f"{self.state.level_prefixes.get(2, '● ')}다음 목차 항목", 4, False)),
            ]

        canvas.create_text(x0 + margin_left * scale, y0 + (margin_top - 30) * scale, text=title_text, anchor=tk.W, font=("Malgun Gothic", title_font, "bold"), fill="#253041")

        items_per_page = max(1, int((page_h - margin_top - margin_bottom) / line_height))
        max_rows = 5 if compact else items_per_page
        preview_items = list(enumerate(filtered_toc[:max_rows]))
        if not preview_items:
            preview_items = placeholder_items

        y = y0 + margin_top * scale
        bottom = y1 - margin_bottom * scale
        sub_item_count = 0
        for loop_idx, (item_index, item) in enumerate(preview_items):
            if y > bottom:
                break

            level = item.level
            title = item.title
            page = item.dest_page
            indent = x0 + (margin_left + max(0, level - 1) * 20) * scale

            if level == 1:
                max_width = max(40, x1 - indent - margin_left * scale)
                text = self._clip_preview_text(str(title), max_width, category_font, bold=True)
                canvas.create_text(indent, y, text=text, anchor=tk.W, font=("Malgun Gothic", category_font, "bold"), fill="#253041")
            else:
                text_x = indent + (move_button_width + move_button_gap) * scale
                max_width = max(40, x1 - text_x - margin_left * scale)
                text = self._clip_preview_text(f"{title} (Page {page})", max_width, item_font)
                text_height = max(item_font + 8, line_height * scale * 0.82)
                row_rect = (indent - 4, y - text_height * 0.58, x1 - margin_left * scale, y + text_height * 0.42)
                button_rect = (indent - 4, y - text_height * 0.58, indent + move_button_width * scale, y + text_height * 0.42)
                
                click_count = self.preview_click_counts.get(self.get_preview_item_key(item), 0) if item_index is not None else 0
                
                # 홀수/짝수 교차 기본 색상 정의 (심리학적 블루/세이지그린 테마 - 대비 강화 버전)
                if sub_item_count % 2 == 0:
                    default_btn_bg = "#e1edfc"
                    default_btn_outline = "#8cb2e6"
                    default_btn_fg = "#1e4a85"
                    default_title_bg = "#edf3fa"
                    default_title_outline = "#cad5e3"
                else:
                    default_btn_bg = "#dfeedf"
                    default_btn_outline = "#94c294"
                    default_btn_fg = "#226322"
                    default_title_bg = "#e8f0e8"
                    default_title_outline = "#c4d5c4"
                
                # [이동 ○] / [이동 ●] 상태 구분 렌더링
                if click_count > 0:
                    btn_text = "[이동 ●]"
                    btn_bg = "#eafaf1"       # 연한 초록 배경
                    btn_outline = "#2ecc71"  # 초록 테두리
                    btn_fg = "#27ae60"       # 진한 초록 텍스트
                    
                    title_bg = "#f2faf5" if click_count == 1 else "#e8f7ee"
                    title_outline = "#c2e9d2"
                else:
                    btn_text = "[이동 ○]"
                    btn_bg = default_btn_bg
                    btn_outline = default_btn_outline
                    btn_fg = default_btn_fg
                    
                    title_bg = default_title_bg
                    title_outline = default_title_outline
                    
                canvas.create_rectangle(*row_rect, fill=title_bg, outline=title_outline)
                canvas.create_rectangle(*button_rect, fill=btn_bg, outline=btn_outline)
                canvas.create_text(indent, y, text=btn_text, anchor=tk.W, font=("Malgun Gothic", item_font), fill=btn_fg)
                canvas.create_text(text_x, y, text=text, anchor=tk.W, font=("Malgun Gothic", item_font), fill="#253041")
                
                if click_count:
                    badge = f"✓ {click_count}회"
                    canvas.create_text(row_rect[2] - 4, y, text=badge, anchor=tk.E, font=("Malgun Gothic", max(7, item_font - 2), "bold"), fill="#27ae60")
                    
                if item_index is not None:
                    canvas._toc_hitboxes.append({
                        "rect": button_rect,
                        "title_rect": row_rect,
                        "item_index": item_index,
                        "item": item,
                    })
                sub_item_count += 1

            y += line_height * scale

        if compact and len(self.state.toc_data) > max_rows:
            canvas.create_text(x0 + margin_left * scale, y1 - 18, text=f"... 외 {len(self.state.toc_data) - max_rows}개 항목", anchor=tk.W, font=("Malgun Gothic", 8), fill="#666666")

        canvas.create_text(x1 - 6, y1 - 6, text="미리보기", anchor=tk.SE, font=("Malgun Gothic", 8), fill="#666666")

    def get_preview_item_key(self, item: TOCItem) -> str:
        return f"{item.file_idx}|{item.dest_page}|{item.title}"

    def handle_toc_preview_click(self, event, canvas):
        for hitbox in getattr(canvas, "_toc_hitboxes", []):
            x0, y0, x1, y1 = hitbox["rect"]
            if x0 <= event.x <= x1 and y0 <= event.y <= y1:
                item = hitbox["item"]
                item_key = self.get_preview_item_key(item)
                self.preview_click_counts[item_key] = self.preview_click_counts.get(item_key, 0) + 1
                
                count = self.preview_click_counts[item_key]
                item_index = hitbox["item_index"]
                self.tree.selection_set(str(item_index))
                self.tree.see(str(item_index))
                self.status_text.set(f"[이동] 클릭 확인: {count}번째 - {item.title}")
                
                self.render_toc_preview(
                    canvas,
                    compact=getattr(canvas, "_toc_compact", True),
                    zoom_scale=getattr(canvas, "_toc_zoom_scale", None),
                )
                # 메인 미리보기 캔버스와 팝업 캔버스를 동기화하여 둘 다 실시간 갱신되도록 처리
                if canvas != getattr(self, "toc_preview_canvas", None):
                    self.update_toc_preview()
                return

    def _clip_preview_text(self, text, max_width, font_size, bold=False):
        font = tkfont.Font(family="Malgun Gothic", size=font_size, weight="bold" if bold else "normal")
        if font.measure(text) <= max_width:
            return text
        ellipsis = "..."
        available = max(1, max_width - font.measure(ellipsis))
        clipped = ""
        for char in text:
            if font.measure(clipped + char) > available:
                break
            clipped += char
        return clipped.rstrip() + ellipsis

    def open_toc_preview_window(self):
        window = tk.Toplevel(self.root)
        window.title("목차 페이지 미리보기 크게 보기")
        window.geometry("920x720")

        top = tk.Frame(window, padx=10, pady=8)
        top.pack(fill=tk.X)
        settings = PDFService.get_toc_layout_settings(self.config)
        tk.Label(top, text=f"{settings['paper_name']} / {settings['orientation_label']} 미리보기", font=("Malgun Gothic", 10, "bold")).pack(side=tk.LEFT)
        tk.Button(top, text="닫기", command=window.destroy, width=10).pack(side=tk.RIGHT)

        frame = tk.Frame(window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        canvas = tk.Canvas(frame, width=860, height=620, bg="#eeeeee", highlightthickness=1, highlightbackground="#bbbbbb")
        vbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        hbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=canvas.xview)
        canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        page_w = settings["page_width"]
        window.update_idletasks()
        scale = max(0.50, (max(canvas.winfo_width(), 860) - 90) / page_w)
        canvas.update_idletasks()
        self.render_toc_preview(canvas, compact=False, zoom_scale=scale)
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

    def select_files(self):
        filetypes = [
            ("지원 파일", "*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff;*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt"),
            ("PDF files", "*.pdf"),
            ("Image files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff"),
            ("Word files", "*.docx;*.doc"),
            ("Excel files", "*.xlsx;*.xls"),
            ("PowerPoint files", "*.pptx;*.ppt"),
            ("All files", "*.*")
        ]
        files = filedialog.askopenfilenames(title="PDF 및 이기종 파일 선택", filetypes=filetypes)
        if not files:
            return
            
        final_files_to_load = []
        custom_filenames = {}
        existing_filenames = [item.filename for item in self.state.pdf_items]
        
        for filepath in files:
            filename = os.path.basename(filepath)
            current_existing = existing_filenames + [custom_filenames.get(fp, os.path.basename(fp)) for fp in final_files_to_load]
            
            if filename in current_existing:
                dialog = DuplicateResolveDialog(self.root, filename)
                self.root.wait_window(dialog)
                choice = dialog.result
                
                if choice == "cancel":
                    continue
                elif choice == "rename":
                    name, ext = os.path.splitext(filename)
                    default_new_name = f"{name}_복사본{ext}"
                    new_name = simpledialog.askstring(
                        "이름 변경",
                        f"'{filename}' 파일의 새로운 이름을 입력하세요:",
                        initialvalue=default_new_name,
                        parent=self.root
                    )
                    if not new_name:
                        continue
                    new_name = new_name.strip()
                    if not new_name:
                        continue
                    if not new_name.lower().endswith(ext.lower()):
                        new_name += ext
                    custom_filenames[filepath] = new_name
                    final_files_to_load.append(filepath)
                else: # "add"
                    name, ext = os.path.splitext(filename)
                    counter = 1
                    new_filename = f"{name}_{counter}{ext}"
                    while new_filename in current_existing:
                        counter += 1
                        new_filename = f"{name}_{counter}{ext}"
                    custom_filenames[filepath] = new_filename
                    final_files_to_load.append(filepath)
            else:
                final_files_to_load.append(filepath)
                
        if not final_files_to_load:
            return
            
        self.set_ui_state("disabled")
        self.load_usecase = PDFLoadUseCase(
            self.state,
            list(final_files_to_load),
            save_preprocessed=self.save_preprocessed.get(),
            compression_mode=self.compression_mode.get(),
            custom_filenames=custom_filenames
        )
        self.load_usecase.run_in_background()
        self.poll_load_progress()

    def poll_load_progress(self):
        if not hasattr(self, "load_usecase") or not self.load_usecase:
            return
        try:
            while True:
                msg = self.load_usecase.progress_queue.get_nowait()
                msg_type = msg["type"]
                percent = msg.get("percent", 0)
                if msg_type == "start":
                    self.progress["value"] = 0
                    self.lbl_percent.config(text="0%")
                    self.status_text.set(msg["message"])
                elif msg_type == "progress":
                    self.progress["value"] = percent
                    self.lbl_percent.config(text=f"{percent}%")
                    self.status_text.set(msg["message"])
                elif msg_type == "ask_continue":
                    result = messagebox.askyesno("전처리 파일 저장 실패", msg["message"])
                    self.load_usecase.response_queue.put(result)
                elif msg_type == "ask_bookmarks":
                    dialog = BookmarkPromptDialog(self.root, msg["filename"])
                    self.load_usecase.response_queue.put(dialog.result)
                elif msg_type == "done":
                    self.progress["value"] = 100
                    self.lbl_percent.config(text="100%")
                    self.status_text.set("파일 불러오기 완료.")
                    
                    new_pdf_items = msg["pdf_items"]
                    new_toc_data = msg["toc_data"]
                    new_page_offsets = msg["page_offsets"]
                    new_page_rotations = msg["page_rotations"]
                    
                    num_new_files = len(new_pdf_items)
                    # Partition new_toc_data by file_idx beforehand to prevent reference sharing index corruption
                    new_toc_by_file = {k: [] for k in range(num_new_files)}
                    for item in new_toc_data:
                        if 0 <= item.file_idx < num_new_files:
                            new_toc_by_file[item.file_idx].append(item)
                            
                    for k in range(num_new_files - 1, -1, -1):
                        self.state.push_undo(list(self.tree.selection()))
                        
                        single_pdf_item = new_pdf_items[k]
                        single_toc_items = new_toc_by_file[k]
                        single_page_count = single_pdf_item.page_count
                        
                        # 1. Shift existing items
                        for item in self.state.pdf_items:
                            item.file_idx += 1
                        for item in self.state.toc_data:
                            item.file_idx += 1
                            
                        # 2. Shift page offsets
                        self.state.page_offsets = [offset + single_page_count for offset in self.state.page_offsets]
                        
                        # 3. Shift page rotations
                        shifted_rotations = { (f + 1, p): r for (f, p), r in self.state.page_rotations.items() }
                        
                        # 4. Shift manually rotated
                        self.state.manually_rotated = [ (f + 1, p) for (f, p) in self.state.manually_rotated ]
                        
                        # 5. Extract new rotations for this file (index 0)
                        new_file_rotations = { (0, p): r for (f, p), r in new_page_rotations.items() if f == k }
                        
                        # 6. Update single item fields
                        single_pdf_item.file_idx = 0
                        for item in single_toc_items:
                            item.file_idx = 0
                            
                        # 7. Prepend to state lists
                        self.state.pdf_items = [single_pdf_item] + self.state.pdf_items
                        self.state.toc_data = single_toc_items + self.state.toc_data
                        self.state.page_offsets = [0] + self.state.page_offsets
                        self.state.page_rotations = { **new_file_rotations, **shifted_rotations }
                    
                    self.state.notify_listeners()
                    
                    self.lbl_files.config(text=f"총 {len(self.state.pdf_items)}개의 파일이 선택되었습니다.")
                    self.status_text.set(f"목차 {len(self.state.toc_data)}개를 불러왔습니다. 병합 전 검수하세요.")
                    self.set_ui_state("normal")
                    self.load_usecase = None
                    return
                elif msg_type == "error":
                    self.progress["value"] = 0
                    self.lbl_percent.config(text="0%")
                    self.status_text.set("파일 로드 중 오류가 발생했습니다.")
                    self.set_ui_state("normal")
                    messagebox.showerror("오류", msg["message"])
                    self.load_usecase = None
                    return
        except queue.Empty:
            pass
        self.root.after(100, self.poll_load_progress)

    def sort_by_column(self, col):
        if not hasattr(self, "_sort_directions"):
            self._sort_directions = {}
        
        reverse = self._sort_directions.get(col, False)
        self._sort_directions[col] = not reverse
        
        self.state.push_undo(list(self.tree.selection()))
        
        if col == "file":
            key_fn = lambda item: item.file_name.lower()
        elif col == "level":
            key_fn = lambda item: item.level
        elif col == "title":
            key_fn = lambda item: item.title.lower()
        elif col == "rotation":
            key_fn = lambda item: self.state.page_rotations.get((item.file_idx, item.dest_page - 1), 0) if item.dest_page > 0 else -1
        elif col == "page":
            key_fn = lambda item: item.dest_page
        else:
            return
            
        self.state.toc_data.sort(key=key_fn, reverse=reverse)
        self.state.notify_listeners()
        
        dir_str = "내림차순" if reverse else "오름차순"
        col_name = {
            "file": "원본 파일",
            "level": "레벨",
            "title": "목차 제목",
            "rotation": "회전",
            "page": "목적 페이지"
        }.get(col, col)
        self.status_text.set(f"'{col_name}' 기준으로 {dir_str} 정렬되었습니다.")

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for idx, item in enumerate(self.state.toc_data):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            if item.dest_page > 0:
                rot = self.state.page_rotations.get((item.file_idx, item.dest_page - 1), 0)
                rot_str = f"{rot}°"
            else:
                rot_str = ""
            self.tree.insert("", "end", iid=str(idx), values=(item.file_name, item.level, item.title, rot_str, item.dest_page), tags=(tag,))
            
        # Update active file count label dynamically
        active_files_count = sum(1 for item in self.state.toc_data if item.is_file_header)
        self.lbl_files.config(text=f"총 {active_files_count}개의 파일이 선택되었습니다.")
        
        self.update_toc_preview()

    def on_double_click(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item_id = selected[0]
        col = self.tree.identify_column(event.x)
        if not col:
            return
        column_id = self.tree.column(col, "id")
        
        idx = int(item_id)
        item = self.state.toc_data[idx]
        
        if column_id == "title":
            current_title = item.title
            new_title = simpledialog.askstring("이름 수정", "새로운 목차 이름을 입력하세요:", initialvalue=current_title, parent=self.root)
            
            if new_title is not None and new_title.strip() != "":
                self.state.update_item_title(idx, new_title.strip())
                self.tree.selection_set(item_id)
        elif column_id == "level":
            if item_id not in selected:
                selected = (item_id,)
            current_level = item.level
            new_level_str = simpledialog.askstring("레벨 수정", "새로운 레벨을 입력하세요 (0 ~ 10):", initialvalue=str(current_level), parent=self.root)
            if new_level_str is not None:
                try:
                    new_level = int(new_level_str)
                    if 0 <= new_level <= 10:
                        self.state.push_undo(list(selected))
                        for sel_id in selected:
                            sel_idx = int(sel_id)
                            self.state.update_item_level(sel_idx, new_level)
                        self.tree.selection_set(*selected)
                    else:
                        messagebox.showwarning("경고", "레벨은 0부터 10 사이의 숫자여야 합니다.")
                except ValueError:
                    messagebox.showerror("오류", "유효한 숫자를 입력해 주세요.")
        elif column_id == "rotation":
            if item.dest_page > 0:
                self.open_rotation_dialog(idx)

    def open_rotation_dialog(self, toc_idx: int):
        # Collect all selected TOC items with valid destination pages
        selected_iids = self.tree.selection()
        toc_indices = [int(iid) for iid in selected_iids]
        if toc_idx not in toc_indices:
            toc_indices.append(toc_idx)
            
        unique_targets = []
        for idx in toc_indices:
            if 0 <= idx < len(self.state.toc_data):
                item = self.state.toc_data[idx]
                if item.dest_page > 0:
                    pair = (item.file_idx, item.dest_page - 1)
                    if pair not in unique_targets:
                        unique_targets.append(pair)
                        
        if not unique_targets:
            return
            
        # Determine initial rotation
        first_file, first_page = unique_targets[0]
        initial_rot = self.state.page_rotations.get((first_file, first_page), 0)
        
        # Check if all selected targets have the same rotation
        all_same = True
        for f_idx, p_idx in unique_targets:
            if self.state.page_rotations.get((f_idx, p_idx), 0) != initial_rot:
                all_same = False
                break
                
        if not all_same:
            clicked_item = self.state.toc_data[toc_idx]
            if clicked_item.dest_page > 0:
                initial_rot = self.state.page_rotations.get((clicked_item.file_idx, clicked_item.dest_page - 1), 0)
            else:
                initial_rot = 0
                
        dialog = tk.Toplevel(self.root)
        dialog.title("페이지 회전 설정")
        dialog.geometry("480x380")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog relative to main window
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 150, self.root.winfo_rooty() + 100))
        
        # Bottom Buttons Area (Pack first at bottom so it is always visible)
        btn_frame = tk.Frame(dialog, padx=15, pady=10, bg="#f1f5f9")
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Main Frame with padding (Pack after bottom frame to fill remaining space)
        main_frame = tk.Frame(dialog, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Information Label
        if len(unique_targets) == 1:
            clicked_item = self.state.toc_data[toc_idx]
            info_text = f"원본: {clicked_item.file_name}\n페이지: {clicked_item.dest_page} 페이지"
        else:
            info_text = f"선택된 대상: 총 {len(unique_targets)}개의 페이지 회전 일괄 설정"
            
        info_lbl = tk.Label(main_frame, text=info_text, font=("Malgun Gothic", 10, "bold"), anchor=tk.W, justify=tk.LEFT)
        info_lbl.pack(fill=tk.X, pady=(0, 12))
        
        # Body frame for controls & preview
        body_frame = tk.Frame(main_frame)
        body_frame.pack(fill=tk.BOTH, expand=True)
        
        left_frame = tk.Frame(body_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        
        right_frame = tk.Frame(body_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Preview Canvas
        canvas = tk.Canvas(right_frame, width=240, height=220, bg="#f8f9fa", highlightthickness=1, highlightbackground="#cbd5e1")
        canvas.pack(fill=tk.BOTH, expand=True)
        
        selected_rot = tk.IntVar(value=initial_rot)
        
        def update_preview(*args):
            rot = selected_rot.get()
            canvas.delete("all")
            cx, cy = 120, 110
            
            # Dimensions based on rotation
            if rot in (0, 180):
                pw, ph = 100, 140
            else:
                pw, ph = 140, 100
                
            # Draw shadow
            canvas.create_rectangle(cx - pw/2 + 4, cy - ph/2 + 4, cx + pw/2 + 4, cy + ph/2 + 4, outline="", fill="#e2e8f0")
            # Draw page outline
            canvas.create_rectangle(cx - pw/2, cy - ph/2, cx + pw/2, cy + ph/2, fill="#ffffff", outline="#475569", width=2)
            
            # Draw arrow pointing to top and text
            arrow_color = "#3b82f6"
            text_color = "#1e3a8a"
            
            if rot == 0:
                canvas.create_line(cx, cy + 25, cx, cy - 25, arrow=tk.LAST, width=3, fill=arrow_color)
                canvas.create_text(cx, cy - 40, text="위쪽 (0°)", font=("Malgun Gothic", 9, "bold"), fill=text_color)
            elif rot == 90:
                canvas.create_line(cx - 25, cy, cx + 25, cy, arrow=tk.LAST, width=3, fill=arrow_color)
                canvas.create_text(cx + 45, cy, text="위쪽\n(90°)", font=("Malgun Gothic", 9, "bold"), fill=text_color, justify=tk.CENTER)
            elif rot == 180:
                canvas.create_line(cx, cy - 25, cx, cy + 25, arrow=tk.LAST, width=3, fill=arrow_color)
                canvas.create_text(cx, cy + 40, text="위쪽 (180°)", font=("Malgun Gothic", 9, "bold"), fill=text_color)
            elif rot == 270:
                canvas.create_line(cx + 25, cy, cx - 25, cy, arrow=tk.LAST, width=3, fill=arrow_color)
                canvas.create_text(cx - 45, cy, text="위쪽\n(270°)", font=("Malgun Gothic", 9, "bold"), fill=text_color, justify=tk.CENTER)

        # Radio button choices
        choices = [
            ("0도 (원래 방향)", 0),
            ("90도 (우회전)", 90),
            ("180도 (상하반전)", 180),
            ("270도 (좌회전)", 270)
        ]
        
        for label_text, angle in choices:
            rb = tk.Radiobutton(left_frame, text=label_text, variable=selected_rot, value=angle, command=update_preview, font=("Malgun Gothic", 9))
            rb.pack(anchor=tk.W, pady=6)
            
        update_preview()
        
        # Populate buttons on the previously packed btn_frame
        def on_confirm():
            new_rot = selected_rot.get()
            
            # Check if any rotation is modified
            any_changed = False
            for f_idx, p_idx in unique_targets:
                if self.state.page_rotations.get((f_idx, p_idx), 0) != new_rot:
                    any_changed = True
                    break
                    
            if any_changed:
                # Save undo point with current tree selection
                self.state.push_undo(list(self.tree.selection()))
                for f_idx, p_idx in unique_targets:
                    self.state.page_rotations[(f_idx, p_idx)] = new_rot
                    if (f_idx, p_idx) not in self.state.manually_rotated:
                        self.state.manually_rotated.append((f_idx, p_idx))
                self.state.notify_listeners()
            dialog.destroy()
            
        confirm_btn = tk.Button(btn_frame, text="확인", command=on_confirm, width=10, bg="#3b82f6", fg="white", font=("Malgun Gothic", 9, "bold"))
        confirm_btn.pack(side=tk.RIGHT, padx=5)
        
        cancel_btn = tk.Button(btn_frame, text="취소", command=dialog.destroy, width=10, font=("Malgun Gothic", 9))
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def delete_item(self):
        selected = self.tree.selection()
        if not selected:
            return
        indices = [int(x) for x in selected]
        self.state.push_undo(selected)
        self.state.delete_items(indices)

    def move_up(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        self.state.push_undo(selected)
        if self.state.move_item_up(idx):
            self.tree.selection_set(str(idx-1))

    def move_down(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        self.state.push_undo(selected)
        if self.state.move_item_down(idx):
            self.tree.selection_set(str(idx+1))

    def open_batch_edit_dialog(self):
        if not self.state.toc_data:
            messagebox.showwarning("경고", "변경할 목차 항목이 없습니다.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("목차명 일괄 변경")
        dialog.geometry("720x460")
        dialog.transient(self.root)

        mode_var = tk.StringVar(value="search")
        action_var = tk.StringVar(value="replace")
        scope_var = tk.StringVar(value="all")
        case_var = tk.BooleanVar(value=False)
        find_var = tk.StringVar()
        value_var = tk.StringVar()
        position_var = tk.StringVar(value="1")
        length_var = tk.StringVar(value="1")

        form = tk.Frame(dialog, padx=12, pady=12)
        form.pack(fill=tk.BOTH, expand=True)

        tk.Label(form, text="편집 기준", font=("Malgun Gothic", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=5)
        tk.Radiobutton(form, text="검색어 기준", variable=mode_var, value="search").grid(row=0, column=1, sticky=tk.W)
        tk.Radiobutton(form, text="앞에서 N번째", variable=mode_var, value="from_start").grid(row=0, column=2, sticky=tk.W)
        tk.Radiobutton(form, text="뒤에서 N번째", variable=mode_var, value="from_end").grid(row=0, column=3, sticky=tk.W)

        tk.Label(form, text="검색어", font=("Malgun Gothic", 9, "bold")).grid(row=1, column=0, sticky=tk.W, pady=5)
        tk.Entry(form, textvariable=find_var).grid(row=1, column=1, columnspan=3, sticky=tk.EW, pady=5)

        tk.Label(form, text="위치/글자 수", font=("Malgun Gothic", 9, "bold")).grid(row=2, column=0, sticky=tk.W, pady=5)
        position_frame = tk.Frame(form)
        position_frame.grid(row=2, column=1, columnspan=3, sticky=tk.W, pady=5)
        tk.Label(position_frame, text="N=").pack(side=tk.LEFT)
        tk.Spinbox(position_frame, from_=1, to=999, width=6, textvariable=position_var).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(position_frame, text="보호 삭제/교체 글자 수=").pack(side=tk.LEFT)
        tk.Spinbox(position_frame, from_=1, to=999, width=6, textvariable=length_var).pack(side=tk.LEFT)

        tk.Label(form, text="변경/삽입할 내용", font=("Malgun Gothic", 9, "bold")).grid(row=3, column=0, sticky=tk.W, pady=5)
        tk.Entry(form, textvariable=value_var).grid(row=3, column=1, columnspan=3, sticky=tk.EW, pady=5)

        tk.Label(form, text="작업", font=("Malgun Gothic", 9, "bold")).grid(row=4, column=0, sticky=tk.W, pady=5)
        tk.Radiobutton(form, text="교체", variable=action_var, value="replace").grid(row=4, column=1, sticky=tk.W)
        tk.Radiobutton(form, text="삭제", variable=action_var, value="delete").grid(row=4, column=2, sticky=tk.W)
        tk.Radiobutton(form, text="앞에 삽입", variable=action_var, value="insert_before").grid(row=5, column=1, sticky=tk.W)
        tk.Radiobutton(form, text="뒤에 삽입", variable=action_var, value="insert_after").grid(row=5, column=2, sticky=tk.W)
        tk.Radiobutton(form, text="위치에 삽입", variable=action_var, value="insert_at", command=lambda: mode_var.set("from_start")).grid(row=5, column=3, sticky=tk.W)

        tk.Label(form, text="범위", font=("Malgun Gothic", 9, "bold")).grid(row=6, column=0, sticky=tk.W, pady=5)
        tk.Radiobutton(form, text="전체 목차", variable=scope_var, value="all").grid(row=6, column=1, sticky=tk.W)
        tk.Radiobutton(form, text="선택 항목만", variable=scope_var, value="selected").grid(row=6, column=2, sticky=tk.W)
        tk.Checkbutton(form, text="대소문자 구분", variable=case_var).grid(row=7, column=1, columnspan=2, sticky=tk.W, pady=5)

        tk.Label(form, text="목차 선택", font=("Malgun Gothic", 9, "bold")).grid(row=8, column=0, sticky=tk.W, pady=5)
        tk.Button(form, text="전체 선택", command=self.select_all_toc_items, width=12).grid(row=8, column=1, sticky=tk.W, pady=5)
        tk.Button(form, text="전체 해제", command=self.clear_toc_selection, width=12).grid(row=8, column=2, sticky=tk.W, pady=5)
        tk.Button(form, text="검색 항목 선택", command=lambda: self.select_matching_toc_items(find_var.get(), case_var.get()), width=14).grid(row=8, column=3, sticky=tk.W, pady=5)

        tk.Label(
            form,
            text="위치 기준: 삽입은 N=1일 때 맨 앞/맨 뒤 문자를 보존하며 바깥쪽에 붙입니다. 삭제/교체도 N=1은 첫/마지막 문자를 보호하고 안쪽부터 적용합니다.",
            fg="#555555",
            font=("Malgun Gothic", 8),
        ).grid(row=9, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))

        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, weight=1)
        form.columnconfigure(3, weight=1)

        def apply_batch_edit():
            changed = self.state.batch_edit_titles(
                search=find_var.get(),
                value=value_var.get(),
                action=action_var.get(),
                scope=scope_var.get(),
                case_sensitive=case_var.get(),
                mode=mode_var.get(),
                position=position_var.get(),
                length=length_var.get(),
                selected_iids=list(self.tree.selection())
            )
            
            if changed is None:
                messagebox.showwarning("경고", "입력값을 다시 확인해주세요.")
                return
            if changed == 0:
                messagebox.showinfo("알림", "변경된 목차명이 없습니다.")
                return
                
            messagebox.showinfo("완료", f"목차명 {changed}개 항목을 변경했습니다.")
            selected_scope = list(self.tree.selection())
            self.tree.selection_remove(self.tree.selection())
            if scope_var.get() == "selected":
                for idx in selected_scope:
                    self.tree.selection_add(idx)
            else:
                for idx in range(min(100, len(self.state.toc_data))):
                    self.tree.selection_add(str(idx))

        button_frame = tk.Frame(dialog, padx=12, pady=8)
        button_frame.pack(fill=tk.X)
        tk.Button(button_frame, text="이전 작업 취소", command=self.undo_last_toc_edit, width=14).pack(side=tk.LEFT, padx=4)
        tk.Button(button_frame, text="적용", command=apply_batch_edit, width=12, bg="#4CAF50", fg="white").pack(side=tk.RIGHT, padx=4)
        tk.Button(button_frame, text="취소", command=dialog.destroy, width=12).pack(side=tk.RIGHT, padx=4)
        dialog.lift()

    def open_prefix_settings_dialog(self):
        PrefixSettingsDialog(self.root, self, self.state)

    def select_all_toc_items(self):
        self.tree.selection_set([str(idx) for idx in range(len(self.state.toc_data))])
        self.status_text.set(f"목차 {len(self.state.toc_data)}개 항목을 선택했습니다.")

    def clear_toc_selection(self):
        self.tree.selection_remove(self.tree.selection())
        self.status_text.set("목차 선택을 해제했습니다.")

    def undo_last_toc_edit(self):
        result = self.state.pop_undo()
        if result is None:
            messagebox.showinfo("알림", "취소할 일괄 변경 작업이 없습니다.")
            return
        previous_selection = result[-1]
        valid_selection = [item_id for item_id in previous_selection if self.tree.exists(item_id)]
        if valid_selection:
            self.tree.selection_set(valid_selection)
            self.tree.see(valid_selection[0])
        self.status_text.set("이전 작업을 취소했습니다.")

    def select_matching_toc_items(self, search, case_sensitive=False):
        search = search.strip()
        if not search:
            messagebox.showwarning("경고", "선택할 항목의 검색어를 입력하세요.")
            return

        if case_sensitive:
            matcher = lambda title: search in title
        else:
            folded_search = search.casefold()
            matcher = lambda title: folded_search in title.casefold()

        matches = [str(idx) for idx, item in enumerate(self.state.toc_data) if matcher(item.title)]
        self.tree.selection_remove(self.tree.selection())
        if matches:
            self.tree.selection_set(matches)
            self.tree.see(matches[0])
        self.status_text.set(f"검색어와 일치하는 목차 {len(matches)}개 항목을 선택했습니다.")

    def merge_pdfs(self):
        if not self.state.pdf_items:
            messagebox.showwarning("경고", "병합할 파일이 선택되지 않았습니다.")
            return
            
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="최종_병합문서.pdf", title="저장할 위치 선택", filetypes=[("PDF files", "*.pdf")])
        if not save_path:
            return
            
        self._sync_config()
        self.set_ui_state("disabled")
        self.usecase = PDFMergeUseCase(self.state, self.config)
        self.usecase.run_in_background(save_path)
        self.poll_progress_queue()

    def poll_progress_queue(self):
        if not self.usecase:
            return
        try:
            while True:
                msg = self.usecase.progress_queue.get_nowait()
                msg_type = msg["type"]
                percent = msg.get("percent", 0)
                if msg_type == "start":
                    self.progress["value"] = 0
                    self.lbl_percent.config(text="0%")
                    self.status_text.set(msg["message"])
                elif msg_type == "progress":
                    self.progress["value"] = percent
                    self.lbl_percent.config(text=f"{percent}%")
                    self.status_text.set(msg["message"])
                elif msg_type == "ask_continue":
                    result = messagebox.askyesno("전처리 파일 저장 실패", msg["message"])
                    self.usecase._response_queue.put(result)
                elif msg_type == "done":
                    self.progress["value"] = 100
                    self.lbl_percent.config(text="100%")
                    self.status_text.set("완료되었습니다.")
                    self.set_ui_state("normal")
                    messagebox.showinfo("완료", msg["message"])
                    self.usecase = None
                    return
                elif msg_type == "error":
                    self.progress["value"] = 0
                    self.lbl_percent.config(text="0%")
                    self.status_text.set("오류가 발생했습니다.")
                    self.set_ui_state("normal")
                    messagebox.showerror("오류", msg["message"])
                    self.usecase = None
                    return
        except queue.Empty:
            pass
        self.root.after(100, self.poll_progress_queue)


# =====================================================================
# 5. TEST RUNNER LAYER (E2E Integration Testing)
# =====================================================================

def run_e2e_tests() -> int:
    print("==================================================")
    print("PDF Master Merge E2E Test Suite Starting...")
    print("==================================================")
    
    # 임시 파일 경로를 상위 .temp_backups 디렉토리로 변경
    backup_dir = get_backup_dir()
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir, exist_ok=True)
        
    mock_files = [
        os.path.join(backup_dir, "test_mock1.pdf"),
        os.path.join(backup_dir, "test_mock2.pdf")
    ]
    output_file = os.path.join(backup_dir, "test_output.pdf")
    
    try:
        # 1. Create mock images and PDFs
        print("[1/5] Generating mock PDFs with images...")
        img_io1 = BytesIO()
        Image.new("RGB", (800, 800), color="red").save(img_io1, format="JPEG")
        img_bytes1 = img_io1.getvalue()
        
        img_io2 = BytesIO()
        Image.new("RGB", (1000, 1000), color="blue").save(img_io2, format="JPEG")
        img_bytes2 = img_io2.getvalue()

        # Doc 1 (2 pages, with images and bookmarks)
        doc1 = fitz.open()
        p1 = doc1.new_page(width=595, height=842)
        p1.insert_text((50, 50), "Document 1 - Page 1", fontsize=20)
        p1.insert_image(fitz.Rect(100, 100, 400, 400), stream=img_bytes1)
        
        p2 = doc1.new_page(width=595, height=842)
        p2.insert_text((50, 50), "Document 1 - Page 2", fontsize=20)
        doc1.set_toc([[1, "Doc1 Bookmark Page 1", 1], [2, "Doc1 Sub-bookmark Page 2", 2]])
        doc1.save(mock_files[0])
        doc1.close()

        # Doc 2 (1 page, with image)
        doc2 = fitz.open()
        p3 = doc2.new_page(width=595, height=842)
        p3.insert_text((50, 50), "Document 2 - Page 1", fontsize=20)
        p3.insert_image(fitz.Rect(100, 100, 500, 500), stream=img_bytes2)
        doc2.save(mock_files[1])
        doc2.close()
        
        print(f"Mock files created: {mock_files}")
        
        # 2. Test AppState
        print("[2/5] Initializing AppState and loading PDFs...")
        state = AppState()
        state.set_pdfs(mock_files)
        
        assert len(state.pdf_items) == 2, "PDF items count mismatch"
        assert state.pdf_items[0].page_count == 2, "PDF 1 page count mismatch"
        assert state.pdf_items[1].page_count == 1, "PDF 2 page count mismatch"
        
        print(f"Initial TOC extracted: {[item.title for item in state.toc_data]}")
        
        # 3. Test TOC Edits
        print("[3/5] Testing AppState state mutations...")
        state.update_item_title(1, "Modified Bookmark 1")
        assert state.toc_data[1].title == "Modified Bookmark 1", "Update title failed"
        
        changed = state.batch_edit_titles(
            search="Bookmark",
            value="Target",
            action="replace",
            scope="all",
            case_sensitive=False,
            mode="search",
            position="1",
            length="1",
            selected_iids=[]
        )
        assert changed == 2, f"Batch edit search-replace failed, changed count: {changed}"
        assert "Target" in state.toc_data[1].title, "Batch edit replace value mismatch"
        
        assert state.move_item_down(0), "Move down failed"
        assert state.toc_data[1].title.startswith("■ test_mock1"), "Item order mismatch after move down"
        
        state.pop_undo()
        assert state.toc_data[0].title.startswith("■ test_mock1"), "Undo failed to restore order"
        
        # Test page rotation state and undo
        assert state.page_rotations[(0, 0)] == 0, "Initial page rotation should be 0"
        state.push_undo([])
        state.page_rotations[(0, 0)] = 90
        state.manually_rotated.append((0, 0))
        assert state.page_rotations[(0, 0)] == 90, "Page rotation change failed"
        
        state.pop_undo()
        assert state.page_rotations[(0, 0)] == 0, "Undo failed to restore page rotation"
        assert (0, 0) not in state.manually_rotated, "Undo failed to restore manually_rotated list"
        
        # Test bulk rotation setting simulation
        state.push_undo([])
        targets = [(0, 0), (1, 0)]
        for f, p in targets:
            state.page_rotations[(f, p)] = 180
            state.manually_rotated.append((f, p))
        assert state.page_rotations[(0, 0)] == 180, "Bulk rotation for page (0,0) failed"
        assert state.page_rotations[(1, 0)] == 180, "Bulk rotation for page (1,0) failed"
        
        state.pop_undo()
        assert state.page_rotations[(0, 0)] == 0, "Bulk rotation undo failed for page (0,0)"
        assert state.page_rotations[(1, 0)] == 0, "Bulk rotation undo failed for page (1,0)"
        assert (0, 0) not in state.manually_rotated, "Bulk rotation undo failed to clear manually_rotated page (0,0)"
        
        # Set page (0,0) to 90 degrees and page (1,0) to 180 degrees for final merge test
        state.page_rotations[(0, 0)] = 90
        state.manually_rotated.append((0, 0))
        state.page_rotations[(1, 0)] = 180
        state.manually_rotated.append((1, 0))
        
        # 4. Test MergeUseCase
        print("[4/5] Running MergeUseCase (balanced compression)...")
        config = MergeConfig(compression_mode="balanced", filename_mode="toc_bookmark")
        usecase = PDFMergeUseCase(state, config)
        
        # Execute synchronously for tests
        usecase._execute(output_file)
        
        # Poll progress messages
        messages = []
        while not usecase.progress_queue.empty():
            messages.append(usecase.progress_queue.get())
            
        for msg in messages:
            if msg["type"] == "error":
                raise RuntimeError(f"Merge execution error: {msg['message']}")
                
        print("Merge process completed without errors.")
        
        # 5. Validate output PDF
        print("[5/5] Validating merged PDF file...")
        assert os.path.exists(output_file), "Output PDF file not created"
        
        final_doc = fitz.open(output_file)
        final_page_count = len(final_doc)
        print(f"Merged PDF Page Count: {final_page_count}")
        assert final_page_count == 4, f"Final page count mismatch, expected 4 but got {final_page_count}"
        
        # Verify page rotation is applied to page index 1 (90 degrees) and page index 3 (180 degrees) of final_doc
        assert final_doc[1].rotation == 90, f"Page rotation not applied on page 1, expected 90 but got {final_doc[1].rotation}"
        assert final_doc[3].rotation == 180, f"Page rotation not applied on page 3, expected 180 but got {final_doc[3].rotation}"
        
        # Verify page dimensions taking rotation into account (aspect ratios must swap dynamically for 90/270 degrees)
        assert final_doc[1].rect.width == 842.0 and final_doc[1].rect.height == 595.0, f"Page 1 dimensions mismatch: {final_doc[1].rect}"
        assert final_doc[3].rect.width == 595.0 and final_doc[3].rect.height == 842.0, f"Page 3 dimensions mismatch: {final_doc[3].rect}"
        
        final_toc = final_doc.get_toc()
        print(f"Final PDF bookmarks: {final_toc}")
        assert len(final_toc) == 4, "Bookmarks count mismatch"
        assert final_toc[0][1] == "■ test_mock1", "First bookmark name mismatch"
        assert final_toc[0][2] == 2, "First bookmark page mismatch"
        
        # Validate navigation widgets exist on all pages (forces hand cursor)
        for page_idx in range(final_page_count):
            widgets = list(final_doc[page_idx].widgets())
            has_home = any("btn_nav_home" in (w.field_name or "") for w in widgets)
            assert has_home, f"Page {page_idx} is missing HOME navigation widget"
            
        final_doc.close()
        
        # 6. Test Heterogeneous files loading (Image + Office)
        print("[6/6] Testing heterogeneous file loading and conversion...")
        mock_png = os.path.join(backup_dir, "test_mock_img.png")
        Image.new("RGB", (300, 300), color="green").save(mock_png)
        
        mock_jpg = os.path.join(backup_dir, "test_mock_img.jpg")
        Image.new("RGB", (300, 300), color="blue").save(mock_jpg)
        
        mock_bmp = os.path.join(backup_dir, "test_mock_img.bmp")
        Image.new("RGB", (300, 300), color="red").save(mock_bmp)
        
        mock_docx = os.path.abspath(os.path.join(backup_dir, "test_mock_doc.docx"))
        def create_mock_docx_thread():
            nonlocal mock_docx
            import pythoncom
            pythoncom.CoInitialize()
            import win32com.client
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Add()
                word.Selection.TypeText("This is a mock DOCX file created for E2E testing.")
                doc.SaveAs(mock_docx, FileFormat=16) # 16 is docx
                doc.Close(0)
                word.Quit()
            except Exception as e:
                print(f"Word mock creation error: {str(e)}")
                mock_docx = ""
            finally:
                pythoncom.CoUninitialize()

        t = threading.Thread(target=create_mock_docx_thread)
        t.daemon = True
        t.start()
        t.join(timeout=5.0)
        if t.is_alive():
            print("Word mock creation timed out (Word COM may be blocked or unactivated). Skipping Word E2E test.")
            mock_docx = ""
            
        load_targets = [mock_files[0], mock_png, mock_jpg, mock_bmp]
        if mock_docx:
            load_targets.append(mock_docx)
            
        state2 = AppState()
        state2.set_pdfs(load_targets)
        
        assert len(state2.pdf_items) == len(load_targets), "Heterogeneous files load count mismatch"
        assert state2.pdf_items[1].filename == "test_mock_img.png", "PNG filename mismatch"
        assert state2.pdf_items[1].page_count == 1, "PNG PDF page count mismatch"
        assert state2.pdf_items[2].filename == "test_mock_img.jpg", "JPG filename mismatch"
        assert state2.pdf_items[2].page_count == 1, "JPG PDF page count mismatch"
        assert state2.pdf_items[3].filename == "test_mock_img.bmp", "BMP filename mismatch"
        assert state2.pdf_items[3].page_count == 1, "BMP PDF page count mismatch"
        
        if mock_docx:
            assert state2.pdf_items[4].filename == "test_mock_doc.docx", "Word filename mismatch"
            assert state2.pdf_items[4].page_count >= 1, "Word PDF page count mismatch"
            
        # Test merging this heterogeneous collection to verify E2E heterogeneous merge works
        print("Testing heterogeneous merge (including JPG, PNG, BMP, PDF, and DOCX)...")
        het_output = os.path.join(backup_dir, "test_het_output.pdf")
        het_config = MergeConfig(compression_mode="strong", filename_mode="toc_bookmark")
        het_usecase = PDFMergeUseCase(state2, het_config)
        het_usecase._execute(het_output)
        
        # Verify het_output pages
        assert os.path.exists(het_output), "Heterogeneous output PDF not created"
        het_doc = fitz.open(het_output)
        print(f"Heterogeneous merged PDF page count: {len(het_doc)}")
        
        # Expected page count = TOC (1) + PDF1 (2) + PNG (1) + JPG (1) + BMP (1) + DOCX (if mock_docx, page_count of docx, else 0)
        expected_pages = 1 + 2 + 1 + 1 + 1
        if mock_docx:
            expected_pages += state2.pdf_items[4].page_count
            
        assert len(het_doc) == expected_pages, f"Heterogeneous page count mismatch, expected {expected_pages} but got {len(het_doc)}"
        het_doc.close()
        os.remove(het_output)
            
        # Clean up heterogeneous mock files
        for f in [mock_png, mock_jpg, mock_bmp, mock_docx]:
            if f and os.path.exists(f):
                os.remove(f)
                
        print("Merged PDF and heterogeneous load validation passed successfully.")
        
        # Clean up
        print("Cleaning up temporary test files...")
        for f in mock_files + [output_file]:
            if os.path.exists(f):
                os.remove(f)
                
        print("==================================================")
        print("ALL E2E TESTS PASSED SUCCESSFULLY! (INTEGRITY SECURED)")
        print("==================================================")
        return 0
        
    except Exception as e:
        print(f"\n[!] E2E TEST FAILED: {str(e)}")
        # Cleanup
        for f in mock_files + [output_file]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except Exception:
                    pass
        import traceback
        traceback.print_exc()
        return 1


# =====================================================================
# 6. RUNNER INITIALIZATION
# =====================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF Master Merge CLI & GUI")
    parser.add_argument("--test", "--e2e", action="store_true", help="Run self-contained end-to-end integration tests")
    args = parser.parse_args()
    
    if args.test:
        sys.exit(run_e2e_tests())
    else:
        root = tk.Tk()
        
        # 강제로 윈도우를 최상단으로 올리고 포커스를 획득하여 작업표시줄에 갇히는 현상 방지
        root.lift()
        root.attributes("-topmost", True)
        root.attributes("-topmost", False)
        root.focus_force()
        
        app = PDFMergeApp(root)
        root.protocol("WM_DELETE_WINDOW", lambda: {cleanup_temp_files(), root.destroy()})
        root.mainloop()
