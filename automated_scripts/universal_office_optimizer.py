"""
================================================================================
 [OPTIMIZER] 범용 오피스 최적화 및 통합 도구 (Universal Office Optimizer) v35.4.18
================================================================================
 - 아키텍처: Clean Layer Architecture (Domain / Application / Infrastructure / UI)
 - 디자인 패턴: Domain-Driven Design (DDD), Repository, Gateway
 - 업데이트 내역 (v35.4.3): 
    1. 통합 병합(Merging) 프로세스 직후 '최적화' 및 '심층 정제' 엔진 자동 연동 파이프라인 구축
    2. 필터 확장자 개별 선택 및 누적 자동 입력 인터페이스(Cumulative UI) 탑재
    3. 병합 결과물의 용량 비대화 현상 선제적 차단 및 리소스 해제 무결성 보증
================================================================================
"""

import os
import shutil
import zipfile
import tempfile
import threading
import time
import re
import gc
import sys
import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple, Dict, Set, Any, cast
from pathlib import Path

# Enforce UTF-8 for process communication
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass

# Dependency: Pillow
try:
    from PIL import Image
except ImportError:
    print("[ERROR] Pillow library is required. Run: pip install Pillow")
    sys.exit(1)

# 지원 확장자 상수
SUPPORTED_EXTS = ['.pptx', '.ppt', '.pptm', '.xlsx', '.xls', '.xlsm', '.xlsb', '.docx', '.doc']
FAMILY_MAP = {
    'PPT': ['.pptx', '.ppt', '.pptm'],
    'EXCEL': ['.xlsx', '.xls', '.xlsm', '.xlsb'],
    'WORD': ['.docx', '.doc']
}
HIGH_EXT_MAP = {'PPT': '.pptx', 'EXCEL': '.xlsx', 'WORD': '.docx'}

# ═══════════════════════════════════════════════════════════════════════════════
# 1. INFRASTRUCTURE LAYER (인프라: 외부 시스템 연동 - COM, FileSystem)
# ═══════════════════════════════════════════════════════════════════════════════

class OfficeAppGateway:
    """MS Office COM 애플리케이션의 생명주기와 세부 조작을 담당하는 게이트웨이"""
    
    def __init__(self):
        self.xl_app = None
        self.ppt_app = None
        self.wd_app = None

    def start_apps(self, needs_excel=False, needs_ppt=False, needs_word=False):
        """[v35.2.0] 필요한 오피스 앱을 지능적으로 바인딩 (Dynamic & Deep Polling)"""
        pythoncom.CoInitialize()
        from win32com.client import DispatchEx
        from win32com.client.dynamic import Dispatch
        
        try:
            if needs_excel and not self.xl_app:
                # 3-Phase Resiliency: DispatchEx -> Dispatch -> Fallback
                try: self.xl_app = DispatchEx("Excel.Application")
                except:
                    try: self.xl_app = Dispatch("Excel.Application")
                    except: self.xl_app = None

                xl = self.xl_app
                if xl:
                    try:
                        xl.Visible = False
                        xl.DisplayAlerts = False
                        xl.Interactive = False # [v35.4.18] 사용자 입력 차단 및 프리징 방지
                        xl.AutomationSecurity = 3 # ForceDisable
                    except: pass
                
            if needs_ppt and not self.ppt_app:
                try: self.ppt_app = DispatchEx("PowerPoint.Application")
                except:
                    try: self.ppt_app = Dispatch("PowerPoint.Application")
                    except: self.ppt_app = None

                ppt = self.ppt_app
                if ppt:
                    try: ppt.DisplayAlerts = 1 # ppAlertsNone
                    except: pass
                    try: ppt.Visible = 0 
                    except:
                        try: ppt.WindowState = 2 # Minimized
                        except: pass
                    
            if needs_word and not self.wd_app:
                try: self.wd_app = DispatchEx("Word.Application")
                except:
                    try: self.wd_app = Dispatch("Word.Application")
                    except: self.wd_app = None

                wd = self.wd_app
                if wd:
                    try:
                        wd.Visible = False
                        wd.DisplayAlerts = 0 
                    except: pass
                
        except Exception as e:
            return f"COM 가동 실패: {e}"
        return None

    def close_apps(self):
        """가동 중인 모든 오피스 앱을 안전하게 종료"""
        if self.xl_app:
            try: self.xl_app.Quit()
            except: pass
            self.xl_app = None
        if self.ppt_app:
            try: self.ppt_app.Quit()
            except: pass
            self.ppt_app = None
        if self.wd_app:
            try: self.wd_app.Quit()
            except: pass
            self.wd_app = None
        pythoncom.CoUninitialize()

    def kill_zombie_processes(self):
        """잔존 프로세스 강제 종료 (무결성 보장)"""
        targets = ["EXCEL.EXE", "POWERPNT.EXE", "WINWORD.EXE"]
        try:
            import psutil
            for p in psutil.process_iter(['name']):
                if p.info['name'] and p.info['name'].upper() in targets:
                    try: p.kill()
                    except: pass
        except:
            # Fallback to taskkill
            for target in targets:
                os.system(f"taskkill /f /im {target} >nul 2>&1")
            
            # [v35.4.13] OS가 리소스를 완전히 해제할 시간을 확보하기 위한 짧은 대기
            time.sleep(1.0)


class FileSystemRepository:
    """파일 스캔, 복사, 임시 디렉토리 관리를 담당"""
    
    def __init__(self):
        self.temp_dirs = []

    @staticmethod
    def _safe(path: str) -> str:
        """[v35.4.2] Windows MAX_PATH 제한을 우회하기 위한 경로 정규화 (UNC 완벽 대응)"""
        if not path: return ""
        try:
            # 1. 경로 정규화 (중복 슬래시 등 제거)
            norm_p = os.path.normpath(os.path.abspath(path))
            
            # 2. Windows에서 이미 접두사가 있거나 길이가 짧으면 그대로 반환 (COM 호환성)
            if os.name != 'nt': return norm_p
            if norm_p.startswith('\\\\?\\'): return norm_p
            
            # 3. 240자 이상이거나 UNC 경로인 경우 접두사 부여
            if len(norm_p) > 240 or norm_p.startswith('\\\\'):
                if norm_p.startswith('\\\\'):
                    return '\\\\?\\UNC\\' + norm_p[2:]
                return '\\\\?\\' + norm_p
            return norm_p
        except: return path

    @staticmethod
    def shorten_path(path: str, max_len: int = 250) -> str:
        """[v35.4.2] 경로가 너무 길 경우 파일명을 강제 단축하며 확장자 보존"""
        if len(path) <= max_len: return path
        
        d = os.path.dirname(path)
        f = os.path.basename(path)
        name, ext = os.path.splitext(f)
        
        # 경로 부분이 이미 너무 길면 어쩔 수 없으나, 파일명 위주로 깎음
        budget = max_len - len(d) - len(ext) - 8
        if budget > 0:
            new_name = name[:budget] + "_..."
            return os.path.join(d, new_name + ext)
        return path

    def create_temp_dir(self) -> str:
        d = FileSystemRepository._safe(tempfile.mkdtemp())
        self.temp_dirs.append(d)
        return d

    def cleanup_temp(self):
        for d in self.temp_dirs:
            d_s = FileSystemRepository._safe(d)
            if os.path.exists(d_s):
                # [v35.4.13] 지연된 점유 해제에 대응하기 위한 재시도 로직 도입
                for _ in range(3):
                    try:
                        shutil.rmtree(d_s, ignore_errors=True)
                        if not os.path.exists(d_s): break
                    except: pass
                    time.sleep(0.3)
        self.temp_dirs = []

    def get_output_dir(self, src_path: str, prefix: str) -> str:
        parent = os.path.dirname(os.path.abspath(src_path))
        ts = time.strftime("%Y%m%d_%H%M%S")
        out_dir = FileSystemRepository._safe(os.path.join(parent, f"{prefix}_{ts}"))
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        return out_dir


# ═══════════════════════════════════════════════════════════════════════════════
# 2. DOMAIN LAYER (도메인: 핵심 비즈니스 로직 및 엔티티)
# ═══════════════════════════════════════════════════════════════════════════════

class OfficeFile:
    """도메인 엔티티: 처리 대상 문서 정보"""
    def __init__(self, path: str):
        # [v35.4.9] 경로 정규화 및 대소문자 무결성 확보 (Deduplication 대비)
        self.path = os.path.normpath(os.path.abspath(path))
        self.filename = os.path.basename(path)
        self.extension = os.path.splitext(path)[1].lower().strip()
        self.is_valid = os.path.exists(path)

class OptimizationDomainService:
    """최적화 알고리즘 서비스 (압축, XML 정밀 정제)"""
    def __init__(self, repo: FileSystemRepository):
        self.repo = repo
    
    def optimize_image(self, img_path: str, quality: int, resize: bool) -> Tuple[int, bool]:
        """개별 이미지 파일 압축 (Pillow)"""
        try:
            with Image.open(img_path) as img:
                old_sz = os.path.getsize(img_path)
                save_needed = False
                
                if resize and (img.width > 2560 or img.height > 2560):
                    img.thumbnail((2560, 2560), Image.Resampling.LANCZOS)
                    save_needed = True
                
                s_fmt = img.format
                if img.mode not in ['RGBA', 'LA'] and 'transparency' not in img.info:
                    s_fmt = 'JPEG'
                    save_needed = True

                if save_needed or img.format in ['JPEG', 'PNG']:
                    img.save(img_path, format=s_fmt, quality=quality, optimize=True)
                    new_sz = os.path.getsize(img_path)
                    if new_sz < old_sz:
                        return (old_sz - new_sz), True
        except: pass
        return 0, False

    def minify_xml(self, xml_path: str) -> bool:
        """XML 파일 경량화"""
        try:
            s_xml = FileSystemRepository._safe(xml_path)
            with open(s_xml, 'r', encoding='utf-8') as f:
                content = f.read()
            nc = re.sub(r'>\s+<', '><', content)
            if len(nc) < len(content):
                with open(s_xml, 'w', encoding='utf-8') as f:
                    f.write(nc)
                return True
        except: pass
        return False


class MergingDomainService:
    """문서 통합 및 목차(TOC) 생성 도메인 서비스"""
    
    def __init__(self, gateway: OfficeAppGateway):
        self.gateway = gateway

    def unify_files(self, files: List[str], target_ext: str) -> List[dict]:
        """[v35.4.9] 확장자 임시 통합 시 원본 이름 보존 및 COM 안전 경로 필터링"""
        results = []
        for f in files:
            f_s = FileSystemRepository._safe(f)
            orig_name = os.path.basename(f)
            ext = os.path.splitext(f)[1].lower().strip()
            
            if ext == target_ext:
                results.append({"path": f_s, "is_temp": False, "orig_name": orig_name})
                continue
            
            # Convert
            t_dir = tempfile.gettempdir()
            t_name = f"tmp_{int(time.time()*1000)}_{orig_name}{target_ext}"
            temp_path = FileSystemRepository._safe(os.path.join(t_dir, t_name))
            temp_path = FileSystemRepository.shorten_path(temp_path)

            try:
                # [v35.4.9] COM Safe Path: 260자 미만이면 \\?\ 접두사 제거 (Office COM 호환성)
                com_path = f_s
                if len(f_s) < 260 and f_s.startswith("\\\\?\\"): com_path = f_s[4:]
                
                com_temp = temp_path
                if len(temp_path) < 260 and temp_path.startswith("\\\\?\\"): com_temp = temp_path[4:]

                if target_ext == '.pptx':
                    p = self.gateway.ppt_app.Presentations.Open(com_path, WithWindow=False)
                    try: p.SaveAs(com_temp, 24)
                    finally: p.Close()
                elif target_ext == '.xlsx':
                    xl = self.gateway.xl_app
                    xl.DisplayAlerts = False
                    w = xl.Workbooks.Open(com_path, UpdateLinks=0, ReadOnly=True)
                    try: w.SaveAs(com_temp, 51)
                    finally: w.Close(False)
                elif target_ext == '.docx':
                    wd = self.gateway.wd_app
                    d = wd.Documents.Open(com_path, Visible=False, ReadOnly=True)
                    try: d.SaveAs2(com_temp, 12)
                    finally: d.Close(False)
                results.append({"path": temp_path, "is_temp": True, "orig_name": orig_name})
            except Exception as e:
                print(f"Conversion Error ({f}): {e}")
                results.append({"path": f_s, "is_temp": False, "orig_name": orig_name})
        return results

    def merge_presentations(self, unified_files: List[dict], output_path: str) -> bool:
        """PPTX 병합: 첫 파일을 베이스로 하여 레이아웃 및 원본 서식 무결성 보존 (v35.4.9)"""
        if len(unified_files) < 2: return False
        ppt = self.gateway.ppt_app
        try:
            out_s = FileSystemRepository._safe(output_path)
            # 1. 첫 번째 파일을 베이스로 복사
            shutil.copy2(unified_files[0]['path'], out_s)
            
            # COM Open 시 260자 미만이면 접두사 제거
            com_out = out_s[4:] if len(out_s) < 260 and out_s.startswith("\\\\?\\") else out_s
            main_pres = ppt.Presentations.Open(com_out, WithWindow=False)
            
            # 2. TOC 슬라이드 추가
            toc_slide = main_pres.Slides.Add(1, 12) 
            title_box = toc_slide.Shapes.AddTextbox(1, 40, 40, 640, 60)
            title_box.TextFrame.TextRange.Text = "통합 프레젠테이션 목록 (TOC)"
            title_box.TextFrame.TextRange.Font.Size = 24
            title_box.TextFrame.TextRange.Font.Bold = True
            
            list_box = toc_slide.Shapes.AddTextbox(1, 40, 110, 640, 400)
            # [v35.4.9] TMP 이름 대신 원본 이름 사용
            toc_text = "\n".join([f"• {f['orig_name']}" for f in unified_files])
            list_box.TextFrame.TextRange.Text = toc_text
            list_box.TextFrame.TextRange.Font.Size = 14
            
            # 3. 삽입
            for f_info in unified_files[1:]:
                f_path = f_info['path']
                f_com = f_path[4:] if len(f_path) < 260 and f_path.startswith("\\\\?\\") else f_path
                main_pres.Slides.InsertFromFile(f_com, main_pres.Slides.Count)
            
            main_pres.Save()
            main_pres.Close()
            return True
        except Exception as e:
            print(f"PPT Merge Error: {e}")
            return False

    def merge_workbooks(self, unified_files: List[dict], output_path: str) -> bool:
        """Excel 병합: 베이스 복제 및 전수 시트 통합 (v35.4.9 TOC 원본 이름 보존)"""
        if len(unified_files) < 1: return False
        xl = self.gateway.xl_app
        try:
            xl.DisplayAlerts = False 
            xl.ScreenUpdating = False 
            
            out_s = FileSystemRepository._safe(output_path)
            shutil.copy2(unified_files[0]['path'], out_s)
            
            com_out = out_s[4:] if len(out_s) < 260 and out_s.startswith("\\\\?\\") else out_s
            main_wb = xl.Workbooks.Open(com_out, UpdateLinks=0)
            
            toc_name = "통합_파일_목록"
            try:
                for _ws in main_wb.Worksheets:
                    if _ws.Name == toc_name:
                        if main_wb.Worksheets.Count > 1: _ws.Delete()
                        else: _ws.Name = f"OLD_{toc_name}"
            except: pass

            toc_ws = main_wb.Worksheets.Add(Before=main_wb.Sheets(1))
            toc_ws.Name = toc_name
            toc_ws.Cells(1, 1).Value = "▣ 통합 엑셀 문서 리스트"
            toc_ws.Cells(1, 1).Font.Bold = True
            toc_ws.Cells(1, 1).Font.Size = 14
            
            used_names = {ws.Name.lower() for ws in main_wb.Worksheets}
            # [v35.4.9] TOC에 원본 이름 기재
            toc_ws.Cells(3, 1).Value = f"1. {unified_files[0]['orig_name']} (기준 파일)"
            curr_row = 4
            
            for f_idx, f_info in enumerate(unified_files[1:], 2):
                src_wb = None
                f_path = f_info['path']
                f_com = f_path[4:] if len(f_path) < 260 and f_path.startswith("\\\\?\\") else f_path
                try:
                    src_wb = xl.Workbooks.Open(f_com, UpdateLinks=0, ReadOnly=True)
                    toc_ws.Cells(curr_row, 1).Value = f"{f_idx}. {f_info['orig_name']}"
                    curr_row += 1
                    
                    for ws in src_wb.Worksheets:
                        # [v35.3.8] 중복 이름 대응전략: 동일 이름 시트 존재 시 [번호] 접두어 부여
                        orig_name = ws.Name
                        target_name = orig_name
                        if target_name.lower() in used_names:
                            prefix = f"[{f_idx}] "
                            target_name = (prefix + orig_name)[:31] # 31자 제한
                        
                        # 시트 복사 (After 인자 명시로 Book X 차단)
                        ws.Copy(None, main_wb.Sheets(main_wb.Sheets.Count))
                        
                        # 복사된 시트 이름 변경 및 used_names 업데이트
                        new_ws = main_wb.Sheets(main_wb.Sheets.Count)
                        try:
                            # 만약 [번호]를 붙여도 중복인 경우 (예외적 상황) 루프 대응
                            retry = 1
                            final_name = target_name
                            while final_name.lower() in used_names:
                                suffix = f"({retry})"
                                final_name = (target_name[:31-len(suffix)] + suffix)
                                retry += 1
                            new_ws.Name = final_name
                            used_names.add(final_name.lower())
                        except: pass
                    
                    src_wb.Close(False)
                except Exception as inner_e:
                    print(f"  [WARN] 시트 통합 실패 ({f_info['orig_name']}): {inner_e}")
                    if src_wb:
                        try: src_wb.Close(False)
                        except: pass
            
            toc_ws.Columns(1).AutoFit()
            main_wb.Save()
            main_wb.Close(True) # 변경사항 저장 확정
            return True
        except Exception as e:
            print(f"Excel Merge Error: {e}")
            return False
        finally:
            try: 
                xl.DisplayAlerts = True
                xl.ScreenUpdating = True
            except: pass

    def merge_documents(self, unified_files: List[dict], output_path: str) -> bool:
        """Word 병합: 베이스 복제 후 내용 삽입 (v35.4.10)"""
        if len(unified_files) < 1: return False
        wd = self.gateway.wd_app
        try:
            # 1. 첫 파일을 베이스로 복제
            shutil.copy2(unified_files[0]['path'], output_path)
            
            com_out = output_path[4:] if len(output_path) < 260 and output_path.startswith("\\\\?\\") else output_path
            main_doc = wd.Documents.Open(com_out, Visible=False)
            
            # 2. 맨 앞에 목차 추가 (서식 적용)
            toc_range = main_doc.Range(0, 0)
            toc_range.InsertAfter("▣ 통합 워드 문서 리스트\n" + "-"*40 + "\n")
            for i, f_info in enumerate(unified_files, 1):
                toc_range.InsertAfter(f"{i}. {f_info['orig_name']}\n")
            toc_range.InsertAfter("-"*40 + "\n\n")
            
            # TOC 부분 강조
            toc_range.Font.Bold = False
            toc_range.Font.Size = 11
            
            # 3. 두 번째 파일부터 순차 삽입
            for f_info in unified_files[1:]:
                f_path = f_info['path']
                f_com = f_path[4:] if len(f_path) < 260 and f_path.startswith("\\\\?\\") else f_path
                # 페이지 나누기 후 삽입
                main_doc.Range(main_doc.Content.End - 1, main_doc.Content.End - 1).InsertBreak(7) # 7: wdPageBreak
                main_doc.Range(main_doc.Content.End - 1, main_doc.Content.End - 1).InsertFile(f_com)
            
            main_doc.Save()
            main_doc.Close()
            return True
        except Exception as e:
            print(f"Word Merge Full Error: {e}")
            return False


class RenamingDomainService:
    """[v35.5.0] 폴더명 기반 파일명 변경 및 규칙 적용 도메인 서비스"""

    def generate_new_name(self, current_path: str, rule: dict, is_merged_result: bool) -> str:
        """규칙에 따른 신규 파일명 생성 (확장자 유지)"""
        d_name = os.path.basename(os.path.dirname(current_path))
        ext = os.path.splitext(current_path)[1]
        
        extract_len = rule.get('extract_len', 2)
        exclude_len = rule.get('exclude_len', 3)
        prefix = rule.get('prefix', '작업요청서_(')
        suffix = rule.get('suffix', ')')
        
        # 추출 부분 (폴더명 앞 N자)
        head = d_name[:extract_len] if len(d_name) >= extract_len else d_name
        
        if is_merged_result:
            # 병합 결과 파일 규칙: [추출] + [접두사] + [제외후 나머지] + [접미사]
            body = d_name[exclude_len:] if len(d_name) > exclude_len else ""
            return f"{head}{prefix}{body}{suffix}{ext}"
        else:
            # 일반 파일 규칙: [추출] + 기존 파일명
            orig_name = os.path.basename(current_path)
            # 이미 추출 글자가 붙어있는지 체크 (중복 방지용 - 선택적 사용 가능)
            return f"{head}{orig_name}"

    def apply_advanced_rules(self, name: str, adv_rule: dict) -> str:
        """[v35.7.0] 고급 정밀 규칙 고도화 (제거/슬라이스대체/문자열교체/삽입)"""
        basename, ext = os.path.splitext(name)
        direction = adv_rule.get('dir', '앞')
        
        # 1. 문자열 교체 (Find & Replace) - [NEW v35.7.0]
        # 범위 규칙 적용 전에 먼저 수행하거나 사용자 의도에 따라 순서 결정 가능하나,
        # '특정 문자열' 매칭은 원본 기준이 명확하므로 최상단 배치.
        find_str = adv_rule.get('find_str', '')
        replace_with = adv_rule.get('replace_with', '')
        if find_str:
            basename = basename.replace(find_str, replace_with)

        res = list(basename)
        L = len(res)
        
        # 2. 제거 규칙
        rem_pos = adv_rule.get('rem_pos', 0)
        rem_len = adv_rule.get('rem_len', 0)
        
        if rem_len > 0 and rem_pos > 0:
            if direction == '앞':
                start = rem_pos - 1
                end = start + rem_len
                if start < L:
                    res[start:end] = []
            else:
                end = L - (rem_pos - 1)
                start = end - rem_len
                if end > 0:
                    res[max(0, start):max(0, end)] = []
        
        L = len(res)
        
        # 3. 범위 대체 규칙 (Slice Replace) - [v35.7.0 확장]
        rep_start = adv_rule.get('rep_start', 0)
        rep_end = adv_rule.get('rep_end', 0)
        rep_str = adv_rule.get('rep_str', '')

        if rep_start > 0:
            # rep_end가 0이거나 rep_start보다 작으면 단일 글자 대체로 간주
            actual_end = max(rep_start, rep_end)
            if direction == '앞':
                s_idx = rep_start - 1
                e_idx = actual_end
                if s_idx < L:
                    res[s_idx:e_idx] = list(rep_str)
            else:
                e_idx = L - (rep_start - 1)
                s_idx = L - actual_end
                if e_idx > 0:
                    res[max(0, s_idx):max(0, e_idx)] = list(rep_str)

        L = len(res)

        # 4. 삽입 규칙 (Insertion) - [NEW v35.7.0]
        ins_start = adv_rule.get('ins_start', 0)
        ins_end = adv_rule.get('ins_end', 0)
        ins_str = adv_rule.get('ins_str', '')

        if ins_start > 0 and ins_str:
            actual_ins_end = max(ins_start, ins_end)
            # 리스트에 삽입 시 인덱스가 밀리는 것을 방지하기 위해 역순 처리
            if direction == '앞':
                # 앞 기준: N번째 '앞'에 삽입
                for i in range(actual_ins_end, ins_start - 1, -1):
                    idx = i - 1
                    if idx <= L:
                        res[idx:idx] = list(ins_str)
            else:
                # 뒤 기준: 뒤에서 N번째 '앞'에 삽입
                for i in range(actual_ins_end, ins_start - 1, -1):
                    idx = L - (i - 1)
                    if 0 <= idx <= len(res):
                        res[idx:idx] = list(ins_str)
                    
        return "".join(res) + ext


# ═══════════════════════════════════════════════════════════════════════════════
# 3. APPLICATION LAYER (응용: 프로세스 제어 및 유스케이스)
# ═══════════════════════════════════════════════════════════════════════════════

class OptimizerApplicationService:
    """UI와 도메인 사이의 가교 역할 (UseCase Orchestrator)"""
    
    def __init__(self, ui_callback):
        self.ui_callback = ui_callback
        self.gateway = OfficeAppGateway()
        self.repo = FileSystemRepository()
        self.opt_service = OptimizationDomainService(self.repo)
        self.merge_service = MergingDomainService(self.gateway)
        self.rename_service = RenamingDomainService() # [v35.5.0] 신규 주입
        self.stop_requested = False
        self.all_finalized_dirs: Set[str] = set() # [v35.5.0] 세션 내 모든 확정 폴더 추적
        self.last_run_results: List[Tuple[str, str]] = []
        self.last_run_dirs: Set[str] = set()
        self.last_merge_sources: List[List[str]] = [] # [v35.4.19] 통합 병합 시 원본 파일 목록 트래킹
        self.last_mode: Optional[str] = None

    def _log(self, msg): self.ui_callback("log", msg)
    def _status(self, msg): self.ui_callback("status", msg)
    def _progress(self, msg): self.ui_callback("progress", msg)

    def _is_excluded(self, ext: str, exclude_val: str, include_val: str = "") -> bool:
        """[v35.4.0] 포함/제외 필터 무결성 강화"""
        # 1. 제외 필터링 (최우선)
        if exclude_val:
            filters = []
            for e in exclude_val.split(','):
                f = e.strip().lower()
                if not f: continue
                filters.append(f)
                if not f.startswith('.'): filters.append('.' + f)
            if ext.lower().strip() in filters: return True
            
        # 2. 포함 필터링 (차우선)
        if include_val:
            filters = []
            for e in include_val.split(','):
                f = e.strip().lower()
                if not f: continue
                filters.append(f)
                if not f.startswith('.'): filters.append('.' + f)
            if ext.lower().strip() not in filters: return True
            
        return False

    def run_optimization(self, files: List[str], options: dict):
        """[v35.4.0] 최적화 프로세스 및 상태 기록"""
        self.stop_requested = False
        self.last_run_results = []
        self.last_run_dirs = set()
        self.last_merge_sources = [] 
        self.last_mode = "optimize"
        
        if options.get('kill_proc', False):
            self._log("[INFO] 기존 오피스 프로세스 정리 중...")
            self.gateway.kill_zombie_processes()
        
        has_xl = any(f.lower().endswith(tuple(FAMILY_MAP['EXCEL'])) for f in files)
        has_ppt = any(f.lower().endswith(tuple(FAMILY_MAP['PPT'])) for f in files)
        has_wd = any(f.lower().endswith(tuple(FAMILY_MAP['WORD'])) for f in files)
        self.gateway.start_apps(needs_excel=has_xl, needs_ppt=has_ppt, needs_word=has_wd)
        
        success, fail = 0, 0
        res_dirs = {} # {src_dir: out_dir}
        
        ex_val = options.get('exclude_ext', '')
        in_val = options.get('include_ext', '')
        
        for i, path in enumerate(files, 1):
            if self.stop_requested: break
            self._progress(f"진행현황: {i}/{len(files)} ({(i/len(files))*100:.1f}%)")
            doc = OfficeFile(path)
            
            if self._is_excluded(doc.extension, ex_val, in_val):
                self._log(f"[{i}/{len(files)}] {doc.path} (필터 제외됨)")
                continue

            self._log(f"[{i}/{len(files)}] {doc.path}")
            
            try:
                src_dir = os.path.dirname(doc.path)
                if src_dir not in res_dirs:
                    out_dir = self.repo.get_output_dir(doc.path, "00_Optimized_Docs")
                    res_dirs[src_dir] = out_dir
                    self.last_run_dirs.add(out_dir)
                
                target_path = FileSystemRepository._safe(os.path.join(res_dirs[src_dir], doc.filename))
                
                # [v35.4.2] 경로 길이 유효성 전수 검점 및 강제 단축 (Budget: 245)
                target_path = FileSystemRepository.shorten_path(target_path, max_len=245)
                
                shutil.copy2(FileSystemRepository._safe(doc.path), target_path)
                
                # [v35.4.16] 최적화 파이프라인 재설계 (Order: COM Cleaning/Conversion -> ZIP Compression)
                # 1. 메타데이터 클리닝 및 포맷 강제 승격 (Legacy binary -> Modern XML)
                # COM 엔진을 먼저 가동하여 구조를 현대화해야 패키지 압축(Pillow)이 작동 가능함.
                if options.get('clean_meta', True):
                    old_target = target_path
                    target_path = self._deep_clean(target_path)
                    # 만약 변환되어 경로가 바뀌었다면 구형 임시 파일은 즉시 제거
                    if target_path != old_target and os.path.exists(old_target):
                        try: os.remove(old_target)
                        except: pass

                # 2. 패키지 최적화 (이제 모든 파일이 XML 구조이므로 압축 효율 극대화)
                is_mod, saved = self._optimize_pkg(target_path, options)
                if is_mod: self._log(f"  [INFO] 압축 완료 (-{saved/1024:.1f}KB)")
                
                # 3. 무결성 검증
                if options.get('verify', True):
                    v_ok, v_msg = self._verify_integrity(target_path)
                    if v_ok:
                        self.last_run_results.append((doc.path, target_path))
                        self.last_merge_sources.append([]) # 최적화 시에는 원본 개별 삭제 안함 (Replace 로직에서 처리)
                        success += 1
                    else:
                        self._log(f"  [ERR] 검증 실패: {v_msg}")
                        fail += 1
                else:
                    self.last_run_results.append((doc.path, target_path))
                    self.last_merge_sources.append([])
                    success += 1
                    
            except Exception as e:
                self._log(f"  [FAIL] {e}")
                fail += 1
        
        self.gateway.close_apps()
        self._status(f"최적화 완료: 성공 {success}, 실패 {fail}")

    def run_merging(self, files: List[str], options: dict):
        """[v35.4.18] 디렉토리별 무결성 통합 병합 및 확정 정리 상태 기록"""
        try:
            self.stop_requested = False
            self.last_run_results = []
            self.last_run_dirs = set()
            self.last_merge_sources = []
            self.last_mode = "merge" # [v35.4.0] 통합 병합 모드 기록
            
            ex_val = options.get('exclude_ext', '')
            in_val = options.get('include_ext', '')

            # 1. 디렉토리별 & 소프트웨어 제품군별 그룹화
            dir_groups = {} # {dir_path: {sw_family: [file_paths]}}
            for f in files:
                doc = OfficeFile(f)
                if self._is_excluded(doc.extension, ex_val, in_val): continue
                
                # Determine SW Family
                sw_family = None
                for family, extensions in FAMILY_MAP.items():
                    if doc.extension in extensions:
                        sw_family = family
                        break
                
                if sw_family:
                    if d_path := os.path.dirname(f):
                        if d_path not in dir_groups: dir_groups[d_path] = {}
                        if sw_family not in dir_groups[d_path]: dir_groups[d_path][sw_family] = []
                        dir_groups[d_path][sw_family].append(f)

            if not dir_groups:
                self._log("[SKIP] 제외 필터 또는 지원하지 않는 확장자로 인해 병합할 대상이 없습니다. (병합 대상 없음)")
                self._status("병합 스킵됨 (대상 없음)")
                return

            # [v35.4.11] 전체 병합 작업 단위(Task) 산출: 폴더별 제품군 그룹 중 2개 이상인 것의 합
            total_tasks = 0
            for d_path, sw_groups_dic in dir_groups.items():
                for sw_fam, g_f in sw_groups_dic.items():
                    if len(g_f) >= 2: total_tasks += 1
            
            if total_tasks == 0:
                self._log("[SKIP] 선택된 파일 중 동일 폴더/제품군 2개 이상의 병합 대상이 없습니다. (병합 스킵)")
                self._status("병합 대상 없음")
                return

            if options.get('kill_proc', False):
                self._log("[INFO] 기존 오피스 프로세스 정리 중 (Clean Start)...")
                self.gateway.kill_zombie_processes()

            self.gateway.start_apps(needs_excel=True, needs_ppt=True, needs_word=True)
            
            # [v35.5.0] 병합 순서 옵션 추출
            order_opt = options.get('merge_order', 'list')

            curr_task_idx = 0
            total_folders = len(dir_groups)
            for idx, (d_path, sw_groups_raw) in enumerate(dir_groups.items(), 1):
                if self.stop_requested: break
                # 폴더별 로그는 유지하되, 전체 진행현황 표시 방식 변경
                sw_groups = cast(Dict[str, List[str]], sw_groups_raw)
                
                # 각 폴더별로 출력 폴더 생성 (v35.3.3 안정화)
                out_dir = self.repo.get_output_dir(os.path.join(d_path, "dummy.txt"), "00_Merged_Docs")
                self.last_run_dirs.add(out_dir)
                
                self._log(f"[{idx}/{total_folders}] 📂 폴더 통합 중: {os.path.basename(d_path)}")
            
                for sw_family, g_files_raw in sw_groups.items():
                    if self.stop_requested: break
                    if len(g_files_raw) < 2:
                        self._log(f"  [SKIP] {sw_family} 파일이 1개뿐이어서 병합하지 않습니다.")
                        continue
                
                    curr_task_idx += 1
                    self._progress(f"진행현황: {curr_task_idx}/{total_tasks} ({(curr_task_idx/total_tasks)*100:.1f}%)")
                    
                    g_files = list(dict.fromkeys(g_files_raw)) # [v35.4.9] 순서 유지하며 중복 제거
                    
                    # [v35.5.0] 병합 순서 적용
                    if order_opt == "reverse":
                        g_files.reverse()
                    elif order_opt == "list":
                        # 기본적으로 리스트뷰에 추가된 순서(g_files_raw 순서)를 따름
                        pass
                    # manual 모드는 이미 self.files가 사용자가 조정한 순서대로이므로 pass
                    
                    actual_exts = set(os.path.splitext(f)[1].lower() for f in g_files)
                    # [v35.4.18] Forced Modernization: ALWAYS merge to modern formats (.pptx, .xlsx, .docx)
                    target_ext = HIGH_EXT_MAP.get(sw_family, list(actual_exts)[0])
                    
                    if any(ext in ['.xls', '.ppt', '.doc'] for ext in actual_exts):
                        self._log(f"  [*] 레거시 포맷 감지 -> 현대적 포맷 {target_ext} 로 변환 필요...")
                    elif len(actual_exts) > 1:
                        self._log(f"  [*] 확장자 혼합 감지 -> {target_ext} 버전으로 임시 통일 중...")
                    
                    self._log(f"  [*] {sw_family} 파일 {len(g_files)}개 분석 및 경로 정규화 중...")
                    unified_info = self.merge_service.unify_files(g_files, target_ext)
                    temp_paths = [info['path'] for info in unified_info if info['is_temp']]
                    
                    first_stem = Path(g_files[0]).stem
                    out_name = f"병합_{first_stem}{target_ext}"
                    
                    # [v35.4.10] Excel/Office COM Name Conflict Defense
                    unique_id = int(time.time() * 1000) % 1000000
                    temp_out_name = f"Merging_{unique_id}_{out_name}"
                    out_path = FileSystemRepository._safe(os.path.join(out_dir, temp_out_name))
                    out_path = FileSystemRepository.shorten_path(out_path)
                    
                    self._log(f"  [*] {sw_family} 병합 엔진 가동 중 (Target: {out_name})...")
                    
                    res = False
                    if sw_family == 'PPT':
                        res = self.merge_service.merge_presentations(unified_info, out_path)
                    elif sw_family == 'EXCEL':
                        res = self.merge_service.merge_workbooks(unified_info, out_path)
                    elif sw_family == 'WORD':
                        res = self.merge_service.merge_documents(unified_info, out_path)
                    
                    # 임시 파일 즉각 제거 (Lock 해제 유도)
                    for tp in temp_paths:
                        try: os.remove(tp)
                        except: pass

                    if res:
                        self._log(f"    [OK] 생성 완료: {out_name}")
                        
                        # [v35.4.3] 통합 병합 결과물에 대해서도 '최적화' 및 '정제' 수행 (심층분석 반영)
                        self._log("    [*] 병합 결과물 최적화 및 정제 중 (Background Processing)...")
                        self._optimize_pkg(out_path, options)
                        if options.get('clean_meta', True):
                            self._deep_clean(out_path)
                        
                        target_final_path = FileSystemRepository._safe(os.path.join(d_path, out_name))
                        self.last_run_results.append((target_final_path, out_path))
                        self.last_merge_sources.append(g_files) # [v35.4.19] 병합에 사용된 원본 파일들 기록
                    else:
                        self._log(f"    [FAIL] {sw_family} 병합 실패")

            self.gateway.close_apps()
            self._status("통합 병합 작업 종료")
        except Exception as e:
            self._log(f"[FATAL ERROR] 통합 병합 중 중단됨: {str(e)}")
            self.gateway.close_apps()
            self._status("병합 오류 발생")

    def get_finalize_info(self) -> Optional[dict]:
        """[v35.4.0] 확정 정리 전 작업 요약 및 무결성 검증 리포트 생성"""
        if not self.last_run_results: return None
        
        mode_name = "최적화 (Optimization)" if self.last_mode == "optimize" else "통합 병합 (Merging)"
        total = len(self.last_run_results)
        v_ok_count = 0
        
        for src, tgt in self.last_run_results:
            s_p, t_p = FileSystemRepository._safe(src), FileSystemRepository._safe(tgt)
            if os.path.exists(t_p):
                ok, _ = self._verify_integrity(t_p)
                if ok: v_ok_count += 1
        
        return {
            "mode": self.last_mode,
            "mode_name": mode_name,
            "count": total,
            "v_status": f"{v_ok_count}/{total} 검증 완료(정상)",
            "all_verify_ok": (v_ok_count == total)
        }

    def finalize_cleanup(self):
        """[v35.4.2] 확정 정리 가로채기 방지(Lock Check) 및 Atomic 백업-교체 아키텍처"""
        if not self.last_run_results: return
        
        self._log("\n[STEP 1] 원본 파일 쓰기 권한 및 점유 상태 사전 체크...")
        for src, tgt in self.last_run_results:
            s_p = FileSystemRepository._safe(src)
            if os.path.exists(s_p):
                try:
                    # 독점적 쓰기 모드로 열기 시도하여 점유 체크
                    with open(s_p, 'r+'): pass
                except OSError:
                    messagebox.showerror("중단", f"파일이 다른 프로그램에서 사용 중입니다:\n{os.path.basename(src)}")
                    self._log(f"  [ABORT] {os.path.basename(src)} 점유로 인한 중단")
                    return

        self._log("[STEP 2] 원자적 교체 프로세스 시작 (Backup -> Replace -> Verify)")
        cnt = 0
        for idx, (src, tgt) in enumerate(self.last_run_results):
            try:
                s_p, t_p = FileSystemRepository._safe(src), FileSystemRepository._safe(tgt)
                if not os.path.exists(t_p): continue
                
                # [v35.4.17] 원자적 확장자 교체 및 레거시 소거 로직 (Extension Aggressive Clean)
                # s_p가 .ppt이고 t_p가 .pptx인 경우, 혹은 s_p가 이미 .pptx인데 .ppt가 존재하는 경우 처리
                src_dir = os.path.dirname(s_p)
                base_stem = os.path.splitext(os.path.basename(s_p))[0]
                tgt_ext = os.path.splitext(t_p)[1].lower()
                final_s_p = FileSystemRepository._safe(os.path.join(src_dir, base_stem + tgt_ext))

                # 레거시 상계 대응: 타겟 확장자와 다른 동일 명칭의 레거시 파일이 있다면 함께 처리
                legacy_targets = []
                if tgt_ext in ['.pptx', '.xlsx', '.docx']:
                    leg_exts = {'.pptx': ['.ppt'], '.xlsx': ['.xls', '.xlsb'], '.docx': ['.doc']}.get(tgt_ext, [])
                    for le in leg_exts:
                        test_leg = FileSystemRepository._safe(os.path.join(src_dir, base_stem + le))
                        if os.path.exists(test_leg): legacy_targets.append(test_leg)

                # 1. 대상 폴더 존재 확인
                os.makedirs(src_dir, exist_ok=True)

                # 2. 백업 생성 (.bak) - 원본 대상(s_p) 및 감지된 레거시 대상
                all_backups = []
                # s_p 백업
                if os.path.exists(s_p):
                    bak_path = s_p + ".bak"
                    if os.path.exists(bak_path): 
                        try: os.remove(bak_path)
                        except: pass
                    os.rename(s_p, bak_path)
                    all_backups.append(bak_path)
                
                # 감지된 레거시들 추가 백업 (s_p와 겹치지 않는 것만)
                for lt in legacy_targets:
                    if lt != s_p and os.path.exists(lt):
                        l_bak = lt + ".bak"
                        if os.path.exists(l_bak): 
                            try: os.remove(l_bak)
                            except: pass
                        os.rename(lt, l_bak)
                        all_backups.append(l_bak)
                
                # 3. 신규 파일 이동 (t_p -> final_s_p)
                # 만약 이미 같은 이름의 파일이 있으면 삭제 후 이동
                if os.path.exists(final_s_p) and final_s_p != s_p:
                    try: os.remove(final_s_p)
                    except: pass
                shutil.move(t_p, final_s_p)
                
                # 4. 검증 및 백업 삭제
                if os.path.exists(final_s_p):
                    for bp in all_backups:
                        try: os.remove(bp)
                        except: pass
                    
                    # [v35.4.19] 통합 병합 모드인 경우 사용된 원본 파일들 소거
                    if self.last_mode == "merge" and idx < len(self.last_merge_sources):
                        sources = self.last_merge_sources[idx]
                        del_cnt = 0
                        for s_file in sources:
                            s_file_safe = FileSystemRepository._safe(s_file)
                            # 합쳐진 결과물 자체(final_s_p)는 삭제하면 안 됨
                            if os.path.exists(s_file_safe) and s_file_safe != final_s_p:
                                try:
                                    os.remove(s_file_safe)
                                    del_cnt += 1
                                except: pass
                        if del_cnt > 0:
                            self._log(f"    [DELETE] 병합 원본 {del_cnt}개 파일 제거 완료")

                    cnt += 1
                    self._log(f"  [SUCCESS] {os.path.basename(src)} -> {os.path.basename(final_s_p)} 교체 완료")
                else:
                    # 실패 시 롤백 시도 (가장 중요한 s_p 복원 우선)
                    if all_backups:
                        main_bak = all_backups[0]
                        if os.path.exists(main_bak): os.rename(main_bak, s_p)
                    self._log(f"  [FAIL] {os.path.basename(src)} 교체 실패 (롤백됨)")

            except Exception as e:
                self._log(f"  [ERROR] {src} 처리 중 예외: {e}")
        
        self._log(f"\n[FINISH] 총 {cnt}개 파일 원본 확정 교환 완료.")
        # [v35.5.0] 확정된 디렉토리 보관 (파일명 변경 대상 식별용)
        for src, _ in self.last_run_results:
            self.all_finalized_dirs.add(os.path.dirname(FileSystemRepository._safe(src)))

        for d in self.last_run_dirs:
            shutil.rmtree(FileSystemRepository._safe(d), ignore_errors=True)
        self.last_run_results = []
        self.last_merge_sources = []
        self.last_mode = None
        self._status("확정 정리 완료")

    def rename_files_by_rule(self, options: dict) -> Tuple[int, int]:
        """[v35.5.0] 규칙 기반 파일명 대량 변경 (응용 레이어 오케스트레이션)"""
        rule_opt = options.get('rename_rule', {})
        target_mode = options.get('target_mode', 'final') # 'final' or 'general'
        
        scope = rule_opt.get('scope', 'A') 
        temporal = rule_opt.get('temporal', '1')
        
        if target_mode == 'general':
            # [v35.5.1] 일반 파일 모드에서도 사용자 설정(rule_opt)을 존중하도록 변경 (기존 강제 고정 해제)
            pass

        # 1. 대상 폴더 선정
        target_dirs = set()
        if target_mode == 'final':
            if temporal == '1':
                target_dirs.update(self.all_finalized_dirs)
            else:
                u_files = options.get('current_files', [])
                for f in u_files:
                    target_dirs.add(os.path.dirname(FileSystemRepository._safe(f)))
        else:
            # 일반 파일 대상: 현황창에 로드된 모든 파일의 폴더
            u_files = options.get('current_files', [])
            for f in u_files:
                target_dirs.add(os.path.dirname(FileSystemRepository._safe(f)))

        if not target_dirs: 
            self._log("[WARN] 변경 대상 폴더를 찾을 수 없습니다.")
            return 0, 0

        success, fail = 0, 0
        for d_path in target_dirs:
            if not os.path.exists(d_path): continue
            self._log(f"[*] 폴더 스캔/변경 중: {os.path.basename(d_path)} (모드:{target_mode})")
            
            for f_name in os.listdir(d_path):
                f_path = FileSystemRepository._safe(os.path.join(d_path, f_name))
                if not os.path.isfile(f_path): continue
                
                # 병합된 파일 조건: f_name에 특정 패턴 포함
                is_merged = f_name.startswith("병합_") or f_name.startswith("Merging_") or "작업요청서_(" in f_name
                
                # 모드 A(병합파일만)인데 병합파일이 아니면 스킵 (일반 파일 모드면 무시됨)
                if scope == 'A' and not is_merged: continue
                
                new_name = self.rename_service.generate_new_name(f_path, rule_opt, is_merged)
                if f_name == new_name: continue
                
                new_path = FileSystemRepository._safe(os.path.join(d_path, new_name))
                try:
                    if os.path.exists(new_path):
                        self._log(f"  [SKIP] 중복 파일 존재: {new_name}")
                        continue
                    os.rename(f_path, new_path)
                    success += 1
                except Exception as e:
                    self._log(f"  [ERR] 이름 변경 실패({f_name}): {e}")
                    fail += 1
                    
        return success, fail

    def run_advanced_renaming(self, files: List[str], adv_rule: dict):
        """[v35.6.0] 고급 이름 편집 실행 (Standalone UseCase)"""
        try:
            self._log(f"\n[START] 고급 명칭 일괄 편집 시작 (대상: {len(files)}개)")
            count = 0
            for i, old_path in enumerate(files):
                if not os.path.exists(old_path): continue
                
                dir_name = os.path.dirname(old_path)
                old_name = os.path.basename(old_path)
                
                new_name = self.rename_service.apply_advanced_rules(old_name, adv_rule)
                if new_name == old_name: continue
                
                new_path = FileSystemRepository._safe(os.path.join(dir_name, new_name))
                
                # 중복 방지
                if os.path.exists(new_path):
                    name_base, ext = os.path.splitext(new_name)
                    new_path = FileSystemRepository._safe(os.path.join(dir_name, f"{name_base}_dup{ext}"))
                
                try:
                    os.rename(old_path, new_path)
                    files[i] = new_path # 리스트 동기화
                    count += 1
                except Exception as e:
                    self._log(f"  [ERR] {old_name} 변경 실패: {e}")
                
            self._status("고급 편집 완료")
            self._log(f"[FINISH] 고급 명칭 편집 완료 (변경: {count}개)")
        except Exception as e:
            self._log(f"[ERROR] 고급 편집 프로세스 오류: {e}")
            raise e

    def _optimize_pkg(self, path: str, options: dict) -> Tuple[bool, int]:
        ext = os.path.splitext(path)[1].lower()
        if ext in ['.xls', '.ppt', '.doc']: return False, 0 # Legacy binary는 ZIP 최적화 제외
        if not zipfile.is_zipfile(path): return False, 0
        temp_dir = self.repo.create_temp_dir()
        is_mod = False
        total_saved = 0
        try:
            with zipfile.ZipFile(path, 'r') as zf:
                zf.extractall(temp_dir)
            for root, _, fs in os.walk(temp_dir):
                for f in fs:
                    fp = os.path.join(root, f)
                    ext = os.path.splitext(f)[1].lower()
                    saved, mod = self.opt_service.optimize_image(fp, options.get('quality', 70), options.get('resize', True))
                    if mod:
                        total_saved += saved
                        is_mod = True
                    if options.get('xml_opt', True) and ext in ['.xml', '.rels']:
                        if self.opt_service.minify_xml(fp): is_mod = True
            if is_mod:
                nz = FileSystemRepository._safe(path + ".tmp")
                with zipfile.ZipFile(nz, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, _, fs in os.walk(temp_dir):
                        for f in fs:
                            p = FileSystemRepository._safe(os.path.join(root, f))
                            # zip 내부는 상대경로를 써야 하므로 prefix 제거된 원본 temp_dir 기준 relpath 사용
                            zf.write(p, os.path.relpath(p, temp_dir).replace('\\\\?\\', '').replace('\\\\?\\UNC\\', ''))
                shutil.move(nz, FileSystemRepository._safe(path))
            return is_mod, total_saved
        except Exception as e:
            self._log(f"  [ERR] 패키지 분석 중 오류: {e}")
            return False, 0
        finally: 
            self.repo.cleanup_temp()

    def _deep_clean(self, path: str) -> str:
        """[v35.4.16] COM 엔진을 이용한 포맷 현대화 및 메타데이터 정제 (결과 경로 반환)"""
        ext = os.path.splitext(path)[1].lower()
        s_p = FileSystemRepository._safe(path)
        final_path = path
        
        try:
            # [v35.4.18] COM Safe Path: 260자 미만이면 \\?\ 접두사 제거
            com_path = s_p
            if len(s_p) < 260 and s_p.startswith("\\\\?\\"): com_path = s_p[4:]

            xl_app = self.gateway.xl_app
            if ext in ['.xlsx', '.xlsm', '.xls', '.xlsb'] and xl_app:
                wb = None
                try:
                    wb = xl_app.Workbooks.Open(com_path, UpdateLinks=False)
                    
                    # [v35.4.16] Format Upgrade: .xls -> .xlsx
                    if ext == '.xls':
                        new_p = os.path.splitext(path)[0] + ".xlsx"
                        s_new = FileSystemRepository._safe(new_p)
                        wb.SaveAs(s_new, 51) # 51: xlOpenXMLWorkbook
                        final_path = new_p
                        self._log(f"    [*] Excel 포맷 현대화 완료 (.xls -> .xlsx)")

                    if ext not in ['.xls', '.xlsb']:
                        try: wb.RemoveDocumentInformation(99)
                        except: pass
                    
                    # 2. 레거시 정밀 삭제: 외부 링크 끊기 및 정의된 이름 삭제
                    try:
                        links = wb.LinkSources(1)
                        if links:
                            for link in links: wb.BreakLink(link, 1)
                    except: pass
                    try:
                        for name in wb.Names:
                            try: name.Delete()
                            except: pass
                    except: pass
                    
                    wb.Save()
                finally:
                    if wb:
                        try: wb.Close()
                        except: pass
                        wb = None
                
            ppt_app = self.gateway.ppt_app
            if ext in ['.pptx', '.pptm', '.ppt'] and ppt_app:
                pres = None
                try:
                    pres = ppt_app.Presentations.Open(s_p, WithWindow=False)
                    
                    # [v35.4.16] Format Upgrade: .ppt -> .pptx
                    if ext == '.ppt':
                        new_p = os.path.splitext(path)[0] + ".pptx"
                        s_new = FileSystemRepository._safe(new_p)
                        pres.SaveAs(s_new, 24) # 24: ppSaveAsOpenXMLPresentation
                        final_path = new_p
                        self._log(f"    [*] PPT 포맷 현대화 완료 (.ppt -> .pptx)")

                    try: pres.RemoveDocumentInformation(99)
                    except: pass
                    pres.Save()
                finally:
                    if pres:
                        try: pres.Close()
                        except: pass
                        pres = None

            wd_app = self.gateway.wd_app
            if ext in ['.docx', '.doc'] and wd_app:
                doc_obj = None
                try:
                    doc_obj = wd_app.Documents.Open(s_p, Visible=False)
                    
                    # [v35.4.16] Format Upgrade: .doc -> .docx
                    if ext == '.doc':
                        new_p = os.path.splitext(path)[0] + ".docx"
                        s_new = FileSystemRepository._safe(new_p)
                        doc_obj.SaveAs2(s_new, 12) # 12: wdFormatXMLDocument
                        final_path = new_p
                        self._log(f"    [*] Word 포맷 현대화 완료 (.doc -> .docx)")

                    try: doc_obj.RemoveDocumentInformation(99)
                    except: pass
                    doc_obj.Save()
                finally:
                    if doc_obj:
                        try: doc_obj.Close()
                        except: pass
                        doc_obj = None
        except Exception as e:
            self._log(f"    [WARN] 심층 정제/변환 중 오류: {e}")
            
        return final_path

    def _verify_integrity(self, path: str) -> Tuple[bool, str]:
        """[v35.3.3] 확장자별 맞춤 무결성 검증 (Legacy vs Modern)"""
        ext = os.path.splitext(path)[1].lower().strip()
        # Legacy 포맷은 Zip 구조가 아니므로 Zip 검증 스킵
        if ext in ['.ppt', '.xls', '.doc']:
            return True, "OK (Legacy Format)"
            
        try:
            with zipfile.ZipFile(path, 'r') as zf:
                if zf.testzip() is not None: return False, "CRC Error"
            return True, "OK"
        except Exception as e: return False, str(e)


# ═══════════════════════════════════════════════════════════════════════════════
# 4. PRESENTATION LAYER (표현: UI 레이어 - v35.1.33 레이아웃 무결성 복원)
# ═══════════════════════════════════════════════════════════════════════════════

class OptimizerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("v35.4.18 - 만능 오피스 최적화 및 통합 도구")
        self.root.geometry("800x980") # 가로폭 상향 (경로 표시 대응)
        
        # [v35.4.18] UI Attributes & Dependency Injection
        self.lbl_status: Optional[tk.Label] = None
        self.lbl_progress: Optional[tk.Label] = None
        self.btn_run: Optional[tk.Button] = None
        self.btn_finish: Optional[tk.Button] = None
        self.btn_rename: Optional[tk.Button] = None # [v35.5.0]
        self.btn_adv_rename: Optional[tk.Button] = None # [v35.6.0]
        self.lst: Optional[tk.Listbox] = None
        self.log_txt: Optional[tk.Text] = None
        
        # State Variables
        self.mode_var = tk.StringVar(value="optimize")
        self.qual_var = tk.IntVar(value=80)
        self.resize_var = tk.BooleanVar(value=True)
        self.xml_opt_var = tk.BooleanVar(value=True)
        self.clean_meta_var = tk.BooleanVar(value=True)
        self.verify_var = tk.BooleanVar(value=True)
        self.kill_proc_var = tk.BooleanVar(value=True)
        self.exclude_ext_var = tk.StringVar(value=".bak, .tmp, .temp")
        self.include_ext_var = tk.StringVar(value=".ppt, .pptx") # [v35.5.0] 파워포인트 기본값 설정
        self.merge_order_var = tk.StringVar(value="list") # (list, reverse, manual)
        
        # [v35.5.0] 파일명 변경 규칙 변수 (1. 확정 정리용)
        self.rename_extract_var = tk.IntVar(value=2)
        self.rename_exclude_var = tk.IntVar(value=3)
        self.rename_prefix_var = tk.StringVar(value="작업요청서_(")
        self.rename_suffix_var = tk.StringVar(value=")")
        self.rename_scope_var = tk.StringVar(value="A")
        self.rename_temporal_var = tk.StringVar(value="1")
        
        # [v35.5.1] 파일명 변경 규칙 변수 (2. 일반 파일용)
        self.gen_rename_extract_var = tk.IntVar(value=2)
        self.gen_rename_exclude_var = tk.IntVar(value=3)
        self.gen_rename_prefix_var = tk.StringVar(value="당사안_(")
        self.gen_rename_suffix_var = tk.StringVar(value=")")
        
        # [v35.6.0] 고급 이름 편집 (Standalone) 변수
        self.adv_rename_dir_var = tk.StringVar(value="앞") # "앞" or "뒤"
        self.adv_remove_pos_var = tk.IntVar(value=0)     # N번째부터
        self.adv_remove_len_var = tk.IntVar(value=0)     # M글자 제거
        
        # [v35.7.0] 고급 이름 편집 변수 고도화 (범위 대체 / 문자열 교차 / 삽입)
        self.adv_replace_start_var = tk.IntVar(value=0)   # N번째부터
        self.adv_replace_end_var = tk.IntVar(value=0)     # M번째까지
        self.adv_replace_str_var = tk.StringVar(value="") # 대체할 문자열
        
        self.adv_find_str_var = tk.StringVar(value="")    # 찾을 문자열
        self.adv_replace_with_var = tk.StringVar(value="") # 바꿀 문자열
        
        self.adv_insert_start_var = tk.IntVar(value=0)    # N번째부터
        self.adv_insert_end_var = tk.IntVar(value=0)      # M번째까지
        self.adv_insert_str_var = tk.StringVar(value="")  # 삽입할 문자열

        self.gen_rename_scope_var = tk.StringVar(value="B") # 일반 파일은 보통 폴더 전체 대상
        self.gen_rename_temporal_var = tk.StringVar(value="2") # 일반 파일은 보통 이전 전수 포함
        
        self.lbl_preview_final: Optional[tk.Label] = None
        self.lbl_preview_gen: Optional[tk.Label] = None
        self.lbl_preview_adv: Optional[tk.Label] = None # [v35.6.0]
        
        # [v35.5.0] 실효적 실시간 위젯 추적 (trace)
        self.merge_order_var.trace_add("write", self._on_merge_order_changed)
        
        v_list = [
            self.rename_extract_var, self.rename_exclude_var, self.rename_prefix_var, self.rename_suffix_var,
            self.gen_rename_extract_var, self.gen_rename_exclude_var, self.gen_rename_prefix_var, self.gen_rename_suffix_var,
            self.adv_rename_dir_var, self.adv_remove_pos_var, self.adv_remove_len_var,
            self.adv_replace_start_var, self.adv_replace_end_var, self.adv_replace_str_var,
            self.adv_find_str_var, self.adv_replace_with_var,
            self.adv_insert_start_var, self.adv_insert_end_var, self.adv_insert_str_var
        ]
        for v in v_list:
            v.trace_add("write", lambda *args: self._update_previews())
        
        self._set_focus()
        self.service = OptimizerApplicationService(self._ui_callback)
        self.files = []
        self._setup_layout()

    def _set_focus(self):
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after_idle(self.root.attributes, "-topmost", False)
        self.root.focus_force()

    def _ui_callback(self, type, msg):
        """[v35.4.18] Thread-safe UI updates via root.after"""
        def _update():
            if type == "log": self.log(msg)
            elif type == "progress":
                if self.lbl_progress: self.lbl_progress.config(text=msg)
            elif type == "status": 
                if self.lbl_status: self.lbl_status.config(text=msg)
                if "완료" in msg or "종료" in msg:
                    if self.btn_run: self.btn_run.config(state="normal")
                    if self.btn_finish: self.btn_finish.config(state="normal")
                    if hasattr(self, 'btn_rename') and self.btn_rename: self.btn_rename.config(state="normal")
                    if self.lbl_progress: self.lbl_progress.config(text="")
        
        if self.root:
            self.root.after(0, _update)

    def _setup_layout(self):
        # Header (Theme: #004D40)
        h_fr = tk.Frame(self.root, bg="#004D40", pady=15)
        h_fr.pack(fill="x", side="top")
        tk.Label(h_fr, text="[OPTIMIZER] Universal Office Optimizer", font=("맑은 고딕", 16, "bold"), bg="#004D40", fg="white").pack()
        tk.Label(h_fr, text="v35.4.18 - Format Hardening & Legacy Replacement Protocol", font=("맑은 고딕", 9), bg="#004D40", fg="#B2DFDB").pack()

        # Action Bar
        act_fr = tk.Frame(self.root, padx=10, pady=10, bg="#E0F2F1")
        act_fr.pack(fill="x")
        act_fr.columnconfigure(0, weight=1); act_fr.columnconfigure(1, weight=1)
        self.btn_run = tk.Button(act_fr, text="[OK] 작업 시작 (Start)", bg="#00695C", fg="white", font=("맑은 고딕", 12, "bold"), height=2, command=self._run)
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        self.btn_finish = tk.Button(act_fr, text="✨ 확정 정리 (Replace)", bg="#455A64", fg="white", font=("맑은 고딕", 12, "bold"), height=2, command=self._on_finalize)
        self.btn_finish.grid(row=0, column=1, sticky="ew", padx=(2, 2))
        # [v35.5.0] 원본 병합 파일명 변경 버튼 복조
        self.btn_rename = tk.Button(act_fr, text="📁 병합 파일명 변경", bg="#1565C0", fg="white", font=("맑은 고딕", 12, "bold"), height=2, command=self._run_renaming)
        self.btn_rename.grid(row=0, column=2, sticky="ew", padx=(2, 2))
        # [v35.6.0] 고급 명칭 일괄 편집 버튼 (Standalone)
        self.btn_adv_rename = tk.Button(act_fr, text="🎯 고급 편집 (Advanced)", bg="#7B1FA2", fg="white", font=("맑은 고딕", 12, "bold"), height=2, command=self._run_advanced_renaming)
        self.btn_adv_rename.grid(row=0, column=3, sticky="ew", padx=(2, 0))
        act_fr.columnconfigure(2, weight=1); act_fr.columnconfigure(3, weight=1)

        # Main Scrollable
        mf = tk.Frame(self.root); mf.pack(fill="both", expand=True)
        cv = tk.Canvas(mf); sb = ttk.Scrollbar(mf, orient="vertical", command=cv.yview)
        sf = ttk.Frame(cv, padding=15)
        sf.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv_win = cv.create_window((0, 0), window=sf, anchor="nw")
        cv.bind("<Configure>", lambda e: cv.itemconfig(cv_win, width=e.width))
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        
        # Mouse Wheel Support (v35.2.6)
        def _on_mousewheel(event):
            cv.yview_scroll(int(-1 * (event.delta / 120)), "units")
        cv.bind_all("<MouseWheel>", _on_mousewheel)
        
        self._build_content(sf)
        self.lbl_status = tk.Label(self.root, text="대기 중...", anchor="w", fg="#64748b", padx=10, pady=5)
        self.lbl_status.pack(fill="x", side="bottom")

    def _build_content(self, p):
        # 0. Mode & Order
        lf0 = ttk.LabelFrame(p, text="0. 작업 모드 및 병합 순서 (Mode & Merge Order)", padding=10)
        lf0.pack(fill="x", pady=5)
        
        m_fr = ttk.Frame(lf0); m_fr.pack(fill="x")
        tk.Label(m_fr, text="● 작업 모드:", font=("맑은 고딕", 9, "bold")).pack(side="left", padx=5)
        tk.Radiobutton(m_fr, text="최적화 (Optimize)", variable=self.mode_var, value="optimize").pack(side="left", padx=10)
        tk.Radiobutton(m_fr, text="통합 병합 (Merge per Folder)", variable=self.mode_var, value="merge").pack(side="left", padx=10)
        
        ttk.Separator(lf0, orient="horizontal").pack(fill="x", pady=5)
        
        o_fr = ttk.Frame(lf0); o_fr.pack(fill="x")
        tk.Label(o_fr, text="● 병합 순서:", font=("맑은 고딕", 9, "bold")).pack(side="left", padx=5)
        tk.Radiobutton(o_fr, text="현황창 순서", variable=self.merge_order_var, value="list").pack(side="left", padx=10)
        tk.Radiobutton(o_fr, text="현황창 역순", variable=self.merge_order_var, value="reverse").pack(side="left", padx=10)
        tk.Radiobutton(o_fr, text="수동(위/아래 이동)", variable=self.merge_order_var, value="manual").pack(side="left", padx=10)

        # 1. Input
        lf1 = ttk.LabelFrame(p, text="1. 파일/폴더 관리 (Source Management)", padding=10)
        lf1.pack(fill="x", pady=5)
        bf = ttk.Frame(lf1); bf.pack(fill="x")
        ttk.Button(bf, text="📂 파일 추가", command=self._add_files).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(bf, text="📁 폴더 추가", command=self._add_folder).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(bf, text="❌ 선택 삭제", command=self._remove_selected).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(bf, text="▲ 위로", width=6, command=lambda: self._move_item(-1)).pack(side="left", padx=2)
        ttk.Button(bf, text="▼ 아래로", width=6, command=lambda: self._move_item(1)).pack(side="left", padx=2)
        ttk.Button(bf, text="✨ 전체 초기화", command=self._clear_all).pack(side="right", padx=2)
        ttk.Button(bf, text="🧹 잔류 폴더 정리", command=self._stale_cleanup).pack(side="right", padx=2)
        
        # Listbox with X/Y Scrollbars (Legacy Sophistication)
        lfr = ttk.Frame(lf1)
        lfr.pack(fill="both", expand=True, pady=5)
        
        sy = ttk.Scrollbar(lfr, orient="vertical")
        sx = ttk.Scrollbar(lfr, orient="horizontal")
        self.lst = tk.Listbox(lfr, height=8, yscrollcommand=sy.set, xscrollcommand=sx.set)
        
        sy.config(command=self.lst.yview)
        sx.config(command=self.lst.xview)
        
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.lst.pack(side="left", fill="both", expand=True)
        
        # [v35.5.2] 현황창 선택 시 미리보기 즉시 동기화
        self.lst.bind("<<ListboxSelect>>", lambda e: self._update_previews())

        # 2. Options
        lf2 = ttk.LabelFrame(p, text="2. 환경 설정 (Configuration)", padding=10)
        lf2.pack(fill="x", pady=5)
        
        # Quality
        qf = ttk.Frame(lf2); qf.pack(fill="x", pady=5)
        ttk.Label(qf, text="이미지 품질:", font=("맑은 고딕", 9, "bold")).pack(side="left")
        self.qual_var = tk.IntVar(value=70)
        tk.Scale(qf, from_=30, to=90, orient="horizontal", variable=self.qual_var, length=120).pack(side="left", padx=10)
        ttk.Button(qf, text="Web", command=lambda: self.qual_var.set(50)).pack(side="left", padx=2)
        ttk.Button(qf, text="Med", command=lambda: self.qual_var.set(70)).pack(side="left", padx=2)
        ttk.Button(qf, text="High", command=lambda: self.qual_var.set(85)).pack(side="left", padx=2)
        
        ttk.Separator(lf2, orient="horizontal").pack(fill="x", pady=10)
        
        # Filters Version: v35.4.8 (2026-03-12)
        fil_fr = ttk.Frame(lf2); fil_fr.pack(fill="x", pady=5)
        
        def _append_to_var(var, ext):
            current = var.get().strip()
            exts = [x.strip().lower() for x in current.split(',') if x.strip()]
            new_ext = ext.strip().lower()
            if new_ext not in exts:
                exts.append(new_ext)
            var.set(", ".join(exts))

        def _create_ext_group(parent, var, extensions, title):
            f = ttk.Frame(parent); f.pack(fill="x", pady=1)
            ttk.Label(f, text=f"{title}: ", font=("맑은 고딕", 8, "bold"), foreground="#666666").pack(side="left")
            for ex in extensions:
                l = tk.Label(f, text=ex, font=("맑은 고딕", 8), fg="#0078D7", cursor="hand2")
                l.pack(side="left", padx=2)
                l.bind("<Button-1>", lambda e, x=ex: _append_to_var(var, x))

        ext_sets = {
            "PPT": [".ppt", ".pptx", ".pptm", ".pps", ".ppsx"],
            "Excel": [".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm"],
            "Word": [".doc", ".docx", ".docm"],
            "Temp": [".bak", ".tmp", ".temp"]
        }

        # 1) 기본 제외 확장자
        ex_row = ttk.Frame(fil_fr); ex_row.pack(fill="x", pady=2)
        ttk.Label(ex_row, text="기본 제외 확장자:", font=("맑은 고딕", 9, "bold"), width=16).pack(side="left")
        ttk.Entry(ex_row, textvariable=self.exclude_ext_var, width=25).pack(side="left", padx=5)
        ex_ex_fr = ttk.Frame(fil_fr); ex_ex_fr.pack(fill="x", padx=120)
        for cat, items in ext_sets.items():
            _create_ext_group(ex_ex_fr, self.exclude_ext_var, items, cat)

        ttk.Separator(fil_fr, orient="horizontal").pack(fill="x", pady=5)

        # 2) 기본 포함 확장자
        in_row = ttk.Frame(fil_fr); in_row.pack(fill="x", pady=2)
        ttk.Label(in_row, text="기본 포함 확장자:", font=("맑은 고딕", 9, "bold"), width=16).pack(side="left")
        ttk.Entry(in_row, textvariable=self.include_ext_var, width=25).pack(side="left", padx=5)
        in_ex_fr = ttk.Frame(fil_fr); in_ex_fr.pack(fill="x", padx=120)
        for cat, items in ext_sets.items():
            _create_ext_group(in_ex_fr, self.include_ext_var, items, cat)

        ttk.Separator(lf2, orient="horizontal").pack(fill="x", pady=10)

        # Swiches
        ttk.Checkbutton(lf2, text="스마트 이미지 리사이징", variable=self.resize_var).pack(anchor="w")
        ttk.Checkbutton(lf2, text="XML 구조 경량화 최적화", variable=self.xml_opt_var).pack(anchor="w")
        ttk.Checkbutton(lf2, text="메타데이터/이름/링크 심층 정제", variable=self.clean_meta_var).pack(anchor="w")
        ttk.Checkbutton(lf2, text="사후 무결성 검증 (Verify CRC)", variable=self.verify_var).pack(anchor="w")
        ttk.Checkbutton(lf2, text="오피스 프로세스 강제 정리 (Kill zombies)", variable=self.kill_proc_var).pack(anchor="w")

        # ════════════════════════════════════════════════════════
        # ════════════════════════════════════════════════════════
        # [v35.5.1] 파일명 변경 세부 규칙 (1. 확정 정리 대상)
        # ════════════════════════════════════════════════════════
        f_final = tk.LabelFrame(lf2, text=" [RULE 1] 확정 정리 대상 파일명 규칙 (Finalized Files) ", fg='#004D40', font=('맑은 고딕', 9, 'bold'))
        f_final.pack(fill='x', padx=5, pady=(10, 5))
        
        fr1 = ttk.Frame(f_final)
        fr1.pack(fill='x', padx=5, pady=2)
        ttk.Label(fr1, text="추출(앞n자):").pack(side='left')
        tk.Spinbox(fr1, from_=0, to=10, width=3, textvariable=self.rename_extract_var).pack(side='left', padx=2)
        ttk.Label(fr1, text=" 제외(앞n자):").pack(side='left')
        tk.Spinbox(fr1, from_=0, to=10, width=3, textvariable=self.rename_exclude_var).pack(side='left', padx=2)
        ttk.Label(fr1, text=" 접두사:").pack(side='left')
        ttk.Entry(fr1, width=12, textvariable=self.rename_prefix_var).pack(side='left', padx=2)
        ttk.Label(fr1, text=" 접미사:").pack(side='left')
        ttk.Entry(fr1, width=4, textvariable=self.rename_suffix_var).pack(side='left', padx=2)
        
        fr2 = ttk.Frame(f_final)
        fr2.pack(fill='x', padx=5, pady=2)
        ttk.Label(fr2, text="대상:").pack(side='left')
        tk.Radiobutton(fr2, text="병합파일만(A)", value="A", variable=self.rename_scope_var, bg="#E0F2F1").pack(side='left')
        tk.Radiobutton(fr2, text="폴더내전체(B)", value="B", variable=self.rename_scope_var, bg="#E0F2F1").pack(side='left')
        ttk.Label(fr2, text=" | 시기:").pack(side='left', padx=(10, 0))
        tk.Radiobutton(fr2, text="방금작업만", value="1", variable=self.rename_temporal_var, bg="#E0F2F1").pack(side='left')
        tk.Radiobutton(fr2, text="이전작업포함", value="2", variable=self.rename_temporal_var, bg="#E0F2F1").pack(side='left')
        
        self.lbl_preview_final = tk.Label(f_final, text="미리보기: 01_folder_name.pptx -> 01작업요청서_(folder_name).pptx", 
                                         font=("Consolas", 8), fg="#00695C", bg="#F1F8E9", anchor="w", padx=5)
        self.lbl_preview_final.pack(fill="x", padx=5, pady=5)

        # ════════════════════════════════════════════════════════
        # [v35.5.1] 파일명 변경 세부 규칙 (2. 일반 파일 대상)
        # ════════════════════════════════════════════════════════
        f_gen = tk.LabelFrame(lf2, text=" [RULE 2] 일반 파일 대상 파일명 규칙 (General Files) ", fg='#1565C0', font=('맑은 고딕', 9, 'bold'))
        f_gen.pack(fill='x', padx=5, pady=(5, 10))
        
        gr1 = ttk.Frame(f_gen)
        gr1.pack(fill='x', padx=5, pady=2)
        ttk.Label(gr1, text="추출(앞n자):").pack(side='left')
        tk.Spinbox(gr1, from_=0, to=10, width=3, textvariable=self.gen_rename_extract_var).pack(side='left', padx=2)
        ttk.Label(gr1, text=" 제외(앞n자):").pack(side='left')
        tk.Spinbox(gr1, from_=0, to=10, width=3, textvariable=self.gen_rename_exclude_var).pack(side='left', padx=2)
        ttk.Label(gr1, text=" 접두사:").pack(side='left')
        ttk.Entry(gr1, width=12, textvariable=self.gen_rename_prefix_var).pack(side='left', padx=2)
        ttk.Label(gr1, text=" 접미사:").pack(side='left')
        ttk.Entry(gr1, width=4, textvariable=self.gen_rename_suffix_var).pack(side='left', padx=2)
        
        gr2 = ttk.Frame(f_gen)
        gr2.pack(fill='x', padx=5, pady=2)
        ttk.Label(gr2, text="대상:").pack(side='left')
        tk.Radiobutton(gr2, text="병합파일만(A)", value="A", variable=self.gen_rename_scope_var, bg="#E3F2FD").pack(side='left')
        tk.Radiobutton(gr2, text="폴더내전체(B)", value="B", variable=self.gen_rename_scope_var, bg="#E3F2FD").pack(side='left')
        ttk.Label(gr2, text=" | 시기:").pack(side='left', padx=(10, 0))
        tk.Radiobutton(gr2, text="방금작업만", value="1", variable=self.gen_rename_temporal_var, bg="#E3F2FD").pack(side='left')
        tk.Radiobutton(gr2, text="이전작업포함", value="2", variable=self.gen_rename_temporal_var, bg="#E3F2FD").pack(side='left')
        
        self.lbl_preview_gen = tk.Label(f_gen, text="미리보기: 01_folder_name.pptx -> 01당사안_(folder_name).pptx", 
                                       font=("Consolas", 8), fg="#1565C0", bg="#E3F2FD", anchor="w", padx=5)
        self.lbl_preview_gen.pack(fill="x", padx=5, pady=5)

        # ════════════════════════════════════════════════════════
        # [v35.6.0] 고급 명칭 편집 규칙 (3. Standalone Editor)
        # ════════════════════════════════════════════════════════
        f_adv = tk.LabelFrame(lf2, text=" [RULE 3] 고급 명칭 편집 (Standalone Advanced Editor) ", fg='#7B1FA2', font=('맑은 고딕', 9, 'bold'))
        f_adv.pack(fill='x', padx=5, pady=(5, 10))
        
        ar1 = ttk.Frame(f_adv); ar1.pack(fill='x', padx=5, pady=2)
        ttk.Label(ar1, text="방향:").pack(side='left')
        ttk.Combobox(ar1, textvariable=self.adv_rename_dir_var, values=["앞", "뒤"], width=4).pack(side='left', padx=5)
        
        ttk.Label(ar1, text=" | [제거] ").pack(side='left', padx=(5,0))
        tk.Spinbox(ar1, from_=0, to=100, width=3, textvariable=self.adv_remove_pos_var).pack(side='left', padx=2)
        ttk.Label(ar1, text="번째부터 ").pack(side='left')
        tk.Spinbox(ar1, from_=0, to=100, width=3, textvariable=self.adv_remove_len_var).pack(side='left', padx=2)
        ttk.Label(ar1, text="글자 삭제").pack(side='left')
        
        ar2 = ttk.Frame(f_adv); ar2.pack(fill='x', padx=5, pady=2)
        ttk.Label(ar2, text="[대체] ").pack(side='left')
        tk.Spinbox(ar2, from_=0, to=100, width=3, textvariable=self.adv_replace_start_var).pack(side='left', padx=2)
        ttk.Label(ar2, text="~").pack(side='left')
        tk.Spinbox(ar2, from_=0, to=100, width=3, textvariable=self.adv_replace_end_var).pack(side='left', padx=2)
        ttk.Label(ar2, text="글자를 ").pack(side='left')
        ttk.Entry(ar2, width=8, textvariable=self.adv_replace_str_var).pack(side='left', padx=2)
        ttk.Label(ar2, text=" 로 변경").pack(side='left')
        
        ttk.Label(ar2, text=" | [교체] ").pack(side='left', padx=(5,0))
        ttk.Entry(ar2, width=8, textvariable=self.adv_find_str_var).pack(side='left', padx=2)
        ttk.Label(ar2, text=" 을 ").pack(side='left')
        ttk.Entry(ar2, width=8, textvariable=self.adv_replace_with_var).pack(side='left', padx=2)
        ttk.Label(ar2, text=" 로").pack(side='left')

        ar3 = ttk.Frame(f_adv); ar3.pack(fill='x', padx=5, pady=2)
        ttk.Label(ar3, text="[삽입] ").pack(side='left')
        tk.Spinbox(ar3, from_=0, to=100, width=3, textvariable=self.adv_insert_start_var).pack(side='left', padx=2)
        ttk.Label(ar3, text="~").pack(side='left')
        tk.Spinbox(ar3, from_=0, to=100, width=3, textvariable=self.adv_insert_end_var).pack(side='left', padx=2)
        ttk.Label(ar3, text="위치에 ").pack(side='left')
        ttk.Entry(ar3, width=15, textvariable=self.adv_insert_str_var).pack(side='left', padx=2)
        ttk.Label(ar3, text=" 삽입").pack(side='left')

        self.lbl_preview_adv = tk.Label(f_adv, text="미리보기: sample.pptx -> sample.pptx", 
                                       font=("Consolas", 8), fg="#7B1FA2", bg="#F3E5F5", anchor="w", padx=5)
        self.lbl_preview_adv.pack(fill="x", padx=5, pady=5)

        # 3. Log
        title_fr3 = ttk.Frame(p)
        ttk.Label(title_fr3, text="3. 작업 모니터링 (Execution Log)", font=("맑은 고딕", 10, "bold")).pack(side="left")
        self.lbl_progress = ttk.Label(title_fr3, text="", foreground="#d32f2f", font=("맑은 고딕", 9, "bold"))
        self.lbl_progress.pack(side="right", padx=10)

        lf3 = ttk.LabelFrame(p, labelwidget=title_fr3, padding=10)
        lf3.pack(fill="both", expand=True, pady=10)
        
        # Log with Vertical Scrollbar (Requirement)
        log_fr = ttk.Frame(lf3)
        log_fr.pack(fill="both", expand=True)
        log_sb = ttk.Scrollbar(log_fr, orient="vertical")
        self.log_txt = tk.Text(log_fr, height=12, state="disabled", font=("Consolas", 9), bg="#FAFAFA", yscrollcommand=log_sb.set)
        log_sb.config(command=self.log_txt.yview)
        
        log_sb.pack(side="right", fill="y")
        self.log_txt.pack(side="left", fill="both", expand=True)

    def log(self, msg):
        if self.log_txt:
            self.log_txt.config(state="normal")
            self.log_txt.insert("end", f"{msg}\n")
            self.log_txt.see("end")
            self.log_txt.config(state="disabled")

    def _add_files(self):
        fs = filedialog.askopenfilenames()
        if fs:
            ex_val = self.exclude_ext_var.get()
            in_val = self.include_ext_var.get()
            for f in fs:
                p = os.path.abspath(f)
                ext = os.path.splitext(p)[1].lower().strip()
                if self.service._is_excluded(ext, ex_val, in_val): continue
                if p not in self.files:
                    self.files.append(p); self.lst.insert("end", p)

    def _add_folder(self):
        d = filedialog.askdirectory()
        if d:
            ex_val = self.exclude_ext_var.get()
            in_val = self.include_ext_var.get()
            for r, _, fs in os.walk(d):
                for f in fs:
                    if f.lower().endswith(tuple(SUPPORTED_EXTS)):
                        path = os.path.abspath(os.path.join(r, f))
                        ext = os.path.splitext(path)[1].lower().strip()
                        if self.service._is_excluded(ext, ex_val, in_val): continue
                        if path not in self.files:
                            self.files.append(path); self.lst.insert("end", path)

    def _remove_selected(self):
        idxs = sorted(self.lst.curselection(), reverse=True)
        for i in idxs: self.lst.delete(i); self.files.pop(i)

    def _move_item(self, direction):
        """[v35.5.0] 목록 내 아이템 순서 수동 조정"""
        idx = self.lst.curselection()
        if not idx: return
        i = idx[0]
        n_idx = i + direction
        if 0 <= n_idx < len(self.files):
            # 내부 목록 스왑
            self.files[i], self.files[n_idx] = self.files[n_idx], self.files[i]
            # UI 목록 스왑
            txt = self.lst.get(i)
            self.lst.delete(i)
            self.lst.insert(n_idx, txt)
            self.lst.selection_set(n_idx)
            # [v35.5.0] 수동 이동 시 "수동" 모드로 자동 전환
            self.merge_order_var.set("manual")

    def _on_merge_order_changed(self, *args):
        """[v35.5.0] 병합 순서 변경 시 현황창 실시간 동기화"""
        order = self.merge_order_var.get()
        if order == "manual": return # 수동 모드는 정렬하지 않음
        
        if order == "list":
            self.files.sort()
        elif order == "reverse":
            self.files.sort(reverse=True)
            
        self._refresh_list()

    def _refresh_list(self):
        """내부 self.files 기반으로 UI 리스트박스 강제 갱신"""
        if not hasattr(self, 'lst') or self.lst is None: return
        self.lst.delete(0, tk.END)
        for f in self.files:
            self.lst.insert(tk.END, f)

    def _update_previews(self):
        """[v35.5.2] 실제 현황창 선택 항목 기반 실시간 미리보기 업데이트"""
        # 0. 실제 샘플 추출 (현황창 선택 항목 우선, 없으면 1순위 항목, 그것도 없으면 기본값)
        sample_path = ""
        try:
            sel = self.lst.curselection()
            if sel:
                sample_path = self.lst.get(sel[0])
            elif self.files:
                sample_path = self.files[0]
        except: pass

        if sample_path:
            folder = os.path.basename(os.path.dirname(sample_path))
            filename = os.path.basename(sample_path)
            name_only, ext = os.path.splitext(filename)
            # 만약 폴더명이 비어있으면 (루트 경로 등) 파일명 자체를 샘플로 사용
            if not folder: folder = name_only
        else:
            folder = "sample_folder"
            name_only = "sample_file"
            ext = ".pptx"

        # 1. 확정 정리 대상 프리뷰
        e1, x1, p1, s1 = self.rename_extract_var.get(), self.rename_exclude_var.get(), \
                         self.rename_prefix_var.get(), self.rename_suffix_var.get()
        extracted1 = folder[:e1] if e1 > 0 else ""
        remainder1 = folder[x1:] if x1 < len(folder) else ""
        res1 = f"{extracted1}{p1}{remainder1}{s1}{ext}"
        if self.lbl_preview_final:
            self.lbl_preview_final.config(text=f"미리보기: {os.path.basename(sample_path) if sample_path else 'sample'} -> {res1}")
            
        # 2. 일반 파일 대상 프리뷰
        e2, x2, p2, s2 = self.gen_rename_extract_var.get(), self.gen_rename_exclude_var.get(), \
                         self.gen_rename_prefix_var.get(), self.gen_rename_suffix_var.get()
        extracted2 = folder[:e2] if e2 > 0 else ""
        # 일반 파일은 '추출' 글자를 기존 파일명 앞에 삽입하는 규칙 (generate_new_name 로직 동기화)
        res2 = f"{extracted2}{name_only}{ext}"
        if self.lbl_preview_gen:
            self.lbl_preview_gen.config(text=f"미리보기: {os.path.basename(sample_path) if sample_path else 'sample'} -> {res2}")

        # 3. [v35.7.0] 고급 편집 미리보기 고도화
        try:
            adv_rule = {
                'dir': self.adv_rename_dir_var.get(),
                'rem_pos': self.adv_remove_pos_var.get(),
                'rem_len': self.adv_remove_len_var.get(),
                'rep_start': self.adv_replace_start_var.get(),
                'rep_end': self.adv_replace_end_var.get(),
                'rep_str': self.adv_replace_str_var.get(),
                'find_str': self.adv_find_str_var.get(),
                'replace_with': self.adv_replace_with_var.get(),
                'ins_start': self.adv_insert_start_var.get(),
                'ins_end': self.adv_insert_end_var.get(),
                'ins_str': self.adv_insert_str_var.get()
            }
            # 실제 파일명 샘플 (확장자 유지)
            fname_adv = os.path.basename(sample_path) if sample_path else "sample_file.pptx"
            new_adv = self.service.rename_service.apply_advanced_rules(fname_adv, adv_rule)
            if self.lbl_preview_adv:
                self.lbl_preview_adv.config(text=f"미리보기: {fname_adv} -> {new_adv}")
        except Exception as e:
            if self.lbl_preview_adv:
                self.lbl_preview_adv.config(text=f"미리보기: (오류-{e})")

    def _run_renaming(self):
        """[v35.5.0] 파일명 변경 절차 진입 (v35.5.1 대상별 옵션 분리 적용)"""
        dialog = RenameTargetDialog(self.root)
        self.root.wait_window(dialog.top)
        target_mode = dialog.result
        
        if not target_mode: return
        
        # [v35.5.1] 대상 모드에 따라 전용 옵션 세트 구성
        if target_mode == 'final':
            rule_opts = {
                'extract_len': self.rename_extract_var.get(),
                'exclude_len': self.rename_exclude_var.get(),
                'prefix': self.rename_prefix_var.get(),
                'suffix': self.rename_suffix_var.get(),
                'scope': self.rename_scope_var.get(),
                'temporal': self.rename_temporal_var.get()
            }
        else: # general
            rule_opts = {
                'extract_len': self.gen_rename_extract_var.get(),
                'exclude_len': self.gen_rename_exclude_var.get(),
                'prefix': self.gen_rename_prefix_var.get(),
                'suffix': self.gen_rename_suffix_var.get(),
                'scope': self.gen_rename_scope_var.get(),
                'temporal': self.gen_rename_temporal_var.get()
            }
        
        options = {
            'rename_rule': rule_opts,
            'current_files': self.files,
            'target_mode': target_mode
        }
        
        s, f = self.service.rename_files_by_rule(options)
        messagebox.showinfo("완료", f"파일명 변경 완료 ({'확정정리대상' if target_mode=='final' else '일반파일대상'})\n성공: {s}건\n실패: {f}건")

    def _run_advanced_renaming(self):
        """[v35.6.0] 고급 명칭 일괄 편집 (Standalone)"""
        if not self.files:
            messagebox.showwarning("경고", "현황창에 처리할 파일이 없습니다.")
            return
            
        msg = f"현황창의 {len(self.files)}개 파일에 대해 고급 이름 편집을 실행하시겠습니까?\n\n"
        msg += "⚠️ 주의: 실제 파일명이 즉시 변경됩니다.\n"
        msg += "(변경 전후 미리보기를 꼭 확인하세요!)"
        if not messagebox.askyesno("확인", msg):
            return
            
        adv_rule = {
            'dir': self.adv_rename_dir_var.get(),
            'rem_pos': self.adv_remove_pos_var.get(),
            'rem_len': self.adv_remove_len_var.get(),
            'rep_start': self.adv_replace_start_var.get(),
            'rep_end': self.adv_replace_end_var.get(),
            'rep_str': self.adv_replace_str_var.get(),
            'find_str': self.adv_find_str_var.get(),
            'replace_with': self.adv_replace_with_var.get(),
            'ins_start': self.adv_insert_start_var.get(),
            'ins_end': self.adv_insert_end_var.get(),
            'ins_str': self.adv_insert_str_var.get()
        }
        
        def task():
            if hasattr(self, 'btn_adv_rename') and self.btn_adv_rename:
                self.btn_adv_rename.config(state="disabled")
            try:
                self.service.run_advanced_renaming(self.files, adv_rule)
                self.root.after(0, self._refresh_list)
                self.root.after(0, lambda: messagebox.showinfo("완료", "고급 명칭 일괄 편집이 완료되었습니다."))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"변경 중 오류 발생: {e}"))
            finally:
                if hasattr(self, 'btn_adv_rename') and self.btn_adv_rename:
                    self.root.after(0, lambda: self.btn_adv_rename.config(state="normal"))
        
        threading.Thread(target=task, daemon=True).start()

    def _clear_all(self):
        """[v35.2.6] 모든 목록 및 로그 초기화"""
        if not self.files and not self.log_txt.get("1.0", "end-1c").strip(): return
        if messagebox.askyesno("초기화", "모든 파일 목록과 작업 로그를 초기화하시겠습니까?"):
            self.files = []
            self.lst.delete(0, tk.END)
            self.log_txt.config(state="normal")
            self.log_txt.delete("1.0", tk.END)
            self.log_txt.config(state="disabled")
            self._ui_callback("status", "초기화 완")

    def _run(self):
        if not self.files: messagebox.showwarning("경고", "대상을 추가해 주세요."); return
        self.btn_run.config(state="disabled"); self.btn_finish.config(state="disabled")
        opts = {
            'quality': self.qual_var.get(), 'resize': self.resize_var.get(), 'xml_opt': self.xml_opt_var.get(),
            'clean_meta': self.clean_meta_var.get(), 'exclude_ext': self.exclude_ext_var.get(),
            'include_ext': self.include_ext_var.get(),
            'verify': self.verify_var.get(), 'kill_proc': self.kill_proc_var.get(),
            'merge_order': self.merge_order_var.get()
        }
        mode = self.mode_var.get()
        def task():
            try:
                if mode == "optimize": self.service.run_optimization(self.files, opts)
                else: self.service.run_merging(self.files, opts)
            finally:
                # [v35.3.8] 쓰레드 종료 전 CoUninitialize 강제로 리소스 완벽 정리
                try: pythoncom.CoUninitialize()
                except: pass
                # UI 버튼 복구
                self.root.after(0, lambda: self.btn_run.config(state="normal"))
                self.root.after(0, lambda: self.btn_finish.config(state="normal"))
        
        threading.Thread(target=task, daemon=True).start()

    def _on_finalize(self):
        """[v35.4.0] 확정 정리 전 무결성 리포트 및 사용자 최종 확인 절차"""
        info = self.service.get_finalize_info()
        if not info:
            messagebox.showwarning("경고", "교체할 완료 파일이 없습니다. 먼저 작업을 수행해 주세요.")
            return
            
        msg = f"▣ 이전에 수행된 작업: {info['mode_name']}\n"
        msg += f"▣ 생성된 파일 개수: {info['count']}개\n"
        msg += f"▣ 무결성 점검 결과: {info['v_status']}\n\n"
        msg += "⚠️ 주의: '확인'을 누르면 원본 파일이 생성된 파일로 교체됩니다.\n\n"
        msg += f"[{info['mode_name']}] 완료된 파일을 실제 열어 이상 유무를 확인하셨습니까?"
        
        if messagebox.askyesno("확정 정리 최종 승인", msg):
            self.service.finalize_cleanup()
            messagebox.showinfo("완료", "원본 교체 및 임시 폴더 정리가 완료되었습니다.")

    def _stale_cleanup(self):
        """[v35.4.0] 잔류 폴더 전수 점검 (깊이 우선 스캔 및 누락 방지)"""
        d = filedialog.askdirectory(title="잔류 폴더 스캔 및 정리 (Deep Scan)")
        if not d: return
        found = []
        prefixes = ("00_Optimized_Docs", "00_Merged_Docs")
        
        self.log(f"[SCAN] '{d}' 폴더 하위 전수 스캔 중...")
        for r, ds, _ in os.walk(d):
            to_remove = []
            for d_name in ds:
                if any(d_name.startswith(p) for p in prefixes):
                    full_p = os.path.join(r, d_name)
                    found.append(full_p)
                    to_remove.append(d_name)
            
            # 발견된 폴더 내부는 더 이상 스캔할 필요 없으므로 리스트에서 제외 (Pruning)
            for tr in to_remove:
                ds.remove(tr)
                
        if not found: 
            messagebox.showinfo("알림", "정리할 잔류 폴더가 발견되지 않았습니다.")
            return
            
        if messagebox.askyesno("전수 정리", f"하위 폴더 포함 총 {len(found)}개의 임시 폴더가 발견되었습니다.\n모두 삭제하시겠습니까?"):
            for p in found:
                try: 
                    shutil.rmtree(p, ignore_errors=True)
                    self.log(f"  [DELETED] {p}")
                except: pass
            messagebox.showinfo("완료", f"{len(found)}개 잔류 폴더 전수 정리 완료")

class RenameTargetDialog:
    """[v35.5.0] 파일명 변경 대상을 선택하기 위한 커스텀 다이얼로그"""
    def __init__(self, parent):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title("대상 선택")
        self.top.geometry("350x180")
        self.top.resizable(False, False)
        self.top.grab_set()
        
        tk.Label(self.top, text="파일명 변경을 적용할 대상을 선택하세요.", font=("맑은 고딕", 10)).pack(pady=20)
        
        btn_fr = ttk.Frame(self.top)
        btn_fr.pack(pady=10)
        
        tk.Button(btn_fr, text="✨ 확정 정리 대상", bg="#455A64", fg="white", width=14, 
                  command=lambda: self._close("final")).pack(side="left", padx=5)
        tk.Button(btn_fr, text="📁 일반 파일 대상", bg="#1976D2", fg="white", width=14, 
                  command=lambda: self._close("general")).pack(side="left", padx=5)
        tk.Button(btn_fr, text="취소", width=8, 
                  command=lambda: self._close(None)).pack(side="left", padx=5)
        
        # 중앙 배치
        self.top.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (self.top.winfo_width() // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.top.winfo_height() // 2)
        self.top.geometry(f"+{x}+{y}")

    def _close(self, val):
        self.result = val
        self.top.destroy()

if __name__ == "__main__":
    tk_root = tk.Tk()
    app = OptimizerGUI(tk_root)
    tk_root.mainloop()
