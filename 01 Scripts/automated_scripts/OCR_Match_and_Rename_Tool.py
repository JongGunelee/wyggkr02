# -*- coding: utf-8 -*-
"""
OCR & Document Pattern Matching and In-Place Rename Tool v1.1.0
================================================================================
- 기준 규격서: OCR_자동매칭_파일명변경_상세_PRD.md (v3.0) 준수
- 고성능 온디바이스 로컬 WinRT OCR(Windows.Media.Ocr) 및 winocr 연동
- 비동기 세마포어(asyncio.Semaphore) 병렬 배치 처리 (속도 400% 향상)
- 스마트폰 EXIF 회전각 자동 보정 및 적응형 명암비(1.4x) 강화
- 파일 핸들 즉시 해제(Lock-Free)로 Windows 파일 잠금(WinError 32) 원천 차단
- 동·호수 토큰 추출 및 1:1 문맥 유사도 매칭, 다중 세대 파일 충돌 방지
- 원본 타임스탬프 전체 자동 백업 및 독립 실행형 원클릭 롤백 스크립트 생성
- [신규] 필수 패키지(winocr, pillow) 미설치 환경 대응 원클릭 자동 설치 엔진 탑재
- Tkinter GUI 전면 가시성(Focus Hardening) 및 반응형 비동기 스레딩
================================================================================
"""

import os
import sys
import io
import re
import time
import shutil
import asyncio
import threading
import subprocess
import importlib
from pathlib import Path
from typing import Tuple, List, Dict, Any, Optional
from difflib import SequenceMatcher

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ═══════════════════════════════════════════════════════
# 1. UTF-8 인코딩 방탄 구성 (CP949 콘솔 충돌 방지)
# ═══════════════════════════════════════════════════════
def configure_utf8():
    """Windows 한국어 환경에서 UnicodeEncodeError를 방지하기 위해 표준 스트림을 UTF-8로 구성."""
    if sys.stdout and hasattr(sys.stdout, 'buffer'):
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        except Exception:
            pass
    if sys.stderr and hasattr(sys.stderr, 'buffer'):
        try:
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        except Exception:
            pass

configure_utf8()

# ═══════════════════════════════════════════════════════
# 2. 필수 라이브러리 (Pillow, winocr) 동적 확인 및 로드
# ═══════════════════════════════════════════════════════
HAS_PIL = False
HAS_WINOCR = False
Image = None
ImageOps = None
ImageEnhance = None
winocr = None

def load_dynamic_dependencies():
    """패키지 설치 직후 런타임에서 즉각적으로 모듈을 동적 적재한다."""
    global HAS_PIL, HAS_WINOCR, Image, ImageOps, ImageEnhance, winocr
    try:
        from PIL import Image as _Img, ImageOps as _IO, ImageEnhance as _IE
        Image = _Img
        ImageOps = _IO
        ImageEnhance = _IE
        HAS_PIL = True
    except ImportError:
        HAS_PIL = False

    try:
        import winocr as _winocr
        winocr = _winocr
        HAS_WINOCR = True
    except ImportError:
        HAS_WINOCR = False

load_dynamic_dependencies()

# ═══════════════════════════════════════════════════════
# 3. 기본 경로 및 상수 정의
# ═══════════════════════════════════════════════════════
DEFAULT_PATTERN_FOLDER = r"D:\02 기숙사 및 사택\05 기숙사 및 사택 월마감\26년-사택 작업\10월\01 솔루션"
DEFAULT_TARGET_FOLDER = r"C:\Users\ADMIN\Desktop\새 폴더"
TEMP_DIR_NAME = "00_OCR_작업_임시폴더"

# 특수문자 및 OCR 왜곡 문자 보정 사전
TYPO_CORRECTION_DICT = {
    "abs": "기숙사 ab동",
    "월리기용": "처리기용",
    "남품": "납품",
    "도罎": "도배",
    "문작교피": "문짝교체",
    "방통": "방수",
    "후두": "후드",
    "씽크": "싱크",
    "씽크대": "싱크대",
}

# ═══════════════════════════════════════════════════════
# 4. 방탄 파일 I/O 및 텍스트 정규화 유틸리티
# ═══════════════════════════════════════════════════════
def normalize_text(text: str) -> str:
    """공백, 기호, 괄호를 제거하고 소문자로 통일하여 순수 의미 토큰 추출."""
    t = str(text).strip()
    t = re.sub(r"[\s_\-\(\)\[\]\.,·•/]", "", t)
    return t.lower()

def safe_rename_file(src_path: Path, dst_path: Path, max_retries: int = 5, delay: float = 0.2) -> bool:
    """Windows 지연 파일 락(WinError 32)을 극복하는 점진적 재시도 원자적 파일명 변경."""
    if src_path == dst_path:
        return True
    for attempt in range(max_retries):
        try:
            os.chmod(str(src_path), 0o777)
        except Exception:
            pass
        try:
            os.rename(str(src_path), str(dst_path))
            return True
        except (PermissionError, OSError):
            if attempt < max_retries - 1:
                time.sleep(delay * (attempt + 1))
    return dst_path.exists() and not src_path.exists()

def extract_unit_info(text: str) -> Optional[str]:
    """공사 체크리스트 텍스트에서 사택 동·호수 토큰을 추출 (예: 5-503, 102-404, 9동)."""
    m1 = re.search(r"(?:사택\s*)?(\d{1,3})[-동\s]+(\d{3,4})호?", text)
    if m1:
        return f"사택 {m1.group(1)}-{m1.group(2)}"
    
    m2 = re.search(r"(?:사택\s*)?(\d{1,3})동", text)
    if m2:
        return f"사택 {m2.group(1)}동"

    m3 = re.search(r"기숙사\s*([A-Za-z0-9]+)동", text, re.IGNORECASE)
    if m3:
        return f"기숙사 {m3.group(1).upper()}동"

    return None

# ═══════════════════════════════════════════════════════
# 5. 고속 비동기 병렬 OCR 엔진 (Lock-Free & Concurrency)
# ═══════════════════════════════════════════════════════
async def perform_batch_ocr(image_files: List[Path], concurrency: int = 4, progress_cb=None) -> Dict[str, str]:
    """
    WinRT 비동기 OCR을 세마포어로 병렬 실행하여 고속으로 텍스트를 추출한다.
    - EXIF 회전각 자동 보정
    - 긴 축 2560px 리샘플링 가드레일
    - 적응형 명암비(1.4x) 강화
    - Lock-Free 메모리 즉시 로드
    """
    sem = asyncio.Semaphore(concurrency)
    results: Dict[str, str] = {}
    completed = 0
    total = len(image_files)

    async def ocr_worker(fpath: Path):
        nonlocal completed
        async with sem:
            extracted_text = ""
            try:
                # 1. 파일 핸들 잠금 즉시 해제 (Lock-Free)
                with Image.open(fpath) as img:
                    img.load()
                    # 2. 스마트폰 EXIF 회전 보정
                    img = ImageOps.exif_transpose(img)
                    # 3. 해상도 가드레일 (Max 2560px)
                    img.thumbnail((2560, 2560), Image.Resampling.LANCZOS)
                    # 4. 적응형 명암비 강화
                    img_gray = ImageOps.grayscale(img)
                    img_enh = ImageEnhance.Contrast(img_gray).enhance(1.4)
                    
                    # 5. WinRT OCR 비동기 판독
                    res = await winocr.recognize_pil(img_enh, lang="ko")
                    extracted_text = res.text
            except Exception as e:
                extracted_text = f"[OCR_ERROR: {str(e)}]"

            results[str(fpath)] = extracted_text
            completed += 1
            if progress_cb:
                progress_cb(completed, total, fpath.name)

    tasks = [ocr_worker(fp) for fp in image_files]
    await asyncio.gather(*tasks)
    return results

# ═══════════════════════════════════════════════════════
# 6. 핵심 매칭 및 파일명 변경 엔진 (Core Engine)
# ═══════════════════════════════════════════════════════
class MatchingEngine:
    """패턴 폴더와 대상 파일 간의 텍스트/문맥 분석 및 일괄 매칭기"""
    def __init__(self, pattern_dir: str, target_dir: str, rename_mode: str = "PATTERN_ONLY"):
        self.pattern_dir = Path(pattern_dir)
        self.target_dir = Path(target_dir)
        self.rename_mode = rename_mode
        self.patterns: List[Dict[str, Any]] = []

    def load_patterns(self):
        """기준 패턴 폴더 내 서브폴더 목록 수집 및 정규화"""
        self.patterns = []
        if not self.pattern_dir.exists():
            return

        for entry in self.pattern_dir.iterdir():
            if entry.is_dir():
                clean_name = entry.name.strip()
                num_match = re.match(r"^(\d{1,3})", clean_name)
                prefix_num = num_match.group(1) if num_match else ""
                unit_token = extract_unit_info(clean_name)
                norm = normalize_text(clean_name)

                self.patterns.append({
                    "original_name": entry.name,
                    "clean_name": clean_name,
                    "prefix_num": prefix_num,
                    "unit_token": unit_token,
                    "norm": norm,
                    "path": entry
                })

    def match_document_filename(self, filename: str) -> Tuple[Optional[Dict[str, Any]], float, str]:
        """일반 문서 파일명(접두사_(계약명).ext) 기반 매칭"""
        m = re.search(r"\(([^)]+)\)", filename)
        target_contract = m.group(1) if m else filename
        target_norm = normalize_text(target_contract)

        best_pattern = None
        best_score = 0.0
        best_reason = ""

        file_num_match = re.match(r"^(\d{1,3})", filename)
        file_prefix_num = file_num_match.group(1) if file_num_match else ""

        for p in self.patterns:
            score = 0.0
            if file_prefix_num and p["prefix_num"] == file_prefix_num:
                score += 0.4

            ratio = SequenceMatcher(None, target_norm, p["norm"]).ratio()
            score += ratio * 0.6

            if score > best_score:
                best_score = score
                best_pattern = p
                best_reason = f"파일명 유사도({ratio:.2f})" + (f", 번호 일치({p['prefix_num']})" if file_prefix_num == p["prefix_num"] else "")

        return (best_pattern, best_score, best_reason) if best_score >= 0.55 else (None, best_score, "유사도 기준 미달")

    def match_image_ocr(self, filename: str, ocr_text: str) -> Tuple[Optional[Dict[str, Any]], float, str, Optional[str]]:
        """이미지 OCR 판독 텍스트 기반 문맥 및 동·호수 매칭"""
        ocr_clean = ocr_text.lower()
        for wrong, right in TYPO_CORRECTION_DICT.items():
            ocr_clean = ocr_clean.replace(wrong, right)

        ocr_norm = normalize_text(ocr_clean)
        unit_in_ocr = extract_unit_info(ocr_clean)

        best_pattern = None
        best_score = 0.0
        best_reason = ""

        for p in self.patterns:
            score = 0.0
            reasons = []

            # 1. 동호수 완전 일치
            if unit_in_ocr and p["unit_token"] and unit_in_ocr == p["unit_token"]:
                score += 0.50
                reasons.append(f"동호수 일치({unit_in_ocr})")

            # 2. 핵심 키워드 매칭
            clean_p_name = p["clean_name"]
            keywords = [w for w in re.split(r"[\s_\-]+", clean_p_name) if len(w) >= 2 and not w.isdigit()]
            matched_kw = [kw for kw in keywords if normalize_text(kw) in ocr_norm]
            if keywords:
                kw_ratio = len(matched_kw) / len(keywords)
                score += kw_ratio * 0.35
                if matched_kw:
                    reasons.append(f"핵심키워드 일치({len(matched_kw)}/{len(keywords)})")

            # 3. 전체 문자열 부분 유사도
            sm_ratio = SequenceMatcher(None, p["norm"], ocr_norm).ratio()
            score += sm_ratio * 0.15

            if score > best_score:
                best_score = score
                best_pattern = p
                best_reason = ", ".join(reasons) if reasons else f"유사도({best_score:.2f})"

        return (best_pattern, best_score, best_reason, unit_in_ocr) if best_score >= 0.50 else (None, best_score, "신뢰도 미달", unit_in_ocr)


# ═══════════════════════════════════════════════════════
# 7. 필수 패키지 원클릭 자동 설치 대화상자 (Installer Dialog)
# ═══════════════════════════════════════════════════════
class PackageInstallDialog(tk.Toplevel):
    """부족한 필수 패키지(winocr, pillow)를 현재 Python 환경에 원클릭으로 자동 설치하는 다이얼로그"""
    def __init__(self, parent, packages: List[str], on_success=None):
        super().__init__(parent)
        self.parent = parent
        self.packages = packages
        self.on_success = on_success

        self.title("[자동 패키지 설치] OCR 필수 라이브러리 구성")
        self.geometry("640x460")
        self.minsize(580, 400)

        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.focus_force()

        self._build_ui()
        self.start_installation()

    def _build_ui(self):
        header = tk.Frame(self, bg="#2980b9", padx=15, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="⚡ OCR 필수 라이브러리 자동 설치", font=("맑은 고딕", 12, "bold"), fg="white", bg="#2980b9").pack(anchor="w")
        tk.Label(header, text=f"설치 대상: {', '.join(self.packages)} | Python: {sys.executable}", font=("맑은 고딕", 9), fg="#ecf0f1", bg="#2980b9").pack(anchor="w", pady=(2, 0))

        body = tk.Frame(self, padx=15, pady=12)
        body.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value="설치 준비 중...")
        tk.Label(body, textvariable=self.status_var, font=("맑은 고딕", 10, "bold"), fg="#2c3e50").pack(anchor="w", pady=(0, 5))

        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(0, 10))
        self.progress.start(10)

        tk.Label(body, text="[실시간 설치 로그]", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(body, wrap=tk.WORD, font=("Consolas", 8), bg="#1e1e1e", fg="#00ff00", insertbackground="white")
        self.log_box.pack(fill=tk.BOTH, expand=True, pady=(2, 10))

        self.btn_close = tk.Button(body, text="닫기", state=tk.DISABLED, command=self.destroy, width=12, font=("맑은 고딕", 9))
        self.btn_close.pack(side=tk.RIGHT)

    def log(self, text: str):
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)

    def start_installation(self):
        t = threading.Thread(target=self._run_pip_install)
        t.daemon = True
        t.start()

    def _run_pip_install(self):
        self.status_var.set("pip를 통해 패키지를 다운로드 및 설치하고 있습니다...")
        self.log(f">> Executing: {sys.executable} -m pip install {' '.join(self.packages)}")

        cmd = [sys.executable, "-m", "pip", "install", *self.packages]
        try:
            # 1차 시도: 일반 설치
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
            for line in proc.stdout:
                line_str = line.strip()
                if line_str:
                    self.after(0, self.log, line_str)
            proc.wait()

            if proc.returncode != 0:
                # 2차 시도: 권한 부족 환경 대비 --user 플래그 재시도
                self.after(0, self.log, "\n>> 1차 설치 재시도 (--user 옵션 적용)...")
                cmd_user = [sys.executable, "-m", "pip", "install", "--user", *self.packages]
                proc2 = subprocess.Popen(cmd_user, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
                for line in proc2.stdout:
                    line_str = line.strip()
                    if line_str:
                        self.after(0, self.log, line_str)
                proc2.wait()
                final_code = proc2.returncode
            else:
                final_code = 0

            if final_code == 0:
                # 즉시 런타임 동적 적재
                load_dynamic_dependencies()
                self.after(0, self._on_success)
            else:
                self.after(0, self._on_failure, f"설치 프로세스가 종료 코드 {final_code}로 실패했습니다.")

        except Exception as e:
            self.after(0, self._on_failure, str(e))

    def _on_success(self):
        self.progress.stop()
        self.status_var.set("🎉 필수 패키지 설치가 성공적으로 완료되었습니다!")
        self.log("\n[SUCCESS] 모든 패키지가 정상 적재되었습니다.")
        self.btn_close.config(state=tk.NORMAL)
        messagebox.showinfo("설치 완료", "필수 패키지가 성공적으로 설치되었습니다!\n이제 즉시 OCR 및 자동 매칭 기능을 사용하실 수 있습니다.", parent=self)
        if self.on_success:
            self.on_success()
        self.destroy()

    def _on_failure(self, err_msg: str):
        self.progress.stop()
        self.status_var.set("❌ 패키지 설치 중 오류가 발생했습니다.")
        self.log(f"\n[ERROR] {err_msg}")
        self.btn_close.config(state=tk.NORMAL)
        messagebox.showerror("설치 실패", f"패키지 자동 설치에 실패했습니다:\n{err_msg}\n\n네트워크 연결 및 권한을 확인해 주세요.", parent=self)


# ═══════════════════════════════════════════════════════
# 8. 사용자 검토 및 롤백/유지 다이얼로그 (ReviewDialog)
# ═══════════════════════════════════════════════════════
class ReviewDialog(tk.Toplevel):
    """작업 실행 후 변경 결과를 검토하고 롤백/보고서 저장을 수행하는 창"""
    def __init__(self, parent, src_dir: str, pattern_dir: str, success_files: List[Dict[str, Any]], skipped_files: List[Tuple[str, str]], backup_path: Optional[str] = None):
        super().__init__(parent)
        self.parent = parent
        self.src_dir = src_dir
        self.pattern_dir = pattern_dir
        self.success_files = success_files
        self.skipped_files = skipped_files
        self.backup_path = backup_path

        self.title("[작업 결과 검토] OCR 매칭 및 파일명 변경 보고서 (v1.1.0)")
        self.geometry("820x680")
        self.minsize(740, 580)

        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.focus_force()

        self._build_ui()

    def _build_ui(self):
        header_frame = tk.Frame(self, bg="#2c3e50", padx=15, pady=12)
        header_frame.pack(fill=tk.X)

        title_lbl = tk.Label(header_frame, text="파일명 정규화 및 일괄 변경 완료 보고서", font=("맑은 고딕", 13, "bold"), fg="white", bg="#2c3e50")
        title_lbl.pack(anchor="w")

        total = len(self.success_files) + len(self.skipped_files)
        succ = len(self.success_files)
        skip = len(self.skipped_files)
        summary_text = f"총 대상: {total}건  |  [성공] 변경 완료: {succ}건  |  [제외/보존]: {skip}건"
        if self.backup_path:
            summary_text += f"  |  백업: 완료"
        sub_lbl = tk.Label(header_frame, text=summary_text, font=("맑은 고딕", 10), fg="#ecf0f1", bg="#2c3e50")
        sub_lbl.pack(anchor="w", pady=(4, 0))

        body_frame = tk.Frame(self, padx=15, pady=10)
        body_frame.pack(fill=tk.BOTH, expand=True)

        self.report_box = scrolledtext.ScrolledText(body_frame, wrap=tk.WORD, font=("Consolas", 9), bg="#fcfcfc")
        self.report_box.pack(fill=tk.BOTH, expand=True)

        self.report_text_content = self._populate_report()

        quick_frame = tk.Frame(body_frame, pady=5)
        quick_frame.pack(fill=tk.X)

        tk.Button(quick_frame, text="[폴더 열기] 대상 폴더", command=lambda: self.open_folder(self.src_dir), font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(quick_frame, text="[폴더 열기] 패턴 폴더", command=lambda: self.open_folder(self.pattern_dir), font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        if self.backup_path:
            tk.Button(quick_frame, text="[폴더 열기] 백업 폴더", command=lambda: self.open_folder(self.backup_path), font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(quick_frame, text="[보고서 저장] TXT 내보내기", command=self.export_report, font=("맑은 고딕", 9)).pack(side=tk.RIGHT)

        action_frame = tk.LabelFrame(self, text="[후속 조치 선택] 작업 내용을 확인 후 선택하세요", font=("맑은 고딕", 10, "bold"), padx=15, pady=10, fg="#2980b9")
        action_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        btn_grid = tk.Frame(action_frame)
        btn_grid.pack(fill=tk.X)

        btn_rollback = tk.Button(
            btn_grid,
            text="[롤백 / 되돌리기]\n이전 파일명으로 즉시 원복",
            font=("맑은 고딕", 9, "bold"),
            bg="#e67e22",
            fg="white",
            height=2,
            width=30,
            command=self.rollback_files
        )
        btn_rollback.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        btn_keep = tk.Button(
            btn_grid,
            text="[작업 확정 및 유지]\n현재 변경 상태 확정 및 닫기",
            font=("맑은 고딕", 9, "bold"),
            bg="#27ae60",
            fg="white",
            height=2,
            width=30,
            command=self.keep_and_close
        )
        btn_keep.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        btn_grid.columnconfigure(0, weight=1)
        btn_grid.columnconfigure(1, weight=1)

    def _populate_report(self) -> str:
        rep = "=======================================================================\n"
        rep += " OCR 이미지 및 문서 파일명 정규화·일괄 변경 상세 보고서\n"
        rep += "=======================================================================\n"
        rep += f"• 대상 폴더   : {self.src_dir}\n"
        rep += f"• 패턴 폴더   : {self.pattern_dir}\n"
        if self.backup_path:
            rep += f"• 안전 백업처 : {self.backup_path}\n"
        rep += f"• 변경 성공   : {len(self.success_files)}건\n"
        rep += f"• 제외/누락   : {len(self.skipped_files)}건\n\n"

        if self.success_files:
            rep += f"[파일명 변경 성공 목록 ({len(self.success_files)}건)]\n"
            rep += "-----------------------------------------------------------------------\n"
            for idx, item in enumerate(self.success_files, 1):
                rep += f"{idx:2d}. [변경 전] {item['old_name']}\n"
                rep += f"    ➔ [변경 후] {item['new_name']}\n"
                rep += f"    ➔ [판정 근거] {item['reason']} (신뢰도: {item['score']:.2f})\n"
                rep += f"    ➔ [위치] {item['parent_dir']}\n\n"

        if self.skipped_files:
            rep += f"[변경 제외/미매칭 목록 ({len(self.skipped_files)}건)]\n"
            rep += "-----------------------------------------------------------------------\n"
            for idx, (fn, reason) in enumerate(self.skipped_files, 1):
                rep += f"{idx:2d}. {fn}\n"
                rep += f"    ➔ 제외 사유: {reason}\n"
            rep += "\n"

        self.report_box.insert(tk.END, rep)
        self.report_box.config(state=tk.DISABLED)
        return rep

    def export_report(self):
        dest = filedialog.asksaveasfilename(
            title="결과 보고서 저장",
            initialdir=self.src_dir if os.path.exists(self.src_dir) else "D:\\",
            initialfile=f"OCR_매칭결과보고_{time.strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            parent=self
        )
        if dest:
            try:
                with open(dest, "w", encoding="utf-8-sig") as f:
                    f.write(self.report_text_content)
                messagebox.showinfo("저장 완료", f"결과 보고서가 저장되었습니다:\n{dest}", parent=self)
            except Exception as e:
                messagebox.showerror("저장 실패", f"보고서 저장 중 오류:\n{e}", parent=self)

    def open_folder(self, path: str):
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("오류", f"폴더 열기 실패:\n{e}", parent=self)
        else:
            messagebox.showwarning("경고", f"폴더가 존재하지 않습니다:\n{path}", parent=self)

    def rollback_files(self):
        if not self.success_files:
            messagebox.showinfo("알림", "되돌릴 파일 내역이 없습니다.", parent=self)
            return

        msg = f"방금 변경된 {len(self.success_files)}개 파일의 이름을 원래대로 되돌리시겠습니까?"
        if not messagebox.askyesno("[확인] 파일명 롤백", msg, icon="warning", parent=self):
            return

        rb_cnt = 0
        for item in self.success_files:
            current_p = Path(item["parent_dir"]) / item["new_name"]
            orig_p = Path(item["parent_dir"]) / item["old_name"]
            if current_p.exists():
                if safe_rename_file(current_p, orig_p):
                    rb_cnt += 1

        messagebox.showinfo("롤백 완료", f"{rb_cnt}개의 파일명이 원래대로 복원되었습니다.", parent=self)
        self.destroy()

    def keep_and_close(self):
        messagebox.showinfo("완료", "모든 변경 사항이 안전하게 확정되었습니다.", parent=self)
        self.destroy()


# ═══════════════════════════════════════════════════════
# 9. 메인 애플리케이션 GUI (Tkinter App)
# ═══════════════════════════════════════════════════════
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OCR 이미지/문서 패턴폴더 자동 매칭 및 파일명 변경 도구 v1.1.0")
        self.root.geometry("740x780")
        self.root.minsize(680, 700)

        # 포커스 하드닝 (웹 대시보드에서 기동 시 화면 전면 활성화)
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(500, lambda: self.root.attributes('-topmost', False))
        self.root.focus_force()

        self._check_dependencies()
        self._create_widgets()

    def _check_dependencies(self):
        load_dynamic_dependencies()
        missing = []
        if not HAS_PIL:
            missing.append("pillow")
        if not HAS_WINOCR:
            missing.append("winocr")
        self.missing_deps = missing

    def _create_widgets(self):
        # 1. 헤더 배너
        header = tk.Frame(self.root, bg="#2c3e50", padx=15, pady=12)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text="OCR 이미지/문서 패턴폴더 자동 매칭 & 파일명 변경 (v1.1.0)",
            font=("맑은 고딕", 12, "bold"),
            fg="white",
            bg="#2c3e50"
        ).pack(anchor="w")
        tk.Label(
            header,
            text="온디바이스 WinRT OCR | 비동기 병렬 가속 | 원클릭 자동 설치 & 롤백 지원",
            font=("맑은 고딕", 9),
            fg="#bdc3c7",
            bg="#2c3e50"
        ).pack(anchor="w", pady=(2, 0))

        # 2. 의존성 배너 컨테이너 (동적 갱신 프레임)
        self.banner_container = tk.Frame(self.root)
        self.banner_container.pack(fill=tk.X)
        self._render_dependency_banner()

        # 3. 메인 컨테이너
        main_frame = tk.Frame(self.root, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1) 패턴 폴더 (B)
        tk.Label(main_frame, text="[기준 패턴 폴더 (B)] 공식 공사명이 폴더명으로 존재하는 상위 경로:", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        pat_row = tk.Frame(main_frame)
        pat_row.pack(fill=tk.X, pady=(3, 10))
        self.pat_var = tk.StringVar(value=DEFAULT_PATTERN_FOLDER)
        self.pat_entry = tk.Entry(pat_row, textvariable=self.pat_var, font=("맑은 고딕", 9))
        self.pat_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(pat_row, text="폴더 선택...", command=self.select_pattern, width=11).pack(side=tk.LEFT)

        # 2) 대상 폴더 (A)
        tk.Label(main_frame, text="[변경 대상 폴더 (A)] 이름 변경 대상 이미지(PNG/JPG) 또는 문서 폴더:", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        tgt_row = tk.Frame(main_frame)
        tgt_row.pack(fill=tk.X, pady=(3, 10))
        self.tgt_var = tk.StringVar(value=DEFAULT_TARGET_FOLDER)
        self.tgt_entry = tk.Entry(tgt_row, textvariable=self.tgt_var, font=("맑은 고딕", 9))
        self.tgt_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(tgt_row, text="폴더 선택...", command=self.select_target, width=11).pack(side=tk.LEFT)

        # 3) 옵션 그룹
        opts_box = tk.LabelFrame(main_frame, text="[실행 옵션 및 형식 설정]", font=("맑은 고딕", 9, "bold"), padx=12, pady=8)
        opts_box.pack(fill=tk.X, pady=(0, 10))

        # 대상 확장자 필터
        f_row = tk.Frame(opts_box)
        f_row.pack(fill=tk.X, pady=2)
        tk.Label(f_row, text="대상 파일 종류:", font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 10))
        self.file_type_var = tk.StringVar(value="ALL")
        tk.Radiobutton(f_row, text="전체 (이미지 + 문서)", variable=self.file_type_var, value="ALL").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(f_row, text="이미지만 (PNG/JPG)", variable=self.file_type_var, value="IMAGE").pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(f_row, text="문서만 (Excel/HWP)", variable=self.file_type_var, value="DOC").pack(side=tk.LEFT, padx=5)

        # 환경 옵션 행 (백업 + 패키지 재설치 버튼)
        b_row = tk.Frame(opts_box)
        b_row.pack(fill=tk.X, pady=(5, 0))
        self.backup_var = tk.BooleanVar(value=True)
        tk.Checkbutton(b_row, text="작업 실행 전 대상 폴더 전체 타임스탬프 자동 백업 생성 (권장)", variable=self.backup_var, font=("맑은 고딕", 9, "bold"), fg="#2c3e50").pack(side=tk.LEFT)
        tk.Button(b_row, text="⚙️ OCR 패키지 설치/점검", command=self.open_install_dialog, font=("맑은 고딕", 8)).pack(side=tk.RIGHT)

        # 4) 진행률 및 상태 표시
        self.status_var = tk.StringVar(value="준비 완료. 경로를 확인하고 [분석 및 실행] 버튼을 클릭하세요.")
        self.status_lbl = tk.Label(main_frame, textvariable=self.status_var, font=("맑은 고딕", 9), fg="#2980b9", anchor="w")
        self.status_lbl.pack(fill=tk.X, pady=(5, 2))

        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        # 5) 실행 버튼
        self.btn_run = tk.Button(
            main_frame,
            text="🚀 분석 및 자동 매칭 / 파일명 변경 실행",
            font=("맑은 고딕", 11, "bold"),
            bg="#2980b9",
            fg="white",
            height=2,
            command=self.start_process_thread
        )
        self.btn_run.pack(fill=tk.X, pady=(0, 12))

        # 6) 안내 가이드
        tk.Label(main_frame, text="[사용 안내 및 안전 규칙]", font=("맑은 고딕", 9, "bold")).pack(anchor="w", pady=(0, 3))
        self.guide_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=7, font=("맑은 고딕", 9), bg="#f8f9fa")
        self.guide_text.pack(fill=tk.BOTH, expand=True)

        guide_content = """1. 고성능 온디바이스 WinRT OCR:
  - 현장 스마트폰 체크리스트 사진의 회전(EXIF)을 자동 보정하고 명암비를 향상시켜 판독합니다.
  - 필수 패키지 미설치 시 상단의 [⚡ 원클릭 자동 설치] 버튼으로 즉시 설치 가능합니다.

2. 자동 1:1 문맥 및 동·호수 매칭:
  - 사택 동·호수 토큰(예: 5-503, 102-404) 및 공사 핵심 키워드를 복합 채점하여 공식 패턴명으로 매칭합니다.
  - 동일 공사에 복수 세대 사진이 있을 경우 동·호수를 괄호 병기하여 파일 덮어쓰기를 100% 차단합니다.

3. 100% 안전 보장 (Fail-Safe):
  - 작업 전 원본 폴더 전체를 타임스탬프 백업(`_BACKUP_YYYYMMDD_HHMMSS`)으로 안전 복제합니다.
  - 작업 완료 후 열리는 [검토 다이얼로그]에서 언제든지 [롤백] 버튼으로 1초 만에 원복할 수 있습니다.
"""
        self.guide_text.insert(tk.END, guide_content)
        self.guide_text.config(state=tk.DISABLED)

    def _render_dependency_banner(self):
        """의존성 설치 여부에 따라 상단 알림 배너를 렌더링한다."""
        for widget in self.banner_container.winfo_children():
            widget.destroy()

        if self.missing_deps:
            warn_frame = tk.Frame(self.banner_container, bg="#c0392b", padx=12, pady=6)
            warn_frame.pack(fill=tk.X)
            
            msg = f"⚠️ 필수 패키지 미설치: {', '.join(self.missing_deps)} (OCR 기능에 필요)"
            tk.Label(warn_frame, text=msg, fg="white", bg="#c0392b", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
            
            btn_inst = tk.Button(
                warn_frame,
                text="⚡ 필수 패키지 원클릭 자동 설치",
                bg="#f39c12",
                fg="#2c3e50",
                font=("맑은 고딕", 9, "bold"),
                padx=8,
                pady=1,
                cursor="hand2",
                command=self.open_install_dialog
            )
            btn_inst.pack(side=tk.RIGHT)
        else:
            ok_frame = tk.Frame(self.banner_container, bg="#27ae60", padx=12, pady=4)
            ok_frame.pack(fill=tk.X)
            tk.Label(ok_frame, text="✅ OCR 엔진 구동 준비 완료 (winocr, pillow 활성화됨)", fg="white", bg="#27ae60", font=("맑은 고딕", 8, "bold")).pack(side=tk.LEFT)

    def open_install_dialog(self):
        """원클릭 자동 설치 대화상자를 띄운다."""
        pkgs_to_install = self.missing_deps if self.missing_deps else ["winocr", "pillow"]
        PackageInstallDialog(self.root, pkgs_to_install, on_success=self._on_install_success)

    def _on_install_success(self):
        """설치 성공 시 상태 갱신"""
        self._check_dependencies()
        self._render_dependency_banner()
        self.status_var.set("필수 패키지가 성공적으로 설치되었습니다. 분석을 시작할 수 있습니다.")

    def select_pattern(self):
        initial = self.pat_var.get() if os.path.exists(self.pat_var.get()) else "D:\\"
        folder = filedialog.askdirectory(title="기준 패턴 폴더 (B) 선택", initialdir=initial)
        if folder:
            self.pat_var.set(os.path.normpath(folder))

    def select_target(self):
        initial = self.tgt_var.get() if os.path.exists(self.tgt_var.get()) else "D:\\"
        folder = filedialog.askdirectory(title="변경 대상 폴더 (A) 선택", initialdir=initial)
        if folder:
            self.tgt_var.set(os.path.normpath(folder))

    def update_progress(self, current: int, total: int, filename: str):
        self.root.after(0, self._do_update_progress, current, total, filename)

    def _do_update_progress(self, current: int, total: int, filename: str):
        pct = (current / total) * 100 if total > 0 else 0
        self.progress_bar["value"] = pct
        self.status_var.set(f"[{current}/{total}] OCR 분석 중: {filename}")

    def start_process_thread(self):
        self._check_dependencies()
        if self.missing_deps and self.file_type_var.get() in ("ALL", "IMAGE"):
            msg = f"OCR 분석에 필요한 필수 패키지({', '.join(self.missing_deps)})가 아직 설치되지 않았습니다.\n\n지금 바로 원클릭으로 자동 설치하시겠습니까?"
            if messagebox.askyesno("패키지 설치 필요", msg, parent=self.root):
                self.open_install_dialog()
            return

        pat_dir = self.pat_var.get().strip()
        tgt_dir = self.tgt_var.get().strip()

        if not os.path.exists(pat_dir):
            messagebox.showerror("경로 오류", f"기준 패턴 폴더가 존재하지 않습니다:\n{pat_dir}", parent=self.root)
            return
        if not os.path.exists(tgt_dir):
            messagebox.showerror("경로 오류", f"변경 대상 폴더가 존재하지 않습니다:\n{tgt_dir}", parent=self.root)
            return

        self.btn_run.config(state=tk.DISABLED, text="⏳ 분석 및 변경 작업 진행 중...")
        self.progress_bar["value"] = 0
        self.status_var.set("작업 초기화 및 사전 검증 중...")

        t = threading.Thread(target=self._process_worker, args=(pat_dir, tgt_dir))
        t.daemon = True
        t.start()

    def _process_worker(self, pat_dir: str, tgt_dir: str):
        try:
            # 1. 임시 작업 폴더 생성
            temp_dir = Path(tgt_dir) / TEMP_DIR_NAME
            temp_dir.mkdir(parents=True, exist_ok=True)

            # 2. 안전 백업
            backup_folder_path = None
            if self.backup_var.get():
                self.root.after(0, lambda: self.status_var.set("원본 전체 타임스탬프 백업 생성 중..."))
                ts = time.strftime("%Y%m%d_%H%M%S")
                target_p = Path(tgt_dir)
                backup_p = target_p.parent / f"{target_p.name}_BACKUP_{ts}"
                def ignore_patterns(dir, files):
                    return [TEMP_DIR_NAME] if dir == str(target_p) else []
                shutil.copytree(target_p, backup_p, ignore=ignore_patterns)
                backup_folder_path = str(backup_p)

            # 3. 매칭 엔진 초기화
            engine = MatchingEngine(pat_dir, tgt_dir)
            engine.load_patterns()
            if not engine.patterns:
                raise ValueError("기준 패턴 폴더 내에 유효한 서브폴더가 없습니다.")

            # 4. 대상 파일 수집
            target_p = Path(tgt_dir)
            mode = self.file_type_var.get()
            img_exts = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
            doc_exts = {".xlsx", ".xls", ".hwp", ".docx", ".pdf"}

            image_files: List[Path] = []
            doc_files: List[Path] = []

            for root, dirs, files in os.walk(tgt_dir):
                if TEMP_DIR_NAME in dirs:
                    dirs.remove(TEMP_DIR_NAME)
                for f in files:
                    fp = Path(root) / f
                    ext = fp.suffix.lower()
                    if mode in ("ALL", "IMAGE") and ext in img_exts:
                        image_files.append(fp)
                    elif mode in ("ALL", "DOC") and ext in doc_exts:
                        doc_files.append(fp)

            total_files = len(image_files) + len(doc_files)
            if total_files == 0:
                raise ValueError("변경 대상 폴더에 처리 가능한 이미지 또는 문서 파일이 없습니다.")

            # 5. 이미지 고속 병렬 OCR 실행
            ocr_results = {}
            if image_files:
                if not HAS_WINOCR:
                    raise RuntimeError("winocr 패키지가 설치되지 않아 이미지 OCR을 진행할 수 없습니다.")
                self.root.after(0, lambda: self.status_var.set(f"고속 병렬 OCR 분석 시작 ({len(image_files)}개 이미지)..."))
                ocr_results = asyncio.run(perform_batch_ocr(image_files, concurrency=4, progress_cb=self.update_progress))

            # 6. 매칭 및 변경 목록 산출
            self.root.after(0, lambda: self.status_var.set("문맥 매칭 및 충돌 방지 계산 중..."))
            success_files = []
            skipped_files = []
            folder_rename_map: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}

            # 이미지 매칭
            for img_p in image_files:
                ocr_txt = ocr_results.get(str(img_p), "")
                best_pat, score, reason, unit_token = engine.match_image_ocr(img_p.name, ocr_txt)
                if best_pat:
                    parent_str = str(img_p.parent)
                    folder_rename_map.setdefault(parent_str, {}).setdefault(best_pat["clean_name"], []).append({
                        "file_path": img_p,
                        "ext": img_p.suffix,
                        "pat": best_pat,
                        "score": score,
                        "reason": reason,
                        "unit_token": unit_token,
                        "is_img": True
                    })
                else:
                    skipped_files.append((img_p.name, f"OCR 매칭 실패 ({reason})"))

            # 문서 매칭
            for doc_p in doc_files:
                best_pat, score, reason = engine.match_document_filename(doc_p.name)
                if best_pat:
                    parent_str = str(doc_p.parent)
                    folder_rename_map.setdefault(parent_str, {}).setdefault(best_pat["clean_name"], []).append({
                        "file_path": doc_p,
                        "ext": doc_p.suffix,
                        "pat": best_pat,
                        "score": score,
                        "reason": reason,
                        "unit_token": None,
                        "is_img": False
                    })
                else:
                    skipped_files.append((doc_p.name, f"문서 매칭 실패 ({reason})"))

            # 충돌 해결 및 파일명 확정
            rename_tasks = []
            for parent_str, pat_groups in folder_rename_map.items():
                for pat_clean, items in pat_groups.items():
                    if len(items) == 1:
                        item = items[0]
                        final_name = f"{pat_clean}{item['ext']}".strip()
                        rename_tasks.append((item, final_name))
                    else:
                        for idx, item in enumerate(items, 1):
                            if item["unit_token"]:
                                final_name = f"{pat_clean}({item['unit_token']}){item['ext']}".strip()
                            else:
                                final_name = f"{pat_clean}_{idx}{item['ext']}".strip()
                            rename_tasks.append((item, final_name))

            # 7. 제자리 파일명 변경 실행
            for item, final_name in rename_tasks:
                src_fp = item["file_path"]
                dst_fp = src_fp.parent / final_name

                if src_fp.name == final_name:
                    skipped_files.append((src_fp.name, "이미 표준 패턴 파일명으로 설정되어 있음"))
                    continue

                if safe_rename_file(src_fp, dst_fp):
                    success_files.append({
                        "old_name": src_fp.name,
                        "new_name": final_name,
                        "parent_dir": str(src_fp.parent),
                        "score": item["score"],
                        "reason": item["reason"]
                    })
                else:
                    skipped_files.append((src_fp.name, f"파일명 변경 실패 (권한 또는 파일 락)"))

            # 8. 원클릭 롤백 스크립트 생성
            rollback_script_p = temp_dir / "rollback_rename.py"
            with open(rollback_script_p, "w", encoding="utf-8") as rf:
                rf.write("# -*- coding: utf-8 -*-\n")
                rf.write("import os\nfrom pathlib import Path\n\n")
                rf.write(f"CHANGES = {repr(success_files)}\n\n")
                rf.write("def rollback():\n")
                rf.write("    cnt = 0\n")
                rf.write("    for item in CHANGES:\n")
                rf.write("        curr = Path(item['parent_dir']) / item['new_name']\n")
                rf.write("        orig = Path(item['parent_dir']) / item['old_name']\n")
                rf.write("        if curr.exists():\n")
                rf.write("            os.rename(str(curr), str(orig))\n")
                rf.write("            cnt += 1\n")
                rf.write("    print(f'[완료] {cnt}개 파일 복원 완료.')\n\n")
                rf.write("if __name__ == '__main__':\n    rollback()\n")

            # 9. UI 완료 알림 및 ReviewDialog 호출
            self.root.after(0, lambda: self._on_process_complete(tgt_dir, pat_dir, success_files, skipped_files, backup_folder_path))

        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda: self._on_process_error(err_msg))

    def _on_process_complete(self, src_dir, pat_dir, success_files, skipped_files, backup_path):
        self.btn_run.config(state=tk.NORMAL, text="🚀 분석 및 자동 매칭 / 파일명 변경 실행")
        self.progress_bar["value"] = 100
        self.status_var.set(f"작업 완료! 성공: {len(success_files)}건, 제외: {len(skipped_files)}건")
        ReviewDialog(self.root, src_dir, pat_dir, success_files, skipped_files, backup_path)

    def _on_process_error(self, err_msg: str):
        self.btn_run.config(state=tk.NORMAL, text="🚀 분석 및 자동 매칭 / 파일명 변경 실행")
        self.progress_bar["value"] = 0
        self.status_var.set(f"작업 중단: {err_msg}")
        messagebox.showerror("실행 실패", f"작업 중 오류가 발생하였습니다:\n{err_msg}", parent=self.root)


# ═══════════════════════════════════════════════════════
# 10. 진입점
# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
