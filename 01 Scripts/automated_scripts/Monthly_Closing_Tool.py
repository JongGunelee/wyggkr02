# -*- coding: utf-8 -*-
"""
Monthly Closing File Match and Copy Tool v39.4.0 (Python version)
- D_Drive_Deep_Clean 아키텍처 및 초격차 무결성 표준 준수
- JSON 파일 일체 생성 없음 (메모리 변수 기반)
- 대상 폴더 (A) 및 타겟 상위 폴더 (B-1) 지속적 변동성 지원
- 복사 후 검토 다이얼로그에서 [원본 삭제] / [복사본 롤백] / [유지] 선택적 실행
- Windows 지연 파일 락 방탄 삭제(safe_delete_file) 프로토콜 탑재
- UTF-8 인코딩 안전장치 및 이모지 배제 텍스트 마커 일원화
- 대시보드 웹 런처 전면 가시성 보장(Focus Hardening)
"""

import os
import sys
import io
import time
import shutil
import re
from pathlib import Path
from typing import Tuple, List, Dict, Any
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# ═══════════════════════════════════════════════════════
# 1. UTF-8 인코딩 방탄 구성 (CP949 콘솔 충돌 방지)
# ═══════════════════════════════════════════════════════
def configure_utf8():
    """Windows 한국어 콘솔 환경에서의 UnicodeEncodeError를 방지하기 위해 표준 스트림을 UTF-8로 재구성한다."""
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
# 2. 기본 경로 및 상수 정의 (JSON 파일 없이 코드 내 기본값 관리)
# ═══════════════════════════════════════════════════════
DEFAULT_SOURCE = r"D:\03 금일작업\00 월마감\일성"
DEFAULT_TARGET = r"D:\02 기숙사 및 사택\05 기숙사 및 사택 월마감\26년-사택 작업\09월\01 솔루션"
DEFAULT_PREFIX = "표준단가_"

# ═══════════════════════════════════════════════════════
# 3. 방탄 파일 삭제 프로토콜 (Windows File Handle Lock 극복)
# ═══════════════════════════════════════════════════════
def safe_delete_file(path_obj: Path, max_retries: int = 5, delay: float = 0.2) -> bool:
    """Windows OS 커널의 지연 핸들 락(WinError 32)을 극복하는 5단계 점진적 백오프 삭제 프로토콜."""
    if not path_obj.exists():
        return True
    for attempt in range(max_retries):
        try:
            try:
                os.chmod(str(path_obj), 0o777)
            except Exception:
                pass
            path_obj.unlink(missing_ok=True)
            return True
        except (PermissionError, OSError):
            if attempt < max_retries - 1:
                time.sleep(delay * (attempt + 1))
    return not path_obj.exists()

# ═══════════════════════════════════════════════════════
# 4. 핵심 매칭 및 복사 처리 엔진 (Core Process)
# ═══════════════════════════════════════════════════════
def run_process(source_dir: str, target_parent_dir: str, prefix_text: str) -> Tuple[bool, Any]:
    """
    대상 폴더(A)의 파일들을 타겟 상위 폴더(B-1)의 서브폴더(01~99) 내 '01 자료' 폴더로 매칭하여 복사한다.
    파일명 앞 2자리 숫자를 prefix_text(기본: '표준단가_')로 치환한다.
    """
    if not source_dir or not target_parent_dir:
        return False, "대상 폴더와 타겟 폴더를 모두 지정해야 합니다."

    src_p = Path(source_dir)
    tgt_p = Path(target_parent_dir)

    if not src_p.exists():
        return False, f"대상 폴더(A)가 존재하지 않습니다:\n{source_dir}"
    if not tgt_p.exists():
        return False, f"타겟 상위 폴더(B-1)가 존재하지 않습니다:\n{target_parent_dir}"

    # 타겟 상위 폴더 내 서브폴더 맵 구성 (예: '01' -> '.../01 기숙사...')
    target_subfolders: Dict[str, Path] = {}
    for entry in tgt_p.iterdir():
        if entry.is_dir():
            match = re.match(r"^(\d{2})", entry.name)
            if match:
                prefix = match.group(1)
                target_subfolders[prefix] = entry

    success_files: List[Dict[str, Any]] = []
    skipped_files: List[Tuple[str, str]] = []

    # 대상 폴더(A)의 파일 전수 스캔 및 정렬
    source_files = sorted([f for f in src_p.iterdir() if f.is_file()], key=lambda x: x.name)

    for entry in source_files:
        filename = entry.name
        match = re.match(r"^(\d{2})", filename)
        if not match:
            skipped_files.append((filename, "파일명이 2자리 숫자로 시작하지 않음"))
            continue

        prefix = match.group(1)
        if prefix not in target_subfolders:
            skipped_files.append((filename, f"매칭되는 타겟 폴더(접두사 '{prefix}') 없음"))
            continue

        target_subfolder_path = target_subfolders[prefix]
        dest_dir = target_subfolder_path / "01 자료"

        if not dest_dir.exists():
            skipped_files.append((filename, f"'{target_subfolder_path.name}' 내 '01 자료' 폴더 없음"))
            continue

        new_filename = prefix_text + filename[2:]
        dest_file_path = dest_dir / new_filename

        try:
            shutil.copy2(str(entry), str(dest_file_path))
            orig_size = entry.stat().st_size
            dest_size = dest_file_path.stat().st_size
            if orig_size == dest_size:
                success_files.append({
                    "filename": filename,
                    "orig_path": str(entry),
                    "dest_path": str(dest_file_path),
                    "target_folder": target_subfolder_path.name,
                    "new_filename": new_filename,
                    "size": dest_size
                })
            else:
                skipped_files.append((filename, f"파일 크기 불일치 (원본: {orig_size}, 복사본: {dest_size})"))
        except Exception as e:
            skipped_files.append((filename, f"복사 중 에러: {e}"))

    return True, (success_files, skipped_files)

# ═══════════════════════════════════════════════════════
# 5. 사용자 검토 및 선택적 조치 다이얼로그 (ReviewDialog)
# ═══════════════════════════════════════════════════════
class ReviewDialog(tk.Toplevel):
    """작업 실행 완료 후 세부 매칭 결과를 확인하고 후속 조치(삭제/롤백/유지)를 결정하는 다이얼로그"""
    def __init__(self, parent, src_dir: str, tgt_dir: str, success_files: List[Dict[str, Any]], skipped_files: List[Tuple[str, str]]):
        super().__init__(parent)
        self.parent = parent
        self.src_dir = src_dir
        self.tgt_dir = tgt_dir
        self.success_files = success_files
        self.skipped_files = skipped_files

        self.title("[결과 검토] 월마감 매칭 복사 및 후속 조치 선택 (v39.4.0)")
        self.geometry("780x660")
        self.minsize(720, 560)

        # 포커스 하드닝: 다이얼로그 창 전면 강제 활성화
        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.focus_force()

        self._build_ui()

    def _build_ui(self):
        header_frame = tk.Frame(self, bg="#2c3e50", padx=15, pady=12)
        header_frame.pack(fill=tk.X)

        title_lbl = tk.Label(header_frame, text="작업 실행 완료 및 검토 보고서", font=("맑은 고딕", 13, "bold"), fg="white", bg="#2c3e50")
        title_lbl.pack(anchor="w")

        total = len(self.success_files) + len(self.skipped_files)
        succ = len(self.success_files)
        skip = len(self.skipped_files)
        summary_text = f"총 처리 대상: {total}건  |  [성공] 복사 완료: {succ}건  |  [제외] 미매칭/누락: {skip}건"
        sub_lbl = tk.Label(header_frame, text=summary_text, font=("맑은 고딕", 10), fg="#ecf0f1", bg="#2c3e50")
        sub_lbl.pack(anchor="w", pady=(4, 0))

        body_frame = tk.Frame(self, padx=15, pady=10)
        body_frame.pack(fill=tk.BOTH, expand=True)

        txt_label = tk.Label(body_frame, text="[세부 매칭 및 처리 결과]", font=("맑은 고딕", 9, "bold"))
        txt_label.pack(anchor="w", pady=(0, 5))

        self.report_box = scrolledtext.ScrolledText(body_frame, wrap=tk.WORD, font=("Consolas", 9), bg="#fcfcfc")
        self.report_box.pack(fill=tk.BOTH, expand=True)

        self.report_text_content = self._populate_report()

        quick_frame = tk.Frame(body_frame, pady=5)
        quick_frame.pack(fill=tk.X)

        tk.Button(quick_frame, text="[폴더 열기] 대상 폴더(A)", command=lambda: self.open_folder(self.src_dir), font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(quick_frame, text="[폴더 열기] 타겟 상위 폴더(B-1)", command=lambda: self.open_folder(self.tgt_dir), font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(quick_frame, text="[보고서 저장] 결과 TXT 내보내기", command=self.export_report, font=("맑은 고딕", 9)).pack(side=tk.RIGHT)

        action_frame = tk.LabelFrame(self, text="[후속 조치 선택] 검토 후 원하시는 작업을 선택하세요", font=("맑은 고딕", 10, "bold"), padx=15, pady=10, fg="#2980b9")
        action_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        btn_grid = tk.Frame(action_frame)
        btn_grid.pack(fill=tk.X)

        # 1) 원본 삭제
        btn_del_src = tk.Button(
            btn_grid,
            text="[원본 정리]\n성공한 원본 파일 삭제",
            font=("맑은 고딕", 9, "bold"),
            bg="#e74c3c",
            fg="white",
            height=2,
            width=24,
            command=self.delete_source_files
        )
        btn_del_src.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        # 2) 롤백
        btn_rollback = tk.Button(
            btn_grid,
            text="[롤백 / 되돌리기]\n복사본 파일 전체 삭제",
            font=("맑은 고딕", 9, "bold"),
            bg="#e67e22",
            fg="white",
            height=2,
            width=24,
            command=self.rollback_copied_files
        )
        btn_rollback.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # 3) 유지
        btn_keep = tk.Button(
            btn_grid,
            text="[작업 완료 및 유지]\n원본/복사본 모두 보존",
            font=("맑은 고딕", 9, "bold"),
            bg="#27ae60",
            fg="white",
            height=2,
            width=24,
            command=self.keep_and_close
        )
        btn_keep.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        btn_grid.columnconfigure(0, weight=1)
        btn_grid.columnconfigure(1, weight=1)
        btn_grid.columnconfigure(2, weight=1)

    def _populate_report(self) -> str:
        rep = "=======================================================================\n"
        rep += " 월마감 파일 패턴폴더 복사 및 파일명 치환 상세 결과 보고\n"
        rep += "=======================================================================\n"
        rep += f"• 대상 폴더 (A)       : {self.src_dir}\n"
        rep += f"• 타겟 상위 폴더 (B-1): {self.tgt_dir}\n"
        rep += f"• 복사 성공           : {len(self.success_files)}건\n"
        rep += f"• 복사 제외(생략)     : {len(self.skipped_files)}건\n\n"

        if self.success_files:
            rep += f"[복사 및 치환 성공 목록 ({len(self.success_files)}건)]\n"
            rep += "-----------------------------------------------------------------------\n"
            for idx, item in enumerate(self.success_files, 1):
                rep += f"{idx:2d}. {item['filename']}\n"
                rep += f"    -> 저장: {item['target_folder']}\\01 자료\\{item['new_filename']}\n"
                rep += f"    -> 크기: {item['size']:,} bytes\n"
            rep += "\n"

        if self.skipped_files:
            rep += f"[복사 제외/누락 목록 ({len(self.skipped_files)}건)]\n"
            rep += "-----------------------------------------------------------------------\n"
            for idx, (fn, reason) in enumerate(self.skipped_files, 1):
                rep += f"{idx:2d}. {fn}\n"
                rep += f"    -> 제외 사유: {reason}\n"
            rep += "\n"

        self.report_box.insert(tk.END, rep)
        self.report_box.config(state=tk.DISABLED)
        return rep

    def export_report(self):
        """결과 리포트를 utf-8-sig 인코딩 텍스트 파일로 저장한다."""
        dest = filedialog.asksaveasfilename(
            title="결과 보고서 저장",
            initialdir=self.src_dir if os.path.exists(self.src_dir) else "D:\\",
            initialfile=f"월마감_매칭결과보고_{time.strftime('%Y%m%d_%H%M%S')}.txt",
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
            messagebox.showwarning("경고", f"폴더가 없습니다:\n{path}", parent=self)

    def delete_source_files(self):
        if not self.success_files:
            messagebox.showinfo("알림", "삭제할 원본 파일이 없습니다.", parent=self)
            return

        count = len(self.success_files)
        msg = f"대상 폴더(A)에서 성공적으로 복사된 원본 파일 {count}개를 삭제하시겠습니까?\n\n" \
              f"[주의] 삭제된 파일은 복구되지 않습니다!\n" \
              f"대상 폴더: {self.src_dir}"

        if not messagebox.askyesno("[확인] 원본 파일 삭제", msg, icon="warning", parent=self):
            return

        del_cnt = 0
        for item in self.success_files:
            orig_p = Path(item["orig_path"])
            if safe_delete_file(orig_p):
                del_cnt += 1

        src_path = Path(self.src_dir)
        remaining = [f for f in src_path.iterdir() if f.is_file()] if src_path.exists() else []
        if not remaining and src_path.exists():
            if messagebox.askyesno("폴더 정리", f"원본 파일 {del_cnt}개가 삭제되었습니다.\n\n폴더가 비어 있습니다. 대상 폴더(A) 자체도 삭제하시겠습니까?", parent=self):
                try:
                    src_path.rmdir()
                    messagebox.showinfo("완료", "대상 폴더(A)가 삭제되었습니다.", parent=self)
                except Exception:
                    pass
        else:
            messagebox.showinfo("완료", f"원본 파일 {del_cnt}개가 삭제되었습니다.\n(제외 파일 {len(remaining)}개는 보존됨)", parent=self)

        self.destroy()

    def rollback_copied_files(self):
        if not self.success_files:
            messagebox.showinfo("알림", "롤백할 복사본 파일이 없습니다.", parent=self)
            return

        count = len(self.success_files)
        msg = f"방금 목적지(01 자료)로 복사된 파일 {count}개를 모두 삭제하여 작업 전 상태로 되돌리시겠습니까?\n\n" \
              f"타겟 상위 폴더: {self.tgt_dir}"

        if not messagebox.askyesno("[확인] 복사본 롤백", msg, icon="warning", parent=self):
            return

        rb_cnt = 0
        for item in self.success_files:
            dest_p = Path(item["dest_path"])
            if safe_delete_file(dest_p):
                rb_cnt += 1

        messagebox.showinfo("롤백 완료", f"복사본 파일 {rb_cnt}개가 안전하게 삭제(롤백)되었습니다.", parent=self)
        self.destroy()

    def keep_and_close(self):
        messagebox.showinfo("완료", "모든 파일이 원본 및 대상 폴더에 안전하게 보존되었습니다.", parent=self)
        self.destroy()

# ═══════════════════════════════════════════════════════
# 6. 메인 애플리케이션 UI (Tkinter App)
# ═══════════════════════════════════════════════════════
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("월마감 파일 패턴폴더 매칭 & 복사/치환 도구 v39.4.0")
        self.root.geometry("700x700")
        self.root.minsize(640, 600)

        # 대시보드 런처 호출 시 창이 뒤로 숨는 현상을 방지하는 포커스 하드닝 시퀀스
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(500, lambda: self.root.attributes('-topmost', False))
        self.root.focus_force()

        self._create_widgets()

    def _create_widgets(self):
        header = tk.Frame(self.root, bg="#34495e", padx=15, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="월마감 파일 패턴폴더 자동 매칭 & 치환 도구 (v39.4.0)", font=("맑은 고딕", 12, "bold"), fg="white", bg="#34495e").pack(anchor="w")
        tk.Label(header, text="지속적 폴더 변동성 지원 | 원본 삭제 / 복사본 롤백 다이얼로그 선택 지원", font=("맑은 고딕", 9), fg="#bdc3c7", bg="#34495e").pack(anchor="w", pady=(2, 0))

        main_frame = tk.Frame(self.root, padx=20, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 대상 폴더(A)
        tk.Label(main_frame, text="[대상 폴더 (A)] 복사할 원본 파일 위치:", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        src_row = tk.Frame(main_frame)
        src_row.pack(fill=tk.X, pady=(3, 12))
        self.src_var = tk.StringVar(value=DEFAULT_SOURCE)
        self.src_entry = tk.Entry(src_row, textvariable=self.src_var, font=("맑은 고딕", 9))
        self.src_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(src_row, text="폴더 선택...", command=self.select_source, width=11).pack(side=tk.LEFT)

        # 타겟 상위 폴더(B-1)
        tk.Label(main_frame, text="[타겟 상위 폴더 (B-1)] 01~99 서브폴더가 위치한 상위 경로:", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        tgt_row = tk.Frame(main_frame)
        tgt_row.pack(fill=tk.X, pady=(3, 12))
        self.tgt_var = tk.StringVar(value=DEFAULT_TARGET)
        self.tgt_entry = tk.Entry(tgt_row, textvariable=self.tgt_var, font=("맑은 고딕", 9))
        self.tgt_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(tgt_row, text="폴더 선택...", command=self.select_target, width=11).pack(side=tk.LEFT)

        # 접두사
        prefix_frame = tk.Frame(main_frame)
        prefix_frame.pack(fill=tk.X, pady=(0, 15))
        tk.Label(prefix_frame, text="치환 접두사:", font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        self.prefix_var = tk.StringVar(value=DEFAULT_PREFIX)
        tk.Entry(prefix_frame, textvariable=self.prefix_var, width=18, font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=10)
        tk.Label(prefix_frame, text="(파일명 앞 숫자 2자리를 이 단어로 치환)", fg="#7f8c8d", font=("맑은 고딕", 8)).pack(side=tk.LEFT)

        # 실행 버튼
        self.btn_run = tk.Button(
            main_frame,
            text="분석 및 복사/치환 실행",
            font=("맑은 고딕", 11, "bold"),
            bg="#2980b9",
            fg="white",
            height=2,
            command=self.execute
        )
        self.btn_run.pack(fill=tk.X, pady=(5, 15))

        # 가이드
        guide_label = tk.Label(main_frame, text="[사용 안내]", font=("맑은 고딕", 9, "bold"))
        guide_label.pack(anchor="w", pady=(5, 3))

        self.guide_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=12, font=("맑은 고딕", 9), bg="#f8f9fa")
        self.guide_text.pack(fill=tk.BOTH, expand=True)

        guide_content = """1. 폴더 경로 변경 지원:
  - '대상 폴더 (A)'와 '타겟 상위 폴더 (B-1)'는 [폴더 선택...] 버튼으로 매월 자유롭게 변경할 수 있습니다.
  - JSON 등 불필요한 외부 파일은 일체 생성되지 않으며 메모리 기반으로 안전하게 동작합니다.

2. 자동 매칭 및 복사 규칙:
  - A 폴더 파일명의 앞 2자리 숫자(예: 01, 02)를 추출합니다.
  - B-1 폴더 내 동일한 숫자로 시작하는 서브폴더를 검색합니다.
  - 해당 서브폴더 내 '01 자료' 폴더가 존재하는 경우에만 복사합니다.
  - 파일명의 앞 2자리 숫자를 입력한 접두사('표준단가_')로 치환합니다.
  - 타겟 폴더가 없거나 '01 자료' 폴더가 없는 파일은 복사를 생략하고 원본을 보존합니다.

3. 검토 후 선택적 조치 다이얼로그:
  - 실행 즉시 결과 보고서 창이 열립니다.
  - 검토 후 3가지 조치 중 선택할 수 있습니다:
    ① [원본 정리] : 성공한 원본 파일을 대상 폴더(A)에서 안전하게 삭제
    ② [복사본 롤백] : 방금 복사된 타겟 파일들을 일괄 삭제하여 이전 상태로 복구
    ③ [유지 및 완료] : 원본과 복사본 모두 그대로 보존하고 완료
"""
        self.guide_text.insert(tk.END, guide_content)
        self.guide_text.config(state=tk.DISABLED)

    def select_source(self):
        initial = self.src_var.get() if os.path.exists(self.src_var.get()) else "D:\\"
        folder = filedialog.askdirectory(title="대상 폴더 (A) 선택", initialdir=initial)
        if folder:
            self.src_var.set(os.path.normpath(folder))

    def select_target(self):
        initial = self.tgt_var.get() if os.path.exists(self.tgt_var.get()) else "D:\\"
        folder = filedialog.askdirectory(title="타겟 상위 폴더 (B-1) 선택", initialdir=initial)
        if folder:
            self.tgt_var.set(os.path.normpath(folder))

    def execute(self):
        src = self.src_var.get().strip()
        tgt = self.tgt_var.get().strip()
        prefix = self.prefix_var.get().strip()

        success, result = run_process(src, tgt, prefix)
        if not success:
            messagebox.showerror("실행 오류", result, parent=self.root)
            return

        success_files, skipped_files = result
        ReviewDialog(self.root, src, tgt, success_files, skipped_files)

# ═══════════════════════════════════════════════════════
# 7. 진입점
# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
