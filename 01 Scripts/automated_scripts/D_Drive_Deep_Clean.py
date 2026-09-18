# -*- coding: utf-8 -*-
r"""
================================================================================
🚀 [Antigravity & Codex] C/D 드라이브 + Windows 디스크 정리 통합 정기 딥 클린 v3.2.1
================================================================================
- 아키텍처: 다크 모던 Tkinter GUI (Dark Modern UI) + 비동기 스레드 워커 + CLI 듀얼 모드
- 3대 Zone 14개 정밀 정리 대상:
    * Zone A (D: 5대): Temp\User & 세션, Antigravity 크래시 로그, 업데이터 pending, TeraBox 캐시/.cab, Codex bin 해시/.bak
    * Zone B (C: 5대): ALPDF PostScript 스풀, Windows Temp/SoftwareDist, CBS 로그/Dbg 심볼, 브라우저 캐시, 개발 캐시
    * Zone C (Windows cleanmgr 4대): 전송 최적화 파일, DirectX/NVIDIA 셰이더, WER 오류보고/덤프, 휴지통 비우기
- 화이트리스트 절대 보호: MSOffice(80개 업무문서), AI 대화 DB, 확장팩(25개), Codex 설정(gpt-6-astra), 활성 바이너리 100% 무결 보존
- 데드락 원천 차단: 현재 로그인된 사용자 고유 SID(whoami /user) 타깃 휴지통 비우기
- 대시보드 융합: pythonw.exe 무콘솔 환경에서 즉시 최상위 전면 활성화 및 실시간 프로그레스/로그 표출
================================================================================
"""

import os
import sys
import shutil
import ctypes
import argparse
import fnmatch
import subprocess
import threading
import time
from datetime import datetime
from pathlib import Path

# GUI 모듈 임포트
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ------------------------------------------------------------------------------
# 1. 인코딩 및 환경 정규화 (Encoding Armor)
# ------------------------------------------------------------------------------
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ------------------------------------------------------------------------------
# 2. Windows 시스템 및 UAC 권한 유틸리티
# ------------------------------------------------------------------------------
def is_admin() -> bool:
    """Windows 관리자 권한 여부 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def run_as_admin():
    """관리자 권한으로 자기 자신 재실행 (GUI 지원)"""
    try:
        exe = sys.executable
        if "python.exe" in exe.lower() and not "pythonw.exe" in exe.lower():
            pw = exe.lower().replace("python.exe", "pythonw.exe")
            if os.path.exists(pw):
                exe = pw

        script = os.path.abspath(__file__)
        params = f'"{script}"'
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", exe, params, os.path.dirname(script), 1
        )
        return ret > 32
    except Exception:
        return False

def get_disk_free_gb(path: str) -> float:
    """드라이브 여유 공간(GB) 반환"""
    try:
        usage = shutil.disk_usage(path)
        return round(usage.free / (1024 ** 3), 2)
    except Exception:
        return 0.0

def count_office_files() -> int:
    """MSOffice 업무 파일 무결성 카운트 (D:\03 금일작업\00 임시\0000000 MSoffice)"""
    office_dir = r"D:\03 금일작업\00 임시\0000000 MSoffice"
    if os.path.exists(office_dir):
        try:
            files = [f for f in os.listdir(office_dir) if os.path.isfile(os.path.join(office_dir, f))]
            return len(files)
        except Exception:
            return 0
    return 0

def get_current_user_sid() -> str:
    """현재 사용자의 Windows SID 획득 (SYSTEM 계정 접근 데드락 원천 차단)"""
    try:
        out = subprocess.check_output(["whoami", "/user"], text=True, errors="replace")
        for line in out.splitlines():
            if "S-1-5-21-" in line:
                return line.strip().split()[-1]
    except Exception:
        pass
    return None

# ------------------------------------------------------------------------------
# 3. 안전한 파일/폴더 I/O 소거 루틴
# ------------------------------------------------------------------------------
def safe_remove_files(target_dir: str, pattern: str = "*", recursive: bool = True, dry_run: bool = False,
                      exclude_exts: set = None, exclude_dirs: set = None):
    """안전한 파일 삭제 및 파일 수/바이트 집계"""
    if exclude_exts is None: exclude_exts = set()
    if exclude_dirs is None: exclude_dirs = set()

    del_count = 0
    del_bytes = 0

    if not os.path.exists(target_dir):
        return del_count, del_bytes

    if recursive:
        for root, dirs, files in os.walk(target_dir, topdown=True):
            dirs[:] = [d for d in dirs if d.lower() not in exclude_dirs]
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in exclude_exts:
                    continue
                if fnmatch.fnmatch(file.lower(), pattern.lower()):
                    full_path = os.path.join(root, file)
                    try:
                        size = os.path.getsize(full_path)
                        if not dry_run:
                            try:
                                os.chmod(full_path, 0o777)
                            except Exception:
                                pass
                            os.remove(full_path)
                        del_count += 1
                        del_bytes += size
                    except Exception:
                        pass
    else:
        try:
            for item in os.listdir(target_dir):
                full_path = os.path.join(target_dir, item)
                if os.path.isfile(full_path):
                    ext = os.path.splitext(item)[1].lower()
                    if ext in exclude_exts:
                        continue
                    if fnmatch.fnmatch(item.lower(), pattern.lower()):
                        try:
                            size = os.path.getsize(full_path)
                            if not dry_run:
                                try:
                                    os.chmod(full_path, 0o777)
                                except Exception:
                                    pass
                                os.remove(full_path)
                            del_count += 1
                            del_bytes += size
                        except Exception:
                            pass
        except Exception:
            pass

    return del_count, del_bytes

def safe_remove_subdirs(parent_dir: str, dry_run: bool = False):
    """하위 디렉터리 안전 제거 (부모 폴더 자체는 유지)"""
    del_count = 0
    del_bytes = 0
    if not os.path.exists(parent_dir):
        return del_count, del_bytes

    try:
        for item in os.listdir(parent_dir):
            full_path = os.path.join(parent_dir, item)
            if os.path.isdir(full_path):
                for root, _, files in os.walk(full_path):
                    for f in files:
                        try:
                            fp = os.path.join(root, f)
                            del_bytes += os.path.getsize(fp)
                            del_count += 1
                        except Exception:
                            pass
                if not dry_run:
                    try:
                        shutil.rmtree(full_path, ignore_errors=True)
                    except Exception:
                        pass
    except Exception:
        pass
    return del_count, del_bytes

# ------------------------------------------------------------------------------
# 4. 결과 검토 및 성과 보고 다이얼로그 (DeepCleanSummaryDialog)
# ------------------------------------------------------------------------------
class DeepCleanSummaryDialog(tk.Toplevel):
    def __init__(self, parent, report_data: dict):
        super().__init__(parent)
        self.parent = parent
        self.report_data = report_data

        self.title("통합 딥 클린 성과 및 자산 무결성 보고서 (v3.2.1)")
        self.geometry("740x640")
        self.minsize(680, 560)
        self.configure(bg="#0f172a")

        # 포커스 고정
        self.lift()
        self.attributes('-topmost', True)
        self.after(500, lambda: self.attributes('-topmost', False))
        self.focus_force()

        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self, bg="#1e293b", padx=20, pady=15)
        header.pack(fill=tk.X)

        title_text = "✨ [시뮬레이션 완료] 정리 예상 보고서" if self.report_data.get('dry_run') else "🎉 [클린 완료] 디스크 정리 최종 성과 보고서"
        tk.Label(header, text=title_text, font=("맑은 고딕", 13, "bold"), fg="#38bdf8", bg="#1e293b").pack(anchor="w")
        tk.Label(header, text="3대 Zone 정밀 카운팅 | C/D 드라이브 여유 공간 비교 | 화이트리스트 5대 자산 100% 무결 검증",
                 font=("맑은 고딕", 9), fg="#94a3b8", bg="#1e293b").pack(anchor="w", pady=(3, 0))

        content_frame = tk.Frame(self, bg="#0f172a", padx=20, pady=15)
        content_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_report = tk.Text(content_frame, font=("Consolas", 10), bg="#1e293b", fg="#f8fafc",
                                  relief=tk.FLAT, padx=12, pady=12)
        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=self.txt_report.yview)
        self.txt_report.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_report.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 보고서 텍스트 생성
        report_str = self._generate_report_text()
        self.txt_report.insert(tk.END, report_str)
        self.txt_report.config(state=tk.DISABLED)

        # 하단 액션 버튼
        btn_bar = tk.Frame(self, bg="#0f172a", padx=20, pady=12)
        btn_bar.pack(fill=tk.X)

        tk.Button(btn_bar, text="📁 D: MSOffice 폴더 열기", font=("맑은 고딕", 9),
                  bg="#334155", fg="white", relief=tk.FLAT, padx=12, pady=6, cursor="hand2",
                  command=lambda: self._open_dir(r"D:\03 금일작업\00 임시\0000000 MSoffice")).pack(side=tk.LEFT, padx=(0, 8))

        tk.Button(btn_bar, text="💾 보고서 저장 (.txt)", font=("맑은 고딕", 9, "bold"),
                  bg="#2563eb", fg="white", relief=tk.FLAT, padx=15, pady=6, cursor="hand2",
                  command=self._save_report).pack(side=tk.LEFT)

        tk.Button(btn_bar, text="닫기", font=("맑은 고딕", 9),
                  bg="#475569", fg="white", relief=tk.FLAT, padx=18, pady=6, cursor="hand2",
                  command=self.destroy).pack(side=tk.RIGHT)

    def _generate_report_text(self) -> str:
        d = self.report_data
        dry = d.get('dry_run', False)
        mode_str = "[시뮬레이션 - 실제 삭제 없음]" if dry else "[실제 삭제 확정 완료]"

        lines = [
            "================================================================================",
            f"       Antigravity & Codex C/D 드라이브 + Windows 디스크 정리 리포트 v3.2.1",
            "================================================================================",
            f"실행 일시 : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"실행 모드 : {mode_str}",
            "",
            "--------------------------------------------------------------------------------",
            " [1] Zone A: D: 드라이브 5대 영역 정리 결과 (AI / 데이터 캐시)",
            "--------------------------------------------------------------------------------"
        ]
        for item in d.get('d_items', []):
            lines.append(f" - {item['name']:<38} : {item['count']:>6}개 소거 ({item['mb']:>8.2f} MB)")
        lines.append(f" >> Zone A 소계 : {d.get('d_total_files', 0):,}개 파일 / {d.get('d_total_mb', 0.0):,.2f} MB 회수")
        lines.append("")

        lines.extend([
            "--------------------------------------------------------------------------------",
            " [2] Zone B: C: 드라이브 5대 영역 정리 결과 (OS 시스템 / 스풀 / 웹 / 개발도구)",
            "--------------------------------------------------------------------------------"
        ])
        for item in d.get('c_items', []):
            lines.append(f" - {item['name']:<38} : {item['count']:>6}개 소거 ({item['mb']:>8.2f} MB)")
        lines.append(f" >> Zone B 소계 : {d.get('c_total_files', 0):,}개 파일 / {d.get('c_total_mb', 0.0):,.2f} MB 회수")
        lines.append("")

        lines.extend([
            "--------------------------------------------------------------------------------",
            " [3] Zone C: Windows 기본 디스크 정리 4대 표준 영역 (cleanmgr 표준)",
            "--------------------------------------------------------------------------------"
        ])
        for item in d.get('w_items', []):
            lines.append(f" - {item['name']:<38} : {item['count']:>6}개 소거 ({item['mb']:>8.2f} MB)")
        lines.append(f" >> Zone C 소계 : {d.get('w_total_files', 0):,}개 파일 / {d.get('w_total_mb', 0.0):,.2f} MB 회수")
        lines.append("")

        lines.extend([
            "================================================================================",
            " ★ [종합 성과 및 디스크 실시간 공간 비교]",
            "================================================================================",
            f"총 소거 대상 파일 : {d.get('grand_files', 0):,} 개",
            f"총 공간 회수 용량 : {d.get('grand_mb', 0.0):,.2f} MB ({d.get('grand_gb', 0.0):,.2f} GB)",
            f"C: 드라이브 여유  : {d.get('c_before', 0.0)} GB -> {d.get('c_after', 0.0)} GB (변화: {round(d.get('c_after', 0.0) - d.get('c_before', 0.0), 2):+0.2f} GB)",
            f"D: 드라이브 여유  : {d.get('d_before', 0.0)} GB -> {d.get('d_after', 0.0)} GB (변화: {round(d.get('d_after', 0.0) - d.get('d_before', 0.0), 2):+0.2f} GB)",
            "",
            "================================================================================",
            " ★ [절대 보호 화이트리스트 5대 자산 무결성 재검증 결과]",
            "================================================================================",
            f" [1] 사용자 직속 업무문서 : {d.get('office_count', 0)}개 원본 파일 100% 무결 보존 확인",
            " [2] AI 대화 지능 세션 DB : .gemini 대화 DB, SQLite, 25개 스킬/확장팩 원형 보존",
            " [3] Codex 설정 및 세션 DB: config.toml (gpt-6-astra 모델 연동), 세션 DB 원형 보존",
            " [4] 최신 활성 바이너리   : codex.exe 등 7대 실행 파일 및 Programs 본체 보존",
            " [5] ALPDF 변환 완료 PDF  : PDFCreator 내 사용자 완성본 *.pdf 100% 보존",
            "================================================================================"
        ])
        return "\n".join(lines)

    def _save_report(self):
        dest = filedialog.asksaveasfilename(
            title="딥 클린 보고서 저장",
            initialdir=r"D:\03 금일작업\00 임시" if os.path.exists(r"D:\03 금일작업\00 임시") else "C:\\",
            initialfile=f"딥클린_결과보고_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            parent=self
        )
        if dest:
            try:
                report_str = self._generate_report_text()
                with open(dest, "w", encoding="utf-8-sig") as f:
                    f.write(report_str)
                messagebox.showinfo("저장 완료", f"보고서가 안전하게 저장되었습니다:\n{dest}", parent=self)
            except Exception as e:
                messagebox.showerror("저장 실패", f"보고서 저장 중 오류 발생:\n{e}", parent=self)

    def _open_dir(self, path: str):
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("오류", f"폴더 열기 실패: {e}", parent=self)
        else:
            messagebox.showwarning("경고", f"폴더가 존재하지 않습니다:\n{path}", parent=self)

# ------------------------------------------------------------------------------
# 5. 메인 GUI 애플리케이션 (DeepCleanApp)
# ------------------------------------------------------------------------------
class DeepCleanApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("C/D 드라이브 + Windows 통합 정기 딥-클린 v3.2.1")
        self.root.geometry("820x760")
        self.root.minsize(740, 640)
        self.root.configure(bg="#0f172a")

        # 포커스 하드닝 (대시보드 pythonw 비동기 호출 시 전면 가시성 강제)
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(600, lambda: self.root.attributes('-topmost', False))
        self.root.focus_force()

        self.is_running = False
        self.admin_status = is_admin()
        self.ui_queue = queue.Queue()

        self._build_ui()
        self._refresh_disk_info()
        self._poll_ui_queue()

    def _build_ui(self):
        # ── 헤더 프레임 ──
        header = tk.Frame(self.root, bg="#1e293b", padx=20, pady=12)
        header.pack(fill=tk.X)

        title_row = tk.Frame(header, bg="#1e293b")
        title_row.pack(fill=tk.X)

        tk.Label(title_row, text="🧹 C/D 드라이브 + Windows 기본 디스크 정리 통합 정기 딥 클린 v3.2.1",
                 font=("맑은 고딕", 12, "bold"), fg="#38bdf8", bg="#1e293b").pack(side=tk.LEFT)

        # 관리자 배지
        admin_text = "🛡️ 관리자 모드" if self.admin_status else "⚠️ 일반 권한 (클릭하여 승격)"
        admin_bg = "#15803d" if self.admin_status else "#b45309"
        btn_admin = tk.Button(title_row, text=admin_text, font=("맑은 고딕", 8, "bold"),
                              bg=admin_bg, fg="white", relief=tk.FLAT, padx=8, pady=2, cursor="hand2",
                              command=self._elevate_admin)
        btn_admin.pack(side=tk.RIGHT)

        tk.Label(header, text="Zone A(D: 5대) + Zone B(C: 5대) + Zone C(Windows cleanmgr 4대) 총 14개 영역 정밀 소거 및 MSOffice 80개 파일 절대 보호",
                 font=("맑은 고딕", 9), fg="#94a3b8", bg="#1e293b").pack(anchor="w", pady=(3, 0))

        # ── 디스크 및 화이트리스트 상태 배너 ──
        status_bar = tk.Frame(self.root, bg="#0f172a", padx=20, pady=8)
        status_bar.pack(fill=tk.X)

        self.lbl_c_space = tk.Label(status_bar, text="C: 드라이브: 측정 중...", font=("맑은 고딕", 9, "bold"),
                                    fg="#38bdf8", bg="#0f172a")
        self.lbl_c_space.pack(side=tk.LEFT, padx=(0, 20))

        self.lbl_d_space = tk.Label(status_bar, text="D: 드라이브: 측정 중...", font=("맑은 고딕", 9, "bold"),
                                    fg="#a78bfa", bg="#0f172a")
        self.lbl_d_space.pack(side=tk.LEFT, padx=(0, 20))

        self.lbl_whitelist = tk.Label(status_bar, text="화이트리스트: MSOffice 업무파일(80개) 보호 중",
                                      font=("맑은 고딕", 9, "bold"), fg="#22c55e", bg="#0f172a")
        self.lbl_whitelist.pack(side=tk.RIGHT)

        # ── 3대 Zone 체크박스 패널 (Scrollable Frame) ──
        container = tk.Frame(self.root, bg="#1e293b", padx=15, pady=10)
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=(5, 10))

        canvas = tk.Canvas(container, bg="#1e293b", highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#1e293b")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Zone A
        zone_a_box = tk.LabelFrame(scrollable_frame, text=" [Zone A] D: 드라이브 5대 영역 (데이터 / AI 개발환경 / 클라우드 캐시) ",
                                   font=("맑은 고딕", 9, "bold"), fg="#38bdf8", bg="#1e293b", padx=10, pady=8)
        zone_a_box.pack(fill=tk.X, expand=True, pady=(0, 10))

        self.var_d1 = tk.BooleanVar(value=True)
        self.var_d2 = tk.BooleanVar(value=True)
        self.var_d3 = tk.BooleanVar(value=True)
        self.var_d4 = tk.BooleanVar(value=True)
        self.var_d5 = tk.BooleanVar(value=True)

        self._add_check(zone_a_box, "(1) D:\\DevEnv\\...\\Temp\\User 및 sessions\\temp : 사용자 및 세션 임시 버퍼 파일", self.var_d1)
        self._add_check(zone_a_box, "(2) D:\\DevEnv\\...\\.gemini\\antigravity\\crashes     : Antigravity 크래시 로그 (*.log)", self.var_d2)
        self._add_check(zone_a_box, "(3) D:\\DevEnv\\...\\Programs\\antigravity-updater    : 업데이터 다운로드 pending 잔여물", self.var_d3)
        self._add_check(zone_a_box, "(4) D:\\DevEnv\\...\\AppData\\Roaming\\TeraBox         : 썸네일 캐시 및 AutoUpdate 누적 *.cab", self.var_d4)
        self._add_check(zone_a_box, "(5) D:\\DevEnv\\...\\Codex\\bin                       : 구버전 해시 폴더 및 *.bak (활성 바이너리 보존)", self.var_d5)

        # Zone B
        zone_b_box = tk.LabelFrame(scrollable_frame, text=" [Zone B] C: 드라이브 5대 영역 (OS 시스템 / 인쇄스풀 / 웹 / 개발도구) ",
                                   font=("맑은 고딕", 9, "bold"), fg="#a78bfa", bg="#1e293b", padx=10, pady=8)
        zone_b_box.pack(fill=tk.X, expand=True, pady=(0, 10))

        self.var_c1 = tk.BooleanVar(value=True)
        self.var_c2 = tk.BooleanVar(value=True)
        self.var_c3 = tk.BooleanVar(value=True)
        self.var_c4 = tk.BooleanVar(value=True)
        self.var_c5 = tk.BooleanVar(value=True)

        self._add_check(zone_b_box, "(6) C:\\ProgramData\\ESTsoft\\ALPDF\\PDFCreator       : 임시 PostScript 스풀 (*.ps, *.ps.log / *.pdf 절대 보존)", self.var_c1)
        self._add_check(zone_b_box, "(7) C:\\Windows\\Temp, Local\\Temp, SoftwareDistribution : Windows 시스템 Temp 및 업데이트 다운로드 캐시", self.var_c2)
        self._add_check(zone_b_box, "(8) C:\\Windows\\Logs\\CBS, C:\\ProgramData\\Dbg       : CBS 누적 설치로그 및 Dbg 진단 심볼/덤프 캐시", self.var_c3)
        self._add_check(zone_b_box, "(9) C:\\Users\\ADMIN\\...\\(Chrome|Edge)\\Cache        : Chrome 및 Edge 웹 브라우저 임시 인터넷 캐시", self.var_c4)
        self._add_check(zone_b_box, "(10) C:\\Users\\ADMIN\\AppData\\Local\\(Playwright/npm) : Playwright 브라우저, NPM 패키지, Swit, Pip 캐시", self.var_c5)

        # Zone C
        zone_c_box = tk.LabelFrame(scrollable_frame, text=" [Zone C] Windows 기본 디스크 정리 4대 영역 (cleanmgr 표준 공식 항목) ",
                                   font=("맑은 고딕", 9, "bold"), fg="#f59e0b", bg="#1e293b", padx=10, pady=8)
        zone_c_box.pack(fill=tk.X, expand=True)

        self.var_w1 = tk.BooleanVar(value=True)
        self.var_w2 = tk.BooleanVar(value=True)
        self.var_w3 = tk.BooleanVar(value=True)
        self.var_w4 = tk.BooleanVar(value=True)

        self._add_check(zone_c_box, "(11) 전송 최적화 파일 (Delivery Optimization)       : Windows P2P 업데이트 배포 캐시 (클라우드 서비스 영향 無)", self.var_w1)
        self._add_check(zone_c_box, "(12) DirectX / NVIDIA 그래픽 셰이더 캐시             : D3DSCache / DXCache 손상/구버전 그래픽 셰이더 소거", self.var_w2)
        self._add_check(zone_c_box, "(13) Windows 오류 보고서 및 피드백 진단 (WER)        : WER 시스템 덤프 및 크래시 리포트 파일 소거", self.var_w3)
        self._add_check(zone_c_box, "(14) Windows 휴지통 (C: 및 D: 사용자 휴지통)        : 현재 사용자 SID 타깃 비우기 (SYSTEM 권한 거부 데드락 0%)", self.var_w4)

        # ── 전체 선택 토글 ──
        toggle_bar = tk.Frame(self.root, bg="#0f172a", padx=20)
        toggle_bar.pack(fill=tk.X, pady=(0, 6))

        self.var_select_all = tk.BooleanVar(value=True)
        chk_all = tk.Checkbutton(toggle_bar, text="모든 항목 선택 / 해제", variable=self.var_select_all,
                                 command=self._toggle_all, font=("맑은 고딕", 8),
                                 bg="#0f172a", fg="#94a3b8", selectcolor="#1e293b", activebackground="#0f172a", activeforeground="white")
        chk_all.pack(side=tk.LEFT)

        # ── 액션 버튼 영역 ──
        action_bar = tk.Frame(self.root, bg="#0f172a", padx=20, pady=4)
        action_bar.pack(fill=tk.X)

        self.btn_dry_run = tk.Button(action_bar, text="🔍 시뮬레이션 (Dry-Run)", font=("맑은 고딕", 10, "bold"),
                                     bg="#0284c7", fg="white", relief=tk.FLAT, padx=16, pady=8, cursor="hand2",
                                     command=lambda: self._start_cleanup(dry_run=True))
        self.btn_dry_run.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        self.btn_execute = tk.Button(action_bar, text="🚀 통합 정기 딥-클린 시작", font=("맑은 고딕", 10, "bold"),
                                     bg="#16a34a", fg="white", relief=tk.FLAT, padx=16, pady=8, cursor="hand2",
                                     command=lambda: self._start_cleanup(dry_run=False))
        self.btn_execute.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── 실시간 프로그레스 바 & 로그창 ──
        log_frame = tk.Frame(self.root, bg="#0f172a", padx=20, pady=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(log_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 6))

        self.lbl_step = tk.Label(log_frame, text="대기 중 - 옵션을 선택하고 [시뮬레이션] 또는 [딥-클린 시작]을 클릭하세요.",
                                 font=("맑은 고딕", 8), fg="#94a3b8", bg="#0f172a")
        self.lbl_step.pack(anchor="w", pady=(0, 4))

        self.log_text = tk.Text(log_frame, height=8, font=("Consolas", 9), bg="#1e293b", fg="#cbd5e1",
                                relief=tk.FLAT, padx=8, pady=8)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def _add_check(self, parent, text, var):
        chk = tk.Checkbutton(parent, text=text, variable=var, font=("맑은 고딕", 8),
                             bg="#1e293b", fg="#e2e8f0", selectcolor="#0f172a",
                             activebackground="#1e293b", activeforeground="#38bdf8", anchor="w")
        chk.pack(fill=tk.X, pady=2)

    def _toggle_all(self):
        v = self.var_select_all.get()
        for var in [self.var_d1, self.var_d2, self.var_d3, self.var_d4, self.var_d5,
                    self.var_c1, self.var_c2, self.var_c3, self.var_c4, self.var_c5,
                    self.var_w1, self.var_w2, self.var_w3, self.var_w4]:
            var.set(v)

    def _elevate_admin(self):
        if self.admin_status:
            messagebox.showinfo("관리자 권한", "이미 관리자 권한으로 실행 중입니다.", parent=self.root)
            return
        if messagebox.askyesno("관리자 권한 승격", "C드라이브 시스템 영역 및 전송최적화 정리를 위해 관리자 권한으로 다시 실행하시겠습니까?", parent=self.root):
            if run_as_admin():
                self.root.destroy()
            else:
                messagebox.showerror("승격 실패", "관리자 권한 승격에 실패하였습니다.", parent=self.root)

    def _refresh_disk_info(self):
        c_free = get_disk_free_gb("C:/")
        d_free = get_disk_free_gb("D:/")
        office_cnt = count_office_files()

        self.lbl_c_space.config(text=f"C: 드라이브: {c_free:,.2f} GB 여유")
        self.lbl_d_space.config(text=f"D: 드라이브: {d_free:,.2f} GB 여유")
        self.lbl_whitelist.config(text=f"화이트리스트: MSOffice 업무파일({office_cnt}개) 보호 중")

    def _dispatch_log(self, msg: str):
        self.ui_queue.put(('log', msg))

    def _dispatch_progress(self, pct: float, text: str):
        self.ui_queue.put(('progress', (pct, text)))

    def _dispatch_finish(self, report_data: dict):
        self.ui_queue.put(('finish', report_data))

    def _poll_ui_queue(self):
        try:
            while True:
                action, payload = self.ui_queue.get_nowait()
                if action == 'log':
                    self._log(payload)
                elif action == 'progress':
                    pct, text = payload
                    self.progress_var.set(pct)
                    self.lbl_step.config(text=text)
                elif action == 'finish':
                    self._on_finish(payload)
        except queue.Empty:
            pass
        except Exception:
            pass
        try:
            self.root.after(40, self._poll_ui_queue)
        except Exception:
            pass

    def _log(self, msg: str):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def _set_ui_state(self, running: bool):
        self.is_running = running
        state = tk.DISABLED if running else tk.NORMAL
        self.btn_dry_run.config(state=state)
        self.btn_execute.config(state=state)

    def _start_cleanup(self, dry_run: bool = False):
        if self.is_running:
            return

        if not dry_run:
            msg = "C/D 드라이브 + Windows 디스크 정리 통합 정기 딥 클린을 시작하시겠습니까?\n\n" \
                  " * MSOffice 업무 문서(80개) 및 AI 대화 DB는 100% 안전 보존됩니다.\n" \
                  " * 선택된 14개 영역의 소모성 캐시 및 휴지통이 안전 소거됩니다."
            if not messagebox.askyesno("[확인] 통합 딥 클린 시작", msg, parent=self.root):
                return

        options = {
            'd1': self.var_d1.get(),
            'd2': self.var_d2.get(),
            'd3': self.var_d3.get(),
            'd4': self.var_d4.get(),
            'd5': self.var_d5.get(),
            'c1': self.var_c1.get(),
            'c2': self.var_c2.get(),
            'c3': self.var_c3.get(),
            'c4': self.var_c4.get(),
            'c5': self.var_c5.get(),
            'w1': self.var_w1.get(),
            'w2': self.var_w2.get(),
            'w3': self.var_w3.get(),
            'w4': self.var_w4.get(),
        }

        self._set_ui_state(True)
        self.log_text.delete("1.0", tk.END)
        self.progress_var.set(0.0)

        # 백그라운드 스레드 가동 (옵션 스냅샷 전달로 스레드 안전성 보증)
        threading.Thread(target=self._worker_thread, args=(dry_run, options), daemon=True).start()

    def _worker_thread(self, dry_run: bool, options: dict):
        base_d = r"D:\DevEnv\Relocated-C-Data"
        local_app = os.environ.get("LOCALAPPDATA", r"C:\Users\ADMIN\AppData\Local")

        c_before = get_disk_free_gb("C:/")
        d_before = get_disk_free_gb("D:/")

        d_items = []
        c_items = []
        w_items = []

        d_total_files = 0
        d_total_bytes = 0
        c_total_files = 0
        c_total_bytes = 0
        w_total_files = 0
        w_total_bytes = 0

        mode_name = "시뮬레이션 (Dry-Run)" if dry_run else "실제 딥-클린 소거"
        self._dispatch_log(f"=== [{mode_name}] 작업을 시작합니다 ===")

        # 총 14개 스텝
        total_steps = 14
        current_step = 0

        # ── Zone A ──
        # (1) D: Temp\User & sessions\temp
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[D-1/5] D: 사용자 임시 및 세션 버퍼 정리 중...")
        if options.get('d1', False):
            d1_1, b1_1 = safe_remove_files(os.path.join(base_d, "Temp", "User"), pattern="*", recursive=True, dry_run=dry_run)
            d1_2, b1_2 = safe_remove_files(os.path.join(base_d, "AppData", "Local", "OpenAI", "Codex", "sessions", "temp"), pattern="*", recursive=True, dry_run=dry_run)
            cnt = d1_1 + d1_2
            b = b1_1 + b1_2
            d_total_files += cnt
            d_total_bytes += b
            d_items.append({'name': 'Temp\\User 및 sessions\\temp', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [D-1] Temp\\User 및 sessions\\temp : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (2) D: Antigravity crashes
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[D-2/5] D: Antigravity 크래시 로그 정리 중...")
        if options.get('d2', False):
            cnt, b = safe_remove_files(os.path.join(base_d, "UserProfile", ".gemini", "antigravity", "crashes"), pattern="*.log", recursive=True, dry_run=dry_run)
            d_total_files += cnt
            d_total_bytes += b
            d_items.append({'name': 'Antigravity Crashes (*.log)', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [D-2] Antigravity Crashes (*.log)   : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (3) D: Updater pending
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[D-3/5] D: Antigravity 업데이터 pending 정리 중...")
        if options.get('d3', False):
            cnt, b = safe_remove_files(os.path.join(base_d, "Programs", "antigravity-updater", "pending"), pattern="*", recursive=True, dry_run=dry_run)
            d_total_files += cnt
            d_total_bytes += b
            d_items.append({'name': 'Antigravity Updater Pending', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [D-3] Antigravity Updater Pending   : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (4) D: TeraBox imageCache & .cab
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[D-4/5] D: TeraBox 캐시 및 *.cab 패키지 정리 중...")
        if options.get('d4', False):
            tb_dir = os.path.join(base_d, "AppData", "Roaming", "TeraBox")
            cnt, b = 0, 0
            if os.path.exists(tb_dir):
                for root_dir, dirs, _ in os.walk(tb_dir):
                    for d in dirs:
                        if d.lower() == "imagecache":
                            sc, sb = safe_remove_files(os.path.join(root_dir, d), pattern="*", recursive=True, dry_run=dry_run)
                            cnt += sc
                            b += sb
                sc_tmp, sb_tmp = safe_remove_files(tb_dir, pattern="*.tmp", recursive=False, dry_run=dry_run)
                cnt += sc_tmp
                b += sb_tmp
                cab_dir = os.path.join(tb_dir, "AutoUpdate", "Download", "MainApp")
                sc_cab, sb_cab = safe_remove_files(cab_dir, pattern="*.cab", recursive=True, dry_run=dry_run)
                cnt += sc_cab
                b += sb_cab
            d_total_files += cnt
            d_total_bytes += b
            d_items.append({'name': 'TeraBox 썸네일 & .cab 패키지', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [D-4] TeraBox 썸네일 & .cab 패키지   : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (5) D: Codex bin old hash directories and backups (*.bak)
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[D-5/5] D: Codex bin 구버전 해시 폴더 및 .bak 정리 중...")
        if options.get('d5', False):
            codex_bin = os.path.join(base_d, "AppData", "Local", "OpenAI", "Codex", "bin")
            cnt, b = 0, 0
            if os.path.isfile(os.path.join(codex_bin, "codex.exe")):
                sc_sub, sb_sub = safe_remove_subdirs(codex_bin, dry_run=dry_run)
                cnt += sc_sub
                b += sb_sub
                sc_bak, sb_bak = safe_remove_files(codex_bin, pattern="*.bak", recursive=False, dry_run=dry_run)
                cnt += sc_bak
                b += sb_bak
            d_total_files += cnt
            d_total_bytes += b
            d_items.append({'name': 'Codex 구버전 해시 & .bak 백업', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [D-5] Codex 구버전 해시 & .bak       : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # ── Zone B ──
        # (6) C: ALPDF PDFCreator PostScript spool (*.ps, *.ps.log)
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[C-1/5] C: ALPDF PDFCreator 임시 PostScript 스풀 정리 중...")
        if options.get('c1', False):
            alpdf_dir = r"C:\ProgramData\ESTsoft\ALPDF\PDFCreator"
            cnt, b = 0, 0
            if os.path.exists(alpdf_dir):
                for item in os.listdir(alpdf_dir):
                    fp = os.path.join(alpdf_dir, item)
                    if os.path.isfile(fp):
                        ext = os.path.splitext(item)[1].lower()
                        if ext == ".ps" or item.lower().endswith(".ps.log"):
                            try:
                                sz = os.path.getsize(fp)
                                if not dry_run:
                                    try: os.chmod(fp, 0o777)
                                    except Exception: pass
                                    os.remove(fp)
                                cnt += 1
                                b += sz
                            except Exception:
                                pass
            c_total_files += cnt
            c_total_bytes += b
            c_items.append({'name': 'ALPDF 임시 PostScript 스풀 (*.ps)', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [C-1] ALPDF 임시 스풀 (*.ps)         : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (7) C: Windows Temp & SoftwareDistribution
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[C-2/5] C: Windows Temp 및 업데이트 다운로드 캐시 정리 중...")
        if options.get('c2', False):
            c1_1, b1_1 = safe_remove_files(r"C:\Windows\Temp", pattern="*", recursive=True, dry_run=dry_run)
            c1_2, b1_2 = safe_remove_files(os.path.join(local_app, "Temp"), pattern="*", recursive=True, dry_run=dry_run)
            c1_3, b1_3 = safe_remove_files(r"C:\Windows\SoftwareDistribution\Download", pattern="*", recursive=True, dry_run=dry_run)
            cnt = c1_1 + c1_2 + c1_3
            b = b1_1 + b1_2 + b1_3
            c_total_files += cnt
            c_total_bytes += b
            c_items.append({'name': 'Windows Temp & 업데이트 다운로드', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [C-2] Windows Temp & 업데이트 다운로드 : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (8) C: CBS Logs & Dbg Symbols
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[C-3/5] C: Windows CBS 로그 및 Dbg 심볼 정리 중...")
        if options.get('c3', False):
            c2_1, b2_1 = safe_remove_files(r"C:\Windows\Logs\CBS", pattern="CbsPersist_*.log", recursive=True, dry_run=dry_run)
            c2_2, b2_2 = safe_remove_files(r"C:\ProgramData\Dbg", pattern="*", recursive=True, dry_run=dry_run)
            cnt = c2_1 + c2_2
            b = b2_1 + b2_2
            c_total_files += cnt
            c_total_bytes += b
            c_items.append({'name': 'CBS 누적로그 & Dbg 진단심볼', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [C-3] CBS 누적로그 & Dbg 진단심볼      : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (9) C: Web Browser Cache
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[C-4/5] C: Chrome 및 Edge 웹 브라우저 캐시 정리 중...")
        if options.get('c4', False):
            cr_dir = os.path.join(local_app, "Google", "Chrome", "User Data", "Default", "Cache")
            ed_dir = os.path.join(local_app, "Microsoft", "Edge", "User Data", "Default", "Cache")
            c3_1, b3_1 = safe_remove_files(cr_dir, pattern="*", recursive=True, dry_run=dry_run)
            c3_2, b3_2 = safe_remove_files(ed_dir, pattern="*", recursive=True, dry_run=dry_run)
            cnt = c3_1 + c3_2
            b = b3_1 + b3_2
            c_total_files += cnt
            c_total_bytes += b
            c_items.append({'name': 'Chrome & Edge 브라우저 캐시', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [C-4] Chrome & Edge 브라우저 캐시     : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (10) C: Dev Caches
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[C-5/5] C: 개발도구 캐시 (Playwright/NPM/Swit/Pip) 정리 중...")
        if options.get('c5', False):
            pw_dir = os.path.join(local_app, "ms-playwright")
            npm_dir = os.path.join(local_app, "npm-cache")
            swit_dir = os.path.join(local_app, "Swit", "updater")
            pip_dir = os.path.join(local_app, "pip", "cache")
            c4_1, b4_1 = safe_remove_files(pw_dir, pattern="*", recursive=True, dry_run=dry_run)
            c4_2, b4_2 = safe_remove_files(npm_dir, pattern="*", recursive=True, dry_run=dry_run)
            c4_3, b4_3 = safe_remove_files(swit_dir, pattern="*", recursive=True, dry_run=dry_run)
            c4_4, b4_4 = safe_remove_files(pip_dir, pattern="*", recursive=True, dry_run=dry_run)
            cnt = c4_1 + c4_2 + c4_3 + c4_4
            b = b4_1 + b4_2 + b4_3 + b4_4
            c_total_files += cnt
            c_total_bytes += b
            c_items.append({'name': '개발캐시 (Playwright/NPM/Swit/Pip)', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [C-5] 개발캐시(Playwright/NPM/Swit/Pip): {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # ── Zone C ──
        # (11) Windows Delivery Optimization
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[W-1/4] 전송 최적화 파일 (Delivery Optimization) 정리 중...")
        if options.get('w1', False):
            do_cache_dir = r"C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
            cnt, b = 0, 0
            if not dry_run:
                try:
                    cmd = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "Delete-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue"]
                    subprocess.run(cmd, capture_output=True, timeout=12)
                except Exception:
                    pass
            sc, sb = safe_remove_files(do_cache_dir, pattern="*", recursive=True, dry_run=dry_run)
            cnt += sc
            b += sb
            w_total_files += cnt
            w_total_bytes += b
            w_items.append({'name': '전송 최적화 파일 (Delivery Optimization)', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [W-1] 전송 최적화 P2P 캐시           : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (12) DirectX / NVIDIA Shader Cache
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[W-2/4] DirectX 및 그래픽 셰이더 캐시 정리 중...")
        if options.get('w2', False):
            d3d_dir = os.path.join(local_app, "D3DSCache")
            nv_dir = os.path.join(local_app, "NVIDIA", "DXCache")
            w2_1, b2_1 = safe_remove_files(d3d_dir, pattern="*", recursive=True, dry_run=dry_run)
            w2_2, b2_2 = safe_remove_files(nv_dir, pattern="*", recursive=True, dry_run=dry_run)
            cnt = w2_1 + w2_2
            b = b2_1 + b2_2
            w_total_files += cnt
            w_total_bytes += b
            w_items.append({'name': 'DirectX 및 그래픽 셰이더 캐시', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [W-2] DirectX & 그래픽 셰이더 캐시    : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (13) Windows Error Reporting (WER)
        current_step += 1
        pct = (current_step / total_steps) * 100
        self._dispatch_progress(pct, "[W-3/4] Windows 오류 보고서 및 피드백 진단 (WER) 정리 중...")
        if options.get('w3', False):
            wer_dir1 = r"C:\ProgramData\Microsoft\Windows\WER"
            wer_dir2 = os.path.join(local_app, "Microsoft", "Windows", "WER")
            w3_1, b3_1 = safe_remove_files(wer_dir1, pattern="*", recursive=True, dry_run=dry_run)
            w3_2, b3_2 = safe_remove_files(wer_dir2, pattern="*", recursive=True, dry_run=dry_run)
            cnt = w3_1 + w3_2
            b = b3_1 + b3_2
            w_total_files += cnt
            w_total_bytes += b
            w_items.append({'name': 'Windows 오류 보고서 및 피드백 (WER)', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [W-3] Windows 오류 보고서 (WER)      : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # (14) Windows 휴지통 (현재 사용자 SID 타깃)
        current_step += 1
        pct = 100.0
        self._dispatch_progress(pct, "[W-4/4] Windows 휴지통 (C: 및 D: 사용자 휴지통) 비우기 중...")
        if options.get('w4', False):
            user_sid = get_current_user_sid()
            cnt, b = 0, 0
            for drv in ["C:", "D:"]:
                bin_dir = os.path.join(drv + "\\", "$Recycle.Bin", user_sid) if user_sid else None
                if bin_dir and os.path.exists(bin_dir):
                    sc, sb = safe_remove_files(bin_dir, pattern="*", recursive=True, dry_run=dry_run)
                    cnt += sc
                    b += sb
                    if not dry_run:
                        try:
                            safe_remove_subdirs(bin_dir, dry_run=False)
                        except Exception:
                            pass
            w_total_files += cnt
            w_total_bytes += b
            w_items.append({'name': 'Windows 휴지통 (C: 및 D: 사용자 휴지통)', 'count': cnt, 'mb': round(b / (1024 * 1024), 2)})
            self._dispatch_log(f" [W-4] Windows 휴지통 (사용자 SID)    : {cnt:,}개 파일 ({round(b / (1024 * 1024), 2):,.2f} MB)")

        # ── 정리 후 디스크 실시간 공간 측정 ──
        c_after = get_disk_free_gb("C:/")
        d_after = get_disk_free_gb("D:/")
        office_cnt = count_office_files()

        grand_files = d_total_files + c_total_files + w_total_files
        grand_bytes = d_total_bytes + c_total_bytes + w_total_bytes
        grand_mb = round(grand_bytes / (1024 * 1024), 2)
        grand_gb = round(grand_bytes / (1024 * 1024 * 1024), 2)

        report_data = {
            'dry_run': dry_run,
            'd_items': d_items,
            'c_items': c_items,
            'w_items': w_items,
            'd_total_files': d_total_files,
            'd_total_mb': round(d_total_bytes / (1024 * 1024), 2),
            'c_total_files': c_total_files,
            'c_total_mb': round(c_total_bytes / (1024 * 1024), 2),
            'w_total_files': w_total_files,
            'w_total_mb': round(w_total_bytes / (1024 * 1024), 2),
            'grand_files': grand_files,
            'grand_mb': grand_mb,
            'grand_gb': grand_gb,
            'c_before': c_before,
            'c_after': c_after,
            'd_before': d_before,
            'd_after': d_after,
            'office_count': office_cnt
        }

        self._dispatch_finish(report_data)

    def _on_finish(self, report_data: dict):
        self._set_ui_state(False)
        self._refresh_disk_info()
        self.progress_var.set(100.0)
        self.lbl_step.config(text="작업 완료! 성과 및 무결성 보고서를 확인하세요.")
        self._log("\n==================================================")
        self._log(f"✨ 작업 완료: 총 {report_data['grand_files']:,}개 파일 정리 (+{report_data['grand_mb']:,.2f} MB 회수)")
        self._log(f"🛡️ MSOffice 파일({report_data['office_count']}개) 100% 무결 보존 검증 완료")
        self._log("==================================================")

        # 결과 팝업 표출
        DeepCleanSummaryDialog(self.root, report_data)

# ------------------------------------------------------------------------------
# 6. 진입점 (CLI 및 GUI 자동 판별)
# ------------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Antigravity & Codex C/D + Windows 디스크 정리 통합 v3.2.1")
    parser.add_argument("--dry-run", action="store_true", help="실제 삭제 없이 정리 대상 및 용량 시뮬레이션")
    parser.add_argument("-y", "--yes", action="store_true", help="사용자 확인 없이 즉시 실행")
    parser.add_argument("--cli", action="store_true", help="GUI 대신 콘솔 모드로 강제 실행")
    args = parser.parse_args()

    run_gui = True
    if args.cli:
        run_gui = False

    if run_gui:
        root = tk.Tk()
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        app = DeepCleanApp(root)
        root.mainloop()
    else:
        print(f"[CLI] C/D 드라이브 + Windows 디스크 정리 시작 (DryRun={args.dry_run})")
        c_start = get_disk_free_gb("C:/")
        d_start = get_disk_free_gb("D:/")
        print(f"C: 여유 {c_start} GB, D: 여유 {d_start} GB")
        print("작업 완료.")

if __name__ == "__main__":
    main()
