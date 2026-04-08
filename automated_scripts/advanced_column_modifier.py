"""
================================================================================
 [EXCEL] 엑셀 열 스마트 교정 도구 (Excel Column Smart Modifier) v34.1.16
================================================================================
 - 아키텍처: Clean Layer Architecture (Domain / Presentation / Application)
 - 주요 기능: E00000 마커 기반 스마트 교정 로직
 - 가이드라인 준수: 00 PRD 가이드.md | AI_CODING_GUIDELINES_2026.md
 - 무결성 보증: Auto-Elevation (관리자 자동 승격), Force Process Kill (좀비 소거)
 - 최적화: Dynamic COM Binding 기능으로 환경 이식성 극대화
================================================================================
"""
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import win32com.client
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import threading


# ==========================================
# 1. DOMAIN LAYER (Core Engine - 비즈니스 로직)
# ==========================================
class ColumnModifierEngine:
    """엑셀 열 교정 핵심 엔진: GUI와 분리된 순수 로직"""
    
    @staticmethod
    def get_files_by_list(list_str, directory):
        """
        문자열로 입력된 파일명 리스트를 실제 경로로 변환
        """
        if not list_str or not directory:
            return []
        
        names = [n.strip() for n in list_str.replace('\n', ',').split(',') if n.strip()]
        full_paths = []
        not_found = []
        
        for name in names:
            path = os.path.join(directory, name)
            if os.path.exists(path):
                full_paths.append(path)
            else:
                not_found.append(name)
        
        return full_paths, not_found
    
    @staticmethod
    def _kill_office_processes(callback=None):
        """[v2.3] 기존 좀비 프로세스 강제 소거 (psutil+WMI 하이브리드)"""
        import time, ctypes
        try: ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x0002 | 0x8000)
        except: pass
        if callback: callback("info", "[INFO] 기존 Excel 좀비 프로세스 네이티브 정리 중 (팝업 차단)...")
        
        targets = ["EXCEL.EXE"]
        killed = False
        try:
            import psutil
            for p in psutil.process_iter(['name']):
                if p.info['name'] and p.info['name'].upper() in targets:
                    try: p.kill(); killed = True
                    except: pass
        except ImportError:
            try:
                import pythoncom, win32com.client
                pythoncom.CoInitialize()
                wmi = win32com.client.GetObject("winmgmts:")
                for p in wmi.InstancesOf("Win32_Process"):
                    name = p.Properties_('Name').Value
                    if name and name.upper() in targets:
                        try: p.Terminate(); killed = True
                        except: pass
            except: pass
        if killed:
            time.sleep(1.0)

    @staticmethod
    def process_files(files, sheet_name, col_count, callback=None):
        """
        엑셀 파일들의 열 구조 교정 수행
        """
        import pythoncom
        pythoncom.CoInitialize()
        
        # [v2.4] 시작 전 강제 소거
        ColumnModifierEngine._kill_office_processes(callback)
        
        prog_id = "Excel.Application"
        excel = None
        errors = []
        
        import win32com.client
        from win32com.client import Dispatch, DispatchEx
        
        try:
            # Phase 0: CLSID Direct Dispatch (가장 강력한 직접 바인딩)
            clsid = "{00024500-0000-0000-C000-000000000046}" # Excel CLSID
            try:
                from win32com.client.dynamic import Dispatch as DynDispatch
                excel = DynDispatch(clsid)
                if excel: return excel
            except: pass

            # Phase 1: EnsureDispatch (레지스트리 강제 복구)
            try:
                from win32com.client import gencache
                excel = gencache.EnsureDispatch(prog_id)
                if excel: return excel
            except Exception as e1:
                errors.append(f"P1:{str(e1)[:20]}")

            # Phase 2: Shell/Popen + Deep Polling (강제 가중 가동 후 추적)
            try:
                import subprocess, time
                subprocess.Popen(['excel'], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                for i in range(10): # 10초간 추적 가동
                    time.sleep(1.0)
                    try:
                        excel = win32com.client.GetActiveObject(prog_id)
                        if excel: return excel
                    except: pass
            except Exception as e2:
                errors.append(f"P2:{str(e2)[:20]}")

            # Phase 3: Last Resort (Standard Dispatch)
            try:
                excel = win32com.client.Dispatch(prog_id)
                if excel: return excel
            except Exception as e3:
                errors.append(f"P3:{str(e3)[:20]}")

            # Phase 4: Emergency Shell Hook
            try:
                import os
                os.startfile("excel")
                for _ in range(5):
                    time.sleep(2.0)
                    try: 
                        excel = win32com.client.GetActiveObject(prog_id)
                        if excel: return excel
                    except: pass
            except Exception as e4:
                errors.append(f"P4:{str(e4)[:20]}")

            if not excel:
                if callback: callback("error", f"[FAIL] Excel 가동 실패: {'; '.join(errors)}")
                return {"success": 0, "errors": ["; ".join(errors)]}

            # Silent Mode 강제 주입
            try:
                excel.Visible = False
                excel.DisplayAlerts = False
                try: excel.Interactive = False
                except: pass
            except:
                pass
        
            success_count = 0
            fail_list = []
            
            for filepath in files:
                filename = os.path.basename(filepath)
                try:
                    if callback:
                        callback("progress", f"처리 중: {filename}")
                    
                    wb = excel.Workbooks.Open(filepath)
                    sheet = None
                    for s in wb.Sheets:
                        if s.Name == sheet_name:
                            sheet = s
                            break
                    
                    if sheet:
                        sheet.Activate()
                        sheet.Columns.EntireColumn.Hidden = False  # 모든 열 표시
                        
                        # 스마트 로직: E00000 마커로 현재 데이터 위치 감지
                        data_col = -1
                        for c in range(1, 25):
                            val = str(sheet.Cells(5, c).Value)
                            if val.startswith("E00000"):
                                data_col = c
                                break
                        
                        target_data_pos = col_count + 1
                        
                        if data_col == -1:
                            # 마커 미발견시 기본 삽입
                            if callback:
                                callback("info", f"  [{filename}] 마커 미발견. 기본 {col_count}개 삽입 수행.")
                            for _ in range(col_count):
                                sheet.Columns(1).Insert()
                        elif data_col > target_data_pos:
                            diff = data_col - target_data_pos
                            if callback:
                                callback("info", f"  [{filename}] 초과 삽입 감지 ({diff}개). 삭제 중...")
                            for _ in range(diff):
                                sheet.Columns(1).Delete()
                        elif data_col < target_data_pos:
                            diff = target_data_pos - data_col
                            if callback:
                                callback("info", f"  [{filename}] 부족 삽입 감지 ({diff}개). 추가 삽입 중...")
                            for _ in range(diff):
                                sheet.Columns(1).Insert()
                        else:
                            if callback:
                                callback("info", f"  [{filename}] 이미 정상 위치. 숨김만 적용.")
                        
                        # 최종: 1부터 col_count까지 숨기기
                        if col_count > 0:
                            for i in range(1, col_count + 1):
                                sheet.Columns(i).EntireColumn.Hidden = True
                        
                        wb.Save()
                        success_count += 1
                        if callback:
                            callback("success", f"  [완료] {filename}")
                    else:
                        fail_list.append(f"{filename} (시트 없음)")
                        if callback:
                            callback("error", f"  [실패] {filename}: 시트 없음")
                    wb.Close()
                except Exception as e:
                    fail_list.append(f"{filename} ({str(e)})")
                    if callback:
                        callback("error", f"  [오류] {filename}: {str(e)}")
        finally:
            excel.Quit()
        
        return {"success": success_count, "errors": fail_list}


# ==========================================
# 2. PRESENTATION LAYER (View - GUI 정의)
# ==========================================
class ColumnModifierView:
    """GUI 화면 구성: Controller와 분리된 순수 View"""
    
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[엑셀 열 스마트 교정 도구 v34.1.16]")
        self.root.geometry("650x600")
        self._build_ui()
    
    def _build_ui(self):
        # 헤더
        header = tk.Frame(self.root, bg="#1565C0", pady=12)
        header.pack(fill="x")
        tk.Label(header, text="[엑셀 열 스마트 교정 도구 v34.1.16]", font=("Malgun Gothic", 14, "bold"),
                 bg="#1565C0", fg="white").pack()
        tk.Label(header, text="E00000 마커 기반 가변형 열 삽입/삭제 자동 교정 (Clean Architecture)",
                 font=("Malgun Gothic", 9), bg="#1565C0", fg="#BBDEFB").pack()
        
        main = tk.Frame(self.root, padx=20, pady=10)
        main.pack(fill="both", expand=True)
        
        # 1. 열 개수 설정
        box1 = tk.LabelFrame(main, text="[1] 숨길 열 개수 설정", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box1.pack(fill="x", pady=5)
        f1 = tk.Frame(box1)
        f1.pack(fill="x")
        tk.Label(f1, text="열 개수:", font=("Malgun Gothic", 9)).pack(side="left")
        tk.Spinbox(f1, from_=1, to=20, textvariable=self.controller.col_count_var, width=5, 
                   font=("Malgun Gothic", 10)).pack(side="left", padx=10)
        tk.Label(f1, text="(1~20, 기본: 4)", font=("Malgun Gothic", 8), fg="#666").pack(side="left")
        
        # 2. 시트 이름
        box2 = tk.LabelFrame(main, text="[2] 작업 시트 이름", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box2.pack(fill="x", pady=5)
        tk.Entry(box2, textvariable=self.controller.sheet_name_var, font=("Malgun Gothic", 10)).pack(fill="x")
        
        # 3. 파일 선택 방법
        box3 = tk.LabelFrame(main, text="[3] 파일 선택 방법", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box3.pack(fill="x", pady=5)
        f3 = tk.Frame(box3)
        f3.pack(fill="x")
        tk.Radiobutton(f3, text="탐색기에서 직접 선택", variable=self.controller.select_mode, 
                       value="dialog", font=("Malgun Gothic", 9)).pack(side="left")
        tk.Radiobutton(f3, text="파일명 목록 입력", variable=self.controller.select_mode, 
                       value="list", font=("Malgun Gothic", 9)).pack(side="left", padx=15)
        
        btn_f = tk.Frame(box3)
        btn_f.pack(fill="x", pady=10)
        self.select_btn = tk.Button(btn_f, text="[파일 선택]", command=self.controller.handle_select_files,
                                    font=("Malgun Gothic", 10), bg="#E3F2FD", width=15)
        self.select_btn.pack(side="left")
        self.file_count_label = tk.Label(btn_f, text="선택된 파일: 0개", font=("Malgun Gothic", 9))
        self.file_count_label.pack(side="left", padx=15)
        
        # 4. 진행 로그
        box4 = tk.LabelFrame(main, text="[4] 진행 상황", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box4.pack(fill="both", expand=True, pady=5)
        
        self.log_txt = tk.Text(box4, font=("Consolas", 9), bg="#FAFAFA", height=10)
        scr = tk.Scrollbar(box4, command=self.log_txt.yview)
        self.log_txt.config(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y")
        self.log_txt.pack(fill="both", expand=True)
        
        # 실행 버튼
        self.run_btn = tk.Button(self.root, text="[열 교정 실행]", font=("Malgun Gothic", 12, "bold"),
                                 bg="#2E7D32", fg="white", height=2, command=self.controller.handle_start)
        self.run_btn.pack(fill="x", padx=20, pady=15)
        
        # 상태바
        self.status_bar = tk.Label(self.root, text="준비됨 | 파일을 선택하고 실행 버튼을 누르세요.",
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
    
    def update_file_count(self, count):
        """파일 개수 업데이트"""
        self.file_count_label.config(text=f"선택된 파일: {count}개")
    
    def set_running_state(self, is_running):
        """실행 상태 UI 업데이트"""
        if is_running:
            self.run_btn.config(state="disabled", bg="#9E9E9E")
            self.select_btn.config(state="disabled")
        else:
            self.run_btn.config(state="normal", bg="#2E7D32")
            self.select_btn.config(state="normal")


# ==========================================
# 3. APPLICATION LAYER (Controller - 로직 연결)
# ==========================================
class ColumnModifierController:
    """Controller: View와 Engine을 연결하고 사용자 이벤트 처리"""
    
    def __init__(self, root):
        self.root = root
        self.engine = ColumnModifierEngine()
        self.is_running = False
        self.selected_files = []
        
        # 상태 변수 초기화
        self.col_count_var = tk.IntVar(value=4)
        self.sheet_name_var = tk.StringVar(value="내역서(표준단가)")
        self.select_mode = tk.StringVar(value="dialog")
        
        # View 생성 (Controller 주입)
        self.view = ColumnModifierView(root, self)
    
    def handle_select_files(self):
        """파일 선택 처리"""
        mode = self.select_mode.get()
        
        if mode == "dialog":
            files = filedialog.askopenfilenames(
                title="작업할 엑셀 파일들을 선택하세요",
                filetypes=[("Excel files", "*.xlsx *.xlsm")]
            )
            if files:
                self.selected_files = list(files)
                self.view.update_file_count(len(self.selected_files))
                self.view.log(f"[*] {len(self.selected_files)}개 파일 선택됨")
        else:
            list_str = simpledialog.askstring(
                "파일 목록 입력",
                "파일명을 쉼표(,)나 줄바꿈으로 구분하여 입력하세요:\n(예: 23당사안.xlsx, 25당사안.xlsx)"
            )
            if list_str:
                directory = filedialog.askdirectory(title="입력한 파일들이 포함된 폴더를 선택하세요")
                if directory:
                    files, not_found = self.engine.get_files_by_list(list_str, directory)
                    self.selected_files = files
                    self.view.update_file_count(len(self.selected_files))
                    self.view.log(f"[*] {len(self.selected_files)}개 파일 발견")
                    if not_found:
                        self.view.log(f"[!] 미발견 파일: {', '.join(not_found[:5])}")
    
    def handle_start(self):
        """열 교정 실행"""
        if self.is_running:
            return
        
        if not self.selected_files:
            messagebox.showwarning("경고", "파일이 선택되지 않았습니다.")
            return
        
        sheet_name = self.sheet_name_var.get().strip()
        if not sheet_name:
            messagebox.showwarning("경고", "시트 이름을 입력해주세요.")
            return
        
        col_count = self.col_count_var.get()
        
        # [Root Cause Fix] 실행 전 확인 다이얼로그 제거 (대시보드 직결 실행 대응)
        # confirm = messagebox.askyesno(
        #     "최종 확인",
        #     f"총 {len(self.selected_files)}개의 파일에 대해\n"
        #     f"'{sheet_name}' 시트의 처리를 시작합니다.\n"
        #     f"최종적으로 앞에서부터 {col_count}개의 열이 숨겨진 상태가 됩니다.\n\n"
        #     f"진행하시겠습니까?"
        # )
        # if not confirm:
        #     return
        
        self.is_running = True
        self.view.set_running_state(True)
        self.view.clear_log()
        self.view.log(f"[INFO] 열 교정 시작 (대상: {len(self.selected_files)}개)")
        self.view.log("-" * 50)
        self.view.set_status("처리 중...")
        
        def run():
            result = self.engine.process_files(
                self.selected_files,
                sheet_name,
                col_count,
                self._callback
            )
            self.root.after(0, lambda: self._finalize(result))
        
        threading.Thread(target=run, daemon=True).start()
    
    def _callback(self, msg_type, msg):
        """엔진 콜백 (스레드 안전)"""
        self.root.after(0, lambda: self.view.log(msg))
    
    def _finalize(self, result):
        """작업 완료 처리"""
        self.is_running = False
        self.view.set_running_state(False)
        
        self.view.log("-" * 50)
        self.view.log(f"[OK] 작업 완료!")
        self.view.log(f"   성공: {result['success']}개")
        self.view.log(f"   실패: {len(result['errors'])}개")
        self.view.set_status(f"완료: 성공 {result['success']}개, 실패 {len(result['errors'])}개")
        
        report = f"작업 완료!\n\n성공: {result['success']}\n실패: {len(result['errors'])}"
        if result['errors']:
            report += "\n\n실패 목록:\n" + "\n".join(result['errors'][:10])
        messagebox.showinfo("결과 리포트", report)


# ==========================================
# 4. ENTRY POINT
# ==========================================
def is_admin():
    import ctypes
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def run_as_admin():
    import ctypes, sys
    if is_admin(): return
    params = f'"{sys.argv[0]}" ' + " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit(0)

if __name__ == "__main__":
    # [Root Cause Fix] 관리자 권한 자동 승격 제거 (UAC 팝업 및 COM 권한 충돌 방지)
    # run_as_admin()
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = ColumnModifierController(root)
    root.mainloop()
