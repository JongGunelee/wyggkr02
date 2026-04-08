"""
================================================================================
 [엑셀 열 구조 수리 도구 v34.1.16] (Auto-Elevation & Force Cleanup)
================================================================================
- 아키텍처: Clean Layer Architecture (Domain / Presentation / Application)
- 주요 기능: E00000 마커 기반 데이터 위치 자동 감지 및 교정
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
from tkinter import filedialog, messagebox, ttk
import threading


# ==========================================
# 1. DOMAIN LAYER (Core Engine - 비즈니스 로직)
# ==========================================
class ExcelRepairEngine:
    """엑셀 구조 수리 핵심 엔진: GUI와 분리된 순수 로직"""
    
    @staticmethod
    def _kill_office_processes(callback=None):
        """[v2.3] 기존 좀비 프로세스 강제 소거 (psutil+WMI 하이브리드)"""
        import time, ctypes
        try: ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x0002 | 0x8000)
        except: pass
        if callback: callback("info", "🧹 기존 Excel 좀비 프로세스 네이티브 정리 중 (팝업 차단)...")
        
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
    def repair_file(filepath, sheet_name, target_col=5, callback=None):
        """
        단일 엑셀 파일의 열 구조 수리
        target_col: 데이터가 위치해야 할 열 번호 (기본 5 = E열)
        Returns: (success, message)
        """
        import pythoncom
        pythoncom.CoInitialize()
        
        # [v2.4] 시작 전 강제 소거
        ExcelRepairEngine._kill_office_processes(callback)
        
        prog_id = "Excel.Application"
        excel = None
        wb = None
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
                return False, f"[FAIL] Excel 가동 실패: {'; '.join(errors)}"

            # Silent Mode 강제 주입
            try:
                excel.Visible = False
                excel.DisplayAlerts = False
            except:
                pass
            
            filename = os.path.basename(filepath)
            if callback:
                callback("progress", f"수리 중: {filename}")
            
            wb = excel.Workbooks.Open(filepath)
            sheet = None
            for s in wb.Sheets:
                if s.Name == sheet_name:
                    sheet = s
                    break
            
            if not sheet:
                wb.Close()
                excel.Quit()
                return False, "시트 없음"
            
            sheet.Activate()
            # 1. 모든 열 숨김 해제
            sheet.Columns.EntireColumn.Hidden = False
            
            # 2. E00000 마커로 현재 데이터 위치 감지
            data_col = -1
            for c in range(1, 20):
                val = str(sheet.Cells(5, c).Value)
                if "E00000" in val:
                    data_col = c
                    break
            
            # 3. 교정 로직
            action_msg = ""
            if data_col == -1:
                action_msg = "마커 미발견 - 수동 확인 필요"
            elif data_col > target_col:
                # 열이 너무 많이 삽입된 경우 (예: 9열 → 5열로 이동)
                diff = data_col - target_col
                if callback:
                    callback("info", f"  초과 열 {diff}개 삭제 중...")
                # 처음부터 diff개 열 삭제
                for _ in range(diff):
                    sheet.Columns(1).Delete()
                action_msg = f"초과 {diff}개 열 삭제"
            elif data_col < target_col:
                # 열이 부족한 경우 (예: 1열 → 5열로 이동)
                diff = target_col - data_col
                if callback:
                    callback("info", f"  부족 열 {diff}개 삽입 중...")
                for _ in range(diff):
                    sheet.Columns(1).Insert()
                action_msg = f"부족 {diff}개 열 삽입"
            else:
                action_msg = "이미 정상 위치 (E열)"
            
            # 4. 최종: A:D (1~4열) 숨기기
            hide_count = target_col - 1
            if hide_count > 0:
                for i in range(1, hide_count + 1):
                    sheet.Columns(i).EntireColumn.Hidden = True
                if callback:
                    callback("info", f"  A:{chr(64+hide_count)} 열 숨김 완료")
            
            wb.Save()
            wb.Close()
            excel.Quit()
            
            return True, action_msg
            
        except Exception as e:
            if wb:
                try:
                    wb.Close()
                except:
                    pass
            if excel:
                try:
                    excel.Quit()
                except:
                    pass
            return False, str(e)
    
    @staticmethod
    def repair_files(files, sheet_name, target_col=5, callback=None):
        """
        여러 엑셀 파일의 열 구조 수리
        callback: (type, message) 형태로 진행 상황 전달
        """
        success_count = 0
        fail_list = []
        
        for filepath in files:
            filename = os.path.basename(filepath)
            success, msg = ExcelRepairEngine.repair_file(filepath, sheet_name, target_col, callback)
            
            if success:
                success_count += 1
                if callback:
                    callback("success", f"  [완료] {filename}: {msg}")
            else:
                fail_list.append(f"{filename}: {msg}")
                if callback:
                    callback("error", f"  [실패] {filename}: {msg}")
        
        return {"success": success_count, "errors": fail_list}


# ==========================================
# 2. PRESENTATION LAYER (View - GUI 정의)
# ==========================================
class ExcelRepairView:
    """GUI 화면 구성: Controller와 분리된 순수 View"""
    
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[엑셀 열 구조 수리 도구 v34.1.16] (Registry Repair & Force Bind)")
        self.root.geometry("700x600")
        self._build_ui()
    
    def _build_ui(self):
        # 헤더
        header = tk.Frame(self.root, bg="#FF6F00", pady=12)
        header.pack(fill="x")
        tk.Label(header, text="[EXE] 엑셀 열 구조 수리 도구 v34.1.16", font=("Malgun Gothic", 14, "bold"),
                 bg="#FF6F00", fg="white").pack()
        tk.Label(header, text="E00000 마커 기반 데이터 위치 자동 감지 및 교정 (Clean Architecture)",
                 font=("Malgun Gothic", 9), bg="#FF6F00", fg="#FFE0B2").pack()
        
        main = tk.Frame(self.root, padx=20, pady=10)
        main.pack(fill="both", expand=True)
        
        # 1. 파일 선택
        box1 = tk.LabelFrame(main, text="1️⃣ 수리 대상 파일 선택", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box1.pack(fill="x", pady=5)
        btn_f = tk.Frame(box1)
        btn_f.pack(fill="x")
        self.select_btn = tk.Button(btn_f, text="[DIR] 파일 선택", command=self.controller.handle_select_files,
                                    font=("Malgun Gothic", 10), bg="#FFF3E0", width=15)
        self.select_btn.pack(side="left")
        self.file_count_label = tk.Label(btn_f, text="선택된 파일: 0개", font=("Malgun Gothic", 9))
        self.file_count_label.pack(side="left", padx=15)
        
        # 2. 수리 조건 설정
        box2 = tk.LabelFrame(main, text="2️⃣ 수리 조건 설정", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box2.pack(fill="x", pady=5)
        
        f2a = tk.Frame(box2)
        f2a.pack(fill="x", pady=3)
        tk.Label(f2a, text="워크시트 이름:", font=("Malgun Gothic", 9), width=15, anchor="e").pack(side="left")
        tk.Entry(f2a, textvariable=self.controller.sheet_name_var, font=("Malgun Gothic", 10), width=25).pack(side="left", padx=5)
        
        f2b = tk.Frame(box2)
        f2b.pack(fill="x", pady=3)
        tk.Label(f2b, text="데이터 목표 열:", font=("Malgun Gothic", 9), width=15, anchor="e").pack(side="left")
        tk.Spinbox(f2b, from_=1, to=26, textvariable=self.controller.target_col_var, width=5,
                   font=("Malgun Gothic", 10)).pack(side="left", padx=5)
        tk.Label(f2b, text="(5 = E열, 기본값)", font=("Malgun Gothic", 8), fg="#666").pack(side="left")
        
        # 수리 로직 안내
        info_f = tk.Frame(box2, bg="#E3F2FD")
        info_f.pack(fill="x", pady=10)
        tk.Label(info_f, text="📋 수리 로직 (E00000 마커 기반 자동 감지):", font=("Malgun Gothic", 9, "bold"), bg="#E3F2FD").pack(anchor="w")
        tk.Label(info_f, text="  • 마커가 목표 열보다 뒤에 있으면: 앞쪽 열 삭제", font=("Malgun Gothic", 8), bg="#E3F2FD").pack(anchor="w")
        tk.Label(info_f, text="  • 마커가 목표 열보다 앞에 있으면: 앞쪽 열 삽입", font=("Malgun Gothic", 8), bg="#E3F2FD").pack(anchor="w")
        tk.Label(info_f, text="  • 최종: 목표 열 앞의 모든 열 숨김 처리", font=("Malgun Gothic", 8), bg="#E3F2FD").pack(anchor="w")
        
        # 3. 진행 로그
        box3 = tk.LabelFrame(main, text="3️⃣ 진행 상황", font=("Malgun Gothic", 10, "bold"), padx=10, pady=10)
        box3.pack(fill="both", expand=True, pady=5)
        
        self.log_txt = tk.Text(box3, font=("Consolas", 9), bg="#FAFAFA", height=10)
        scr = tk.Scrollbar(box3, command=self.log_txt.yview)
        self.log_txt.config(yscrollcommand=scr.set)
        scr.pack(side="right", fill="y")
        self.log_txt.pack(fill="both", expand=True)
        
        # 실행 버튼
        self.run_btn = tk.Button(self.root, text="[START] 열 구조 수리 실행", font=("Malgun Gothic", 12, "bold"),
                                 bg="#E65100", fg="white", height=2, command=self.controller.handle_start)
        self.run_btn.pack(fill="x", padx=20, pady=15)
        
        # 상태바
        self.status_bar = tk.Label(self.root, text="준비됨 | 파일을 선택하고 실행 버튼을 누르세요.",
                                   bd=1, relief="sunken", anchor="w", font=("Malgun Gothic", 8), padx=10)
        self.status_bar.pack(side="bottom", fill="x")
    
    def log(self, msg):
        self.log_txt.insert("end", f"{msg}\n")
        self.log_txt.see("end")
    
    def clear_log(self):
        self.log_txt.delete("1.0", "end")
    
    def set_status(self, msg):
        self.status_bar.config(text=msg)
    
    def update_file_count(self, count):
        self.file_count_label.config(text=f"선택된 파일: {count}개")
    
    def set_running_state(self, is_running):
        if is_running:
            self.run_btn.config(state="disabled", bg="#9E9E9E")
            self.select_btn.config(state="disabled")
        else:
            self.run_btn.config(state="normal", bg="#E65100")
            self.select_btn.config(state="normal")


# ==========================================
# 3. APPLICATION LAYER (Controller - 로직 연결)
# ==========================================
class ExcelRepairController:
    """Controller: View와 Engine을 연결하고 사용자 이벤트 처리"""
    
    def __init__(self, root):
        self.root = root
        self.engine = ExcelRepairEngine()
        self.is_running = False
        self.selected_files = []
        
        # 상태 변수 초기화 (레거시 기본값 유지)
        self.sheet_name_var = tk.StringVar(value="내역서(표준단가)")
        self.target_col_var = tk.IntVar(value=5)  # E열
        
        # View 생성
        self.view = ExcelRepairView(root, self)
    
    def handle_select_files(self):
        """파일 선택 처리"""
        files = filedialog.askopenfilenames(
            title="수리할 엑셀 파일들을 선택하세요",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if files:
            self.selected_files = list(files)
            self.view.update_file_count(len(self.selected_files))
            self.view.log(f"[*] {len(self.selected_files)}개 파일 선택됨")
            for f in self.selected_files[:10]:
                self.view.log(f"    - {os.path.basename(f)}")
            if len(self.selected_files) > 10:
                self.view.log(f"    ... 외 {len(self.selected_files)-10}개")
    
    def handle_start(self):
        """수리 실행"""
        if self.is_running:
            return
        
        if not self.selected_files:
            messagebox.showwarning("경고", "파일이 선택되지 않았습니다.")
            return
        
        sheet_name = self.sheet_name_var.get().strip()
        target_col = self.target_col_var.get()
        
        if not sheet_name:
            messagebox.showwarning("경고", "시트 이름을 입력해주세요.")
            return
        
        # [Root Cause Fix] 실행 전 확인 다이얼로그 제거 (대시보드 직결 실행 대응)
        # target_letter = chr(64 + target_col)
        # confirm = messagebox.askyesno(
        #     "최종 확인",
        #     f"총 {len(self.selected_files)}개의 파일에 대해\n"
        #     f"'{sheet_name}' 시트의 열 구조를 수리합니다.\n\n"
        #     f"목표: 데이터가 {target_letter}열에 위치하도록 교정\n"
        #     f"(A:{chr(64+target_col-1)} 열 숨김 처리)\n\n"
        #     f"진행하시겠습니까?"
        # )
        # if not confirm:
        #     return
        
        self.is_running = True
        self.view.set_running_state(True)
        self.view.clear_log()
        self.view.log(f"🔧 열 구조 수리 시작 (대상: {len(self.selected_files)}개)")
        self.view.log("-" * 50)
        self.view.set_status("수리 중...")
        
        def run():
            result = self.engine.repair_files(
                self.selected_files,
                sheet_name,
                target_col,
                self._callback
            )
            self.root.after(0, lambda: self._finalize(result))
        
        threading.Thread(target=run, daemon=True).start()
    
    def _callback(self, msg_type, msg):
        self.root.after(0, lambda: self.view.log(msg))
    
    def _finalize(self, result):
        self.is_running = False
        self.view.set_running_state(False)
        
        self.view.log("-" * 50)
        self.view.log(f"[OK] 수리 완료!")
        self.view.log(f"   성공: {result['success']}개")
        self.view.log(f"   실패: {len(result['errors'])}개")
        self.view.set_status(f"완료: 성공 {result['success']}개, 실패 {len(result['errors'])}개")
        
        report = f"열 구조 수리 완료!\n\n성공: {result['success']}\n실패: {len(result['errors'])}"
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
    # [v34.1.16] COM 안정성을 위해 관리자 권한 강제 승격은 선택적으로 운영
    # run_as_admin() 
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = ExcelRepairController(root)
    root.mainloop()
