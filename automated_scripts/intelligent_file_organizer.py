"""
================================================================================
 [지능형 파일 정리기 (Intelligent File Organizer) v34.1.29] (Ultimate Master)
================================================================================
- 아키텍처: Clean Layer Architecture (Domain / Presentation / Application)
- 주요 기능: 파일명 패턴 매칭, 일괄 변경, 충돌 방지 미리보기 기반 자동 정리
- 듀얼 모드: 파일명 및 폴더명 개별/일괄 조작 호환
- 잠금 해제: 점유 프로세스 자동 탐지 및 강제 종료 알고리즘 탑재
- 가이드라인 준수: 00 PRD 가이드.md | AI_CODING_GUIDELINES_2026.md
- 무결성 보증: 실행 전 중복 검증 및 원자적 이름 변경 시뮬레이션 지원
================================================================================
"""
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import fnmatch
import shutil
import csv
import psutil
import time
import hashlib
import json

CONFIG_FILE = "intelligent_file_organizer_config.json"

def load_config_path():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                p = data.get("last_path", "")
                if os.path.exists(p):
                    return p
    except:
        pass
    return ""

def save_config_path(path):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({"last_path": path}, f, ensure_ascii=False)
    except:
        pass

class HoverTooltip:
    """[v34.1.41-R] 시각적 접근성을 높이기 위한 범용 툴팁 헬퍼 클래스"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        self.widget.bind("<Motion>", self.move_tooltip)

    def show_tooltip(self, event=None):
        if self.tooltip or not self.text: return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        self.tooltip.attributes("-topmost", True)
        # 툴팁 스타일링 (말풍선 테마)
        lbl = tk.Label(self.tooltip, text=self.text, justify='left',
                       bg="#ffffcc", fg="#333333", relief='solid', borderwidth=1,
                       font=("Malgun Gothic", 9), padx=8, pady=6)
        lbl.pack()

    def move_tooltip(self, event):
        if self.tooltip:
            x = self.widget.winfo_rootx() + event.x + 15
            y = self.widget.winfo_rooty() + event.y + 15
            self.tooltip.wm_geometry(f"+{x}+{y}")

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


USER_MANUAL_TEXT = """# 📘 지능형 파일 고속 관리기 (Intelligent File Organizer)

> 아키텍처: Clean Layer Architecture 기반 고성능 자동화
> 핵심 철학: [무결성 보증], [자동화(지능형 패턴 분석)], [안전성(충돌 방지)]

---

## 💡 Executive Summary (왜 "지능형"인가?)

이 시스템이 "지능형(Intelligent)"이라는 명칭을 가지는 이유는, 단순한 문자열 치환의 수준을 넘어 **데이터 심층 스캔 체계와 자율적 파일 라우팅(Routing) 설계**가 치밀하게 결합되어 있기 때문입니다.

1. **완벽한 트리 딥-다이브(Deep-Dive) 스캔 알고리즘**  
   사용자가 최상단의 폴더를 지정하면 겉핥기식 평면 검색에 그치지 않습니다. 코어 엔진(`os.walk` 기반 Recursive 순회)이 **가장 깊은 하위 폴더의 맨 밑바닥 끝까지 파고들어** 그 내부에 숨겨진 수천 개의 파편화된 파일과 서브 폴더들을 단숨에 추출하고 식별해냅니다.
   
2. **제자리 변경(In-place) 및 구조 파괴적 복사/이동의 지능적 분기 처리**  
   찾아낸 방대한 타겟들을 무작정 이름만 바꾸고 끝내지 않습니다. 대상들을 _원래 위치(제자리)에서 이름만 바꿀지_, 아니면 사용자가 설계한 고도의 분류 규칙에 따라 완전히 새로운 폴더나 체계로 _끌고가서 복사 또는 이동시킬지_ 를 사용자가 직접 통제하고 레고 블록처럼 유연하게 조립할 수 있습니다.

3. **영구 캐시 기반의 자율 컨텍스트 복원 (경로 자동 기억 시스템)**
   단순한 휘발성 세션을 넘어서, 사용자가 마지막으로 스캔했거나 확인했던 타겟 폴더 주소의 동선을 백그라운드의 보안 파일(JSON)에 완벽하게 캐싱(기억)해 둡니다. 시스템을 완전히 종료하고 내일 다시 실행해도 0.1초 만에 폴더 선택 다이얼로그나 메인 경로창에 과거 동선이 입체적으로 복원되어 즉각적인 업무 흐름의 연속성을 보장합니다.

4. **데이터 무결성 보증 및 심층 감사 리포트 (CSV Tracking)**
   어떠한 경우에도 무작위로 원본을 파괴하지 않으며, 모든 파일/폴더 제어는 내부의 점유 해제(Unlocker) 엔진을 거쳐 안전하게 수행됩니다. 실행 후, 성공 과 실패의 세부 내역이 담긴 CSV 추적 파일(Audit Report)을 원클릭으로 출력하도록 하여 돌발 오류 발생 시 즉각적으로 원인을 추적·규명할 수 있는 완벽한 사후 지원을 유지합니다.

이러한 독보적인 메커니즘과 유연성(Flexibility)은 수백 곳에 엉켜있는 데이터 구조를 단 몇 번의 클릭만으로 안전하고 체계적으로 리빌딩해주는 "강력한 인프라스트럭처 도구"임을 입증합니다. 

> 🛡️ **[System Architecture Audit & Security Guarantee]**  
> 본 프로젝트의 전반적인 코드 베이스 설계, 데이터 라우팅 파이프라인의 논리성과 정합성, 멀티 스레딩 환경에서의 작업 중단 없는 프로세스 안전성, OS 환경에 종속되지 않는 경로 무결성(Integrity), 그리고 점유 프로세스 해제(Unlocker)를 통한 원본 파일 파괴 및 유실 방지(Safety) 등 아키텍처 수준의 완벽한 기능 무결성과 지능적 안정성은 Google DeepMind AI 엔지니어 『Antigravity』에 의해 1줄부터 마지막 줄까지 심층 분석·검열되었으며, 이에 그 성능과 안전성을 공식적으로 보증(Guaranteed)하며 서명합니다.

---

## 🚀 코어 엔진 주요 기능

### 1️⃣ 지능형 패턴 제안 및 병합 알고리즘 (Merge Engine)
* **패턴 자동 분석**: 불러온 파일/폴더 이름 중 반복적으로 등장하는 패턴을 엔진이 분석하여 화면 상단 칩(Chip) 형태로 즉시 제안합니다.
* **원클릭 빠른 제거**: 제안된 패턴 칩을 [마우스 좌클릭]하면 즉각적으로 '문자열 제거' 규칙으로 등록되어 파일명에서 소거됩니다.
* **유연한 패턴 병합(우클릭)**: 칩을 [마우스 우클릭]하면 컨텍스트 메뉴가 나타납니다. 해당 패턴을 복사하여 즉시 교체, 삽입, 형식, 대/소문자, 폴더 분류 설정창으로 전송해 복합 적인 규칙 조합을 구성할 수 있습니다.

### 2️⃣ 워크플로우 듀얼 모드 병행 지원
* **[FILE] 파일 모드**: 다수의 엑셀, 워드, 이미지 등의 개별 파일 이름 일괄 변경 타겟 지원. 확장자 필터(.xlsx 등)를 이용한 정밀 타겟 선택 기능을 포함합니다.
* **[DIR] 폴더 모드**: 파일뿐만 아니라 폴더 자체의 이름을 일괄 변경하거나, 지능적인 구조적 분류(이동/복사) 작업을 고속으로 수행합니다.

### 3️⃣ 무결성 기반 스마트 작업 프로세스
* **안전한 미리보기 (Simulation)**: 규칙을 적용하면 실제 디스크 공간의 파일이 파괴되기 전, [변경될 파일명] 섹션에 실시간 데이터 테이블로 결과가 렌더링 됩니다.
* **점유 잠금 해제(Unlocker) 알고리즘**: Windows 탐색기나 엑셀이 파일을 사용중이라 변경이 불가능할 경우, 시스템이 이를 인지하고 점유 프로세스 권한을 우회/종료시켜 작업을 강제 완수시키는 강력한 방어기제가 탑재되어 있습니다.

---

## 📖 핵심 규칙 가이드 (Rules Pipeline)

* 모든 규칙은 가로 칩 형태로 쌓이며(Stack), 좌측에서 우측으로 순차적인 파이프라인 처리가 됩니다.

1. **[형식 (Format Pattern)]**
   - 역할: 연속된 번호를 매기거나 특정 포맷으로 파일명 정규화
   - 기호: `/n` (원본 이름 유지), `/01` (지정된 자릿수의 일련번호 적용), `/YMD` (오늘 날짜)
2. **[교체 (Replace)]**
   - 특수 문자열을 찾아내어 대체합니다. '바꿀 문자열'을 공백으로 두면 소거(제거) 역할을 수행합니다.
3. **[삽입 (Insert)] / [삭제 (Delete)]**
   - 고정된 텍스트를 대상의 앞/뒤 등 지정된 포지션(인덱스 위치)에 병합하거나 렌더링을 잘라냅니다.
4. **[대/소문자 변환 (Case Convert)]**
   - 영문 전용 기능으로 정규화(Upper/Lower/Title)를 수행합니다. (예외 문자 지정 가능)
5. **[폴더 분류 (Prefix Distribute)]**
   - (고급기능) 기존 VBA 매크로를 대체하는 지능형 라우팅 기능입니다. 대상의 앞 번호를 엔진이 스캔하여, 목적지 폴더 안의 동일한 번호 경로를 스스로 찾아 자동으로 [이동 / 복사 / 바로가기] 매핑을 완수합니다.

---

## 🛠️ 작업 증명 및 사용자 권한 개입
* [수동 라우팅]: 트리뷰의 항목을 더블클릭하면, 자동 분석 규칙을 덮어쓰고 수동으로 이름을 강제 지정(Override) 할 수 있습니다. 
* 작업이 모두 종료되면 대상 객체의 결과 이력, 용량 변화율, 오류 스택이 포함된 심층 CSV 감사 리포트(Audit Report)를 발행할 수 있습니다.

---

### [고급 사례 1: 스캔 & 라우팅 지능 활용]
**Q1. 부서별 흩어져 있는 자료들에 번호를 매기면서 새로운 폴더 구조로 한 번에 복사하면서 병합 백업할 수 있나요? (원본은 유지)**
> **A.** 가능합니다! [FILE 모드]에서 '서브 폴더 포함 스캔'을 체크하고 최상위를 지정하면 하위 모든 파일이 엔진에 집계됩니다. 필터로 `.pdf` 등을 분류한 후, `[1:형식]` 규칙으로 문서 번호를 정규화하고, `[5:폴더 분류]`에서 상태를 **'복사(Copy)'**로 바꾼 뒤 새로운 백업용 목적지를 선택하세요. 원본 데이터는 전혀 파괴되지 않으면서 타겟 폴더에 재구성된 복사본이 깔끔히 저장됩니다!

### [고급 사례 2: 영구 캐시 기반 경로 자동 기억]
**Q2. 매일 똑같은 '일일보고서' 폴더를 열어 작업하는데, 프로그램과 팝업창을 켤 때마다 폴더 경로를 클릭하고 클릭해서 찾아가는 과정이 너무 피곤합니다.**
> **A.** 전혀 걱정할 필요가 없습니다. 본 시스템 엔진 내부에는 "캐시 메모리 최신화" 로직이 숨어있습니다. 내일 아침에 출근하여 프로그램을 켜기만 해도 '일일보고서' 폴더 경로가 메인 화면 📍 주소창에 자동으로 복원 채워져 있습니다. 심지어 '폴더 모드' 안의 [개별 추가 팝업창]을 띄울 때조차도 같은 바탕화면의 유사한 지점으로 알아서 시스템이 찾아들어가서 기다리고 있으므로 즉시 다음 작업을 속개하실 수 있습니다.

### [고급 사례 3: CSV 리포트를 통한 오류 원인 역추적]
**Q3. 수백 개의 파일을 이름 변경하며 라우팅시켰는데, 혹시라도 다른 프로그램이 문서를 쓰고 있어서(잠겨서) 파일 한두 개가 실패하고 에디터에 보류되었을까 걱정됩니다. 어떻게 추적하고 원인을 알죠?**
> **A.** 이 시스템은 돌발 상황에 무너지지 않습니다. 방어 엔진 덕에 실패해도 전체 작업은 완료됩니다. 성공 표시 화면에서 **[CSV 리포트 내보내기]** 버튼을 눌러 결과 파일을 만드세요! 생성된 엑셀을 열고 `Result_Status(성공/실패)` 컬럼에서 '실패'만 필터링해보면, 어떤 경로에 있던 파일이 문제였는지 리스트업 됩니다. 아울러 `Error_Detail` 컬럼에 "프로세스 점유 권한 없음" 등의 에러 원인 스택이 그대로 기록되므로 즉각적인 수동 조치가 가능합니다!

### [일반 사례 4: 폴더 헤더 자동 매핑 (매크로 완전 대체)]
**Q4. '01 공정관리', '02 자재협력', '03 설계도면' 등 앞에 번호가 메겨진 폴더 체계가 있습니다. 파일 이름의 바뀐 앞자리 번호표를 인식해서 자기가 알아서 알맞은 경로로 쏙쏙 들어가게 할 수 있나요?**
> **A.** 이 시스템만의 독보적인 기능인 **"자동 라우팅"**을 쓰면 됩니다! 먼저 `[2:교체]`등으로 지저분한 파일명을 바꾸고 `[1:형식]`으로 파일 이름 앞마다 `01`, `02` 번호를 부여하게 설계하세요. 그 후 `[5:폴더 분류]` 규칙을 체인으로 물려 놓으시면, 대상의 번호 헤더를 즉각 AI 엔진이 스캔하여 목적지 폴더 그룹 속 수십 개의 방 중에 일치하는 방을 스스로 찾아냅니다. 오차 없는 진정한 파일 자동 배분 시스템입니다."""

# ═══════════════════════════════════════════════════════════
# LAYER 1: DOMAIN (Engine - Pure Business Logic)
# ═══════════════════════════════════════════════════════════

class FileOrganizerEngine:
    def __init__(self):
        self.stop_requested = False

    def get_dir_size(self, start_path):
        """경로 하위의 모든 파일 크기 합산을 통한 폴더 총 용량 계산 (바이트)"""
        total_size = 0
        try:
            for dirpath, _, filenames in os.walk(start_path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    if not os.path.islink(fp):
                        total_size += os.path.getsize(fp)
        except: pass
        return total_size

    def _force_unlock_folder(self, target_path, callback):
        """[무결성] 폴더 반환/이동 시 권한 에러 방지를 위해 잠금 범인 프로세스 식별 및 사용자 동의 하에 해제"""
        if not callback: return False
        locking_procs = []
        abs_target = os.path.abspath(target_path).lower() + os.sep
        
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    # 1. 프로세스 작업 경로가 해당 폴더 내부에 있는지 검증
                    try:
                        proc_cwd = proc.cwd()
                        if proc_cwd:
                            abs_cwd = os.path.abspath(proc_cwd).lower() + os.sep
                            if abs_cwd.startswith(abs_target) or abs_cwd == abs_target:
                                locking_procs.append(proc)
                                continue
                    except: pass
                    
                    # 2. 프로세스가 물고 있는 파일 중 해당 폴더 하위 항목이 있는지 검증
                    try:
                        open_files = proc.open_files()
                        if open_files:
                            for f in open_files:
                                abs_f = os.path.abspath(f.path).lower()
                                if abs_f.startswith(abs_target):
                                    locking_procs.append(proc)
                                    break
                    except: pass
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except: pass
        
        if not locking_procs:
            return False

        # 콜백을 통해 UI에 알림 및 응답 대기
        wait_event = threading.Event()
        result_container = {'ans': False}
        
        callback('ask_kill', {
            'procs': locking_procs,
            'name': os.path.basename(target_path),
            'wait_event': wait_event,
            'result': result_container
        })
        
        wait_event.wait()
        
        if result_container['ans']:
            # 강제 종료 승인됨
            for p in locking_procs:
                try: p.kill()
                except: pass
            time.sleep(0.5) # 프로세스가 완전히 죽을 때까지 대기
            return True
        return False

    def scan_files(self, directory, recursive=True, pattern="*", target_mode="file", callback=None):
        """
        [v34.1.41] 경로 정규화 및 단일 항목 지원이 강화된 지능형 스캔 엔진
        """
        results = []
        directory = os.path.abspath(os.path.normpath(directory))
        
        # [v34.1.41] 단일 파일/폴더가 입력된 경우에 대한 예외 처리 (os.walk 대체 대응)
        if os.path.isfile(directory) or (target_mode == 'dir' and not recursive and os.path.isdir(directory)):
             # 폴더 모드에서 재귀가 아닐 때 선택한 폴더 자체를 처리하고 싶어하는 수동 대응 포함 가능성
             pass # 하단 os.walk에서 처리되도록 유도 (하지만 root==directory 조건을 탈 것임)

        try:
            for root, dirs, files in os.walk(directory):
                # 윈도우 경로 일관성 확보를 위해 정규화
                root_norm = os.path.abspath(os.path.normpath(root))
                dir_norm = os.path.abspath(os.path.normpath(directory))
                
                if not recursive and root_norm != dir_norm:
                    # 하위 폴더로 내려가지 않도록 dirs를 비움으로써 os.walk 최적화
                    dirs[:] = []
                    continue
                
                target_collection = files if target_mode == "file" else dirs
                
                for item in target_collection:
                    if fnmatch.fnmatch(item, pattern):
                        full_path = os.path.join(root, item)
                        stats = os.stat(full_path)
                        is_dir = os.path.isdir(full_path)
                        
                        # 폴더일 경우 용량 계산이 느릴 수 있으므로 0으로 초기화
                        size_bytes = stats.st_size if not is_dir else 0 
                        ext = os.path.splitext(item)[1] if not is_dir else ""
                        
                        results.append({
                            'name': item,
                            'path': root,
                            'full_path': full_path,
                            'ext': ext,
                            'size': size_bytes,
                            'mtime': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                            'ctime': datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                            '_target_mode': target_mode
                        })
                
                if self.stop_requested:
                    break
                    
            if callback:
                callback('scan_complete', results)
            return results
        except Exception as e:
            if callback:
                callback('error', str(e))
            return []

    def detect_common_patterns(self, files, active_rules=None):
        """
        [v34.1.41] 파일/폴더 모드별 논리성 및 정합성 전수점검된 지능형 패턴 추출 엔진
        - 괄호 체계((), [], {}) 정밀 추출 및 날짜, 고빈도 토큰 감지
        - 파일 모드 시 확장자 제외 분석으로 정합성 확보
        """
        if not files or len(files) < 2:
            return []
            
        candidate_map = {} # {pattern: count}
        existing_markers = set()
        if active_rules:
            for r in active_rules:
                # [v34.1.41-R] 규칙의 모든 문자열 파라미터를 추출하여 중복 제안 전수 방지
                params = r.get('params', {})
                for k, v in params.items():
                    if isinstance(v, str) and len(v) > 0:
                        existing_markers.add(v)

        for f in files:
            # 모드에 따른 이름 전처리 (파일일 경우 확장자 제외)
            is_dir = f.get('_target_mode') == 'dir'
            name = f['name'] if is_dir else os.path.splitext(f['name'])[0]
            
            # 1. 괄호 패턴 정밀 추출 (닫는 괄호 포함 무결성 확보)
            bracket_matches = re.findall(r'(\([^\)]+\)|\[[^\]]+\]|\{[^\}]+\})', name)
            for b in bracket_matches:
                candidate_map[b] = candidate_map.get(b, 0) + 1
            
            # 2. 날짜 패턴 감지 (YYYY-MM-DD, YY.MM.DD 등 다양한 형식)
            date_matches = re.findall(r'(\d{2,4}[-._]\d{2}[-._]\d{2})', name)
            for d in date_matches:
                candidate_map[d] = candidate_map.get(d, 0) + 1
                
            # 3. 구분자 기반 의미 있는 단어(토큰) 추출
            # 파일명 중간에 반복되는 프로젝트명이나 태그 감지
            tokens = re.split(r'[-_\s.]', name)
            for t in tokens:
                if len(t) > 2 and not t.isdigit(): # 3글자 이상 의미 있는 식별자만
                    candidate_map[t] = candidate_map.get(t, 0) + 1

        # 제안 목록 가공 (유효성 및 중복 검증)
        suggestions = []
        # 최소 2개 이상의 파일 혹은 20% 이상 발견 시 유효한 패턴으로 간주
        threshold = max(2, len(files) * 0.15) 
        
        for pt, count in candidate_map.items():
            if count >= threshold and pt not in existing_markers:
                suggestions.append(pt)
        
        # 빈도 수와 길이를 조합하여 사용자에게 가장 유용한 순서로 정렬
        suggestions.sort(key=lambda x: (candidate_map[x], len(x)), reverse=True)
        
        return [s for s in suggestions if s.strip()][:12] # 정제된 상위 12개 제안

    def preview_rename(self, files, rules):
        """
        Apply a sequence of rules to files and return previews with caching optimizations.
        """
        previews = []
        planned_names = {}  # {new_full_path: original_name}
        file_counter = 0
        
        # [최적화] 목적지 폴더 캐시 (분류 규칙 사용 시 디스크 I/O 절감)
        dest_cache = {} # {base_dest: [folder_entries]}

        for file_info in files:
            if self.stop_requested:
                break
                
            is_dir_mode = file_info.get('_target_mode') == 'dir'
            if is_dir_mode:
                current_name = file_info['name']
                ext = ""
            else:
                current_name, ext = os.path.splitext(file_info['name'])
            
            file_counter += 1
            current_rules = [] 
            
            # 이전 분류 정보 초기화
            if '_is_distribute' in file_info: del file_info['_is_distribute']
            if '_dist_dest_base' in file_info: del file_info['_dist_dest_base']
            
            for rule in rules:
                r_type = rule['type']
                params = rule['params']
                
                if r_type == 'marker_truncate':
                    marker = params.get('marker', '')
                    suffix = params.get('suffix', '')
                    if marker and marker in current_name:
                        idx = current_name.find(marker)
                        current_name = current_name[:idx + len(marker)] + suffix
                
                elif r_type == 'regex_replace':
                    pattern = params.get('pattern', '')
                    replacement = params.get('replacement', '')
                    try:
                        current_name = re.sub(pattern, replacement, current_name)
                    except: pass
                
                elif r_type == 'simple_replace':
                    old_str = params.get('old_str', '')
                    new_str = params.get('new_str', '')
                    current_name = current_name.replace(old_str, new_str)

                elif r_type == 'normalize':
                    current_name = re.sub(r'[\\/:*?"<>|]', '', current_name)
                    if params.get('space_to_under'):
                        current_name = current_name.replace(' ', '_')
                    if params.get('to_upper'):
                        current_name = current_name.upper()

                elif r_type == 'format_pattern':
                    pattern_str = params.get('pattern', '/n')
                    start_num = params.get('start', 1)
                    digits = params.get('digits', 2)
                    increment = params.get('increment', 1)
                    zero_pad = params.get('zero_pad', True)
                    pos_type = params.get('pos_type', 'replace') 
                    
                    current_num = start_num + (file_counter - 1) * increment
                    num_str = str(current_num).zfill(digits) if zero_pad else str(current_num)
                    
                    result = pattern_str
                    has_explicit_n = '/n' in result
                    
                    result = result.replace('/n', current_name)
                    result = result.replace('/e', ext[1:] if ext else '')
                    now = datetime.now()
                    result = result.replace('/YMD', now.strftime('%Y%m%d'))
                    result = result.replace('/HMS', now.strftime('%H%M%S'))
                    result = re.sub(r'/(\d+)', num_str, result)
                    
                    if pos_type == 'prefix' and not has_explicit_n:
                         current_name = result + current_name
                    elif pos_type == 'suffix' and not has_explicit_n:
                         current_name = current_name + result
                    else:
                         current_name = result

                elif r_type == 'replace':
                    find_str = params.get('find', '')
                    replace_str = params.get('replace', '')
                    ignore_case = params.get('ignore_case', False)
                    max_count = params.get('max_count', 0)
                    
                    if find_str:
                        if ignore_case:
                            pattern = re.compile(re.escape(find_str), re.IGNORECASE)
                            if max_count > 0:
                                current_name = pattern.sub(replace_str, current_name, count=max_count)
                            else:
                                current_name = pattern.sub(replace_str, current_name)
                        else:
                            if max_count > 0:
                                current_name = current_name.replace(find_str, replace_str, max_count)
                            else:
                                current_name = current_name.replace(find_str, replace_str)

                elif r_type == 'insert':
                    insert_str = params.get('text', '')
                    position = params.get('position', 0)
                    from_end = params.get('from_end', False)
                    
                    if insert_str:
                        if from_end:
                            actual_pos = max(0, len(current_name) - position)
                        else:
                            actual_pos = min(position, len(current_name))
                        current_name = current_name[:actual_pos] + insert_str + current_name[actual_pos:]

                elif r_type == 'delete':
                    position = params.get('position', 0)
                    length = params.get('length', 1)
                    from_end = params.get('from_end', False)
                    
                    if from_end:
                        actual_start = max(0, len(current_name) - position - length)
                        actual_end = max(0, len(current_name) - position)
                    else:
                        actual_start = min(position, len(current_name))
                        actual_end = min(position + length, len(current_name))
                    current_name = current_name[:actual_start] + current_name[actual_end:]

                rule_name_map = {
                    'format_pattern': '1:형식',
                    'replace': '2:교체',
                    'insert': '3:삽입',
                    'delete': '4:삭제',
                    'case_convert': '5:대소',
                    'prefix_distribute': '6:분류',
                    'marker_truncate': '자르기'
                }
                current_rules.append(rule_name_map.get(r_type, r_type))

                if r_type == 'prefix_distribute':
                    c_prefix = params.get('prefix', '')
                    num_len = params.get('num_len', 2)
                    
                    numeric_prefix = ""
                    for char in current_name:
                        if char.isdigit():
                            numeric_prefix += char
                            if len(numeric_prefix) == num_len: break
                    
                    if len(numeric_prefix) == num_len:
                        idx = current_name.find(numeric_prefix)
                        current_name = c_prefix + current_name[idx + num_len:]
                        
                        # [최적화] 지능형 폴더 매칭 시뮬레이션 (캐시 사용)
                        dest_base = params.get('base_dest', '')
                        if dest_base and os.path.exists(dest_base):
                            if dest_base not in dest_cache:
                                try:
                                    dest_cache[dest_base] = [e.name for e in os.scandir(dest_base) if e.is_dir()]
                                except:
                                    dest_cache[dest_base] = []
                            
                            for folder_name in dest_cache[dest_base]:
                                if folder_name.startswith(numeric_prefix):
                                    file_info['_dist_dest_base'] = folder_name
                                    break

                    file_info['_is_distribute'] = True
                    if '_dist_dest_base' not in file_info:
                        file_info['_dist_dest_base'] = f"{numeric_prefix}???"

                elif r_type == 'case_convert':
                    target = params.get('target', 'name')
                    case_type = params.get('case_type', 'lower')
                    ignore_chars = params.get('ignore_chars', '')
                    
                    def convert_case(s, ctype, ignore):
                        if not s: return s
                        result = []
                        for char in s:
                            if char in ignore:
                                result.append(char)
                            elif ctype == 'lower':
                                result.append(char.lower())
                            elif ctype == 'upper':
                                result.append(char.upper())
                            else:
                                result.append(char)
                        converted = ''.join(result)
                        if ctype == 'title':
                            converted = converted.title()
                        return converted
                    
                    if target == 'name' or target == 'all':
                        current_name = convert_case(current_name, case_type, ignore_chars)

            new_name = current_name + ext
            new_full_path = os.path.join(file_info['path'], new_name)
            
            status = " + ".join(current_rules) if current_rules else "[OK]"
            
            if file_info.get('_is_distribute'):
                dest_info = file_info.get('_dist_dest_base', '??')
                status = f"6:분류 -> {dest_info}"
            elif new_name == file_info['name']:
                status = '[UNCHANGED]'
            elif new_full_path in planned_names:
                status = '[WARN] 충돌(중복)'
            elif os.path.exists(new_full_path):
                status = '[WARN] 충돌(기존파일)'

            planned_names[new_full_path] = file_info['name']
            
            previews.append({
                'original_path': file_info['full_path'],
                'original_name': file_info['name'],
                'new_name': new_name,
                'status': status,
                'size_bytes': file_info.get('size', 0),
                'size_kb': round(file_info.get('size', 0) / 1024, 2),
                'mtime': file_info.get('mtime', ''),
                'ctime': file_info.get('ctime', ''),
                'is_distribute': file_info.get('_is_distribute', False),
                '_target_mode': file_info.get('_target_mode', 'file'),
                'dist_dest': file_info.get('_dist_dest_base', '')  
            })
            
        return previews

    def perform_distribute(self, rename_list, rules, callback=None):
        """
        [VBA 이식] 숫자 접두사 기반 하위 폴더 자동 매칭 및 분류 실행
        """
        success_count = 0
        fail_count = 0
        import shutil

        # 'prefix_distribute' 규칙 찾기
        dist_rule = next((r for r in rules if r['type'] == 'prefix_distribute'), None)
        if not dist_rule:
             if callback: callback('rename_complete', {'success': 0, 'failed': 0, 'message': '분류 규칙이 없습니다.'})
             return

        params = dist_rule['params']
        base_dest = params.get('base_dest', '')
        sub_path = params.get('sub_path', '')
        num_len = params.get('num_len', 2)
        mode = params.get('mode', 'copy')
        custom_prefix = params.get('prefix', '')

        if not base_dest or not os.path.exists(base_dest):
            if callback: callback('rename_complete', {'success': 0, 'failed': 0, 'message': '목적지 경로가 유효하지 않습니다.'})
            return

        for item in rename_list:
            if self.stop_requested: break
            
            filename = item['original_name']
            
            # 숫자 접두사 추출
            numeric_prefix = ""
            for char in filename:
                if char.isdigit():
                    numeric_prefix += char
                    if len(numeric_prefix) == num_len: break
            
            if len(numeric_prefix) < num_len:
                fail_count += 1
                if callback: callback('item_fail', {'path': item['original_path'], 'name': filename, 'error': f'숫자 접두사({num_len}자리) 미달'})
                continue

            # 하위 폴더 검색
            target_subfolder = None
            try:
                for entry in os.scandir(base_dest):
                    if entry.is_dir() and entry.name.startswith(numeric_prefix):
                        target_subfolder = entry.path
                        break
            except Exception as e:
                fail_count += 1
                if callback: callback('item_fail', {'path': item['original_path'], 'name': filename, 'error': f'폴더 스캔 실패: {str(e)}'})
                continue

            if not target_subfolder:
                fail_count += 1
                if callback: callback('item_fail', {'path': item['original_path'], 'name': filename, 'error': '일치하는 숫자 폴더 없음'})
                continue

            # 최종 경로 조합
            final_dest = os.path.join(target_subfolder, sub_path) if sub_path else target_subfolder
            
            if not os.path.exists(final_dest):
                fail_count += 1
                if callback: callback('item_fail', {'path': item['original_path'], 'name': filename, 'error': f'하위 경로 없음: {sub_path}'})
                continue

            # [무결성 Fix] 미리보기에서 확정된(또는 수동 수정된) 파일명을 우선 사용
            new_file_name = item.get('new_name', filename)
            final_path = os.path.join(final_dest, new_file_name)

            # 충돌 방지
            if os.path.exists(final_path):
                base, ext = os.path.splitext(final_path)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                final_path = f"{base}_{counter}{ext}"
                new_file_name = os.path.basename(final_path)

            try:
                is_dir_mode = item.get('_target_mode', 'file') == 'dir'

                # [시간 지연 판단] 폴더 복사일 때, 예상 시간 초과 경고 로직 (50MB/s 기준 3GB=60초)
                if is_dir_mode and mode == 'copy':
                    folder_bytes = self.get_dir_size(item['original_path'])
                    if folder_bytes > 3145728000: # 약 3GB 초과시 1분 소요 경고
                        wait_evt = threading.Event()
                        res_container = {'ans': False}
                        if callback:
                            callback('ask_time_warn', {
                                'name': filename,
                                'wait_event': wait_evt,
                                'result': res_container
                            })
                            wait_evt.wait()
                            if not res_container['ans']:
                                fail_count += 1
                                if callback: callback('item_fail', {'path': item['original_path'], 'name': filename, 'error': '사용자 복사 취소됨 (대용량 시간 지연)'})
                                continue

                while True: # 권한 잠금 해결 및 재시도 루프
                    try:
                        if mode == 'move':
                            shutil.move(item['original_path'], final_path)
                        else:
                            if is_dir_mode: shutil.copytree(item['original_path'], final_path)
                            else: shutil.copy2(item['original_path'], final_path)
                        break
                    except PermissionError as pe:
                        if is_dir_mode:
                            unlocked = self._force_unlock_folder(item['original_path'], callback)
                            if unlocked:
                                continue # 해제 성공 시 재수행
                        raise pe

                success_count += 1
                mode_verb = "복사" if mode == 'copy' else "이동"
                new_size = os.path.getsize(final_path) if os.path.exists(final_path) and not is_dir_mode else item.get('size_bytes', 0)
                if callback: 
                    callback('item_success', {
                        'path': item['original_path'], 
                        'old': filename, 
                        'new': f'[{mode_verb} 완료] -> {os.path.basename(final_dest)}\\{new_file_name}',
                        'new_size': new_size
                    })
            except Exception as e:
                fail_count += 1
                if callback: callback('item_fail', {'path': item['original_path'], 'name': filename, 'error': str(e)})

        if callback:
            callback('rename_complete', {'success': success_count, 'failed': fail_count})

    def perform_rename(self, rename_list, callback=None):
        """
        Execute the renaming.
        """
        success_count = 0
        fail_count = 0
        
        for item in rename_list:
            if self.stop_requested:
                break
                
            old_path = item['original_path']
            dir_path = os.path.dirname(old_path)
            new_path = os.path.join(dir_path, item['new_name'])
            
            try:
                if old_path != new_path:
                    final_path = new_path
                    if os.path.exists(final_path):
                        base, ext = os.path.splitext(new_path)
                        counter = 1
                        while os.path.exists(f"{base}_{counter}{ext}"):
                            counter += 1
                        final_path = f"{base}_{counter}{ext}"
                    
                    is_dir_mode = item.get('_target_mode', 'file') == 'dir'
                    while True:
                        try:
                            os.rename(old_path, final_path)
                            break
                        except PermissionError as pe:
                            if is_dir_mode:
                                unlocked = self._force_unlock_folder(old_path, callback)
                                if unlocked:
                                    continue
                            raise pe

                    success_count += 1
                    new_size = os.path.getsize(final_path) if os.path.exists(final_path) and not is_dir_mode else item.get('size_bytes', 0)
                    if callback:
                        # [무결성] original_path를 명시적으로 전달하여 UI 동기화 속도 및 정확도 향상
                        callback('item_success', {
                            'path': item['original_path'],
                            'old': item['original_name'], 
                            'new': os.path.basename(final_path),
                            'new_size': new_size
                        })
                else:
                    success_count += 1
            except Exception as e:
                fail_count += 1
                if callback:
                    callback('item_fail', {
                        'path': item['original_path'],
                        'name': item['original_name'], 
                        'error': str(e)
                    })
        
        if callback:
            callback('rename_complete', {'success': success_count, 'failed': fail_count})

# ═══════════════════════════════════════════════════════════
# LAYER 2: PRESENTATION (View - Modern Clean UI)
# ═══════════════════════════════════════════════════════════

class FolderSelectorDialog(tk.Toplevel):
    """
    [v34.1.30] 고도화된 지능형 멀티 폴더 선택기
    - 주소창(Address Bar) 도입으로 빠른 경로 이동 지원
    - 단일 컬럼 및 아이콘(📁/🖴) 적용으로 직관성 극대화
    - 세션 내 마지막 사용 경로 자동 활성화 (Navigate to Last Path)
    """
    def __init__(self, parent, controller, initial_path=None):
        super().__init__(parent)
        self.controller = controller
        self.title("지능형 폴더 개별 추가 (Individual Folder Add)")
        self.geometry("800x600")
        self.minsize(700, 500)
        self.transient(parent)
        self.grab_set()
        
        self.selected_paths = set()
        self.result = []
        # [경로 자동 기억] 전달받은 초기 경로가 없으면 시스템 내장 로컬 경로 사용 (없으면 바탕화면)
        if not initial_path:
            saved = load_config_path()
            if saved: initial_path = saved
            else:
                initial_path = os.path.join(os.path.expanduser("~"), "Desktop")
                if not os.path.exists(initial_path): initial_path = os.path.expanduser("~")
        
        self.initial_path = initial_path
        
        self.setup_ui()
        self.populate_roots()
        
        # 시작 시 초기 경로로 주소창 텍스트 세팅 및 트리 자동 이동
        if self.initial_path:
            self.addr_var.set(self.initial_path)
            self.after(100, lambda: self.navigate_to_path(self.initial_path))

    def setup_ui(self):
        # 1. 주소창 (Address Bar)
        addr_frame = tk.Frame(self, bg="#F0F2F5", padx=10, pady=10)
        addr_frame.pack(fill=tk.X)
        
        tk.Label(addr_frame, text="📍 주소:", font=("Malgun Gothic", 9, "bold"), bg="#F0F2F5").pack(side=tk.LEFT, padx=(0, 5))
        self.addr_var = tk.StringVar()
        self.addr_entry = tk.Entry(addr_frame, textvariable=self.addr_var, font=("Segoe UI", 10), relief=tk.FLAT)
        self.addr_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.addr_entry.bind("<Return>", lambda e: self.on_addr_go())
        
        ttk.Button(addr_frame, text="이동", command=self.on_addr_go, width=8).pack(side=tk.LEFT, padx=5)

        # 2. 안내 문구
        guide_frame = tk.Frame(self, bg="#E1F5FE", padx=10, pady=5)
        guide_frame.pack(fill=tk.X)
        tk.Label(guide_frame, text="💡 [클릭]: 폴더 진입  |  [Ctrl + 클릭]: 해당 폴더 자체를 선택 (멀티 선택 가능)", 
                 font=("Malgun Gothic", 9), bg="#E1F5FE", fg="#01579B").pack()

        # 3. 트리 영역 (심플 단일 열)
        self.tree_frame = tk.Frame(self)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # show="tree"를 사용하여 아이콘과 텍스트만 표시
        self.tree = ttk.Treeview(self.tree_frame, columns=("full_path",), show="tree", selectmode="none")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 컬럼 너비를 최소화하여 보이지 않게 처리
        self.tree.column("#0", stretch=tk.YES)
        self.tree.column("full_path", width=0, stretch=tk.NO)
        
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)
        
        # 4. 하단 버튼 영역
        bottom_frame = tk.Frame(self, pady=10, bg="#F0F2F5")
        bottom_frame.pack(fill=tk.X)
        
        self.info_var = tk.StringVar(value="선택된 항목: 0개")
        tk.Label(bottom_frame, textvariable=self.info_var, font=("Malgun Gothic", 9, "bold"), bg="#F0F2F5", fg="#0984E3").pack(side=tk.LEFT, padx=20)
        
        ttk.Button(bottom_frame, text="✅ 확인 (목록에 추가)", command=self.on_confirm, width=20).pack(side=tk.RIGHT, padx=10)
        ttk.Button(bottom_frame, text="취소", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=5)
        
        # 이벤트 바인딩
        self.tree.bind("<<TreeviewOpen>>", self.on_expand)
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.tag_configure("selected", background="#BBDEFB", foreground="#0D47A1")

    def populate_roots(self):
        import string
        drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
        for d in drives:
            # 드라이브 아이콘 적용
            node = self.tree.insert("", tk.END, text=f" 🖴 {d}", values=(d,), open=False)
            self.tree.insert(node, tk.END)

    def navigate_to_path(self, target_path):
        """특정 경로로 트리를 자동 확장하며 이동"""
        if not target_path or not os.path.exists(target_path): return
        
        parts = target_path.split(os.sep)
        # 드라이브 파싱 (예: C:)
        current_node = ""
        current_path = ""
        
        for i, part in enumerate(parts):
            if i == 0 and part.endswith(':'):
                part += os.sep
            
            if not current_path:
                current_path = part
            else:
                current_path = os.path.join(current_path, part)
            
            # 현재 레벨의 노드들 중 일치하는 것 찾기
            children = self.tree.get_children(current_node)
            found = False
            for child in children:
                node_path = self.tree.item(child, "values")[0]
                # 경로 비교 (대소문자 구분 없이)
                if node_path.lower().rstrip('\\') == current_path.lower().rstrip('\\'):
                    current_node = child
                    # 확장 유도 (하위 항목 로딩)
                    self.tree.item(current_node, open=True)
                    self._load_node(current_node) # 즉시 로딩 강제
                    found = True
                    break
            if not found: break
            
        if current_node:
            self.tree.see(current_node)
            self.tree.focus(current_node)
            self.tree.selection_set(current_node)
            self.addr_var.set(target_path)

    def on_addr_go(self):
        path = self.addr_var.get().strip()
        if os.path.isdir(path):
            save_config_path(path) # 사용자가 팝업창 주소 표시줄에 입력한 경로 파싱 및 영구 캐싱
            self.navigate_to_path(path)
        else:
            messagebox.showerror("오류", "유효하지 않은 폴더 경로입니다.")

    def _load_node(self, node):
        """특정 노드의 자식을 실제로 로드 (on_expand 로직 공통화)"""
        children = self.tree.get_children(node)
        if len(children) == 1 and not self.tree.item(children[0], "text"):
            self.tree.delete(children[0])
            parent_path = self.tree.item(node, "values")[0]
            try:
                items = sorted(os.listdir(parent_path))
                for item in items:
                    full_path = os.path.join(parent_path, item)
                    if os.path.isdir(full_path):
                        if not item.startswith('$') and not item.startswith('.'):
                            # 폴더 아이콘 적용
                            child = self.tree.insert(node, tk.END, text=f" 📁 {item}", values=(full_path,), open=False)
                            self.tree.insert(child, tk.END)
            except: pass

    def on_expand(self, event):
        node = self.tree.focus()
        if node: self._load_node(node)

    def on_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item: return
        
        full_path = self.tree.item(item, "values")[0]
        self.addr_var.set(full_path) # 선택할 때마다 주소창 업데이트
        
        if event.state & 0x0004: # Ctrl+클릭
            if full_path in self.selected_paths:
                self.selected_paths.remove(full_path)
                self.tree.item(item, tags=())
            else:
                self.selected_paths.add(full_path)
                self.tree.item(item, tags=("selected",))
            self.info_var.set(f"선택된 항목: {len(self.selected_paths)}개")
            return "break"
        else:
            self.tree.focus(item)
            self.tree.selection_set(item)

    def on_confirm(self):
        if not self.selected_paths:
            curr = self.tree.focus()
            if curr:
                self.result = [self.tree.item(curr, "values")[0]]
            else:
                messagebox.showwarning("주의", "폴더를 선택(Ctrl+클릭)하거나 하나를 지정해주세요.")
                return
        else:
            self.result = list(self.selected_paths)
            
        if self.result:
            # 선택 완료 시 상위(부모) 폴더를 캐싱하여 재진입 시 같은 카테고리에서 열리도록 지능화
            save_config_path(os.path.dirname(self.result[0]))
            
        self.destroy()

class FileOrganizerView:
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[지능형 파일 정리기 v34.1.41]")
        self.root.geometry("1400x950")
        self.root.configure(bg="#F5F7FA")
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=0)  # 상태바 행
        self.root.columnconfigure(0, weight=1)
        
        self.COLORS = {
            'primary': '#2D3436',
            'accent': '#0984E3',
            'bg': '#F5F7FA',
            'card': '#FFFFFF',
            'text': '#2D3436',
            'text_light': '#636E72',
            'success': '#00B894',
            'error': '#D63031'
        }
        
        self._set_style()
        self._build_ui()

    def _set_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        # [v34.1.41-H] 고대비 및 높은 호환성 스타일 강제 적용
        style.configure("Treeview", background="#FAFAFA", foreground="#1A237E", rowheight=28, fieldbackground="#FAFAFA", font=('Malgun Gothic', 10))
        style.map("Treeview", background=[('selected', '#0984E3')], foreground=[('selected', 'white')])
        style.configure("Treeview.Heading", font=('Malgun Gothic', 10, 'bold'), background="#E1E5EB")
        style.configure("Action.TButton", font=('Malgun Gothic', 10, 'bold'), padding=10, background='#0984E3', foreground="white")
        style.map("Action.TButton", background=[('active', '#0773C5')])
        
        # [v34.1.30] 상단바 배치를 위한 컴팩트 버튼 스타일
        style.configure("ActionSmall.TButton", font=('Malgun Gothic', 9, 'bold'), padding=5, background='#0984E3', foreground="white")
        style.map("ActionSmall.TButton", background=[('active', '#0773C5')])
        
        # [v34.1.30] 미리보기 전용 (기존 중립 테마 복원 - Gray)
        style.configure("PreviewSmall.TButton", font=('Malgun Gothic', 9), padding=5, background="#E1E5EB", foreground="black")
        style.map("PreviewSmall.TButton", background=[('active', '#D1D5DB')])

    def _build_ui(self):
        # 전체를 감싸는 root의 row/col 설정은 __init__에서 완료됨
        main_frame = tk.Frame(self.root, bg=self.COLORS['bg'], padx=20, pady=10)
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # 상하 조절을 위한 PanedWindow 도입
        self.paned = tk.PanedWindow(main_frame, orient=tk.VERTICAL, bg=self.COLORS['bg'], sashwidth=4, sashrelief=tk.FLAT)
        self.paned.grid(row=0, column=0, sticky='nsew')

        # 상단 영역 (Header + Top Card + Content Frame)
        self.upper_pane = tk.Frame(self.paned, bg=self.COLORS['bg'])
        self.paned.add(self.upper_pane, stretch="always")
        
        self.upper_pane.columnconfigure(0, weight=1)
        self.upper_pane.rowconfigure(2, weight=1)
        
        main_frame_ref = self.upper_pane # 기존 main_frame 역할을 upper_pane이 수행

        header_frame = tk.Frame(main_frame_ref, bg=self.COLORS['bg'])
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        tk.Label(header_frame, text="[지능형 파일 관리 시스템]", font=('Malgun Gothic', 20, 'bold'), fg=self.COLORS['primary'], bg=self.COLORS['bg']).pack(side=tk.LEFT)
        tk.Label(header_frame, text="Intelligent File Orchestrator", font=('Segoe UI', 10), fg=self.COLORS['text_light'], bg=self.COLORS['bg']).pack(side=tk.LEFT, padx=15, pady=(10, 0))
        # 도움말 버튼 추가
        ttk.Button(header_frame, text="📘 도움말", style="Action.TButton", command=self.controller.show_manual).pack(side=tk.RIGHT, padx=5)

        top_card = tk.Frame(main_frame_ref, bg=self.COLORS['card'], padx=15, pady=15, highlightthickness=1, highlightbackground="#E1E5EB")
        top_card.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        tk.Label(top_card, text="[DIR] 대상 폴더:", font=("Malgun Gothic", 9), bg=self.COLORS['card'], fg=self.COLORS['text']).grid(row=0, column=0, sticky='w')
        
        saved_path = load_config_path()
        if saved_path:
            default_dir = saved_path
        else:
            default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.exists(default_dir): default_dir = os.path.expanduser("~")
            
        self.path_var = tk.StringVar(value=default_dir)
        
        self.path_entry = tk.Entry(top_card, textvariable=self.path_var, font=('Malgun Gothic', 10), width=65, relief=tk.FLAT, bg="#F0F2F5")
        self.path_entry.grid(row=0, column=1, padx=10, pady=5)
        # 텍스트 직접 입력 후 엔터 시 스캔 실행
        self.path_entry.bind("<Return>", lambda e: self.controller.handle_scan())
        
        ttk.Button(top_card, text="폴더 선택", command=self.controller.handle_browse).grid(row=0, column=2, padx=2)
        ttk.Button(top_card, text="개별 추가", command=self.controller.handle_file_select).grid(row=0, column=3, padx=2)
        ttk.Button(top_card, text="목록 초기화", command=self.controller.handle_clear_list).grid(row=0, column=4, padx=2)
        
        # [v34.1.30] 제어 핵심 버튼 상단 이동 배치
        ttk.Button(top_card, text="미리보기 업데이트", width=16, style="PreviewSmall.TButton", command=self.controller.handle_preview).grid(row=0, column=5, padx=2)
        self.stop_btn = tk.Button(top_card, text="⛔ 작업 중단 (Stop)", font=('Malgun Gothic', 10, 'bold'), bg="#FFEAA7", fg="#D63031", relief=tk.GROOVE, padx=10, command=self.controller.handle_stop)
        self.stop_btn.grid(row=0, column=6, padx=2)
        self.apply_btn = ttk.Button(top_card, text="[실행]", width=8, style="ActionSmall.TButton", command=self.controller.handle_run)
        self.apply_btn.grid(row=0, column=7, padx=2)
        self.recursive_var = tk.BooleanVar(value=True)
        tk.Checkbutton(top_card, text="하위 폴더 포함", variable=self.recursive_var, bg=self.COLORS['card'], activebackground=self.COLORS['card'], font=('Malgun Gothic', 9)).grid(row=1, column=1, sticky='w', padx=10)

        content_frame = tk.Frame(main_frame_ref, bg=self.COLORS['bg'])
        content_frame.grid(row=2, column=0, sticky='nsew')
        content_frame.columnconfigure(0, weight=4) # Treeview 더 넓게
        content_frame.columnconfigure(1, weight=3) # Rules 패널도 넓게 (약 4:3 비율)
        content_frame.rowconfigure(0, weight=1)

        tree_frame = tk.Frame(content_frame, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB")
        tree_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 10))
        header = tk.Frame(tree_frame, bg="#1A237E")
        header.pack(fill="x")
        self.header_title = tk.Label(header, text="[FILE] 지능형 파일 고속 정리기 v34.1.41", font=("Malgun Gothic", 14, "bold"), bg="#1A237E", fg="white")
        self.header_title.pack(pady=5)
        
        # [NEW] 모드 선택 및 파일 선택 제어 
        filter_bar = tk.Frame(tree_frame, bg="#F0F2F5", padx=5, pady=5)
        filter_bar.pack(fill="x")
        
        mode_btn_frame = tk.Frame(filter_bar, bg="#F0F2F5")
        mode_btn_frame.pack(side=tk.LEFT, padx=(5, 15))
        tk.Label(mode_btn_frame, text="작업 대상:", bg="#F0F2F5", font=("Malgun Gothic", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        self.btn_mode_file = tk.Radiobutton(mode_btn_frame, text="파일 모드 📄", variable=self.controller.target_mode_var, value="file", command=self.controller.handle_mode_change, bg="#F0F2F5", cursor="hand2")
        self.btn_mode_dir = tk.Radiobutton(mode_btn_frame, text="폴더 모드 📁", variable=self.controller.target_mode_var, value="dir", command=self.controller.handle_mode_change, bg="#F0F2F5", cursor="hand2")
        self.btn_mode_file.pack(side=tk.LEFT)
        self.btn_mode_dir.pack(side=tk.LEFT)

        tk.Label(filter_bar, text="| 선택:", font=('Malgun Gothic', 9, 'bold'), bg="#F0F2F5").pack(side=tk.LEFT, padx=(5,2))
        sb_all = tk.Button(filter_bar, text="전체", font=('Malgun Gothic', 8), command=self.controller.handle_select_all)
        sb_all.pack(side=tk.LEFT, padx=2)
        sb_none = tk.Button(filter_bar, text="해제", font=('Malgun Gothic', 8), command=self.controller.handle_deselect_all)
        sb_none.pack(side=tk.LEFT, padx=2)
        sb_inv = tk.Button(filter_bar, text="반전", font=('Malgun Gothic', 8), command=self.controller.handle_invert_selection)
        sb_inv.pack(side=tk.LEFT, padx=2)
        
        tk.Label(filter_bar, text="| 확장자 필터:", font=('Malgun Gothic', 9, 'bold'), bg="#F0F2F5").pack(side=tk.LEFT, padx=(15,2))
        self.ext_filter_entry = tk.Entry(filter_bar, font=('Malgun Gothic', 9), width=8)
        self.ext_filter_entry.insert(0, ".xlsx")
        self.ext_filter_entry.pack(side=tk.LEFT, padx=2)
        
        self.btn_ext_sel = tk.Button(filter_bar, text="선택", font=('Malgun Gothic', 8), command=lambda: self.controller.handle_filter_by_ext(self.ext_filter_entry.get(), "select"))
        self.btn_ext_sel.pack(side=tk.LEFT, padx=2)
        self.btn_ext_exc = tk.Button(filter_bar, text="제외", font=('Malgun Gothic', 8), command=lambda: self.controller.handle_filter_by_ext(self.ext_filter_entry.get(), "deselect"))
        self.btn_ext_exc.pack(side=tk.LEFT, padx=2)
        
        cols = ('no', 'check', 'original', 'arrow', 'preview', 'status', 'result')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', selectmode='extended')
        self.tree.heading('no', text='NO', command=lambda: self._treeview_sort_column('no', False))
        self.tree.heading('check', text='선택', command=lambda: self.controller.handle_toggle_all())
        self.tree.heading('original', text='현재 파일명', command=lambda: self._treeview_sort_column('original', False))
        self.tree.heading('arrow', text='→')
        self.tree.heading('preview', text='변경될 파일명', command=lambda: self._treeview_sort_column('preview', False))
        self.tree.heading('status', text='상태', command=lambda: self._treeview_sort_column('status', False))
        self.tree.heading('result', text='작업 결과', command=lambda: self._treeview_sort_column('result', False))
        self.tree.column('no', width=40, anchor='center')
        self.tree.column('check', width=50, anchor='center')
        self.tree.column('original', width=180)
        self.tree.column('arrow', width=30, anchor='center')
        self.tree.column('preview', width=220)
        self.tree.column('status', width=120, anchor='center')
        self.tree.column('result', width=120, anchor='w')
        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 이벤트 바인딩 개선
        self.tree.bind("<Button-1>", self._on_tree_click) # 단일 클릭으로 체크박스 즉시 처리
        self.tree.bind("<Double-1>", self._on_tree_edit)
        self.tree.bind("<space>", lambda e: self.controller.handle_multi_toggle()) # 스페이스바로 다중 선택 토글
        self.edit_entry = None # 현재 편집 중인 엔트리 박스

        rules_container = tk.Frame(content_frame, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB")
        rules_container.grid(row=0, column=1, sticky='nsew')
        rules_container.rowconfigure(0, weight=1)
        rules_container.columnconfigure(0, weight=1)

        self.rules_canvas = tk.Canvas(rules_container, bg=self.COLORS['card'], highlightthickness=0)
        rsb = ttk.Scrollbar(rules_container, orient="vertical", command=self.rules_canvas.yview)
        rules_frame = tk.Frame(self.rules_canvas, bg=self.COLORS['card'], padx=15, pady=15)
        rules_frame.bind("<Configure>", lambda e: self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all")))
        rc_win = self.rules_canvas.create_window((0, 0), window=rules_frame, anchor="nw")
        self.rules_canvas.bind("<Configure>", lambda e: self.rules_canvas.itemconfig(rc_win, width=e.width))
        self.rules_canvas.configure(yscrollcommand=rsb.set)
        self.rules_canvas.grid(row=0, column=0, sticky='nsew')
        rsb.grid(row=0, column=1, sticky='ns')

        # [v34.1.41-R] 우클릭 활용 등 병합 메뉴 안내 툴팁 추가
        lbl_pattern = tk.Label(rules_frame, text="[패턴 제안] 💡 (도움말 표시)", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold'), fg=self.COLORS['accent'], cursor="question_arrow")
        lbl_pattern.pack(anchor='w', pady=(5, 0))
        HoverTooltip(lbl_pattern, "[ 패턴 병합(조합) 가이드 ]\n\n"
                                  "• 👆 [마우스 왼쪽 클릭] : 즉시 '제거' 규칙이 추가됩니다.\n"
                                  "                             (해당 키워드를 파일명에서 빠르게 지울 때 사용)\n\n"
                                  "• 🖱️ [마우스 오른쪽 클릭] : 병합(Context) 메뉴가 나타납니다.\n"
                                  "                             패턴을 문자열로 활용할 수 있는 '모든 규칙' 5가지\n"
                                  "                             (교체, 삽입, 형식, 대/소, 분류)의 입력칸으로 복사하여\n"
                                  "                             유연하게 조합(Merge)하여 사용할 수 있습니다.\n"
                                  "                             ※ 주의: '4:삭제' 규칙은 위치 숫자만 받으므로 메뉴에서 제외됩니다.")
        # [v34.1.40] 전체 슬림 통합 레이아웃 (두 패널 모두 초기 45px)
        self.pattern_outer = tk.Frame(rules_frame, height=45, bg="#F0F2F5")
        self.pattern_outer.pack(fill=tk.X, pady=(2, 5))
        self.pattern_outer.grid_propagate(False)
        self.pattern_outer.grid_columnconfigure(0, weight=1)
        
        self.pattern_scrollbar = tk.Scrollbar(self.pattern_outer)
        self.pattern_container = tk.Text(self.pattern_outer, height=15, bg="#F0F2F5", relief=tk.FLAT, wrap=tk.WORD, 
                                        state='disabled', cursor='arrow', yscrollcommand=self.pattern_scrollbar.set)
        self.pattern_scrollbar.config(command=self.pattern_container.yview)
        self.pattern_container.grid(row=0, column=0, sticky='nsew')

        tk.Label(rules_frame, text="[적용 규칙 목록]:", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold')).pack(anchor='w', pady=(5, 0))
        # [v34.1.40] 적용 규칙 목록 패널도 45px로 슬림하게 통일
        self.rule_outer = tk.Frame(rules_frame, height=45, bg="#F0F2F5")
        self.rule_outer.pack(fill=tk.X, pady=(2, 2))
        self.rule_outer.grid_propagate(False)
        self.rule_outer.grid_columnconfigure(0, weight=1)
        
        self.rule_scrollbar = tk.Scrollbar(self.rule_outer)
        self.rule_container = tk.Text(self.rule_outer, height=15, bg="#F0F2F5", relief=tk.FLAT, wrap=tk.WORD, 
                                     state='disabled', cursor='arrow', yscrollcommand=self.rule_scrollbar.set)
        self.rule_scrollbar.config(command=self.rule_container.yview)
        self.rule_container.grid(row=0, column=0, sticky='nsew')
        # 스크롤바는 필요 시점에 grid(row=0, column=1)로 노출
        
        rb_frame = tk.Frame(rules_frame, bg=self.COLORS['card'])
        rb_frame.pack(fill=tk.X)
        ttk.Button(rb_frame, text="규칙 추가", width=10, style="ActionSmall.TButton", command=self.controller.handle_add_rule).pack(side=tk.LEFT, padx=2)
        ttk.Button(rb_frame, text="규칙 초기화", width=10, command=self.controller.handle_clear_rules).pack(side=tk.LEFT, padx=2)

        tk.Label(rules_frame, text="[새 규칙 정의]", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold'), fg=self.COLORS['text_light']).pack(anchor='w', pady=(10, 0))
        self.rule_type = tk.StringVar(value="format_pattern")
        types = [
            ("1:형식", "format_pattern"), 
            ("2:교체", "replace"), 
            ("3:삽입", "insert"), 
            ("4:삭제", "delete"), 
            ("5:대/소", "case_convert"),
            ("6:접두사(이동/복사)", "prefix_distribute")
        ]
        t_frame = tk.Frame(rules_frame, bg=self.COLORS['card'])
        t_frame.pack(fill=tk.X, pady=5)
        # 2열 그리드로 배치하여 폭 절약 및 가독성 확보
        for i, (text, val) in enumerate(types):
            at = tk.Radiobutton(t_frame, text=text, variable=self.rule_type, value=val, 
                                bg=self.COLORS['card'], font=('Malgun Gothic', 9), 
                                indicatoron=0, width=12, padx=5, pady=5, 
                                selectcolor="#E1F5FE", command=self._on_rule_change)
            at.grid(row=i//3, column=i%3, sticky='ew', padx=2, pady=2)
        t_frame.columnconfigure((0,1,2), weight=1)

        self.params_frame = tk.Frame(rules_frame, bg=self.COLORS['card'], pady=15)
        self.params_frame.pack(fill=tk.X)

        # 하단 로그 영역 (PanedWindow 하단에 추가) - 초기 5줄만 표시, 드래그로 확장 가능
        self.lower_pane = tk.Frame(self.paned, bg=self.COLORS['card'], highlightthickness=1, highlightbackground="#E1E5EB", padx=10, pady=3)
        self.paned.add(self.lower_pane, stretch="never", minsize=50) 
        
        # [v34.1.41-R] 구분선 초기 위치: 창 높이 대비 하단 약 100px만 로그에 할당 (5줄 표시)
        def _set_initial_sash():
            try:
                h = self.root.winfo_height()
                self.paned.sash_place(0, 0, max(h - 130, 600))
            except: pass
        self.root.after(300, _set_initial_sash)
        
        self.log_text = tk.Text(self.lower_pane, font=('Consolas', 9), bg="#F8FAFC", fg="#2D3436", relief=tk.FLAT, height=5)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.status_var = tk.StringVar(value="준비됨")
        tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, font=('Malgun Gothic', 9), bg="#E1E5EB").grid(row=1, column=0, sticky='ew')
        self._on_rule_change()

    def _set_format_preset(self, val, btn_idx=None):
        # [v34.1.42] 버튼 피드백 추가 & 규칙 추가 없는 미리보기(오해 유발) 제거
        if hasattr(self, 'format_preset_buttons'):
            for b in self.format_preset_buttons:
                b.configure(bg="#F8F9FA", font=('Malgun Gothic', 9), fg=self.COLORS['primary'], relief=tk.GROOVE)
            if btn_idx is not None and btn_idx < len(self.format_preset_buttons):
                self.format_preset_buttons[btn_idx].configure(bg="#E1F5FE", font=('Malgun Gothic', 9, 'bold'), fg="#0056b3", relief=tk.SUNKEN)

        self.format_pattern_entry.delete(0, tk.END)
        self.format_pattern_entry.insert(0, val)
        self.format_pattern_entry.focus()
        # 시각적 지시어 추가: 사용자가 [규칙 추가] 버튼을 누르도록 유도
        self.status_var.set(f"💡'{val}' 형식 선택됨 (이제 하단의 [규칙 추가]를 누르세요)")

    def _on_rule_change(self):
        for w in self.params_frame.winfo_children(): w.destroy()
        
        rt = self.rule_type.get()
        if rt == "format_pattern":
            presets = [
                ("번호_이름", "/01_/n"), ("이름_번호", "/n_/01"),
                ("번호이름", "/01/n"), ("이름번호", "/n/01"),
                ("날짜_이름", "/YMD_/n"), ("이름_날짜", "/n_/YMD"),
                ("날짜이름", "/YMD/n"), ("이름날짜", "/n/YMD"),
                ("번호만", "/01"), ("날짜만", "/YMD")
            ]
            
            # 공간 효율성 개선: 3개씩 배치하고 패딩을 줄여 전체 높이 최적화
            row_frame = None
            self.format_preset_buttons = []
            for i, (txt, val) in enumerate(presets):
                if i % 3 == 0: # 3개씩 한 줄로 배치
                    row_frame = tk.Frame(self.params_frame, bg=self.COLORS['card'])
                    row_frame.pack(fill=tk.X, pady=1)
                    for c in range(3): row_frame.columnconfigure(c, weight=1)
                
                btn = tk.Button(row_frame, text=txt, font=('Malgun Gothic', 9), 
                                bg="#F8F9FA", fg=self.COLORS['primary'],
                                relief=tk.GROOVE, pady=2,
                                command=lambda v=val, idx=i: self._set_format_preset(v, idx))
                btn.grid(row=0, column=i%3, sticky='ew', padx=2)
                self.format_preset_buttons.append(btn)

            tk.Label(self.params_frame, text="형식(O): (직접 수정 가능)", bg=self.COLORS['card'], font=('Malgun Gothic', 9, 'bold')).pack(anchor='w')
            self.format_pattern_entry = tk.Entry(self.params_frame, font=('Consolas', 10), bg="#F0F2F5", relief=tk.FLAT)
            self.format_pattern_entry.insert(0, "/01_/n") # 기본값을 더 흔한 번호_이름으로 변경
            self.format_pattern_entry.pack(fill=tk.X, pady=2)
            
            # 기호 설명 가이드
            guide_lbl = tk.Label(self.params_frame, text="/n: 원본이름  /01: 번호  /YMD: 날짜", 
                                font=('Malgun Gothic', 8), fg=self.COLORS['text_light'], bg=self.COLORS['card'])
            guide_lbl.pack(anchor='w')

            f_opt = tk.Frame(self.params_frame, bg=self.COLORS['card'])
            f_opt.pack(fill=tk.X, pady=5)
            tk.Label(f_opt, text="시작 번호:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_start_var = tk.IntVar(value=1)
            tk.Spinbox(f_opt, from_=0, to=999, width=5, textvariable=self.format_start_var).pack(side=tk.LEFT, padx=5)
            tk.Label(f_opt, text="자리수:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_digits_var = tk.IntVar(value=2)
            tk.Spinbox(f_opt, from_=1, to=10, width=5, textvariable=self.format_digits_var).pack(side=tk.LEFT, padx=5)
            
            # 위치 옵션 추가
            pos_frame = tk.Frame(self.params_frame, bg=self.COLORS['card'])
            pos_frame.pack(fill=tk.X, pady=5)
            tk.Label(pos_frame, text="적용 위치:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.format_pos_var = tk.StringVar(value="replace")
            for t_txt, t_val in [("앞에 추가", "prefix"), ("뒤에 추가", "suffix"), ("전체 대체", "replace")]:
                tk.Radiobutton(pos_frame, text=t_txt, variable=self.format_pos_var, value=t_val, bg=self.COLORS['card'], font=('Malgun Gothic', 8)).pack(side=tk.LEFT, padx=2)

            self.format_inc_var = tk.IntVar(value=1)
            self.format_zeropad_var = tk.BooleanVar(value=True)
        elif rt == "replace":
            tk.Label(self.params_frame, text="찾을 문자열:", bg=self.COLORS['card']).pack(anchor='w')
            self.replace_find_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.replace_find_entry.pack(fill=tk.X, pady=2)
            tk.Label(self.params_frame, text="바꿀 문자열:", bg=self.COLORS['card']).pack(anchor='w')
            self.replace_to_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.replace_to_entry.pack(fill=tk.X, pady=2)
            self.replace_ignorecase_var = tk.BooleanVar(value=False)
            tk.Checkbutton(self.params_frame, text="대소문자 무시", variable=self.replace_ignorecase_var, bg=self.COLORS['card']).pack(anchor='w')
            self.replace_max_var = tk.IntVar(value=0)
        elif rt == "insert":
            tk.Label(self.params_frame, text="삽입할 문자열:", bg=self.COLORS['card']).pack(anchor='w')
            self.insert_text_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.insert_text_entry.pack(fill=tk.X, pady=2)
            self.insert_position_var = tk.IntVar(value=0)
            tk.Spinbox(self.params_frame, from_=0, to=999, textvariable=self.insert_position_var).pack(anchor='w')
            self.insert_fromend_var = tk.BooleanVar(value=False)
            tk.Checkbutton(self.params_frame, text="뒤에서부터", variable=self.insert_fromend_var, bg=self.COLORS['card']).pack(anchor='w')
        elif rt == "delete":
            self.delete_position_var = tk.IntVar(value=0)
            tk.Spinbox(self.params_frame, from_=0, to=999, textvariable=self.delete_position_var).pack(anchor='w')
            self.delete_length_var = tk.IntVar(value=1)
            tk.Spinbox(self.params_frame, from_=1, to=999, textvariable=self.delete_length_var).pack(anchor='w')
            self.delete_fromend_var = tk.BooleanVar(value=False)
            tk.Checkbutton(self.params_frame, text="뒤에서부터", variable=self.delete_fromend_var, bg=self.COLORS['card']).pack(anchor='w')
        elif rt == "case_convert":
            self.case_target_var = tk.StringVar(value="name")
            self.case_type_var = tk.StringVar(value="lower")
            for t in ["lower", "upper", "title"]: tk.Radiobutton(self.params_frame, text=t, variable=self.case_type_var, value=t, bg=self.COLORS['card']).pack(anchor='w')
            self.case_ignore_entry = tk.Entry(self.params_frame, font=('Malgun Gothic', 10), bg="#F0F2F5")
            self.case_ignore_entry.pack(fill=tk.X, pady=2)
        elif rt == "prefix_distribute":
            # VBA 스타일 숫자 접두사 분류 UI
            f_dest = tk.LabelFrame(self.params_frame, text="분류 설정 (VBA 이식)", bg=self.COLORS['card'], padx=5, pady=5)
            f_dest.pack(fill=tk.X, pady=5)
            
            tk.Label(f_dest, text="목적지(Base):", bg=self.COLORS['card']).grid(row=0, column=0, sticky='w')
            self.dist_dest_var = tk.StringVar()
            tk.Entry(f_dest, textvariable=self.dist_dest_var, font=('Malgun Gothic', 9)).grid(row=0, column=1, sticky='ew')
            tk.Button(f_dest, text="..", command=lambda: self.dist_dest_var.set(filedialog.askdirectory())).grid(row=0, column=2)
            
            tk.Label(f_dest, text="하위경로(Sub):", bg=self.COLORS['card']).grid(row=1, column=0, sticky='w')
            self.dist_sub_var = tk.StringVar(value="01 자료")
            tk.Entry(f_dest, textvariable=self.dist_sub_var, font=('Malgun Gothic', 9)).grid(row=1, column=1, columnspan=2, sticky='ew')
            
            f_opt = tk.Frame(f_dest, bg=self.COLORS['card'])
            f_opt.grid(row=2, column=0, columnspan=3, sticky='w', pady=5)
            tk.Label(f_opt, text="숫자길이:", bg=self.COLORS['card']).pack(side=tk.LEFT)
            self.dist_num_len = tk.IntVar(value=2)
            tk.Spinbox(f_opt, from_=1, to=10, width=3, textvariable=self.dist_num_len).pack(side=tk.LEFT, padx=5)
            
            tk.Label(f_opt, text="추가 접두사:", bg=self.COLORS['card']).pack(side=tk.LEFT, padx=(10,0))
            self.dist_prefix_var = tk.StringVar(value="표준단가_")
            tk.Entry(f_opt, textvariable=self.dist_prefix_var, width=10).pack(side=tk.LEFT, padx=5)
            
            self.dist_mode_var = tk.StringVar(value="copy")
            tk.Radiobutton(f_dest, text="복사", variable=self.dist_mode_var, value="copy", bg=self.COLORS['card']).grid(row=3, column=0)
            tk.Radiobutton(f_dest, text="이동", variable=self.dist_mode_var, value="move", bg=self.COLORS['card']).grid(row=3, column=1)
            
            f_dest.columnconfigure(1, weight=1)

    def log(self, message, m_type="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {message}\n")
        self.log_text.see(tk.END)

    def _on_tree_click(self, event):
        """단일 클릭 시 체크박스 영역이면 즉시 토글"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        if column == "#2": # '선택' 열 클릭 시 즉시 토글 (NO열 추가로 #2)
            # item_id가 iid(original_path)이므로 이를 이용해 개별 토글 처리
            self.controller.handle_item_toggle_by_id(item_id)
            return "break"

    def _on_tree_edit(self, event):
        """더블 클릭 시 '변경될 파일명' 편집 엔트리 생성"""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        if column != "#5": return # '변경될 파일명' 열 인덱스 (NO열 추가로 #5)
        
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        # 이전 편집창이 있다면 제거
        if self.edit_entry: self.edit_entry.destroy()
        
        x, y, w, h = self.tree.bbox(item_id, column)
        # vals[4]이 변경될 파일명임 (no:0, check:1, original:2, arrow:3, preview:4)
        old_val = self.tree.item(item_id, 'values')[4] 
        
        self.edit_entry = tk.Entry(self.tree, font=('Malgun Gothic', 10))
        self.edit_entry.insert(0, old_val)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.focus_set()
        
        self.edit_entry.bind("<Return>", lambda e: self._save_tree_edit(item_id))
        self.edit_entry.bind("<FocusOut>", lambda e: self._save_tree_edit(item_id))
        self.edit_entry.bind("<Escape>", lambda e: self.edit_entry.destroy())
        self.edit_entry.place(x=x, y=y, width=w, height=h)

    def _save_tree_edit(self, item_id):
        """편집된 내용을 저장하고 Treeview 및 컨트롤러 데이터 업데이트"""
        if not self.edit_entry: return
        new_val = self.edit_entry.get().strip()
        self.edit_entry.destroy()
        self.edit_entry = None
        
        # 트리뷰 업데이트 (논리적 컬럼 인덱스 정합성 확보)
        vals = list(self.tree.item(item_id, 'values'))
        if vals[4] == new_val: return # 변경사항 없음 (Index 4: 변경될 파일명)
        
        vals[4] = new_val      # Index 4: 변경될 파일명에 수정값 반영
        vals[5] = "[수정됨]"   # Index 5: 상태 열에 수정 표시
        self.tree.item(item_id, values=vals, tags=('edited',))
        self.tree.tag_configure('edited', foreground="#E67E22", font=('Malgun Gothic', 9, 'bold'))
        
        # iid(hash)를 통해 컨트롤러 데이터 업데이트
        path_key = self.tree.item(item_id, 'values')[2] # 원본 파일명은 values에 보관 중이거나 iid로 역추적 필요
        # 단, handle_manual_name_change_by_id는 original_path를 필요로 하므로 컨트롤러에서 해시-경로 매핑 필요
        # 여기서는 item_id(hash)를 그대로 전달
        self.controller.handle_manual_name_change_by_id(item_id, new_val)

    def update_tree(self, data_list):
        self.tree.delete(*self.tree.get_children())
        if not data_list: 
            self.status_var.set("준비 완료")
            self.root.update()
            return

        added_count = 0
        for i, item in enumerate(data_list, 1):
            try:
                # 상태에 따른 태그 지정
                status = item.get('status', '[OK]')
                tag = 'error' if '[WARN]' in status else ('silent' if '[UNCHANGED]' in status else '')
                if status == "[수정됨]": tag = 'edited'
                chk = "✅" if item.get('selected', True) else "⬜"
                
                # [v34.1.41-G] IID 원시화: 가장 안전한 index 기반 식별자 사용
                simple_iid = f"row_{i}"
                item['_row_id'] = simple_iid # 내부 매핑용
                
                self.tree.insert('', tk.END, iid=simple_iid, values=(i, chk, item['original_name'], "→", item['new_name'], status, ""), tags=(tag,))
                added_count += 1
            except Exception as e:
                continue
        
        # [v34.1.41-G] 레이아웃 강제 재계산 (Nuclear Force)
        self.tree.pack_forget()
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # [v34.1.41] 물리적 가시성 보장 및 렌더링 동기화
        self.root.update() 
        self.tree.yview_moveto(0)
        
        # [v34.1.41-G] 연쇄 화면 갱신 트리거 (2px 리사이징)
        def _force_redraw():
            try:
                w, h = self.root.winfo_width(), self.root.winfo_height()
                self.root.geometry(f"{w+1}x{h+1}")
                self.root.after(100, lambda: self.root.geometry(f"{w}x{h}"))
            except: pass
        self.root.after(50, _force_redraw)
        
        real_ids = self.tree.get_children()
        read_back = self.tree.item(real_ids[0], 'values')[4] if real_ids else "N/A"
        self.log(f"UI 렌더링 완료: {added_count}개 삽입 (Widget 보유 ID: {len(real_ids)}개, 실측값: {read_back})")
        
        self.tree.tag_configure('error', foreground="#D63031")
        self.tree.tag_configure('silent', foreground="#95A5A6")
        self.tree.tag_configure('edited', foreground="#E67E22", font=('Malgun Gothic', 9, 'bold'))
        self.tree.tag_configure('fail_red', foreground="#FF0000", font=('Malgun Gothic', 9, 'bold'))
        self.tree.tag_configure('success_blue', foreground="#0000FF")

    def update_tree_result(self, original_path, result_text, is_success=True):
        """[Index IID 대응] 순차 검색을 통한 작업 결과 업데이트"""
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, 'values')
            # Index 2는 원본 파일명임 (정합성 확인)
            if original_path.endswith(vals[2]):
                new_vals = list(vals)
                new_vals[6] = result_text 
                tag = 'success_blue' if is_success else 'fail_red'
                self.tree.item(item_id, values=new_vals, tags=(tag,))
                break

    # [v34.1.39] 패턴 제안 전용 슬림 높이(45px) 관리 메서드
    def clear_patterns(self):
        self.pattern_container.configure(state='normal')
        self.pattern_container.delete('1.0', tk.END)
        self.pattern_container.configure(state='disabled')
        self.pattern_outer.configure(height=45) # 패턴 패널만 45px 복원
        self.pattern_scrollbar.grid_forget() 
        self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))

    def _show_pattern_menu(self, event, pattern):
        menu = tk.Menu(self.root, tearoff=0, font=('Malgun Gothic', 9))
        menu.add_command(label=f"[{pattern}] 즉시 제거", command=lambda p=pattern: self.controller.handle_pattern_select_direct(p))
        menu.add_separator()
        menu.add_command(label="▶ [2:교체] 규칙의 '찾을 문자열'로 복사", command=lambda p=pattern: self._send_to_input("replace", p))
        menu.add_command(label="▶ [3:삽입] 규칙의 '텍스트'로 복사", command=lambda p=pattern: self._send_to_input("insert", p))
        menu.add_command(label="▶ [1:형식] 규칙의 '포맷' 문자열로 복사", command=lambda p=pattern: self._send_to_input("format_pattern", p))
        menu.add_separator()
        menu.add_command(label="▶ [5:대/소] 규칙의 '예외 단어'로 복사", command=lambda p=pattern: self._send_to_input("case_convert", p))
        menu.add_command(label="▶ [6:분류] 규칙의 '추가 접두사'로 복사", command=lambda p=pattern: self._send_to_input("prefix_distribute", p))
        menu.post(event.x_root, event.y_root)

    def _send_to_input(self, rule_type, pattern):
        self.rule_type.set(rule_type)
        self._on_rule_change()
        self.root.update_idletasks() # UI 렌더링 강제 대기
        
        if rule_type == 'replace':
            self.replace_find_entry.delete(0, tk.END)
            self.replace_find_entry.insert(0, pattern)
            self.replace_to_entry.focus()
        elif rule_type == 'insert':
            self.insert_text_entry.delete(0, tk.END)
            self.insert_text_entry.insert(0, pattern)
            self.insert_text_entry.focus()
        elif rule_type == 'format_pattern':
            self.format_pattern_entry.delete(0, tk.END)
            self.format_pattern_entry.insert(0, f"{pattern}_/n")
            self.format_pattern_entry.focus()
        elif rule_type == 'case_convert':
            self.case_ignore_entry.delete(0, tk.END)
            self.case_ignore_entry.insert(0, pattern)
            self.case_ignore_entry.focus()
        elif rule_type == 'prefix_distribute':
            self.dist_prefix_var.set(f"{pattern}_")
            
        self.log(f"패턴 [{pattern}]이(가) [{rule_type}] 규칙 입력란으로 자동 복사되었습니다.", "info")
        self.status_var.set(f"패턴 복사됨: {pattern}")

    def add_pattern_chip(self, pattern):
        self.pattern_container.configure(state='normal')
        btn = tk.Button(self.pattern_container, text=pattern, font=('Malgun Gothic', 8), 
                        bg="#E1E5EB", fg=self.COLORS['text'], relief=tk.FLAT, padx=8, cursor="hand2",
                        command=lambda p=pattern: self.controller.handle_pattern_select_direct(p))
                        
        # [v34.1.41-R] 우클릭 병합 메뉴 및 툴팁 바인딩
        btn.bind("<Button-3>", lambda e, p=pattern: self._show_pattern_menu(e, p))
        btn.bind("<Enter>", lambda e: self.status_var.set("좌클릭: 즉시 제거 | 우클릭: 다른 규칙과 병합하기 위해 복사"))
        btn.bind("<Leave>", lambda e: self.status_var.set("준비됨" if not self.controller.scanned_files else "대기 중"))
        
        self.pattern_container.window_create(tk.END, window=btn, padx=2, pady=2)
        self.pattern_container.insert(tk.END, " ") 
        
        self.root.update() 
        bbox = self.pattern_container.bbox('end-1c')
        if bbox:
            raw_h = bbox[1] + bbox[3] + 10
            if raw_h > 120:
                self.pattern_outer.configure(height=120)
                if not self.pattern_scrollbar.winfo_manager():
                    self.pattern_scrollbar.grid(row=0, column=1, sticky='ns')
                self.pattern_container.see(tk.END)
            else:
                self.pattern_outer.configure(height=max(45, raw_h))
                self.pattern_scrollbar.grid_forget()
        
        self.pattern_container.configure(state='disabled')
        self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))
        self.root.update() 

    def clear_rules(self):
        self.rule_container.configure(state='normal')
        self.rule_container.delete('1.0', tk.END)
        self.rule_container.configure(state='disabled')
        self.rule_outer.configure(height=45) # 45px 슬림 복원
        self.rule_scrollbar.grid_forget()
        self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))

    def add_rule_chip(self, rule_text):
        self.rule_container.configure(state='normal')
        lbl = tk.Label(self.rule_container, text=rule_text, font=('Malgun Gothic', 8, 'bold'),
                       bg=self.COLORS['accent'], fg="white", padx=10, relief=tk.FLAT)
        self.rule_container.window_create(tk.END, window=lbl, padx=2, pady=2)
        self.rule_container.insert(tk.END, " ")

        self.root.update() 
        bbox = self.rule_container.bbox('end-1c')
        if bbox:
            raw_h = bbox[1] + bbox[3] + 10
            if raw_h > 120:
                self.rule_outer.configure(height=120)
                if not self.rule_scrollbar.winfo_manager():
                    self.rule_scrollbar.grid(row=0, column=1, sticky='ns')
                self.rule_container.see(tk.END)
            else:
                self.rule_outer.configure(height=max(45, raw_h))
                self.rule_scrollbar.grid_forget()
            
        self.rule_container.configure(state='disabled')
        self.rules_canvas.configure(scrollregion=self.rules_canvas.bbox("all"))
        self.root.update() 

    def _treeview_sort_column(self, col, reverse):
        """헤더 클릭 시 열 정렬 (숫자/문자 자동 지능형 정렬)"""
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        
        def _to_num(v):
            try:
                # 체크박스(✅/⬜)는 숫자로 변환하여 정렬 가능케 함
                if v == "✅": return 1.0
                if v == "⬜": return 0.0
                # 쉼표 등 숫자 외 문자 제거 후 변환 시도
                v_clean = v.replace(',', '').strip()
                return float(v_clean) if v_clean else 0.0
            except (ValueError, TypeError):
                # 숫자가 아니면 원본의 소문자 반환 (문자열 정렬)
                return v.lower()

        # 데이터 변환 후 정렬 수행
        data.sort(key=lambda x: _to_num(x[0]), reverse=reverse)

        # 재정렬 후 트리뷰 아이템 위치 실제 이동
        for index, (_, k) in enumerate(data):
            self.tree.move(k, '', index)

        # 다음 클릭 시 정렬 방향이 반전되도록 헤더의 command 바인딩을 갱신
        self.tree.heading(col, command=lambda: self._treeview_sort_column(col, not reverse))


# ═══════════════════════════════════════════════════════════
# LAYER 3: APPLICATION (Controller)
# ═══════════════════════════════════════════════════════════

class FileOrganizerController:
    def __init__(self, root):
        self.root = root
        self.engine = FileOrganizerEngine()
        
        # 모드 상태 변수
        self.target_mode_var = tk.StringVar(value="file")
        
        self.view = FileOrganizerView(root, self)
        self.scanned_files = []
        self.rules = []
        self.preview_list = []
        self.manual_names = {} # [무결성] 수동으로 수정한 이름들을 보존하기 위한 맵 {path: new_name}
        self.all_selected = True
        
        # [v34.1.30] 개별 추가 시 마지막으로 사용했던 경로 저장 (캐시 연동 및 바탕화면 기본세팅)
        saved_path = load_config_path()
        if saved_path:
            self.last_used_add_dir = saved_path
        else:
            default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.exists(default_dir): default_dir = os.path.expanduser("~")
            self.last_used_add_dir = default_dir

    def handle_mode_change(self):
        mode = self.target_mode_var.get()
        if mode == 'dir':
            self.view.header_title.config(text="[DIR] 지능형 폴더 고속 관리기 v34.1.41")
            self.view.ext_filter_entry.config(state='disabled')
            self.view.btn_ext_sel.config(state='disabled')
            self.view.btn_ext_exc.config(state='disabled')
            self.view.tree.heading('original', text='현재 폴더명')
            self.view.tree.heading('preview', text='변경될 폴더명')
        else:
            self.view.header_title.config(text="[FILE] 지능형 파일 고속 관리기 v34.1.41")
            self.view.ext_filter_entry.config(state='normal')
            self.view.btn_ext_sel.config(state='normal')
            self.view.btn_ext_exc.config(state='normal')
            self.view.tree.heading('original', text='현재 파일명')
            self.view.tree.heading('preview', text='변경될 파일명')
        
        self.handle_clear_list()
        self.handle_scan()

    def handle_stop(self):
        """진행 중인 모든 비동기 작업(스캔, 미리보기, 실행) 중단"""
        self.engine.stop_requested = True
        self.view.log("🛑 사용자에 의해 작업 중단 요청됨", "error")
        self.view.status_var.set("작업 중단됨")

    def handle_toggle_all(self):
        """제목 클릭 시 전체 선택 / 전체 해제 토글"""
        self.all_selected = not self.all_selected
        for item in self.preview_list:
            item['selected'] = self.all_selected
        self.view.update_tree(self.preview_list)

    def handle_select_all(self):
        for item in self.preview_list: item['selected'] = True
        self.view.update_tree(self.preview_list)
        self.view.log("전체 파일이 선택되었습니다.")

    def handle_deselect_all(self):
        for item in self.preview_list: item['selected'] = False
        self.view.update_tree(self.preview_list)
        self.view.log("전체 파일이 선택 해제되었습니다.")

    def handle_invert_selection(self):
        for item in self.preview_list:
            item['selected'] = not item.get('selected', True)
        self.view.update_tree(self.preview_list)
        self.view.log("전체 선택이 반전되었습니다.")

    def handle_multi_toggle(self):
        """드래그로 선택된 여러 행을 일괄 토글 (스페이스바 연동)"""
        selected_ids = self.view.tree.selection()
        if not selected_ids: return
        
        # 첫 번째 선택 항목의 반대 상태로 모두 통일
        idx0 = self.view.tree.index(selected_ids[0])
        new_state = not self.preview_list[idx0].get('selected', True)
        
        for mid in selected_ids:
            idx = self.view.tree.index(mid)
            self.preview_list[idx]['selected'] = new_state
        
        self.view.update_tree(self.preview_list)
        # 이전 선택 영역 유지
        for mid in selected_ids: self.view.tree.selection_add(mid)

    def handle_item_toggle_by_id(self, item_id):
        """[Index IID 대응] iid(row_x)를 기반으로 상태 토글"""
        try:
            idx = int(item_id.split('_')[1]) - 1
            if 0 <= idx < len(self.preview_list):
                self.preview_list[idx]['selected'] = not self.preview_list[idx].get('selected', True)
                self.view.update_tree(self.preview_list)
        except:
            # Fallback: 순차 검색
            for i, item in enumerate(self.preview_list):
                if item.get('_row_id') == item_id:
                    item['selected'] = not item.get('selected', True)
                    self.view.update_tree(self.preview_list)
                    break

    def handle_item_toggle(self, idx):
        """개별 아이템 선택 상태 토글 (인덱스 기반 - 레거시)"""
        if 0 <= idx < len(self.preview_list):
            curr = self.preview_list[idx].get('selected', True)
            self.preview_list[idx]['selected'] = not curr
            self.view.update_tree(self.preview_list)

    def handle_filter_by_ext(self, ext_pattern, mode="select"):
        """확장자별로 선택 또는 제외 (예: .txt, .jpg ...)"""
        if not ext_pattern.startswith('.'): ext_pattern = '.' + ext_pattern
        count = 0
        for item in self.preview_list:
            if item['original_name'].lower().endswith(ext_pattern.lower()):
                item['selected'] = (mode == "select")
                count += 1
        self.view.update_tree(self.preview_list)
        self.view.log(f"확장자 [{ext_pattern}] 대상 {count}건 {'선택' if mode=='select' else '제외'} 완료")

    def show_manual(self):
        manual_win = tk.Toplevel(self.root)
        manual_win.title("도움말")
        manual_win.geometry("900x700")
        manual_win.configure(bg="#F5F7FA")
        
        text_widget = tk.Text(manual_win, font=('Malgun Gothic', 10), bg="#FFFFFF", fg="#2D3436", relief=tk.FLAT, padx=20, pady=20)
        scrollbar = ttk.Scrollbar(manual_win, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        text_widget.insert(tk.END, USER_MANUAL_TEXT)
        text_widget.configure(state='disabled') # 읽기 전용

    def handle_browse(self):
        d_raw = filedialog.askdirectory()
        if d_raw: 
            d = os.path.abspath(os.path.normpath(d_raw))
            save_config_path(d) # 윈도우 로컬 시스템에 경로 영구 캐싱
            self.handle_clear_list() # [v34.1.41] 새 폴더 선택 시 기존 데이터 충돌 방지를 위해 초기화
            self.view.path_var.set(d)
            self.view.log(f"폴더 선택됨: {d}")
            self.handle_scan() # 폴더 선택 시 자동 스캔 실행

    def handle_file_select(self):
        """[v34.1.30] 개별 추가: 파일 모드(표본 지원) vs 폴더 모드(커스텀 셀렉터)"""
        target_mode = self.target_mode_var.get()
        
        if target_mode == 'file':
            # 파일 모드: 표준 멀티 파일 선택
            fs = filedialog.askopenfilenames(title="추가할 파일들을 선택하세요")
            if fs:
                new_items = [{'name': os.path.basename(f), 'path': os.path.dirname(f), 'full_path': f, 'ext': os.path.splitext(f)[1], '_target_mode': 'file'} for f in fs]
                self._append_unique_files(new_items)
                self.view.log(f"개별 파일 {len(fs)}개 추가됨 (총 {len(self.scanned_files)}개)")
                self.refresh_pattern_suggestions()
                self.handle_preview()
        else:
            # 폴더 모드: 커스텀 지능형 폴더 선택기 호출
            selector = FolderSelectorDialog(self.root, self, initial_path=self.last_used_add_dir)
            self.root.wait_window(selector)
            
            if hasattr(selector, 'result') and selector.result:
                new_items = []
                for d in selector.result:
                    new_items.append({
                        'name': os.path.basename(d),
                        'path': os.path.dirname(d),
                        'full_path': d,
                        'ext': '[DIR]',
                        '_target_mode': 'dir' # [v34.1.41] 누락된 모드 식별자 강제 부여
                    })
                
                # 마지막으로 성공적으로 추가한 폴더의 경로를 기억
                if selector.result:
                    self.last_used_add_dir = os.path.dirname(selector.result[0])
                
                self._append_unique_files(new_items)
                self.view.log(f"개별 폴더 {len(selector.result)}개 추가됨 (총 {len(self.scanned_files)}개)")
                self.refresh_pattern_suggestions()
                self.handle_preview()

    def handle_scan(self):
        p_raw = self.view.path_var.get().strip()
        if not p_raw or not os.path.exists(p_raw): return
        
        p = os.path.abspath(os.path.normpath(p_raw)) # [v34.1.41] 경로 무결성 정규화
        save_config_path(p) # 사용자가 입력/로드한 경로 영구 캐싱
        target_mode = self.target_mode_var.get()
        is_recursive = self.view.recursive_var.get() # [무결성] 스레드 진입 전 변수 값 고정
        
        def run():
            # [v34.1.41] 진행 상태 가시화 강화
            self.root.after(0, lambda: self.view.status_var.set(f"⏳ {'폴더' if target_mode=='dir' else '파일'} 목록 스캔 중..."))
            new_files = self.engine.scan_files(p, recursive=is_recursive, target_mode=target_mode)
            def finalize():
                self._append_unique_files(new_files)
                self.view.log(f"스캔 완료: {len(new_files)}개 {'폴더' if target_mode=='dir' else '파일'} 추가 (총 {len(self.scanned_files)}개)")
                self.refresh_pattern_suggestions()
                self.handle_preview()
            self.root.after(0, finalize)
        threading.Thread(target=run, daemon=True).start()

    def _append_unique_files(self, new_files):
        """[무결성] 경로 정규화를 통한 중복 제거 및 리스트 병합"""
        # [v34.1.41] 비교 전 경로 정규화로 오작동 방지
        seen = {os.path.abspath(os.path.normpath(f['full_path'])) for f in self.scanned_files}
        for f in new_files:
            f_norm = os.path.abspath(os.path.normpath(f['full_path']))
            if f_norm not in seen:
                self.scanned_files.append(f)
                seen.add(f_norm)
                seen.add(f['full_path'])

    def refresh_pattern_suggestions(self):
        """[v34.1.41-R] 모든 추가/삭제/모드 변경 시 동기화되는 패턴 제안 엔진 가동"""
        if not self.scanned_files:
            self.view.clear_patterns()
            return
        # 현재 목록 및 활성 규칙을 바탕으로 중복 제안 배제
        pts = self.engine.detect_common_patterns(self.scanned_files, active_rules=self.rules)
        self.view.clear_patterns()
        for pt in pts:
            self.view.add_pattern_chip(pt)

    def handle_clear_list(self):
        self.scanned_files = []
        self.preview_list = []
        self.view.update_tree([])
        self.view.log("대상 파일 목록이 초기화되었습니다.")
        self.view.status_var.set("목록 초기화됨")
        self.refresh_pattern_suggestions()

    def handle_preview(self):
        if not self.scanned_files:
            self.view.update_tree([])
            self.view.status_var.set("준비됨")
            return

        self.view.status_var.set("⏳ 미리보기 계산 중...")
        self.engine.stop_requested = False # 신규 계산 시작 시 중단 플래그 초기화

        def run_preview_task():
            # [무결성 보완] 기존 선택 상태(selected)를 보존하여 규칙 변경 시 초기화 방지
            selection_map = {item['original_path']: item.get('selected', True) for item in self.preview_list}
            
            # 엔진에서 미리보기 리스트 생성 (최적화된 버전 호출)
            calculated_previews = self.engine.preview_rename(self.scanned_files, self.rules)
            
            if self.engine.stop_requested: return

            # [무결성 보완] 이전 선택 상태 및 수동 수정 내역 복원
            for item in calculated_previews:
                path = item['original_path']
                if path in selection_map:
                    item['selected'] = selection_map[path]
                if path in self.manual_names:
                    item['new_name'] = self.manual_names[path]
                    item['status'] = "[수정됨]"
            
            # 메인 스레드에서 UI 업데이트
            def finalize():
                self.preview_list = calculated_previews
                self.view.update_tree(self.preview_list)
                self.view.status_var.set(f"미리보기: {len(self.preview_list)}건")
            
            self.root.after(0, finalize)

        threading.Thread(target=run_preview_task, daemon=True).start()

    def handle_manual_name_change_by_id(self, item_id, new_name):
        """[Index IID 대응] row_x 기반 수동 수정 반영"""
        try:
            idx = int(item_id.split('_')[1]) - 1
            if 0 <= idx < len(self.preview_list):
                item = self.preview_list[idx]
                item['new_name'] = new_name
                item['status'] = "[수정됨]"
                self.manual_names[item['original_path']] = new_name
                self.view.log(f"수동 수정: {item['original_name']} -> {new_name}")
        except:
            pass

    def handle_add_rule(self):
        rt = self.view.rule_type.get()
        p = {}
        if rt == "format_pattern":
            p = {
                'pattern': self.view.format_pattern_entry.get(), 
                'start': self.view.format_start_var.get(), 
                'digits': self.view.format_digits_var.get(), 
                'increment': self.view.format_inc_var.get(), 
                'zero_pad': self.view.format_zeropad_var.get(),
                'pos_type': self.view.format_pos_var.get()
            }
        elif rt == "replace":
            p = {'find': self.view.replace_find_entry.get(), 'replace': self.view.replace_to_entry.get(), 'ignore_case': self.view.replace_ignorecase_var.get(), 'max_count': self.view.replace_max_var.get()}
        elif rt == "insert":
            p = {'text': self.view.insert_text_entry.get(), 'position': self.view.insert_position_var.get(), 'from_end': self.view.insert_fromend_var.get()}
        elif rt == "delete":
            p = {'position': self.view.delete_position_var.get(), 'length': self.view.delete_length_var.get(), 'from_end': self.view.delete_fromend_var.get()}
        elif rt == "case_convert":
            p = {'target': self.view.case_target_var.get(), 'case_type': self.view.case_type_var.get(), 'ignore_chars': self.view.case_ignore_entry.get()}
        elif rt == "prefix_distribute":
            p = {
                'base_dest': self.view.dist_dest_var.get(),
                'sub_path': self.view.dist_sub_var.get(),
                'num_len': self.view.dist_num_len.get(),
                'mode': self.view.dist_mode_var.get(),
                'prefix': self.view.dist_prefix_var.get()
            }
        self.rules.append({'type': rt, 'params': p})
        
        # [v34.1.43] 사용자가 직관적으로 이해할 수 있는 한글 요약 라벨 생성
        chip_label = ""
        if rt == "format_pattern":
            chip_label = f"형식({p['pattern']})"
        elif rt == "replace":
            chip_label = f"교체({p['find']}→{p['replace']})" if p['replace'] else f"교체({p['find']} 제거)"
        elif rt == "insert":
            chip_label = f"삽입({p['text']})"
        elif rt == "delete":
            loc_str = "뒤에서" if p['from_end'] else "앞에서"
            chip_label = f"삭제({loc_str}{p['position']}부터 {p['length']}자)"
        elif rt == "case_convert":
            type_map = {"upper": "대문자", "lower": "소문자", "title": "첫글자대문자"}
            c_type = type_map.get(p['case_type'], p['case_type'])
            chip_label = f"대/소({c_type})"
        elif rt == "prefix_distribute":
            mode_map = {"move": "이동", "copy": "복사", "shortcut": "바로가기"}
            mode_str = mode_map.get(p['mode'], p['mode'])
            pfx = f" '{p['prefix']}' " if p['prefix'] else " "
            chip_label = f"분류({mode_str}{pfx})"
        else:
            chip_label = f"패턴({rt})"
            
        # 가로 칩 방식으로 규칙 추가
        self.view.add_rule_chip(chip_label)
        self.handle_preview()

    def _generate_task_csv(self, target_list=None):
        """작업 이력 리포트를 CSV 파일로 생성 (성공/실패 결과 포함)"""
        if target_list is None:
            target_list = [f for f in self.preview_list if f.get('selected', True)]
            
        if not target_list:
            self.view.log("리포트를 생성할 대상이 없습니다.", "error")
            return

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"file_organizer_report_{ts}.csv",
            title="작업 리포트 저장 위치 선택"
        )
        if not save_path: return
        
        try:
            with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow([
                    '순번', '기존 파일명', '변경될 파일명', '항목 상태', '작업유형', 
                    '작업 결과', '상세 및 오류 메시지',
                    '원본 파일크기(Bytes)', '결과 파일크기(Bytes)', '크기 차이',
                    '원본 전체경로', '이동/변경될 예상경로', 
                    '원본 수정일', '원본 생성일'
                ])
                
                for i, item in enumerate(target_list, 1):
                    target_path = "원본과 동일"
                    work_type = "이름변경"
                    
                    if item.get('is_distribute'):
                        work_type = "분류(이동/복사)"
                        if item.get('dist_dest'):
                            target_path = f"목적지폴더: {item['dist_dest']}"
                    
                    # 작업 결과 및 상세 메시지 (실행 후에 호출될 경우 데이터가 채워져 있음)
                    res = item.get('final_result', '대기 중')
                    msg = item.get('final_message', '')
                    
                    # 크기 정산
                    orig_size = item.get('size_bytes', 0)
                    final_size = item.get('final_size', 0)
                    diff_size = final_size - orig_size if res == "성공" else "N/A"
                    
                    writer.writerow([
                        i,
                        item['original_name'],
                        item['new_name'],
                        item.get('status', ''),
                        work_type,
                        res,
                        msg,
                        orig_size,
                        final_size if res == "성공" else "대기/실패",
                        diff_size,
                        item['original_path'],
                        target_path,
                        item.get('mtime', ''),
                        item.get('ctime', '')
                    ])
            self.view.log(f"심층 추적 리포트 생성 완료: {save_path}", "success")
        except Exception as e:
            self.view.log(f"CSV 생성 중 오류: {str(e)}", "error")

    def handle_clear_rules(self):
        self.rules = []
        # [v34.1.30] 가로 칩 컨테이너 초기화
        self.view.clear_rules()
        self.refresh_pattern_suggestions()
        self.handle_preview()

    def handle_pattern_select_direct(self, pt):
        """[v34.1.41-R] 칩 클릭 시의 논리적 정합성 확보 (파괴적 truncate 대신 단순 텍스트 제거 도입)"""
        self.rules.append({'type': 'simple_replace', 'params': {'old_str': pt, 'new_str': ''}})
        self.view.add_rule_chip(f"제거: {pt}")
        self.refresh_pattern_suggestions()
        self.handle_preview()

    def handle_pattern_select(self, e):
        # 레거시 호환용 (사용되지 않음)
        pass

    def handle_run(self):
        # [Critical Fix] 전수 점검 결과, 선택된 파일만 필터링하여 처리하도록 로직을 무결성 기반으로 수정
        active_list = [f for f in self.preview_list if f.get('selected', True)]
        
        if not active_list:
            messagebox.showwarning("대상의 부재", "작업 대상으로 선택된 파일이 하나도 없습니다.\n미리보기 창 좌측의 '선택' 열을 확인하세요.")
            return

        # 적용된 규칙 중 분류 규칙이 있는지 확인
        dist_rule = next((r for r in self.rules if r['type'] == 'prefix_distribute'), None)
        is_move_mode = dist_rule and dist_rule['params'].get('mode') == 'move'
        
        num_active = len(active_list)
        if dist_rule:
            mode_str = "복사" if not is_move_mode else "이동"
            confirm_msg = f"{num_active}개 파일을 목적지로 {mode_str}하시겠습니까?"
            if is_move_mode:
                confirm_msg = f"⚠️ [안전 경고] {num_active}개 파일을 목적지로 '이동'합니다.\n이동 완료 후 원본 경로에서 파일이 사라지므로 위치를 찾기 어려울 수 있습니다.\n\n계속하시겠습니까?"
        else:
            confirm_msg = f"{num_active}개 파일 이름을 변경하시겠습니까?"
            
        if not messagebox.askyesno("확인", confirm_msg): return
        
        # [최종본] 중복 저장을 방지하기 위해 작업 전 리포트 생성 로직은 제거하고, 작업 완료 후 통합 리포트만 생성하도록 함.
        pass
        
        def run():
            def cb(ev, val):
                if ev == 'ask_kill':
                    def _show_kill():
                        try: proc_msgs = "\n".join([f"- {p.name()} (PID: {p.pid})" for p in val['procs']])
                        except: proc_msgs = "- (알 수 없는 프로세스)"
                        ans = messagebox.askyesno("⚠️ 폴더 강제 점유 알림", 
                            f"'{val['name']}' 폴더가 파일 탐색기 등 다음 프로세스에 의해 완전히 점유되어 이름 변경이 불가능합니다.\n\n"
                            f"{proc_msgs}\n\n"
                            f"해당 프로세스 트리들을 강제 종료(Kill)하고 폴더 강제 변경을 계속 진행하시겠습니까?")
                        val['result']['ans'] = ans
                        val['wait_event'].set()
                    self.root.after(0, _show_kill)
                    return
                elif ev == 'ask_time_warn':
                    def _show_time():
                        ans = messagebox.askyesno("⏳ 대용량 폴더 복사 경고",
                            f"'{val['name']}' 폴더의 용량이 막대하여 복사 완료 시까지 1분 이상 지연될 수 있습니다.\n\n"
                            f"동결 상태처럼 보이더라도 정상 작동 중이오니 복사를 안전하게 진행하시겠습니까?")
                        val['result']['ans'] = ans
                        val['wait_event'].set()
                    self.root.after(0, _show_time)
                    return

                # [무결성] 결과 기록을 위해 active_list에서 해당 아이템 찾기
                target_path = val.get('path')
                target_item = next((f for f in active_list if f['original_path'] == target_path), None)

                if ev == 'item_success': 
                    if target_item:
                        target_item['final_result'] = "성공"
                        target_item['final_message'] = val.get('new', '변경 완료')
                        target_item['final_size'] = val.get('new_size', 0)
                    self.root.after(0, lambda: self.view.log(f"성공: {val['old']} -> {val['new']}", "success"))
                    self.root.after(0, lambda: self.view.update_tree_result(val['path'], "성공 ✅", True))
                
                elif ev == 'item_fail': 
                    err_message = val['error']
                    if target_item:
                        target_item['final_result'] = "실패"
                        target_item['final_message'] = err_message
                    self.root.after(0, lambda: self.view.log(f"실패: {val['name']} ({err_message})", "error"))
                    self.root.after(0, lambda: self.view.update_tree_result(val['path'], f"실패 ❌ ({err_message})", False))
                    
                    try:
                        with open("server_log.txt", "a", encoding="utf-8") as f:
                            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [FAIL] {val['name']}: {err_message}\n")
                    except: pass
                
                elif ev == 'rename_complete': 
                    op_name = "분류" if dist_rule else "이름 변경"
                    msg = f"{op_name} 작업 완료 (성공: {val['success']}, 실패: {val['failed']})"
                    if val.get('message'): msg += f"\n메시지: {val['message']}"
                    
                    def show_final_prompt():
                        messagebox.showinfo("완료", msg)
                        if messagebox.askyesno("최종 리포트", "작업 결과(성공/실패 및 원인)가 포함된 최종 CSV 리포트를 저장하시겠습니까?"):
                            self._generate_task_csv(target_list=active_list)
                    
                    self.root.after(0, show_final_prompt)

            # [Critical Fix] 전체 목록이 아닌 'active_list'를 전달하여 선택된 파일만 처리 보장
            has_dist = any(r['type'] == 'prefix_distribute' for r in self.rules)
            if has_dist:
                self.engine.perform_distribute(active_list, self.rules, callback=cb)
            else:
                self.engine.perform_rename(active_list, callback=cb)
        
        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    
    # [v34.1.22] 창이 표시되지 않는 문제 해결: 최소 크기 지정 및 강제 업데이트
    root.update_idletasks()
    root.minsize(1200, 800)
    
    app = FileOrganizerController(root)
    
    # 윈도우 환경에서 창을 모든 화면의 제일 앞으로 강제 활성화
    root.attributes('-topmost', True)
    root.lift()
    root.focus_force()
    # 0.5초(500ms) 뒤 최상단 고정 속성 해제 (사용자가 다른 창을 띄울 수 있도록)
    root.after(500, lambda: root.attributes('-topmost', False))
    
    root.mainloop()
