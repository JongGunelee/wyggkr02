import http.server
import socketserver
import json
import subprocess
import os
import threading
import time
import sys
import atexit

"""
================================================================================
🚀 대시보드 서버 엔진 (Dashboard Server Engine) v2.7 [Web Hybrid Ready]
================================================================================
- 아키텍처: Interface Adapter (HTTP Request Handling)
- 주요 기능: 로컬 스크립트 실행(Popen), 대시보드 무결성 유지, 하트비트 모니터링
- 무결성 보증: 서버 종료 시 임시 로그(server_log.txt) atexit 핸들러를 통한 자동 전수 소거
- v2.7 업데이트: GitHub Pages/Web Dashboard 연동용 CORS/OPTIONS/PNA 대응 추가
================================================================================
"""

# Configuration
PORT = 8501  # Use a non-standard port to avoid conflicts
HTML_FILE = "00 dashboard.html"

class RequestHandler(http.server.SimpleHTTPRequestHandler):
    """
    대시보드 웹 인터페이스와 로컬 OS 간의 통신을 담당하는 인터페이스 어댑터.
    HTTP GET/POST 요청을 처리하여 스크립트 실행 및 관리를 수행합니다.
    """
    
    def do_OPTIONS(self):
        """GitHub Pages 등 외부 웹 UI의 CORS preflight 요청을 처리합니다."""
        self.send_response(204)
        self.end_headers()

    def do_GET(self):
        """정적 HTML 파일 제공 및 서버 생존 확인(Health Check) 처리."""
        # 모든 요청 시 하트비트 시간 갱신하여 서버 생존 연장
        global LAST_HEARTBEAT_TIME
        LAST_HEARTBEAT_TIME = time.time()

        if self.path == '/':
            self.path = HTML_FILE
        
        if self.path == '/health':
            self._send_json({'status': 'running'})
            return

        if self.path == '/heartbeat':
            self._send_json({'status': 'alive'})
            return

        return super().do_GET()

    def end_headers(self):
        """모든 응답에 브라우저 캐시 방지 헤더를 추가하여 대시보드 강제 최신화."""
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Access-Control-Allow-Private-Network', 'true')
        super().end_headers()

    def do_POST(self):
        """
        API 엔드포인트 라우팅.
        - /run: 개별 스크립트 실행
        - /add_script: 대시보드 스크립트 리스트에 수동 등록
        - /shutdown: 서버 종료
        """
        # 모든 요청 시 하트비트 시간 갱신
        global LAST_HEARTBEAT_TIME
        LAST_HEARTBEAT_TIME = time.time()

        if self.path == '/run':
            self._handle_run()
        elif self.path == '/add_script':
            self._handle_add_script()
        elif self.path == '/shutdown':
            self._handle_shutdown()
        else:
            self.send_error(404)

    def _handle_run(self):
        """
        로컬 스크립트를 무콘솔(Background) 및 전면 활성화 모드로 실행합니다.
        보안을 위해 현재 작업 디렉토리 내의 파일만 실행을 허용합니다.
        """
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        
        try:
            data = json.loads(post_data.decode('utf-8'))
            script_name = data.get('script')
            
            if not script_name:
                raise ValueError("실행할 스크립트 이름이 없습니다.")

            # 보안 검사 1: 상위 디렉토리 접근(Path Traversal) 및 절대 경로 차단
            if ".." in script_name or os.path.isabs(script_name):
                raise PermissionError("보안상 상위 경로 또는 절대 경로는 실행할 수 없습니다.")
            
            # 보안 검사 2: 커맨드 인젝션(Command Injection) 방어 (특수문자 제한)
            # 파일 이름에 흔히 쓰이지 않는 쉘 메타문자 차단
            forbidden_chars = ['&', '|', ';', '$', '>', '<', '`', '\\']
            if any(char in script_name for char in forbidden_chars):
                raise PermissionError("파일명에 허용되지 않는 특수문자가 포함되어 있습니다.")

            # [이식성 강화] 현재 서버 파일의 위치를 기준으로 절대 경로 산정
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            
            # 검색 순서 정의 (Smart Path Resolver)
            search_paths = [
                os.path.join(BASE_DIR, script_name),                     # 1. 루트
                os.path.join(BASE_DIR, "automated_scripts", script_name), # 2. 도구 폴더
                os.path.join(BASE_DIR, "system_guides", script_name)     # 3. 가이드 폴더
            ]
            
            file_path = None
            for p in search_paths:
                if os.path.exists(p):
                    file_path = p
                    break
            
            if not file_path:
                self._send_json({'success': False, 'message': f'파일을 찾을 수 없습니다: {script_name}'})
                return

            # 실제 파일이 위치한 디렉토리를 작업 디렉토리로 설정 (내부 상대경로 유지용)
            script_cwd = os.path.dirname(file_path)
            print(f"[-] Executing request: {file_path} (CWD: {script_cwd})")

            # Windows에서 CMD 창 생성을 봉쇄하면서도 프로그램은 전면에 띄우는 설정
            CREATE_NO_WINDOW = 0x08000000
            
            # 현재 실행 중인 파이썬 경로를 기반으로 'python.exe' 및 'pythonw.exe' 경로 정규화
            # dashboard가 pythonw로 실행되었더라도, 하위 스크립트는 터미널을 보여주기 위해 python.exe 선호
            current_exe = sys.executable
            if "pythonw.exe" in current_exe.lower():
                python_exe = current_exe.lower().replace("pythonw.exe", "python.exe")
                pythonw_exe = current_exe
            else:
                python_exe = current_exe
                pythonw_exe = current_exe.lower().replace("python.exe", "pythonw.exe")

            # 대체 경로 확인 (가상환경 등에서 파일이 없을 경우 대비)
            if not os.path.exists(python_exe): python_exe = "python"
            if not os.path.exists(pythonw_exe): pythonw_exe = "pythonw"

            if (script_name.endswith('.py')):
                # [v2.8] 자동화 도구가 화면 전면에 즉시 활성화 되도록 개선 (터미널 은닉 유지)
                try:
                    # GUI 스크립트의 경우 pythonw.exe를 선호하여 콘솔 없이 기동
                    use_exe = pythonw_exe if os.path.exists(pythonw_exe) else python_exe
                    print(f"[-] Foreground GUI Launch: {file_path} using {use_exe}")
                    
                    # STARTUPINFO를 사용하여 윈도우가 '최소화'가 아닌 '정상' 상태로 전면 부각되도록 강제
                    si = subprocess.STARTUPINFO()
                    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    si.wShowWindow = 1  # SW_SHOWNORMAL (정상 창 표시)
                    
                    creation_flags = 0
                    if use_exe == python_exe:
                        creation_flags = 0x08000000  # python.exe 사용 시에만 콘솔 숨김 플래그 적용
                    
                    subprocess.Popen([str(use_exe), str(file_path)], 
                                     creationflags=creation_flags, 
                                     startupinfo=si, 
                                     cwd=script_cwd)
                except Exception as e:
                    print(f"[-] Launch Error: {e}")
                    subprocess.Popen([str(python_exe), str(file_path)], creationflags=0x08000000, cwd=script_cwd)
                
            elif script_name.endswith('.ps1'):
                # PowerShell 역시 'start'를 통해 전면 활성화를 유도하며, 배경 실행 유지
                subprocess.Popen(f'powershell -WindowStyle Normal -File "{file_path}"', 
                                 shell=True,
                                 cwd=script_cwd)
            
            elif script_name.endswith(('.md', '.txt')):
                # [v2.7] 문서 자산(MD/TXT)은 기본 앱으로 전면 활성화를 유도하여 연다
                try:
                    subprocess.Popen(f'start "" "{file_path}"', shell=True, cwd=script_cwd)
                except Exception as e:
                    # 예외 발생 시에는 시스템 기본 호출 시도 (Fallback)
                    if hasattr(os, 'startfile'):
                        os.startfile(file_path)
                    else:
                        print(f"[-] Fallback error: os.startfile not available on this platform. {e}")

            self._send_json({'success': True, 'message': '백그라운드에서 실행을 시작했습니다.'})
            
        except Exception as e:
            self._send_json({'success': False, 'message': str(e)})

    def _handle_add_script(self):
        """
        대시보드 HTML 파일의 스크립트 배열 부분을 서버 사이드에서 직접 수정하여 반영합니다.
        (Self-Modifying Logic)
        """
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        
        try:
            entry = json.loads(post_data.decode('utf-8'))
            
            # 유효성 검사 (PRD 자산 명세 준수)
            required = ['name', 'title', 'desc', 'usage', 'type', 'icon']
            if not all(k in entry for k in required):
                raise ValueError("스크립트 등록을 위한 필수 항목이 누락되었습니다.")

            if not os.path.exists(HTML_FILE):
                 self._send_json({'success': False, 'message': f'{HTML_FILE} 파일을 찾을 수 없습니다.'})
                 return
                 
            with open(HTML_FILE, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            # 삽입 위치 탐색: scripts 배열의 끝(];)을 찾아 그 앞에 삽입
            marker = "const scripts = ["
            start_idx = html_content.find(marker)
            if start_idx == -1:
                raise ValueError("HTML 내 스크립트 목록 선언(const scripts)을 찾을 수 없습니다.")
            
            script_end_tag = html_content.find("</script>", start_idx)
            insert_pos = html_content.rfind("];", start_idx, script_end_tag)
            
            if insert_pos == -1:
                 raise ValueError("스크립트 배열의 닫는 괄호(];)를 찾을 수 없습니다.")

            # 새 항목 포맷팅
            new_item_str = f"""
            {{
                name: "{entry['name']}",
                title: "{entry['title']}",
                desc: "{entry['desc']}",
                usage: "{entry['usage']}",
                type: "{entry['type']}",
                icon: "{entry['icon']}"
            }},"""
            
            new_content = html_content[:insert_pos].rstrip()
            
            # Trailing Comma 안정을 위해 이전 항목 끝에 콤마가 없는 경우 추가
            if new_content[-1] != ',':
                new_content += ","
            
            final_content = new_content + new_item_str + "\n        " + html_content[insert_pos:]
            
            # 원본 보존을 위한 백업 생성 후 덮어쓰기
            backup_name = HTML_FILE + ".bak"
            with open(backup_name, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            with open(HTML_FILE, 'w', encoding='utf-8') as f:
                f.write(final_content)

            self._send_json({'success': True, 'message': '새 스크립트가 영구적으로 등록되었습니다.'})
            
        except Exception as e:
            print(f"Error adding script: {e}")
            self._send_json({'success': False, 'message': str(e)})

    def _handle_shutdown(self):
        """
        서버를 안전하게 종료합니다.
        대시보드에서 서버 제어를 위해 사용됩니다.
        """
        print("[-] Shutdown request received. Server will terminate.")
        self._send_json({'success': True, 'message': '서버가 종료됩니다.'})
        
        # 응답 전송 후 서버 종료
        def shutdown_server():
            time.sleep(1.0)  # 응답이 충분히 전송될 시간 확보
            # os._exit는 finally 및 atexit를 무시하므로 명시적 호출
            cleanup()
            os._exit(0)
        
        threading.Thread(target=shutdown_server).start()

    def _send_json(self, data):
        """데이터를 JSON 포맷으로 패키징하여 HTTP 응답으로 전송합니다."""
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        if self.path == '/heartbeat':
             # 하트비트는 캐시되지 않도록 헤더 추가
             self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))
        
    def log_message(self, format, *args):
        """불필요한 HTTP 서버 로그가 터미널에 남지 않도록 오버라이드합니다."""
        pass

def log_to_file(msg):
    """서버 로그를 파일에 기록합니다 (분석용)."""
    try:
        with open("server_log.txt", "a", encoding="utf-8") as f:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {msg}\n")
    except:
        pass

def cleanup():
    """[v2.5] 서버 종료 시 임시 로그 및 불필요한 파일/폴더를 전수 소거합니다."""
    # 삭제 대상 목록 (상대 경로)
    targets = [
        "server_log.txt",
        "final_v3_1.log",
        "out.log",
        "__pycache__",
        os.path.join("automated_scripts", "__pycache__")
    ]
    
    print(f"[-] Starting system cleanup...")
    for target in targets:
        try:
            if os.path.isdir(target):
                import shutil
                shutil.rmtree(target)
                print(f"[-] Directory cleaned: {target}")
            elif os.path.exists(target):
                os.remove(target)
                print(f"[-] File cleaned: {target}")
        except Exception as e:
            # 사용 중인 파일 등은 스킵
            pass
    print("[-] Cleanup sequence complete.")

# ─── 글로벌 상태: 마지막 하트비트 시간 ───
LAST_HEARTBEAT_TIME = time.time()
HEARTBEAT_TIMEOUT = 300  # 5분간 신호 없으면 종료 (안정성 위주 상향 조정)

def monitor_heartbeat():
    """주기적으로 하트비트 수신 여부를 확인하고, 타임아웃 시 서버를 종료합니다."""
    log_to_file("Auto-shutdown monitor started")
    
    while True:
        time.sleep(10)
        elapsed = time.time() - LAST_HEARTBEAT_TIME
        if elapsed > HEARTBEAT_TIMEOUT:
            log_to_file(f"Shutdown due to timeout ({int(elapsed)}s)")
            cleanup() # [v2.2] 자동 종료 시에도 로그 삭제
            os._exit(0)

class ThreadingHTTPServer(socketserver.ThreadingMixIn, socketserver.TCPServer):
    """멀티 스레드 HTTP 서버. 브라우저의 동시 요청(HTML, favicon, health 등)을 병렬 처리."""
    daemon_threads = True
    allow_reuse_address = True

def start_server():
    """드라이버 서버를 초기화하고 브라우저 대시보드를 실행합니다."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
    
    # [v2.2] 하트비트 모니터링 스레드 즉시 가동
    threading.Thread(target=monitor_heartbeat, daemon=True).start()
    
    log_to_file("Server starting process initiated")
    
    # 루프백 주소(127.0.0.1)를 우선적으로 사용하도록 함 (방화벽 및 로컬 접속 안정성 확보)
    server_address = ('127.0.0.1', PORT)
    
    try:
        # [v2.0] ThreadingHTTPServer: 브라우저 동시 요청을 병렬 처리하여 무한 로딩 방지
        with ThreadingHTTPServer(server_address, RequestHandler) as httpd:
            log_to_file(f"Dashboard running at http://127.0.0.1:{PORT}")
            print(f"[-] Dashboard running at http://127.0.0.1:{PORT}")
            print("[-] Multi-threaded server started. Use dashboard to stop.")
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[-] Shutdown request received (KeyboardInterrupt)")
    except Exception as e:
        log_to_file(f"Server crash error: {e}")
        # 만약 127.0.0.1 바인딩이 실패할 경우, 모든 인터페이스("" 혹은 0.0.0.0) 시도
        try:
            with ThreadingHTTPServer(("", PORT), RequestHandler) as httpd:
                log_to_file(f"Dashboard running at all interfaces on port {PORT}")
                print(f"[-] Dashboard running at port {PORT}")
                httpd.serve_forever()
        except Exception as e2:
            log_to_file(f"Fallback server crash error: {e2}")
            print(f"❌ 서버 기동 실패: {e2}")
    finally:
        cleanup() # [v2.2] 서버 프로세스 종료 시 최종 클린업

if __name__ == "__main__":
    # [v2.5] 종료 시 로그 파일 삭제 예약
    atexit.register(lambda: cleanup())
    start_server()
