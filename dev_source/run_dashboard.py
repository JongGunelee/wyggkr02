import http.server
import socketserver
import json
import subprocess
import os
import threading
import time
import sys
import atexit
import shutil
import runpy
import hashlib
import urllib.request
import urllib.error
import urllib.parse

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
APP_HOME_NAME = "WYGGKR02_Dashboard_Agent"
GITHUB_REPO_OWNER = "jonggunelee"
GITHUB_REPO_NAME = "wyggkr02"
GITHUB_REPO_BRANCH = "main"
GITHUB_SCRIPTS_PATH = "automated_scripts"
GITHUB_RUNTIME_PATH = "dev_source/runtime_store"
REMOTE_SYNC_INTERVAL_SEC = 120
REMOTE_REQUEST_TIMEOUT_SEC = 12
PYTHON_INSTALL_TIMEOUT_SEC = 1800
PIP_INSTALL_TIMEOUT_SEC = 1800
REMOTE_SCRIPT_EXTENSIONS = (".py", ".ps1")
REMOTE_META_EXTENSIONS = (".md", ".txt")
REMOTE_ALLOWED_EXTENSIONS = REMOTE_SCRIPT_EXTENSIONS + REMOTE_META_EXTENSIONS
RUNTIME_TEMP_SUBDIR = os.path.join("dev_source", "__temp_runtime__")
REMOTE_SCRIPT_INDEX_FILE = os.path.join(RUNTIME_TEMP_SUBDIR, ".remote_scripts_index.json")
RUNTIME_MANIFEST_RELATIVE_PATH = os.path.join("runtime_store", "runtime_manifest.json")
RUNTIME_PACKAGE_CACHE_DIR = os.path.join("runtime_store", "packages")
REMOTE_SCRIPTS_API_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/"
    f"{GITHUB_SCRIPTS_PATH}?ref={GITHUB_REPO_BRANCH}"
)
REMOTE_SCRIPTS_RAW_BASE = (
    f"https://raw.githubusercontent.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/"
    f"{GITHUB_REPO_BRANCH}/{GITHUB_SCRIPTS_PATH}"
)
REMOTE_RUNTIME_MANIFEST_URL = (
    f"https://raw.githubusercontent.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/"
    f"{GITHUB_REPO_BRANCH}/{GITHUB_RUNTIME_PATH}/runtime_manifest.json"
)
REMOTE_RUNTIME_RAW_BASE = (
    f"https://raw.githubusercontent.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/"
    f"{GITHUB_REPO_BRANCH}/{GITHUB_RUNTIME_PATH}"
)
_REMOTE_SYNC_LOCK = threading.Lock()
_RUNTIME_MANIFEST_LOCK = threading.Lock()
_LAST_REMOTE_SYNC_TS = 0.0
_RUNTIME_MANIFEST_CACHE = None
_DIRECT_HTTP_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def get_source_root():
    """소스 실행 시 실제 프로젝트 루트를 추론합니다."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        script_dir,
        os.path.dirname(script_dir),
    ]
    for candidate in candidates:
        if os.path.exists(os.path.join(candidate, HTML_FILE)):
            return candidate
    return script_dir


def get_runtime_root():
    """번들 내부(또는 소스 루트)의 읽기 전용 자산 경로를 반환합니다."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_work_root():
    """실제 런타임이 쓰기를 수행할 작업 루트를 반환합니다."""
    if getattr(sys, "frozen", False):
        local_app_data = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(local_app_data, APP_HOME_NAME)
    return get_source_root()


RUNTIME_ROOT = get_runtime_root()
WORK_ROOT = get_work_root()
HTML_FILE_PATH = os.path.join(WORK_ROOT, HTML_FILE)


def ensure_parent_dir(path):
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def sync_file(src, dst):
    ensure_parent_dir(dst)
    if (not os.path.exists(dst)) or os.path.getmtime(src) > os.path.getmtime(dst):
        shutil.copy2(src, dst)


def sync_tree(src_dir, dst_dir):
    if not os.path.isdir(src_dir):
        return

    for root, dirs, files in os.walk(src_dir):
        rel_root = os.path.relpath(root, src_dir)
        target_root = dst_dir if rel_root == "." else os.path.join(dst_dir, rel_root)
        os.makedirs(target_root, exist_ok=True)

        for filename in files:
            src_file = os.path.join(root, filename)
            dst_file = os.path.join(target_root, filename)
            sync_file(src_file, dst_file)


def _http_get_json(url, timeout):
    request = urllib.request.Request(url, headers={"User-Agent": "WYGGKR02-Dashboard-Agent"})
    with _DIRECT_HTTP_OPENER.open(request, timeout=timeout) as response:
        payload = response.read().decode("utf-8")
    return json.loads(payload)


def _http_get_text(url, timeout):
    request = urllib.request.Request(url, headers={"User-Agent": "WYGGKR02-Dashboard-Agent"})
    with _DIRECT_HTTP_OPENER.open(request, timeout=timeout) as response:
        return response.read().decode("utf-8")


def _download_remote_file(url, destination_path, timeout):
    request = urllib.request.Request(url, headers={"User-Agent": "WYGGKR02-Dashboard-Agent"})
    with _DIRECT_HTTP_OPENER.open(request, timeout=timeout) as response:
        data = response.read()
    ensure_parent_dir(destination_path)
    with open(destination_path, "wb") as file_obj:
        file_obj.write(data)


def _sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as file_obj:
        for chunk in iter(lambda: file_obj.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _source_runtime_store_path(*parts):
    return os.path.join(get_source_root(), "dev_source", "runtime_store", *parts)


def _load_local_runtime_manifest():
    manifest_path = resolve_work_path(RUNTIME_MANIFEST_RELATIVE_PATH)
    if not os.path.exists(manifest_path):
        return None
    try:
        with open(manifest_path, "r", encoding="utf-8") as file_obj:
            payload = json.load(file_obj)
        if isinstance(payload, dict):
            return payload
    except Exception:
        pass
    return None


def _save_local_runtime_manifest(payload):
    manifest_path = resolve_work_path(RUNTIME_MANIFEST_RELATIVE_PATH)
    ensure_parent_dir(manifest_path)
    with open(manifest_path, "w", encoding="utf-8") as file_obj:
        json.dump(payload, file_obj, ensure_ascii=False, indent=2)


def load_runtime_manifest(force_remote=False):
    global _RUNTIME_MANIFEST_CACHE
    if _RUNTIME_MANIFEST_CACHE is not None and not force_remote:
        return _RUNTIME_MANIFEST_CACHE

    with _RUNTIME_MANIFEST_LOCK:
        if _RUNTIME_MANIFEST_CACHE is not None and not force_remote:
            return _RUNTIME_MANIFEST_CACHE

        manifest = None
        if not force_remote:
            manifest = _load_local_runtime_manifest()
            if manifest is not None:
                _RUNTIME_MANIFEST_CACHE = manifest
                return manifest

        try:
            manifest = json.loads(_http_get_text(REMOTE_RUNTIME_MANIFEST_URL, REMOTE_REQUEST_TIMEOUT_SEC))
            if isinstance(manifest, dict):
                _save_local_runtime_manifest(manifest)
                _RUNTIME_MANIFEST_CACHE = manifest
                return manifest
        except Exception:
            pass

        manifest = _load_local_runtime_manifest()
        if manifest is None:
            source_manifest_path = _source_runtime_store_path("runtime_manifest.json")
            if os.path.exists(source_manifest_path):
                with open(source_manifest_path, "r", encoding="utf-8") as file_obj:
                    manifest = json.load(file_obj)

        if not isinstance(manifest, dict):
            raise RuntimeError("런타임 매니페스트를 불러올 수 없습니다.")

        _RUNTIME_MANIFEST_CACHE = manifest
        return manifest


def _runtime_remote_url(relative_path):
    safe_parts = [urllib.parse.quote(part) for part in relative_path.replace("\\", "/").split("/") if part]
    return f"{REMOTE_RUNTIME_RAW_BASE}/{'/'.join(safe_parts)}"


def _ensure_runtime_asset(relative_path, destination_path, sha256_hash=""):
    if os.path.exists(destination_path):
        if (not sha256_hash) or (_sha256_file(destination_path).lower() == sha256_hash.lower()):
            return destination_path

    source_candidate = _source_runtime_store_path(*relative_path.replace("\\", "/").split("/"))
    if os.path.exists(source_candidate):
        ensure_parent_dir(destination_path)
        shutil.copy2(source_candidate, destination_path)
    else:
        _download_remote_file(_runtime_remote_url(relative_path), destination_path, REMOTE_REQUEST_TIMEOUT_SEC)

    if sha256_hash and _sha256_file(destination_path).lower() != sha256_hash.lower():
        raise RuntimeError(f"런타임 자산 해시 검증 실패: {relative_path}")

    return destination_path


def ensure_python_runtime():
    if not getattr(sys, "frozen", False):
        return sys.executable

    manifest = load_runtime_manifest()
    python_meta = manifest.get("python", {})
    required_prefix = str(python_meta.get("major_minor", "3.13"))
    python_dir = os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
        "Programs",
        "Python",
        f"Python{required_prefix.replace('.', '')}",
    )
    python_exe = os.path.join(python_dir, "python.exe")

    def python_matches(exe_path):
        if not os.path.exists(exe_path):
            return False
        try:
            proc = subprocess.run(
                [exe_path, "-c", "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"],
                capture_output=True,
                text=True,
                timeout=15,
                check=False,
            )
            return proc.returncode == 0 and proc.stdout.strip().startswith(required_prefix)
        except Exception:
            return False

    if python_matches(python_exe):
        return python_exe

    install_url = python_meta.get("install_url", "")
    install_args = python_meta.get(
        "install_args",
        ["/quiet", "InstallAllUsers=0", "PrependPath=0", "Include_pip=1", "Include_launcher=0", "Shortcuts=0"],
    )
    if not install_url:
        raise RuntimeError("Python 설치 URL이 런타임 매니페스트에 없습니다.")

    installer_path = resolve_work_path(RUNTIME_TEMP_SUBDIR, "python-installer.exe")
    _download_remote_file(install_url, installer_path, PYTHON_INSTALL_TIMEOUT_SEC)
    subprocess.run([installer_path] + list(install_args), check=True, timeout=PYTHON_INSTALL_TIMEOUT_SEC)

    if not python_matches(python_exe):
        raise RuntimeError("Python 런타임 설치 후 python.exe를 찾지 못했습니다.")

    return python_exe


def _python_imports_available(python_exe, imports_to_check):
    if not imports_to_check:
        return True

    probe_code = "\n".join(
        [
            "import importlib, json, sys",
            f"mods = {json.dumps(sorted(set(imports_to_check)), ensure_ascii=True)}",
            "missing = []",
            "for name in mods:",
            "    try:",
            "        importlib.import_module(name)",
            "    except Exception:",
            "        missing.append(name)",
            "print(json.dumps(missing))",
            "sys.exit(0 if not missing else 1)",
        ]
    )
    proc = subprocess.run(
        [python_exe, "-c", probe_code],
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
    )
    return proc.returncode == 0


def _ensure_runtime_package_cached(package_id, package_meta):
    cache_name = package_meta.get("cache_name") or os.path.basename(package_meta.get("file", ""))
    if not cache_name:
        raise RuntimeError(f"패키지 파일명이 없습니다: {package_id}")

    target_path = resolve_work_path(RUNTIME_PACKAGE_CACHE_DIR, cache_name)
    expected_hash = package_meta.get("sha256", "")

    if "parts" in package_meta:
        if os.path.exists(target_path) and ((not expected_hash) or _sha256_file(target_path).lower() == expected_hash.lower()):
            return target_path

        ensure_parent_dir(target_path)
        temp_parts = []
        try:
            with open(target_path, "wb") as output_file:
                for part_meta in package_meta.get("parts", []):
                    part_path = resolve_work_path(RUNTIME_PACKAGE_CACHE_DIR, os.path.basename(part_meta["path"]))
                    _ensure_runtime_asset(part_meta["path"], part_path, part_meta.get("sha256", ""))
                    temp_parts.append(part_path)
                    with open(part_path, "rb") as part_file:
                        shutil.copyfileobj(part_file, output_file)
        finally:
            for part_path in temp_parts:
                try:
                    if os.path.exists(part_path):
                        os.remove(part_path)
                except OSError:
                    pass
    else:
        _ensure_runtime_asset(package_meta["file"], target_path, expected_hash)

    if expected_hash and _sha256_file(target_path).lower() != expected_hash.lower():
        raise RuntimeError(f"패키지 해시 검증 실패: {package_id}")

    return target_path


def ensure_script_runtime(script_name, python_exe=None):
    manifest = load_runtime_manifest()
    script_meta = manifest.get("scripts", {}).get(script_name, {})
    package_ids = script_meta.get("packages", [])
    if not package_ids:
        return True

    python_exe = python_exe or sys.executable
    package_table = manifest.get("packages", {})
    imports_to_check = []
    wheel_paths = []

    for package_id in package_ids:
        package_meta = package_table.get(package_id)
        if not isinstance(package_meta, dict):
            raise RuntimeError(f"패키지 정의가 없습니다: {package_id}")
        imports_to_check.extend(package_meta.get("imports", []))

    if _python_imports_available(python_exe, imports_to_check):
        return True

    for package_id in package_ids:
        wheel_paths.append(_ensure_runtime_package_cached(package_id, package_table[package_id]))

    subprocess.run(
        [python_exe, "-m", "pip", "install", "--upgrade", "--disable-pip-version-check", "--no-warn-script-location"]
        + wheel_paths,
        check=True,
        timeout=PIP_INSTALL_TIMEOUT_SEC,
    )

    if not _python_imports_available(python_exe, imports_to_check):
        raise RuntimeError(f"필수 런타임 패키지 설치 후에도 import 검증에 실패했습니다: {script_name}")

    return True


def _load_remote_index():
    index_path = resolve_work_path(REMOTE_SCRIPT_INDEX_FILE)
    if not os.path.exists(index_path):
        return {}
    try:
        with open(index_path, "r", encoding="utf-8") as file_obj:
            payload = json.load(file_obj)
        if isinstance(payload, dict):
            return payload
    except Exception:
        pass
    return {}


def _save_remote_index(index_data):
    index_path = resolve_work_path(REMOTE_SCRIPT_INDEX_FILE)
    ensure_parent_dir(index_path)
    with open(index_path, "w", encoding="utf-8") as file_obj:
        json.dump(index_data, file_obj, ensure_ascii=False, indent=2)


def _list_remote_script_entries():
    payload = _http_get_json(REMOTE_SCRIPTS_API_URL, REMOTE_REQUEST_TIMEOUT_SEC)
    if not isinstance(payload, list):
        raise ValueError("원격 스크립트 목록 응답 형식이 올바르지 않습니다.")

    entries = []
    for item in payload:
        if item.get("type") != "file":
            continue
        name = item.get("name", "")
        lowered = name.lower()
        if not lowered.endswith(REMOTE_ALLOWED_EXTENSIONS):
            continue

        download_url = item.get("download_url")
        if not download_url:
            encoded = urllib.parse.quote(name)
            download_url = f"{REMOTE_SCRIPTS_RAW_BASE}/{encoded}"

        entries.append(
            {
                "name": name,
                "sha": item.get("sha", ""),
                "download_url": download_url,
            }
        )
    return entries


def sync_remote_automated_scripts(force=False):
    """
    GitHub의 automated_scripts 최신본을 WORK_ROOT로 동기화합니다.
    - force=False: SHA 변경분/누락분만 다운로드
    - force=True: 전수 재다운로드
    """
    target_dir = resolve_work_path("automated_scripts")
    os.makedirs(target_dir, exist_ok=True)

    cached_index = _load_remote_index()
    remote_entries = _list_remote_script_entries()
    next_index = {}
    updated_files = []

    for entry in remote_entries:
        name = entry["name"]
        sha = entry.get("sha", "")
        destination = os.path.join(target_dir, name)

        previous_sha = cached_index.get(name, "")
        should_download = force or (not os.path.exists(destination))
        if sha:
            should_download = should_download or (previous_sha != sha)

        if should_download:
            _download_remote_file(entry["download_url"], destination, REMOTE_REQUEST_TIMEOUT_SEC)
            updated_files.append(name)

        next_index[name] = sha or previous_sha

    _save_remote_index(next_index)
    return True, updated_files


def maybe_sync_remote_automated_scripts(force=False):
    global _LAST_REMOTE_SYNC_TS
    now = time.time()
    if (not force) and (now - _LAST_REMOTE_SYNC_TS < REMOTE_SYNC_INTERVAL_SEC):
        return True, []

    with _REMOTE_SYNC_LOCK:
        now = time.time()
        if (not force) and (now - _LAST_REMOTE_SYNC_TS < REMOTE_SYNC_INTERVAL_SEC):
            return True, []

        try:
            ok, updated = sync_remote_automated_scripts(force=force)
            _LAST_REMOTE_SYNC_TS = time.time()
            if updated:
                print(f"[-] Remote scripts synced ({len(updated)} files updated)")
            return ok, updated
        except Exception as exc:
            print(f"[!] Remote script sync skipped: {exc}")
            return False, []


def download_single_remote_script(script_name):
    safe_name = os.path.basename(script_name)
    if safe_name != script_name:
        return None

    lowered = safe_name.lower()
    if not lowered.endswith(REMOTE_SCRIPT_EXTENSIONS):
        return None

    target_path = resolve_work_path("automated_scripts", safe_name)
    encoded_name = urllib.parse.quote(safe_name)
    url = f"{REMOTE_SCRIPTS_RAW_BASE}/{encoded_name}"

    try:
        _download_remote_file(url, target_path, REMOTE_REQUEST_TIMEOUT_SEC)
        return target_path
    except Exception as exc:
        print(f"[!] Remote single script download failed ({safe_name}): {exc}")
        return None


def bootstrap_runtime_assets():
    """
    PyInstaller onefile 환경에서도 대시보드/스크립트 자산을
    사용자 쓰기 가능 루트로 동기화합니다.
    """
    if not getattr(sys, "frozen", False):
        return

    os.makedirs(WORK_ROOT, exist_ok=True)

    single_files = [
        "00 dashboard.html",
        "000 Launch_dashboard.bat",
        "웹접속 주소.txt",
        "index.html",
        "manifest.webmanifest",
        "service-worker.js",
        os.path.join("dev_source", "run_dashboard.py"),
        os.path.join("runtime_store", "runtime_manifest.json"),
    ]
    for name in single_files:
        src = os.path.join(RUNTIME_ROOT, name)
        dst = os.path.join(WORK_ROOT, name)
        if os.path.exists(src):
            sync_file(src, dst)

    # automated_scripts는 GitHub 원격 동기화를 우선하되, 오프라인 대응을 위해
    # 번들에 존재할 경우 로컬 fallback 사본도 유지합니다.
    sync_tree(os.path.join(RUNTIME_ROOT, "automated_scripts"), os.path.join(WORK_ROOT, "automated_scripts"))
    sync_tree(os.path.join(RUNTIME_ROOT, "system_guides"), os.path.join(WORK_ROOT, "system_guides"))


def resolve_work_path(*parts):
    return os.path.join(WORK_ROOT, *parts)


def resolve_safe_work_path(relative_path):
    normalized = os.path.normpath(relative_path)
    candidate = os.path.abspath(os.path.join(WORK_ROOT, normalized))
    work_root_abs = os.path.abspath(WORK_ROOT)

    if os.path.commonpath([candidate, work_root_abs]) != work_root_abs:
        raise PermissionError("보안상 작업 루트 밖의 경로는 접근할 수 없습니다.")
    return candidate


def run_embedded_script(relative_path):
    """
    패키징된 EXE가 자기 자신을 다시 호출했을 때,
    내장된 Python 스크립트를 사용자 작업 루트 기준으로 실행합니다.
    """
    bootstrap_runtime_assets()
    target_path = resolve_safe_work_path(relative_path)
    if not os.path.exists(target_path):
        raise FileNotFoundError(f"내장 스크립트를 찾을 수 없습니다: {relative_path}")

    script_dir = os.path.dirname(target_path)
    previous_cwd = os.getcwd()
    previous_argv = list(sys.argv)
    try:
        os.chdir(script_dir)
        sys.argv = [target_path]
        runpy.run_path(target_path, run_name="__main__")
    finally:
        sys.argv = previous_argv
        os.chdir(previous_cwd)

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
            self.path = os.path.basename(HTML_FILE_PATH)
        
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
            BASE_DIR = WORK_ROOT
            
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

            # 자동화 스크립트(.py/.ps1)는 GitHub 원격 동기화 경로를 추가 시도
            if (not file_path) and script_name.lower().endswith(REMOTE_SCRIPT_EXTENSIONS):
                maybe_sync_remote_automated_scripts(force=False)
                for p in search_paths:
                    if os.path.exists(p):
                        file_path = p
                        break

            # 그래도 없으면 단건 다운로드 시도 (최소 회복 경로)
            if (not file_path) and script_name.lower().endswith(REMOTE_SCRIPT_EXTENSIONS):
                downloaded = download_single_remote_script(script_name)
                if downloaded and os.path.exists(downloaded):
                    file_path = downloaded
            
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
            current_exe = ensure_python_runtime() if getattr(sys, "frozen", False) else sys.executable
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
                    ensure_script_runtime(script_name, python_exe=python_exe)
                    use_exe = pythonw_exe if os.path.exists(pythonw_exe) else python_exe
                    launch_args = [str(use_exe), str(file_path)]
                    print(f"[-] Foreground GUI Launch: {file_path} using {use_exe}")
                    
                    # STARTUPINFO를 사용하여 윈도우가 '최소화'가 아닌 '정상' 상태로 전면 부각되도록 강제
                    si = subprocess.STARTUPINFO()
                    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    si.wShowWindow = 1  # SW_SHOWNORMAL (정상 창 표시)
                    
                    creation_flags = 0
                    if (not getattr(sys, "frozen", False)) and use_exe == python_exe:
                        creation_flags = 0x08000000  # python.exe 사용 시에만 콘솔 숨김 플래그 적용
                    
                    subprocess.Popen(launch_args,
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

            if not os.path.exists(HTML_FILE_PATH):
                 self._send_json({'success': False, 'message': f'{HTML_FILE_PATH} 파일을 찾을 수 없습니다.'})
                 return
                 
            with open(HTML_FILE_PATH, 'r', encoding='utf-8') as f:
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
            backup_name = HTML_FILE_PATH + ".bak"
            with open(backup_name, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            with open(HTML_FILE_PATH, 'w', encoding='utf-8') as f:
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
        log_path = resolve_work_path("server_log.txt")
        ensure_parent_dir(log_path)
        with open(log_path, "a", encoding="utf-8") as f:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {msg}\n")
    except:
        pass

def cleanup():
    """[v2.5] 서버 종료 시 임시 로그 및 불필요한 파일/폴더를 전수 소거합니다."""
    # 삭제 대상 목록 (상대 경로)
    targets = [
        resolve_work_path("server_log.txt"),
        resolve_work_path("final_v3_1.log"),
        resolve_work_path("out.log"),
        resolve_work_path("__pycache__"),
        resolve_work_path("automated_scripts", "__pycache__")
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
    bootstrap_runtime_assets()
    os.chdir(WORK_ROOT)
    maybe_sync_remote_automated_scripts(force=False)
    
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
