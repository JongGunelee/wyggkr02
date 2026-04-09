import os
import sys
import urllib.request
import ctypes
import subprocess

import run_dashboard


DEFAULT_WEB_URL = "https://jonggunelee.github.io/wyggkr02/"
WEB_URL_FILE = "웹접속 주소.txt"
SINGLE_INSTANCE_MUTEX = "Global\\WYGGKR02_Dashboard_Agent_Launcher"
ERROR_ALREADY_EXISTS = 183
_MUTEX_HANDLE = None
CREATE_NO_WINDOW = 0x08000000


def acquire_single_instance():
    global _MUTEX_HANDLE
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.CreateMutexW(None, False, SINGLE_INSTANCE_MUTEX)
    if not handle:
        return True

    _MUTEX_HANDLE = handle
    if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        return False
    return True


def load_web_url():
    candidates = [
        os.path.join(run_dashboard.WORK_ROOT, WEB_URL_FILE),
        os.path.join(run_dashboard.RUNTIME_ROOT, WEB_URL_FILE),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), WEB_URL_FILE),
    ]

    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as file_obj:
                    value = file_obj.read().strip()
                if value:
                    return value
            except OSError:
                pass
    return DEFAULT_WEB_URL


def is_agent_running(timeout=1.0):
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{run_dashboard.PORT}/health",
            timeout=timeout,
        ) as response:
            return response.status == 200
    except Exception:
        return False


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--run-script":
        run_dashboard.run_embedded_script(sys.argv[2])
        return

    if not acquire_single_instance():
        return

    run_dashboard.bootstrap_runtime_assets()

    if is_agent_running():
        return

    python_exe = run_dashboard.ensure_python_runtime()
    server_script = os.path.join(run_dashboard.WORK_ROOT, "dev_source", "run_dashboard.py")
    if not os.path.exists(server_script):
        raise FileNotFoundError(f"런타임 서버 스크립트를 찾을 수 없습니다: {server_script}")

    subprocess.Popen(
        [python_exe, server_script],
        cwd=run_dashboard.WORK_ROOT,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", CREATE_NO_WINDOW),
        close_fds=False,
    )


if __name__ == "__main__":
    main()
