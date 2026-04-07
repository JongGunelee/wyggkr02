import os
import sys
import time
import threading
import urllib.request
import webbrowser

import run_dashboard


DEFAULT_WEB_URL = "https://jonggunelee.github.io/wyggkr02/"
WEB_URL_FILE = "웹접속 주소.txt"


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


def open_web_dashboard(url):
    try:
        webbrowser.open(url, new=2)
    except Exception:
        pass


def open_web_dashboard_when_ready(url, wait_seconds=20.0):
    deadline = time.time() + wait_seconds
    while time.time() < deadline:
        if is_agent_running(timeout=0.8):
            break
        time.sleep(0.5)
    open_web_dashboard(url)


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--run-script":
        run_dashboard.run_embedded_script(sys.argv[2])
        return

    run_dashboard.bootstrap_runtime_assets()
    target_url = load_web_url()

    if is_agent_running():
        open_web_dashboard(target_url)
        return

    threading.Thread(
        target=open_web_dashboard_when_ready,
        args=(target_url,),
        daemon=True,
    ).start()
    run_dashboard.start_server()


if __name__ == "__main__":
    main()
