"""
launcher.py — 打包入口點
負責：啟動 Streamlit server、等待就緒、自動開啟瀏覽器
"""
import os
import sys
import time
import socket
import threading
import webbrowser
import subprocess

PORT = 8501
URL = f"http://localhost:{PORT}"


def is_port_open(port: int, timeout: float = 0.5) -> bool:
    """檢查 port 是否已在監聽（即 server 是否就緒）。"""
    try:
        with socket.create_connection(("localhost", port), timeout=timeout):
            return True
    except OSError:
        return False


def wait_and_open_browser(max_wait: int = 30) -> None:
    """背景執行緒：等 server 就緒後開瀏覽器，最多等 max_wait 秒。"""
    for _ in range(max_wait * 2):       # 每 0.5 秒檢查一次
        if is_port_open(PORT):
            webbrowser.open(URL)
            return
        time.sleep(0.5)
    webbrowser.open(URL)  # 超時仍嘗試開啟


def get_app_path() -> str:
    """取得 app.py 的絕對路徑（相容 PyInstaller 打包後的路徑）。"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "app", "app.py")


def main() -> None:
    app_path = get_app_path()

    if not os.path.exists(app_path):
        input(f"[錯誤] 找不到 app.py：{app_path}\n按 Enter 關閉...")
        sys.exit(1)

    threading.Thread(target=wait_and_open_browser, daemon=True).start()

    try:
        subprocess.run(
            [
                sys.executable, "-m", "streamlit", "run", app_path,
                "--server.port", str(PORT),
                "--server.headless", "true",
                "--browser.gatherUsageStats", "false",
            ],
            check=True,
        )
    except KeyboardInterrupt:
        pass
    except subprocess.CalledProcessError as e:
        input(f"[錯誤] Streamlit 啟動失敗（exit code {e.returncode}）\n按 Enter 關閉...")


if __name__ == "__main__":
    main()
