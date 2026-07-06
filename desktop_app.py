from __future__ import annotations

import os
import sys
import threading
import time
import webbrowser

from app import create_server


APP_TITLE = "Handwrite Ledger AI"


def start_backend() -> tuple[object, str]:
    server = create_server(0)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, url


def run_browser_fallback(url: str) -> None:
    print(f"{APP_TITLE} 已启动：{url}")
    webbrowser.open(url)
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        return


def main() -> int:
    server, url = start_backend()
    try:
        try:
            import webview
        except ImportError:
            run_browser_fallback(url)
            return 0

        webview.create_window(
            APP_TITLE,
            url,
            width=1320,
            height=880,
            min_size=(980, 680),
        )
        webview.start()
        return 0
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    raise SystemExit(main())
