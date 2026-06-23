from __future__ import annotations

import socket
import threading
import time
import traceback
import webbrowser
from pathlib import Path

import uvicorn


HOST = "127.0.0.1"
PORT = 8787


def _log_path() -> Path:
    import os

    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or str(Path.home())
    path = Path(base) / "PIPLYiDianTeng"
    path.mkdir(parents=True, exist_ok=True)
    return path / "launcher.log"


def _write_log(message: str) -> None:
    with _log_path().open("a", encoding="utf-8") as file:
        file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")


def _port_is_free(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((host, port)) != 0


def _open_when_ready(url: str) -> None:
    for _ in range(60):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            if sock.connect_ex((HOST, PORT)) == 0:
                webbrowser.open(url)
                return
        time.sleep(0.5)


def main() -> None:
    try:
        url = f"http://{HOST}:{PORT}/"
        _write_log("launcher starting")
        if not _port_is_free(HOST, PORT):
            _write_log("port already in use, opening existing app")
            webbrowser.open(url)
            return

        from trade_docs.api_server import app

        threading.Thread(target=_open_when_ready, args=(url,), daemon=True).start()
        _write_log("starting uvicorn")
        uvicorn.run(
            app,
            host=HOST,
            port=PORT,
            log_level="warning",
            access_log=False,
            log_config=None,
        )
    except Exception:  # noqa: BLE001
        _write_log(traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
