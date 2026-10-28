"""Start the app and verify it serves the page and static assets.

This script uses only the standard library so it can run unchanged on
Windows, macOS, and Linux. It is used by the GitHub Actions matrix.
"""

from __future__ import annotations

import os
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASSETS = (
    "/app/static/word-ui.js",
    "/app/static/bliss.jpg",
    "/app/static/clippy.png",
)
STARTUP_TIMEOUT = 90


def free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def fetch(url: str) -> tuple[int, bytes]:
    with urllib.request.urlopen(url, timeout=10) as response:
        return response.status, response.read()


def wait_for(url: str) -> None:
    deadline = time.monotonic() + STARTUP_TIMEOUT
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            status, _ = fetch(url)
            if status == 200:
                return
        except Exception as exc:  # Connection is expected to fail while booting.
            last_error = exc
        time.sleep(1)
    raise RuntimeError(f"Timed out waiting for {url}: {last_error}")


def main() -> int:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

    print("Importing app.py")
    imported = subprocess.run(
        [sys.executable, "-c", "import app"],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
        timeout=120,
    )
    if imported.returncode != 0:
        print(imported.stdout)
        print(imported.stderr)
        return imported.returncode

    port = free_port()
    base = f"http://127.0.0.1:{port}"
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "app.py",
        f"--server.port={port}",
        "--server.address=127.0.0.1",
        "--server.headless=true",
    ]

    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as log:
        print("Starting:", " ".join(command))
        process = subprocess.Popen(
            command,
            cwd=ROOT,
            env=env,
            stdout=log,
            stderr=subprocess.STDOUT,
        )
        failure: Exception | None = None
        try:
            wait_for(f"{base}/_stcore/health")
            print("Health endpoint: 200")
            for asset in ASSETS:
                status, body = fetch(base + asset)
                if status != 200 or not body:
                    raise RuntimeError(f"{asset} returned {status} with {len(body)} bytes")
                print(f"{asset}: 200, {len(body)} bytes")
            print("SMOKE_OK")
        except Exception as exc:
            failure = exc
        finally:
            process.terminate()
            try:
                process.wait(timeout=15)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=10)

        if failure is not None:
            log.seek(0)
            print("\n--- Streamlit output ---")
            print(log.read())
            print("--- end output ---")
            raise failure

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
