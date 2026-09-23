"""Windows MCP launcher: start Rancher on demand, then reuse the existing container.

No build, secret provisioning or certificate changes here. The host adapter reads
only explicitly requested local image files; document writes remain in Word.
Only MCP frames reach stdout. The client must allow 210 seconds for a cold start.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

CONTAINER = "word-ai-secure"


def run(command, timeout=10):
    return subprocess.run(command, stdin=subprocess.DEVNULL, capture_output=True,
                          text=True, timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW)


def engine_ready(docker):
    try:
        result = run([docker, "version", "--format", "{{.Server.Version}}"], timeout=8)
        return result.returncode == 0 and bool(result.stdout.strip())
    except subprocess.TimeoutExpired:
        return False


def start_rancher():
    rancher = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Rancher Desktop/Rancher Desktop.exe"
    if not rancher.is_file():
        raise RuntimeError("Rancher Desktop is not installed at its expected path")
    processes = run(["tasklist.exe", "/FI", "IMAGENAME eq Rancher Desktop.exe", "/FO", "CSV", "/NH"])
    if "rancher desktop.exe" in processes.stdout.lower():
        return
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = subprocess.SW_HIDE
    subprocess.Popen([str(rancher)], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                     stderr=subprocess.DEVNULL, startupinfo=startup,
                     # Rancher is a user-owned desktop service, not an MCP child
                     # to tear down when Codex replaces/cancels one MCP client.
                     # Windows must permit breakaway; never alter job limits or
                     # security settings to force it if the OS refuses.
                     creationflags=subprocess.CREATE_NO_WINDOW | subprocess.CREATE_BREAKAWAY_FROM_JOB)


def ensure_ready(docker, timeout=180):
    deadline = time.monotonic() + timeout
    if not engine_ready(docker):
        start_rancher()
        while not engine_ready(docker):
            if time.monotonic() >= deadline:
                raise RuntimeError("Docker did not become ready within 180 seconds; inspect Rancher Desktop")
            time.sleep(3)
    found = run([docker, "inspect", "--type", "container", CONTAINER])
    if found.returncode:
        raise RuntimeError("Existing word-ai-secure container not found; run Start-SecureWordAi.ps1 first")
    container = json.loads(found.stdout)[0]
    labels = container.get("Config", {}).get("Labels") or {}
    if labels.get("com.docker.compose.project") != "word-ai-secure" or labels.get("com.docker.compose.service") != "word-ai":
        raise RuntimeError("Container name belongs to an unexpected deployment; refusing to attach")
    if not container["State"]["Running"]:
        started = run([docker, "start", CONTAINER], timeout=20)
        if started.returncode:
            raise RuntimeError("Could not start the existing Word AI container")
    probe = "import socket; s=socket.socket(socket.AF_UNIX); s.settimeout(2); s.connect('/run/word-ai/mcp.sock'); s.close()"
    while True:
        try:
            ready = run([docker, "exec", CONTAINER, "python", "-c", probe], timeout=6)
            if ready.returncode == 0:
                return
        except subprocess.TimeoutExpired:
            pass
        if time.monotonic() >= deadline:
            raise RuntimeError("Word AI socket did not become ready; inspect the container logs")
        time.sleep(1)


def main():
    if sys.platform != "win32":
        raise RuntimeError("This launcher is for Windows with Rancher Desktop")
    docker = shutil.which("docker")
    if not docker:
        raise RuntimeError("Docker CLI is missing from PATH")
    ensure_ready(docker)
    from image_mcp_adapter import relay
    return relay([docker, "exec", "-i", CONTAINER, "python", "/app/secure_live.py", "stdio"])


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, ValueError, RuntimeError, subprocess.TimeoutExpired) as error:
        print(f"Word AI startup failed: {error}", file=sys.stderr)
        sys.exit(1)
