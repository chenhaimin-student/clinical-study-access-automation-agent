#!/usr/bin/env python3
"""Remote Mac watchdog for the Hermes Gateway and DingTalk Stream."""

import fcntl
import json
import os
import re
import subprocess
import time
from datetime import datetime
from pathlib import Path


UID = os.getuid()
GATEWAY_LABEL = f"gui/{UID}/ai.hermes.gateway"
HERMES_HOME = Path.home() / ".hermes"
ERROR_LOG = HERMES_HOME / "logs" / "gateway.error.log"
MONITOR_DIR = HERMES_HOME / "monitor"
STATE_FILE = MONITOR_DIR / "gateway-monitor-state.json"
LOCK_FILE = MONITOR_DIR / "gateway-monitor.lock"
MONITOR_LOG = HERMES_HOME / "logs" / "gateway-monitor.log"


def write_log(message: str) -> None:
    MONITOR_LOG.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with MONITOR_LOG.open("a", encoding="utf-8") as handle:
        handle.write(f"{stamp} {message}\n")


def run(command: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(command, text=True, capture_output=True, timeout=20)


def load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(state: dict) -> None:
    MONITOR_DIR.mkdir(parents=True, exist_ok=True)
    temporary = STATE_FILE.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    temporary.replace(STATE_FILE)


def gateway_status() -> tuple[bool, int | None, str]:
    result = run(["/bin/launchctl", "print", GATEWAY_LABEL])
    output = (result.stdout or "") + (result.stderr or "")
    state_match = re.search(r"\bstate = (\w+)", output)
    pid_match = re.search(r"\bpid = (\d+)", output)
    state = state_match.group(1) if state_match else "missing"
    pid = int(pid_match.group(1)) if pid_match else None
    return result.returncode == 0 and state == "running" and bool(pid), pid, state


def has_established_connection(pid: int) -> bool:
    result = run([
        "/usr/sbin/lsof", "-nP", "-a", "-p", str(pid),
        "-iTCP", "-sTCP:ESTABLISHED",
    ])
    return result.returncode == 0 and "TCP" in (result.stdout or "")


def read_new_errors(state: dict) -> str:
    if not ERROR_LOG.exists():
        state["error_offset"] = 0
        return ""
    size = ERROR_LOG.stat().st_size
    offset = int(state.get("error_offset", size))
    if offset < 0 or offset > size:
        offset = 0
    with ERROR_LOG.open("rb") as handle:
        handle.seek(offset)
        chunk = handle.read()
        state["error_offset"] = handle.tell()
    return chunk.decode("utf-8", errors="replace")


def restart_gateway(reason: str, state: dict) -> None:
    write_log(f"RESTART reason={reason}")
    result = run(["/bin/launchctl", "kickstart", "-k", GATEWAY_LABEL])
    state["last_restart"] = int(time.time())
    state["socket_failures"] = 0
    if result.returncode != 0:
        write_log(f"RESTART_FAILED code={result.returncode} error={(result.stderr or '').strip()[:300]}")
    else:
        write_log("RESTART_REQUESTED")


def main() -> None:
    MONITOR_DIR.mkdir(parents=True, exist_ok=True)
    with LOCK_FILE.open("a+") as lock:
        try:
            fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return

        state = load_state()
        running, pid, launch_state = gateway_status()
        new_errors = read_new_errors(state)

        if not running:
            restart_gateway(f"gateway_not_running state={launch_state}", state)
            save_state(state)
            return

        if "keepalive ping timeout" in new_errors.lower():
            restart_gateway("new_dingtalk_keepalive_timeout", state)
            save_state(state)
            return

        recently_restarted = time.time() - int(state.get("last_restart", 0)) < 45
        if not recently_restarted and pid is not None:
            if has_established_connection(pid):
                state["socket_failures"] = 0
            else:
                failures = int(state.get("socket_failures", 0)) + 1
                state["socket_failures"] = failures
                write_log(f"WARN no_established_socket pid={pid} count={failures}")
                if failures >= 2:
                    restart_gateway("no_established_tcp_socket_twice", state)
                    save_state(state)
                    return

        previous_pid = state.get("pid")
        state["pid"] = pid
        state["last_check"] = int(time.time())
        if previous_pid != pid:
            write_log(f"HEALTHY state=running pid={pid}")
        save_state(state)


if __name__ == "__main__":
    main()
