"""Subprocess execution with cross-platform process-tree termination."""

from __future__ import annotations

import os
import signal
import subprocess
from collections.abc import Sequence


def _terminate_tree(process: subprocess.Popen[bytes]) -> None:
    if os.name == "nt":
        try:
            subprocess.run(  # noqa: S603  # nosec B603
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                capture_output=True,
                check=False,
                timeout=5,
            )
        except subprocess.SubprocessError:
            process.kill()
            return
    else:
        try:
            os.killpg(process.pid, signal.SIGKILL)  # type: ignore[attr-defined]
        except ProcessLookupError:
            return
    try:
        process.kill()
    except OSError:
        return


def run_process(
    command: Sequence[str], timeout: int | None
) -> subprocess.CompletedProcess[bytes]:
    """Run a command and kill its complete process tree on timeout."""
    if os.name == "nt":
        process = subprocess.Popen(  # noqa: S603
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )
    else:
        process = subprocess.Popen(  # noqa: S603
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
        )
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _terminate_tree(process)
        stdout, stderr = process.communicate()
        raise subprocess.TimeoutExpired(
            command, timeout or 0.0, stdout, stderr
        ) from None
    return subprocess.CompletedProcess(command, process.returncode, stdout, stderr)
