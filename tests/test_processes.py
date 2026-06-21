"""Process-tree timeout behavior."""

import subprocess

import pytest

from docx2pdf_py import processes


def test_timeout_terminates_process_tree(monkeypatch):
    calls = []

    class FakeProcess:
        pid = 4321
        returncode = None

        def __init__(self):
            self.attempt = 0
            self.killed = False

        def communicate(self, timeout=None):
            self.attempt += 1
            if self.attempt == 1:
                raise subprocess.TimeoutExpired(["command"], timeout)
            return b"out", b"err"

        def kill(self):
            self.killed = True

    fake = FakeProcess()
    monkeypatch.setattr(processes.subprocess, "Popen", lambda *args, **kwargs: fake)
    monkeypatch.setattr(
        processes.subprocess,
        "run",
        lambda command, **kwargs: calls.append(command),
    )
    if processes.os.name != "nt":
        monkeypatch.setattr(processes.os, "killpg", lambda *args: calls.append(args))

    with pytest.raises(subprocess.TimeoutExpired):
        processes.run_process(["command"], timeout=1)

    assert fake.killed
    if processes.os.name == "nt":
        assert calls and calls[0][:2] == ["taskkill", "/PID"]
    else:
        assert calls
