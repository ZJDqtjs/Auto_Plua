from __future__ import annotations

import shlex
import subprocess
import threading
from pathlib import Path
from typing import Callable

import psutil

from autoplua.models import ManagedProgram


class ProcessService:
    def __init__(self) -> None:
        self._children: dict[str, subprocess.Popen] = {}
        self._output_listener: Callable[[str, str], None] | None = None
        self._exit_listener: Callable[[str, int | None], None] | None = None

    def set_output_listener(self, listener: Callable[[str, str], None] | None) -> None:
        self._output_listener = listener

    def set_exit_listener(self, listener: Callable[[str, int | None], None] | None) -> None:
        self._exit_listener = listener

    def start(self, program: ManagedProgram) -> None:
        if program.name in self._children and self._children[program.name].poll() is None:
            return

        cmd = [program.command, *program.args]
        cwd = program.cwd if program.cwd else None
        process = subprocess.Popen(
            cmd,
            cwd=cwd,
            shell=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        self._children[program.name] = process

        cmdline = " ".join(cmd)
        self._emit_output(program.name, f"[进程启动] pid={process.pid} cmd={cmdline}")

        monitor = threading.Thread(
            target=self._monitor_process_output,
            args=(program.name, process),
            daemon=True,
        )
        monitor.start()

    def stop(self, program_name: str) -> bool:
        process = self._children.get(program_name)
        if process and process.poll() is None:
            process.terminate()
            self._emit_output(program_name, "[进程终止] 已发送 terminate 信号")
            return True

        for proc in psutil.process_iter(["name", "pid", "cmdline"]):
            try:
                cmdline = " ".join(proc.info.get("cmdline") or [])
                if program_name.lower() in cmdline.lower() or program_name.lower() in (proc.info.get("name") or "").lower():
                    proc.terminate()
                    self._emit_output(program_name, f"[进程终止] 已终止匹配进程 pid={proc.pid}")
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False

    def restart(self, program: ManagedProgram) -> None:
        self.stop(program.name)
        self.start(program)

    def is_running(self, program_name: str) -> bool:
        process = self._children.get(program_name)
        if process and process.poll() is None:
            return True

        for proc in psutil.process_iter(["name", "cmdline"]):
            try:
                name = (proc.info.get("name") or "").lower()
                cmdline = " ".join(proc.info.get("cmdline") or []).lower()
                target = program_name.lower()
                if target in name or target in cmdline:
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False

    @staticmethod
    def parse_command(raw: str) -> tuple[str, list[str]]:
        parts = shlex.split(raw)
        if not parts:
            raise ValueError("命令不能为空")
        command = parts[0]
        args = parts[1:]
        if Path(command).suffix and not Path(command).exists() and "\\" in command:
            raise ValueError(f"找不到可执行文件: {command}")
        return command, args

    def _emit_output(self, program_name: str, line: str) -> None:
        if self._output_listener is None:
            return
        self._output_listener(program_name, line)

    def _emit_exit(self, program_name: str, code: int | None) -> None:
        if self._exit_listener is None:
            return
        self._exit_listener(program_name, code)

    def _monitor_process_output(self, program_name: str, process: subprocess.Popen) -> None:
        try:
            if process.stdout is not None:
                for raw_line in iter(process.stdout.readline, ""):
                    line = raw_line.rstrip("\r\n")
                    if line:
                        self._emit_output(program_name, line)
        finally:
            try:
                code = process.wait(timeout=1)
            except subprocess.TimeoutExpired:
                code = process.poll()
            self._emit_exit(program_name, code)
