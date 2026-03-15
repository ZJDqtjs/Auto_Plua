from __future__ import annotations

import os
import locale
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
            text=False,
            bufsize=0,
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

    def stop(self, program_name: str, command: str | None = None) -> bool:
        process = self._children.get(program_name)
        if process and process.poll() is None:
            stopped = self._terminate_process_tree(process.pid, program_name)
            if stopped:
                self._children.pop(program_name, None)
            return stopped

        command_norm = self._normalize_path(command) if command else ""
        for proc in psutil.process_iter(["name", "pid", "cmdline"]):
            try:
                cmdline_parts = proc.info.get("cmdline") or []
                cmdline = " ".join(cmdline_parts)
                matches_name = (
                    program_name.lower() in cmdline.lower()
                    or program_name.lower() in (proc.info.get("name") or "").lower()
                )
                matches_command = False
                if command_norm and cmdline_parts:
                    first = self._normalize_path(str(cmdline_parts[0]))
                    matches_command = bool(first) and first == command_norm

                if matches_command or matches_name:
                    self._terminate_process_tree(proc.pid, program_name)
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False

    def restart(self, program: ManagedProgram) -> None:
        self.stop(program.name, command=program.command)
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

    def get_running_pid(self, program_name: str) -> int | None:
        process = self._children.get(program_name)
        if process and process.poll() is None:
            return int(process.pid)

        target = program_name.lower()
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                name = (proc.info.get("name") or "").lower()
                cmdline = " ".join(proc.info.get("cmdline") or []).lower()
                if target in name or target in cmdline:
                    return int(proc.info.get("pid") or 0)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return None

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
                for raw_line in iter(process.stdout.readline, b""):
                    line = self._decode_output_line(raw_line)
                    if line:
                        self._emit_output(program_name, line)
        finally:
            try:
                code = process.wait(timeout=1)
            except subprocess.TimeoutExpired:
                code = process.poll()
            cached = self._children.get(program_name)
            if cached is process:
                self._children.pop(program_name, None)
            self._emit_exit(program_name, code)

    @staticmethod
    def _decode_output_line(raw_line: bytes) -> str:
        # Windows child processes may emit GBK/CP936 while Python often emits UTF-8.
        data = raw_line.rstrip(b"\r\n")
        if not data:
            return ""

        candidates = ["utf-8", locale.getpreferredencoding(False), "gbk", "cp936"]
        seen: set[str] = set()
        for enc in candidates:
            if not enc:
                continue
            key = enc.lower()
            if key in seen:
                continue
            seen.add(key)
            try:
                return data.decode(enc)
            except UnicodeDecodeError:
                continue

        return data.decode("utf-8", errors="replace")

    def _terminate_process_tree(self, pid: int, program_name: str) -> bool:
        try:
            root = psutil.Process(pid)
        except psutil.NoSuchProcess:
            return False

        children = root.children(recursive=True)
        targets = children + [root]

        for proc in targets:
            try:
                proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        gone, alive = psutil.wait_procs(targets, timeout=4)

        for proc in alive:
            try:
                proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

        self._emit_output(program_name, f"[进程终止] 已结束进程树 pid={pid} 子进程数={len(children)}")
        return True

    @staticmethod
    def _normalize_path(value: str | None) -> str:
        if not value:
            return ""
        try:
            return os.path.normcase(os.path.normpath(value.strip().strip('"')))
        except Exception:
            return ""
