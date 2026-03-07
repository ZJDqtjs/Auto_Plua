from __future__ import annotations

import shlex
import subprocess
from pathlib import Path

import psutil

from autoplua.models import ManagedProgram


class ProcessService:
    def __init__(self) -> None:
        self._children: dict[str, subprocess.Popen] = {}

    def start(self, program: ManagedProgram) -> None:
        if program.name in self._children and self._children[program.name].poll() is None:
            return

        cmd = [program.command, *program.args]
        cwd = program.cwd if program.cwd else None
        process = subprocess.Popen(cmd, cwd=cwd, shell=False)
        self._children[program.name] = process

    def stop(self, program_name: str) -> bool:
        process = self._children.get(program_name)
        if process and process.poll() is None:
            process.terminate()
            return True

        for proc in psutil.process_iter(["name", "pid", "cmdline"]):
            try:
                cmdline = " ".join(proc.info.get("cmdline") or [])
                if program_name.lower() in cmdline.lower() or program_name.lower() in (proc.info.get("name") or "").lower():
                    proc.terminate()
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
