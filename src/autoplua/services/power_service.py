from __future__ import annotations

import ctypes
import os
import subprocess


class PowerService:
    @staticmethod
    def shutdown() -> None:
        os.system("shutdown /s /t 0")

    @staticmethod
    def restart() -> None:
        os.system("shutdown /r /t 0")

    @staticmethod
    def sleep() -> None:
        ctypes.windll.powrprof.SetSuspendState(False, True, False)

    @staticmethod
    def lock() -> None:
        ctypes.windll.user32.LockWorkStation()

    @staticmethod
    def cancel_shutdown() -> None:
        subprocess.run(["shutdown", "/a"], check=False)
