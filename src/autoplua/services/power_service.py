from __future__ import annotations

import ctypes
import os
import subprocess
from datetime import datetime, timezone
from typing import Any


_WINDOWS_TICKS_PER_SECOND = 10_000_000
_WINDOWS_EPOCH_DIFF_SECONDS = 11_644_473_600


class PowerService:
    def __init__(self) -> None:
        self._wake_timer_handle = None
        self._wake_task_name = "AutoPlua_WakeTask"

    @staticmethod
    def shutdown() -> None:
        os.system("shutdown /s /t 0")

    @staticmethod
    def restart() -> None:
        os.system("shutdown /r /t 0")

    @staticmethod
    def sleep() -> None:
        # Explicitly request sleep (S3) instead of hibernation and avoid force-critical path.
        ctypes.windll.powrprof.SetSuspendState(False, False, False)

    @staticmethod
    def lock() -> None:
        ctypes.windll.user32.LockWorkStation()

    @staticmethod
    def cancel_shutdown() -> None:
        subprocess.run(["shutdown", "/a"], check=False)

    def cancel_wake_timer(self) -> bool:
        if self._wake_timer_handle is None:
            return False

        ctypes.windll.kernel32.CancelWaitableTimer(self._wake_timer_handle)
        ctypes.windll.kernel32.CloseHandle(self._wake_timer_handle)
        self._wake_timer_handle = None
        return True

    @staticmethod
    def _connect_scheduler() -> Any:
        import win32com.client  # type: ignore

        scheduler = win32com.client.Dispatch("Schedule.Service")
        scheduler.Connect()
        return scheduler

    def cancel_wake_task(self) -> bool:
        try:
            scheduler = self._connect_scheduler()
            root_folder = scheduler.GetFolder("\\")
            root_folder.DeleteTask(self._wake_task_name, 0)
            return True
        except Exception:
            return False

    def _schedule_wake_task(self, wake_time: datetime) -> tuple[bool, str]:
        try:
            scheduler = self._connect_scheduler()
            root_folder = scheduler.GetFolder("\\")
            task_def = scheduler.NewTask(0)

            task_def.RegistrationInfo.Description = "AutoPlua wake task"
            task_def.Settings.Enabled = True
            task_def.Settings.WakeToRun = True
            task_def.Settings.DisallowStartIfOnBatteries = False
            task_def.Settings.StopIfGoingOnBatteries = False
            task_def.Settings.StartWhenAvailable = True
            task_def.Settings.Hidden = True

            trigger = task_def.Triggers.Create(1)  # TASK_TRIGGER_TIME
            trigger.StartBoundary = wake_time.strftime("%Y-%m-%dT%H:%M:%S")
            trigger.EndBoundary = wake_time.strftime("%Y-%m-%dT%H:%M:%S")
            trigger.Enabled = True

            action = task_def.Actions.Create(0)  # TASK_ACTION_EXEC
            action.Path = "cmd.exe"
            action.Arguments = "/c exit 0"

            # TASK_CREATE_OR_UPDATE=6, TASK_LOGON_INTERACTIVE_TOKEN=3
            root_folder.RegisterTaskDefinition(
                self._wake_task_name,
                task_def,
                6,
                "",
                "",
                3,
            )
            return True, "task-scheduler"
        except Exception as exc:
            return False, f"task-scheduler-failed: {exc}"

    def schedule_wake(self, wake_time: datetime) -> bool:
        # Rebuild wake schedule whenever settings change to ensure only one wake alarm is active.
        self.cancel_wake()

        task_ok, _ = self._schedule_wake_task(wake_time)
        if task_ok:
            return True

        handle = ctypes.windll.kernel32.CreateWaitableTimerW(None, True, None)
        if not handle:
            return False

        wake_utc = wake_time.astimezone(timezone.utc)
        unix_seconds = wake_utc.timestamp()
        filetime_ticks = int((unix_seconds + _WINDOWS_EPOCH_DIFF_SECONDS) * _WINDOWS_TICKS_PER_SECOND)
        due_time = ctypes.c_longlong(filetime_ticks)

        ok = ctypes.windll.kernel32.SetWaitableTimer(
            handle,
            ctypes.byref(due_time),
            0,
            None,
            None,
            True,
        )

        if ok:
            self._wake_timer_handle = handle
            return True

        ctypes.windll.kernel32.CloseHandle(handle)
        return False

    def cancel_wake(self) -> bool:
        task_deleted = self.cancel_wake_task()
        timer_deleted = self.cancel_wake_timer()
        return task_deleted or timer_deleted

    @staticmethod
    def get_wake_timers_report() -> tuple[bool, str]:
        try:
            result = subprocess.run(
                ["powercfg", "/waketimers"],
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
            )
            output = (result.stdout or "").strip()
            if output:
                return True, output
            err = (result.stderr or "").strip()
            return False, err or "no-output"
        except OSError as exc:
            return False, str(exc)
