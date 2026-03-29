from __future__ import annotations

import ctypes
import locale
import os
import re
import subprocess
from datetime import datetime, timedelta
from typing import Any

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover
    win32com = None  # type: ignore


TASK_TRIGGER_DAILY = 2
TASK_TRIGGER_WEEKLY = 3
TASK_ACTION_EXEC = 0
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_PASSWORD = 1
TASK_LOGON_SERVICE_ACCOUNT = 5
TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD = 6
TASK_RUNLEVEL_HIGHEST = 1
TASK_COMPATIBILITY_WIN10 = 6
WEEKDAY_MASK = 2 + 4 + 8 + 16 + 32
WEEKEND_MASK = 1 + 64

class PowerService:
    def __init__(self) -> None:
        self._wake_task_name = "AutoPlua_WakeTask"
        self._wake_anchor_task_name = "AutoPlua_WakeAnchorTask"
        self._shutdown_task_name = "AutoPlua_ShutdownTask"

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

    def apply_windows_power_schedule(self, settings: dict[str, Any]) -> tuple[bool, list[str]]:
        messages: list[str] = []
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        messages.append(f"开始同步 Windows 电源任务，当前时间：{now}")
        messages.append(f"当前供电状态：{self._get_power_source_text()}")
        if win32com is not None:
            messages.append("唤醒任务通道：pywin32(COM) 优先")
        else:
            messages.append("唤醒任务通道：pywin32 不可用，自动回退 PowerShell ScheduledTasks")

        policy_ok, policy_msg = self._ensure_wake_timer_policy_enabled()
        messages.append(policy_msg)
        messages.extend(self._diagnose_wake_prerequisites())

        wake_ok, wake_msg = self._register_wake_task(
            frequency=str(settings.get("boot_frequency", "每天")),
            hhmm=str(settings.get("boot_time", "06:30")),
            login_user=str(settings.get("login_user", "")).strip(),
            login_password=str(settings.get("login_password", "")),
        )
        messages.append(wake_msg)

        shutdown_ok, shutdown_msg = self._register_shutdown_task(
            frequency=str(settings.get("shutdown_frequency", "每天")),
            hhmm=str(settings.get("shutdown_time", "23:00")),
            action=str(settings.get("shutdown_action", "关机")),
        )
        messages.append(shutdown_msg)

        return policy_ok and wake_ok and shutdown_ok, messages

    def clear_windows_power_schedule(self) -> tuple[bool, list[str]]:
        ok1, msg1 = self._delete_task(self._wake_task_name)
        ok2, msg2 = self._delete_task(self._wake_anchor_task_name)
        ok3, msg3 = self._delete_task(self._shutdown_task_name)
        return ok1 and ok2 and ok3, [msg1, msg2, msg3]

    def _delete_task(self, task_name: str) -> tuple[bool, str]:
        result = subprocess.run(
            ["schtasks", "/Delete", "/TN", task_name, "/F"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode == 0:
            return True, f"已删除计划任务：{task_name}"

        output = ((result.stdout or "") + "\n" + (result.stderr or "")).strip().lower()
        if "cannot find the file specified" in output or "找不到" in output:
            return True, f"计划任务不存在，无需删除：{task_name}"
        return False, f"删除计划任务失败：{task_name}"

    def _register_shutdown_task(self, frequency: str, hhmm: str, action: str) -> tuple[bool, str]:
        if not self._is_valid_hhmm(hhmm):
            return False, f"关机任务时间格式无效：{hhmm}"

        cmd = self._shutdown_command(action)
        if not cmd:
            return False, f"不支持的关机动作：{action}"

        schedule_args = self._schtasks_schedule_args(frequency)
        if not schedule_args:
            return False, f"不支持的关机频率：{frequency}"

        args = [
            "schtasks",
            "/Create",
            "/TN",
            self._shutdown_task_name,
            "/TR",
            cmd,
            *schedule_args,
            "/ST",
            hhmm,
            "/RU",
            "SYSTEM",
            "/RL",
            "HIGHEST",
            "/F",
        ]
        result = subprocess.run(
            args,
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode == 0:
            verify = self._query_task_brief(self._shutdown_task_name)
            return True, (
                f"关机任务已更新：{frequency} {hhmm} {action}"
                + (f" | {verify}" if verify else "")
            )

        detail = ((result.stdout or "") + "\n" + (result.stderr or "")).strip() or "unknown-error"
        return False, f"关机任务创建失败：{detail}"

    def _register_wake_task(
        self,
        frequency: str,
        hhmm: str,
        login_user: str,
        login_password: str,
    ) -> tuple[bool, str]:
        if not self._is_valid_hhmm(hhmm):
            return False, f"唤醒任务时间格式无效：{hhmm}"

        interactive_time = hhmm
        wake_anchor_time = self._minus_one_minute(hhmm)
        if self._frequency_kind(frequency) is None:
            return False, f"不支持的唤醒频率：{frequency}"

        task_user = self._normalize_task_user(login_user)
        if not task_user:
            return False, "唤醒任务创建失败：请在电源页填写 Windows 登录账号（示例：PCNAME\\User）"
        if not login_password:
            return False, "唤醒任务创建失败：请在电源页填写登录密码（已移除 S4U 模式）"

        anchor_ok, anchor_msg = self._register_wake_anchor_task(frequency, wake_anchor_time)
        if not anchor_ok:
            return False, f"唤醒锚点任务创建失败：{anchor_msg}"

        # Prefer COM (pywin32) for feature parity with C# Task Scheduler integration.
        if win32com is not None:
            ok, com_msg = self._register_wake_task_via_com(
                frequency,
                interactive_time,
                task_user,
                login_password,
            )
            if ok:
                verify = self._query_wake_task_details(self._wake_task_name)
                anchor_verify = self._query_wake_task_details(self._wake_anchor_task_name)
                lead_warning = self._warn_if_wake_too_close(wake_anchor_time)
                detail_parts = [
                    f"唤醒锚点任务已更新：{frequency} {wake_anchor_time} (SYSTEM, WakeToRun)",
                    anchor_msg,
                    anchor_verify,
                    f"交互探针任务已更新(COM)：{frequency} {interactive_time}",
                    com_msg,
                    verify,
                ]
                if lead_warning:
                    detail_parts.append(lead_warning)
                detail_parts.append("提示：唤醒仅对睡眠/休眠有效，关机(S5)通常无法由任务计划唤醒")
                detail_parts.append("提示：交互探针任务依赖用户会话；唤醒锚点任务由 SYSTEM 保证唤醒可靠性")
                detail_parts.append("验证路径：C:\\ProgramData\\AutoPlua\\wake_probe.log")
                return True, " | ".join(part for part in detail_parts if part)

            # COM failed, continue with PowerShell fallback.
            fallback_ok, fallback_msg = self._register_wake_task_via_powershell(
                frequency,
                interactive_time,
                task_user,
                login_password,
            )
            if fallback_ok:
                verify = self._query_wake_task_details(self._wake_task_name)
                anchor_verify = self._query_wake_task_details(self._wake_anchor_task_name)
                lead_warning = self._warn_if_wake_too_close(wake_anchor_time)
                detail_parts = [
                    f"唤醒锚点任务已更新：{frequency} {wake_anchor_time} (SYSTEM, WakeToRun)",
                    anchor_msg,
                    anchor_verify,
                    f"交互探针任务已更新(PowerShell回退)：{frequency} {interactive_time}",
                    f"COM失败原因：{com_msg}",
                    verify,
                ]
                if lead_warning:
                    detail_parts.append(lead_warning)
                detail_parts.append("提示：唤醒仅对睡眠/休眠有效，关机(S5)通常无法由任务计划唤醒")
                detail_parts.append("提示：交互探针任务依赖用户会话；唤醒锚点任务由 SYSTEM 保证唤醒可靠性")
                detail_parts.append("验证路径：C:\\ProgramData\\AutoPlua\\wake_probe.log")
                return True, " | ".join(part for part in detail_parts if part)
            return False, f"唤醒任务创建失败：COM与PowerShell都失败。COM={com_msg}; PS={fallback_msg}"

        fallback_ok, fallback_msg = self._register_wake_task_via_powershell(
            frequency,
            interactive_time,
            task_user,
            login_password,
        )
        if fallback_ok:
            verify = self._query_wake_task_details(self._wake_task_name)
            anchor_verify = self._query_wake_task_details(self._wake_anchor_task_name)
            lead_warning = self._warn_if_wake_too_close(wake_anchor_time)
            detail_parts = [
                f"唤醒锚点任务已更新：{frequency} {wake_anchor_time} (SYSTEM, WakeToRun)",
                anchor_msg,
                anchor_verify,
                f"交互探针任务已更新(PowerShell)：{frequency} {interactive_time}",
                verify,
            ]
            if lead_warning:
                detail_parts.append(lead_warning)
            detail_parts.append("提示：唤醒仅对睡眠/休眠有效，关机(S5)通常无法由任务计划唤醒")
            detail_parts.append("提示：交互探针任务依赖用户会话；唤醒锚点任务由 SYSTEM 保证唤醒可靠性")
            detail_parts.append("验证路径：C:\\ProgramData\\AutoPlua\\wake_probe.log")
            return True, " | ".join(part for part in detail_parts if part)
        return False, f"唤醒任务创建失败：{fallback_msg}"

    def _register_wake_anchor_task(self, frequency: str, wake_time: str) -> tuple[bool, str]:
        if win32com is not None:
            ok, msg = self._register_wake_anchor_task_via_com(frequency, wake_time)
            if ok:
                return True, msg

        return self._register_wake_anchor_task_via_powershell(frequency, wake_time)

    def _register_wake_anchor_task_via_com(self, frequency: str, wake_time: str) -> tuple[bool, str]:
        if win32com is None:
            return False, "pywin32 不可用"

        try:
            scheduler = win32com.client.Dispatch("Schedule.Service")  # type: ignore[attr-defined]
            scheduler.Connect()
            root_folder = scheduler.GetFolder("\\")
            task_def = scheduler.NewTask(0)

            task_def.RegistrationInfo.Description = "AutoPlua wake anchor task (SYSTEM)"
            task_def.Settings.Enabled = True
            task_def.Settings.WakeToRun = True
            task_def.Settings.DisallowStartIfOnBatteries = False
            task_def.Settings.StopIfGoingOnBatteries = False
            task_def.Settings.StartWhenAvailable = True
            task_def.Settings.Hidden = True
            task_def.Settings.Compatibility = TASK_COMPATIBILITY_WIN10

            task_def.Principal.UserId = "SYSTEM"
            task_def.Principal.LogonType = TASK_LOGON_SERVICE_ACCOUNT
            task_def.Principal.RunLevel = TASK_RUNLEVEL_HIGHEST

            start_boundary = self._next_start_boundary(frequency, wake_time)
            kind = self._frequency_kind(frequency)
            if kind == "daily":
                trigger = task_def.Triggers.Create(TASK_TRIGGER_DAILY)
                trigger.DaysInterval = 1
            else:
                trigger = task_def.Triggers.Create(TASK_TRIGGER_WEEKLY)
                trigger.WeeksInterval = 1
                trigger.DaysOfWeek = self._weekly_mask(frequency)

            trigger.StartBoundary = start_boundary
            trigger.Enabled = True

            action = task_def.Actions.Create(TASK_ACTION_EXEC)
            action.Path = "cmd.exe"
            action.Arguments = "/c exit 0"

            root_folder.RegisterTaskDefinition(
                self._wake_anchor_task_name,
                task_def,
                TASK_CREATE_OR_UPDATE,
                "SYSTEM",
                "",
                TASK_LOGON_SERVICE_ACCOUNT,
                "",
            )
            return True, f"Anchor(COM)：StartBoundary={start_boundary}"
        except Exception as exc:
            return False, self._format_com_error(exc)

    def _register_wake_anchor_task_via_powershell(self, frequency: str, wake_time: str) -> tuple[bool, str]:
        days_of_week = self._powershell_days_of_week(frequency)
        if days_of_week is None:
            return False, f"不支持的唤醒频率：{frequency}"

        if days_of_week:
            trigger_script = (
                "$trigger = New-ScheduledTaskTrigger -Weekly -At '{time}' -DaysOfWeek {days}".format(
                    time=wake_time,
                    days=days_of_week,
                )
            )
        else:
            trigger_script = f"$trigger = New-ScheduledTaskTrigger -Daily -At '{wake_time}'"

        script = (
            "$ErrorActionPreference = 'Stop';"
            f"$taskName = '{self._wake_anchor_task_name}';"
            "$action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument '/c exit 0';"
            f"{trigger_script};"
            "$settings = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable "
            "-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden;"
            "$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest;"
            "Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger "
            "-Settings $settings -Principal $principal -Force | Out-Null"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding=self._preferred_encoding(),
            errors="ignore",
        )
        if result.returncode == 0:
            return True, "Anchor(PowerShell)：注册成功"

        detail = ((result.stdout or "") + "\n" + (result.stderr or "")).strip() or "unknown-error"
        return False, detail

    def _register_wake_task_via_powershell(
        self,
        frequency: str,
        wake_time: str,
        task_user: str,
        login_password: str,
    ) -> tuple[bool, str]:
        days_of_week = self._powershell_days_of_week(frequency)
        if days_of_week is None:
            return False, f"不支持的唤醒频率：{frequency}"

        if days_of_week:
            trigger_script = (
                "$trigger = New-ScheduledTaskTrigger -Weekly -At '{time}' -DaysOfWeek {days}".format(
                    time=wake_time,
                    days=days_of_week,
                )
            )
        else:
            trigger_script = f"$trigger = New-ScheduledTaskTrigger -Daily -At '{wake_time}'"

        script = (
            "$ErrorActionPreference = 'Stop';"
            f"$taskName = '{self._wake_task_name}';"
            f"$action = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument '{self._wake_probe_cmd_arguments()}';"
            f"{trigger_script};"
            "$settings = New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable "
            "-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;"
            f"Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger "
            f"-Settings $settings -User '{task_user}' -Password '{self._ps_single_quote(login_password)}' "
            "-RunLevel Highest -Force | Out-Null"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode == 0:
            return True, "PowerShell ScheduledTasks 注册成功"

        detail = ((result.stdout or "") + "\n" + (result.stderr or "")).strip() or "unknown-error"
        return False, detail

    def _register_wake_task_via_com(
        self,
        frequency: str,
        wake_time: str,
        task_user: str,
        login_password: str,
    ) -> tuple[bool, str]:
        if win32com is None:
            return False, "pywin32 不可用"

        try:
            scheduler = win32com.client.Dispatch("Schedule.Service")  # type: ignore[attr-defined]
            scheduler.Connect()
            root_folder = scheduler.GetFolder("\\")
            task_def = scheduler.NewTask(0)

            task_def.RegistrationInfo.Description = "AutoPlua wake task (pywin32 COM)"
            task_def.Settings.Enabled = True
            task_def.Settings.WakeToRun = True
            task_def.Settings.DisallowStartIfOnBatteries = False
            task_def.Settings.StopIfGoingOnBatteries = False
            task_def.Settings.StartWhenAvailable = True
            task_def.Settings.Hidden = False
            task_def.Settings.Compatibility = TASK_COMPATIBILITY_WIN10

            task_def.Principal.UserId = task_user
            task_def.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD
            task_def.Principal.RunLevel = TASK_RUNLEVEL_HIGHEST

            start_boundary = self._next_start_boundary(frequency, wake_time)
            kind = self._frequency_kind(frequency)
            if kind == "daily":
                trigger = task_def.Triggers.Create(TASK_TRIGGER_DAILY)
                trigger.DaysInterval = 1
            else:
                trigger = task_def.Triggers.Create(TASK_TRIGGER_WEEKLY)
                trigger.WeeksInterval = 1
                trigger.DaysOfWeek = self._weekly_mask(frequency)

            trigger.StartBoundary = start_boundary
            trigger.Enabled = True

            action = task_def.Actions.Create(TASK_ACTION_EXEC)
            action.Path = "cmd.exe"
            action.Arguments = self._wake_probe_cmd_arguments()

            root_folder.RegisterTaskDefinition(
                self._wake_task_name,
                task_def,
                TASK_CREATE_OR_UPDATE,
                task_user,
                login_password,
                TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD,
                "",
            )
            return True, f"COM校验：StartBoundary={start_boundary}; User={task_user}; Compatibility=Windows10"
        except Exception as exc:
            return False, self._format_com_error(exc)

    def _next_start_boundary(self, frequency: str, hhmm: str) -> str:
        now = datetime.now()
        hour_text, minute_text = hhmm.split(":", 1)
        hour = int(hour_text)
        minute = int(minute_text)

        for offset in range(0, 8):
            candidate = (now + timedelta(days=offset)).replace(
                hour=hour,
                minute=minute,
                second=0,
                microsecond=0,
            )
            if candidate <= now:
                continue
            if self._candidate_match_frequency(frequency, candidate):
                return candidate.strftime("%Y-%m-%dT%H:%M:%S")

        fallback = (now + timedelta(days=1)).replace(hour=hour, minute=minute, second=0, microsecond=0)
        return fallback.strftime("%Y-%m-%dT%H:%M:%S")

    @staticmethod
    def _candidate_match_frequency(frequency: str, when: datetime) -> bool:
        weekday = when.weekday()  # Mon=0 ... Sun=6
        if frequency == "每天":
            return True
        if frequency == "工作日":
            return weekday < 5
        if frequency == "周末":
            return weekday >= 5
        return False

    @staticmethod
    def _frequency_kind(frequency: str) -> str | None:
        if frequency == "每天":
            return "daily"
        if frequency in {"工作日", "周末"}:
            return "weekly"
        return None

    @staticmethod
    def _weekly_mask(frequency: str) -> int:
        if frequency == "工作日":
            return WEEKDAY_MASK
        if frequency == "周末":
            return WEEKEND_MASK
        # daily mode does not use weekly mask, but keep a safe fallback.
        return WEEKDAY_MASK + WEEKEND_MASK

    @staticmethod
    def _is_valid_hhmm(hhmm: str) -> bool:
        try:
            hour_text, minute_text = hhmm.split(":", 1)
            hour = int(hour_text)
            minute = int(minute_text)
            return 0 <= hour <= 23 and 0 <= minute <= 59
        except (ValueError, AttributeError):
            return False

    @staticmethod
    def _minus_one_minute(hhmm: str) -> str:
        hour_text, minute_text = hhmm.split(":", 1)
        hour = int(hour_text)
        minute = int(minute_text)
        total = (hour * 60 + minute - 1) % (24 * 60)
        new_hour = total // 60
        new_minute = total % 60
        return f"{new_hour:02d}:{new_minute:02d}"

    @staticmethod
    def _schtasks_schedule_args(frequency: str) -> list[str] | None:
        if frequency == "每天":
            return ["/SC", "DAILY"]
        if frequency == "工作日":
            return ["/SC", "WEEKLY", "/D", "MON,TUE,WED,THU,FRI"]
        if frequency == "周末":
            return ["/SC", "WEEKLY", "/D", "SAT,SUN"]
        return None

    @staticmethod
    def _powershell_days_of_week(frequency: str) -> str | None:
        if frequency == "每天":
            return ""
        if frequency == "工作日":
            return "Monday,Tuesday,Wednesday,Thursday,Friday"
        if frequency == "周末":
            return "Saturday,Sunday"
        return None

    @staticmethod
    def _shutdown_command(action: str) -> str:
        if action == "关机":
            return "shutdown.exe /s /f /t 0"
        if action == "注销":
            return "shutdown.exe /l /f"
        if action == "重启":
            return "shutdown.exe /r /f /t 0"
        if action == "睡眠":
            return "rundll32.exe powrprof.dll,SetSuspendState 0,1,0"
        if action == "锁屏":
            return "rundll32.exe user32.dll,LockWorkStation"
        return ""

    @staticmethod
    def _wake_probe_cmd_arguments() -> str:
        # Keep cmd alive so wake execution is visible to the user.
        return (
            '/k "title AutoPlua Wake Probe'
            ' & echo [AutoPlua] Wake task triggered at %date% %time%'
            ' & echo This window stays open for verification.'
            ' & if not exist C:\\ProgramData\\AutoPlua mkdir C:\\ProgramData\\AutoPlua'
            ' & echo Probe log: C:\\ProgramData\\AutoPlua\\wake_probe.log'
            ' & echo [%date% %time%] Wake task triggered>>C:\\ProgramData\\AutoPlua\\wake_probe.log"'
        )

    def _query_task_brief(self, task_name: str) -> str:
        script = (
            "$task = Get-ScheduledTask -TaskName '" + task_name + "' -ErrorAction Stop;"
            "$info = Get-ScheduledTaskInfo -TaskName '" + task_name + "' -ErrorAction Stop;"
            "$action = $task.Actions | Select-Object -First 1;"
            "$obj = [PSCustomObject]@{"
            "NextRunTime = $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss');"
            "State = $task.State.ToString();"
            "LastTaskResult = $info.LastTaskResult;"
            "Action = ($action.Execute + ' ' + $action.Arguments).Trim()"
            "};"
            "$obj | ConvertTo-Json -Compress"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding=self._preferred_encoding(),
            errors="ignore",
        )
        if result.returncode != 0:
            detail = ((result.stdout or "") + "\n" + (result.stderr or "")).strip() or "unknown-error"
            return f"任务校验失败（无法读取任务详情）：{detail}"

        payload = (result.stdout or "").strip().replace("\r", "").replace("\n", "")
        if not payload:
            return "任务校验失败（查询结果为空）"
        return f"任务校验：{payload}"

    def _query_wake_task_details(self, task_name: str) -> str:
        script = (
            f"$task = Get-ScheduledTask -TaskName '{task_name}' -ErrorAction Stop;"
            f"$info = Get-ScheduledTaskInfo -TaskName '{task_name}' -ErrorAction Stop;"
            "$trigger = $task.Triggers | Select-Object -First 1;"
            "$days = ($trigger.DaysOfWeek | ForEach-Object { $_.ToString() }) -join ',';"
            "$obj = [PSCustomObject]@{"
            "UserId = $task.Principal.UserId;"
            "LogonType = $task.Principal.LogonType.ToString();"
            "WakeToRun = $task.Settings.WakeToRun;"
            "NextRunTime = $info.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss');"
            "LastTaskResult = $info.LastTaskResult;"
            "StartBoundary = $trigger.StartBoundary;"
            "State = $task.State.ToString();"
            "DaysOfWeek = $days"
            "};"
            "$obj | ConvertTo-Json -Compress"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            encoding=self._preferred_encoding(),
            errors="ignore",
        )
        if result.returncode != 0:
            detail = ((result.stdout or "") + "\n" + (result.stderr or "")).strip() or "unknown-error"
            return f"唤醒任务校验失败({task_name})：{detail}"

        payload = (result.stdout or "").strip().replace("\r", "").replace("\n", "")
        if not payload:
            return f"唤醒任务校验失败({task_name})：查询结果为空"
        return f"唤醒任务校验({task_name})：{payload}"

    def _warn_if_wake_too_close(self, wake_time_hhmm: str) -> str:
        try:
            now = datetime.now()
            hour_text, minute_text = wake_time_hhmm.split(":", 1)
            target = now.replace(hour=int(hour_text), minute=int(minute_text), second=0, microsecond=0)
            if target <= now:
                target = target + timedelta(days=1)
            delta_seconds = (target - now).total_seconds()
            if delta_seconds <= 120:
                return "警告：唤醒任务距离当前时间过近（<=120秒），可能来不及进入睡眠状态"
            return ""
        except Exception:
            return ""

    def _diagnose_wake_prerequisites(self) -> list[str]:
        messages: list[str] = []
        messages.append("前置检查：唤醒依赖系统允许 Wake Timers，且设备需进入睡眠/休眠")

        result = subprocess.run(
            ["powercfg", "/a"],
            check=False,
            capture_output=True,
            text=True,
            encoding=self._preferred_encoding(),
            errors="ignore",
        )
        if result.returncode == 0:
            summary = (result.stdout or "").strip().replace("\r", " ").replace("\n", " | ")
            if summary:
                messages.append(f"电源状态支持：{summary[:380]}")
        else:
            messages.append("前置检查失败：无法读取 powercfg /a")

        waketimers = subprocess.run(
            ["powercfg", "/waketimers"],
            check=False,
            capture_output=True,
            text=True,
            encoding=self._preferred_encoding(),
            errors="ignore",
        )
        if waketimers.returncode == 0:
            wt = (waketimers.stdout or "").strip().replace("\r", " ").replace("\n", " | ")
            if wt:
                messages.append(f"当前 Wake Timers：{wt[:380]}")
        return messages

    def _ensure_wake_timer_policy_enabled(self) -> tuple[bool, str]:
        before = self._query_rtcwake_indices()
        if before is None:
            return False, "唤醒策略检查失败：无法读取 RTCWAKE 当前值"

        before_ac, before_dc = before
        steps = [
            ["powercfg", "/SETACVALUEINDEX", "SCHEME_CURRENT", "SUB_SLEEP", "RTCWAKE", "1"],
            ["powercfg", "/SETDCVALUEINDEX", "SCHEME_CURRENT", "SUB_SLEEP", "RTCWAKE", "1"],
            ["powercfg", "/SETACTIVE", "SCHEME_CURRENT"],
        ]
        for cmd in steps:
            result = subprocess.run(
                cmd,
                check=False,
                capture_output=True,
                text=True,
                encoding=self._preferred_encoding(),
                errors="ignore",
            )
            if result.returncode != 0:
                detail = ((result.stdout or "") + "\n" + (result.stderr or "")).strip() or "unknown-error"
                return False, f"唤醒策略修复失败：{' '.join(cmd)} -> {detail}"

        after = self._query_rtcwake_indices()
        if after is None:
            return False, "唤醒策略修复后校验失败：无法读取 RTCWAKE 值"

        after_ac, after_dc = after
        return True, (
            "已校准唤醒定时器策略 RTCWAKE："
            f"AC {before_ac}-> {after_ac}, DC {before_dc}-> {after_dc}"
        )

    def _query_rtcwake_indices(self) -> tuple[int, int] | None:
        result = subprocess.run(
            ["powercfg", "/query", "SCHEME_CURRENT", "SUB_SLEEP", "RTCWAKE"],
            check=False,
            capture_output=True,
            text=True,
            encoding=self._preferred_encoding(),
            errors="ignore",
        )
        if result.returncode != 0:
            return None

        text = result.stdout or ""
        ac = self._extract_hex_index(
            text,
            [r"当前交流电源设置索引:\s*0x([0-9a-fA-F]+)", r"Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)"],
        )
        dc = self._extract_hex_index(
            text,
            [r"当前直流电源设置索引:\s*0x([0-9a-fA-F]+)", r"Current DC Power Setting Index:\s*0x([0-9a-fA-F]+)"],
        )
        if ac is None or dc is None:
            return None
        return ac, dc

    @staticmethod
    def _extract_hex_index(text: str, patterns: list[str]) -> int | None:
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                return int(match.group(1), 16)
        return None

    @staticmethod
    def _get_power_source_text() -> str:
        class SYSTEM_POWER_STATUS(ctypes.Structure):
            _fields_ = [
                ("ACLineStatus", ctypes.c_ubyte),
                ("BatteryFlag", ctypes.c_ubyte),
                ("BatteryLifePercent", ctypes.c_ubyte),
                ("Reserved1", ctypes.c_ubyte),
                ("BatteryLifeTime", ctypes.c_uint32),
                ("BatteryFullLifeTime", ctypes.c_uint32),
            ]

        status = SYSTEM_POWER_STATUS()
        ok = ctypes.windll.kernel32.GetSystemPowerStatus(ctypes.byref(status))
        if not ok:
            return "unknown"

        if status.ACLineStatus == 1:
            return "AC(外接电源)"
        if status.ACLineStatus == 0:
            return "DC(电池供电)"
        return "unknown"

    @staticmethod
    def _resolve_current_user_id() -> str:
        domain = (os.getenv("USERDOMAIN") or "").strip()
        username = (os.getenv("USERNAME") or "").strip()
        if not username:
            return ""
        if domain:
            return f"{domain}\\{username}"
        return username

    def _normalize_task_user(self, login_user: str) -> str:
        value = (login_user or "").strip()
        if not value:
            return ""

        if "\\" in value or "@" in value:
            return value

        domain = (os.getenv("USERDOMAIN") or "").strip()
        if domain:
            return f"{domain}\\{value}"
        return value

    @staticmethod
    def _ps_single_quote(text: str) -> str:
        return text.replace("'", "''")

    def cancel_wake(self) -> bool:
        ok, _ = self._delete_task(self._wake_task_name)
        return ok

    @staticmethod
    def get_wake_timers_report() -> tuple[bool, str]:
        try:
            result = subprocess.run(
                ["powercfg", "/waketimers"],
                check=False,
                capture_output=True,
                text=True,
                encoding=PowerService._preferred_encoding(),
                errors="ignore",
            )
            output = (result.stdout or "").strip()
            if output:
                return True, output
            err = (result.stderr or "").strip()
            return False, err or "no-output"
        except OSError as exc:
            return False, str(exc)

    @staticmethod
    def _preferred_encoding() -> str:
        # Use host locale to avoid mojibake on non-UTF8 consoles (e.g., zh-CN cp936).
        return locale.getpreferredencoding(False) or "utf-8"

    @staticmethod
    def _format_com_error(exc: Exception) -> str:
        text = str(exc)
        hresult_text = ""
        if isinstance(exc.args, tuple) and exc.args:
            first = exc.args[0]
            if isinstance(first, int):
                unsigned = first & 0xFFFFFFFF
                hresult_text = f" HRESULT=0x{unsigned:08X}"
                if unsigned == 0x80070522:
                    hresult_text += "(权限不足，建议管理员运行)"
                elif unsigned == 0x80070005:
                    hresult_text += "(访问被拒绝，建议管理员运行)"
        return (text + hresult_text).strip()
