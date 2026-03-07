from __future__ import annotations

from collections.abc import Callable

from apscheduler.schedulers.background import BackgroundScheduler


class SchedulerService:
    def __init__(self) -> None:
        self.scheduler = BackgroundScheduler()
        self._started = False

    def start(self) -> None:
        if not self._started:
            self.scheduler.start()
            self._started = True

    def shutdown(self) -> None:
        if self._started:
            self.scheduler.shutdown(wait=False)
            self._started = False

    def add_interval_job(self, job_id: str, seconds: int, func: Callable[[], None]) -> None:
        self.scheduler.add_job(
            func,
            trigger="interval",
            seconds=seconds,
            id=job_id,
            replace_existing=True,
            coalesce=True,
            max_instances=1,
        )

    def remove_job(self, job_id: str) -> None:
        self.scheduler.remove_job(job_id)
