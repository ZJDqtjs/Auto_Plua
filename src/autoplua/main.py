from __future__ import annotations

import sys
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtWidgets import QApplication

from autoplua.logger import setup_logger
from autoplua.services.power_service import PowerService
from autoplua.services.process_service import ProcessService
from autoplua.services.scheduler_service import SchedulerService
from autoplua.ui.main_window import MainWindow


def run() -> int:
    logger = setup_logger()
    app = QApplication(sys.argv)

    process_service = ProcessService()
    scheduler_service = SchedulerService()
    power_service = PowerService()

    scheduler_service.start()

    window = MainWindow(
        logger=logger,
        process_service=process_service,
        scheduler_service=scheduler_service,
        power_service=power_service,
    )
    window.show()

    exit_code = app.exec()
    scheduler_service.shutdown()
    return exit_code


if __name__ == "__main__":
    raise SystemExit(run())
