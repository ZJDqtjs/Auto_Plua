from __future__ import annotations

import logging
import os
import shlex
import socket
import struct
import subprocess
from datetime import datetime, timedelta
from functools import partial
from pathlib import Path

from PySide6.QtCore import Qt, QSize, QTime, Signal, QTimer
from PySide6.QtGui import QColor, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QStackedWidget,
    QStyle,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

from autoplua.config import load_config, save_config
from autoplua.models import ManagedProgram
from autoplua.services.opencv_service import OpenCVFlowService
from autoplua.services.power_service import PowerService
from autoplua.services.process_service import ProcessService
from autoplua.services.scheduler_service import SchedulerService
from autoplua.services.virtual_display_service import VirtualDisplayService
from autoplua.ui.program_config_dialog import ProgramConfigDialog
from autoplua.ui.program_list_item import DailyLoopItemWidget
from autoplua.ui.styles import SIDEBAR_COLLAPSED_STYLE, SIDEBAR_EXPANDED_STYLE


class MainWindow(QMainWindow):
    app_log_signal = Signal(str)
    program_log_signal = Signal(str)
    schedule_start_signal = Signal(object)
    schedule_stop_signal = Signal(object)

    def __init__(
        self,
        logger: logging.Logger,
        process_service: ProcessService,
        scheduler_service: SchedulerService,
        power_service: PowerService,
        opencv_flow_service: OpenCVFlowService,
        virtual_display_service: VirtualDisplayService,
    ) -> None:
        super().__init__()
        self.logger = logger
        self.process_service = process_service
        self.scheduler_service = scheduler_service
        self.power_service = power_service
        self.opencv_flow_service = opencv_flow_service
        self.virtual_display_service = virtual_display_service

        self.config = load_config()
        self.program_map: dict[str, ManagedProgram] = {}
        self.program_entries: list[dict] = []
        self.runtime_logs: list[str] = []
        self.program_runtime_logs: list[str] = []
        self._power_state: dict[str, str] = {"boot": "", "shutdown": ""}
        self._project_runtime_active = False
        self._project_runtime_marks: dict[str, str] = {}
        self._program_runtime_job_id = "program_runtime_tick"
        self._app_log_follow_tail = True
        self._program_log_follow_tail = True
        self._last_power_tick_at: datetime | None = None
        self._scheduled_wake_marker = ""

        self.app_log_signal.connect(self._handle_append_log)
        self.program_log_signal.connect(self._handle_append_program_log)
        self.schedule_start_signal.connect(self._handle_schedule_start)
        self.schedule_stop_signal.connect(self._handle_schedule_stop)

        self.process_service.set_output_listener(self._on_program_output)
        self.process_service.set_exit_listener(self._on_program_exit)

        self.setWindowTitle("AutoPlua")
        self.resize(1000, 700)
        self.setStyleSheet("QMainWindow { background-color: #f5f5f5; }")
        self._build_ui()

    def _build_ui(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Left Sidebar
        self.is_sidebar_collapsed = False
        self.sidebarWidget = QFrame()
        self.sidebarWidget.setFixedWidth(160)
        self.sidebarWidget.setStyleSheet("background-color: #ffffff; border-right: 1px solid #ddd;")
        sidebar_layout = QVBoxLayout(self.sidebarWidget)
        sidebar_layout.setContentsMargins(0, 20, 0, 0)
        sidebar_layout.setSpacing(10)

        # Menu Icon
        self.menu_btn = QPushButton("☰")
        self.menu_btn.setFlat(True)
        self.menu_btn.setCursor(Qt.PointingHandCursor)
        self.menu_btn.setFixedSize(50, 40)
        self.menu_btn.setStyleSheet("font-size: 24px; border: none; margin-left: 6px;")
        self.menu_btn.clicked.connect(self._toggle_sidebar)
        sidebar_layout.addWidget(self.menu_btn, 0, Qt.AlignLeft)

        # Nav Buttons
        self.nav_buttons = []
        self.nav_items_text = ["开始", "电源", "日志", "关于"]
        self.nav_item_icons = [
            self.style().standardIcon(QStyle.SP_MediaPlay),
            self._create_power_icon(),
            self.style().standardIcon(QStyle.SP_FileDialogDetailedView),
            self.style().standardIcon(QStyle.SP_MessageBoxInformation),
        ]
        for i, item in enumerate(self.nav_items_text):
            btn = QPushButton(item)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setMinimumHeight(54)
            btn.setIcon(self.nav_item_icons[i])
            btn.setIconSize(QSize(22, 22))
            btn.setToolTip(item)
            btn.setStyleSheet(SIDEBAR_EXPANDED_STYLE)
            btn.clicked.connect(partial(self._switch_page, i))
            sidebar_layout.addWidget(btn)
            self.nav_buttons.append(btn)

        sidebar_layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))
        main_layout.addWidget(self.sidebarWidget)

        # Stacked Widget
        self.stacked_widget = QStackedWidget()
        main_layout.addWidget(self.stacked_widget)

        # Build Pages
        self.start_page = self._build_start_page()
        self.stacked_widget.addWidget(self.start_page)

        self.power_page = self._build_power_page()
        self.stacked_widget.addWidget(self.power_page)

        self.log_page = self._build_log_page()
        self.stacked_widget.addWidget(self.log_page)

        self.about_page = self._build_about_page()
        self.stacked_widget.addWidget(self.about_page)

        # Set Startup state
        self.nav_buttons[0].setChecked(True)
        self.stacked_widget.setCurrentIndex(0)
        self._load_programs_from_config()
        self._position_start_add_button()
        QTimer.singleShot(1000, self._bootstrap_virtual_display_on_app_start)

    def _build_power_page(self) -> QWidget:
        self.power_stack = QStackedWidget()
        self.power_home_page = self._build_power_home_page()
        self.power_task_page = self._build_power_placeholder_page("Windows 任务计划")
        self.power_wol_page = self._build_power_placeholder_page("WOL 方式")

        self.power_stack.addWidget(self.power_home_page)
        self.power_stack.addWidget(self.power_task_page)
        self.power_stack.addWidget(self.power_wol_page)
        return self.power_stack

    def _build_power_home_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(16)

        power_settings = self._get_power_settings()

        title = QLabel("电源自动化管理")
        title.setStyleSheet("font-size: 34px; font-weight: 700; color: #222;")
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #ccc; max-height: 1px;")
        layout.addWidget(line)

        enable_row = QFrame()
        enable_row.setStyleSheet(
            "QFrame { background: #fff; border: 1px solid #e0e0e0; border-radius: 10px; }"
        )
        enable_layout = QHBoxLayout(enable_row)
        enable_layout.setContentsMargins(16, 14, 16, 14)
        enable_layout.setSpacing(12)

        self.power_enable_checkbox = QCheckBox("启用电源自动化功能")
        self.power_enable_checkbox.setStyleSheet(
            "QCheckBox { font-size: 18px; font-weight: 600; color: #222; }"
            "QCheckBox::indicator { width: 20px; height: 20px; }"
        )
        self.power_enable_checkbox.setChecked(bool(self.config.get("power_enabled", False)))
        self.power_enable_checkbox.toggled.connect(self._on_power_enabled_toggled)
        enable_layout.addWidget(self.power_enable_checkbox)
        enable_layout.addStretch()
        layout.addWidget(enable_row)

        panel = QFrame()
        panel.setStyleSheet("QFrame { background: #fff; border: 1px solid #e0e0e0; border-radius: 10px; }")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(18, 16, 18, 16)
        panel_layout.setSpacing(12)

        row_boot = QHBoxLayout()
        row_boot.setSpacing(10)
        boot_label = QLabel("定时开机登录")
        boot_label.setStyleSheet("font-size: 17px; color: #222;")
        self.boot_freq_combo = QComboBox()
        self.boot_freq_combo.addItems(["每天", "工作日", "周末"])
        self.boot_freq_combo.setCurrentText(power_settings.get("boot_frequency", "每天"))
        self.boot_time_edit = QTimeEdit()
        self.boot_time_edit.setDisplayFormat("HH:mm")
        self.boot_time_edit.setTime(self._parse_hhmm(power_settings.get("boot_time", "06:30"), QTime(6, 30)))
        self.boot_freq_combo.currentTextChanged.connect(self._save_power_settings)
        self.boot_time_edit.timeChanged.connect(self._save_power_settings)
        row_boot.addWidget(boot_label)
        row_boot.addWidget(self.boot_freq_combo)
        row_boot.addWidget(self.boot_time_edit)
        row_boot.addStretch()
        panel_layout.addLayout(row_boot)

        row_shutdown = QHBoxLayout()
        row_shutdown.setSpacing(10)
        shutdown_label = QLabel("定时关机注销")
        shutdown_label.setStyleSheet("font-size: 17px; color: #222;")
        self.shutdown_freq_combo = QComboBox()
        self.shutdown_freq_combo.addItems(["每天", "工作日", "周末"])
        self.shutdown_freq_combo.setCurrentText(power_settings.get("shutdown_frequency", "每天"))
        self.shutdown_time_edit = QTimeEdit()
        self.shutdown_time_edit.setDisplayFormat("HH:mm")
        self.shutdown_time_edit.setTime(self._parse_hhmm(power_settings.get("shutdown_time", "23:00"), QTime(23, 0)))
        self.shutdown_action_combo = QComboBox()
        self.shutdown_action_combo.addItems(["关机", "注销", "重启", "睡眠", "锁屏"])
        self.shutdown_action_combo.setCurrentText(power_settings.get("shutdown_action", "关机"))
        self.shutdown_freq_combo.currentTextChanged.connect(self._save_power_settings)
        self.shutdown_time_edit.timeChanged.connect(self._save_power_settings)
        self.shutdown_action_combo.currentTextChanged.connect(self._save_power_settings)
        row_shutdown.addWidget(shutdown_label)
        row_shutdown.addWidget(self.shutdown_freq_combo)
        row_shutdown.addWidget(self.shutdown_time_edit)
        row_shutdown.addWidget(self.shutdown_action_combo)
        row_shutdown.addStretch()
        panel_layout.addLayout(row_shutdown)

        row_user = QHBoxLayout()
        row_user.setSpacing(10)
        user_label = QLabel("登录账号")
        user_label.setStyleSheet("font-size: 17px; color: #222;")
        self.login_user_input = QLineEdit(power_settings.get("login_user", ""))
        self.login_user_input.setPlaceholderText("Windows 用户名")
        self.login_user_input.editingFinished.connect(self._save_power_settings)
        row_user.addWidget(user_label)
        row_user.addWidget(self.login_user_input)
        row_user.addStretch()
        panel_layout.addLayout(row_user)

        row_pwd = QHBoxLayout()
        row_pwd.setSpacing(10)
        pwd_label = QLabel("登录密码")
        pwd_label.setStyleSheet("font-size: 17px; color: #222;")
        self.login_password_input = QLineEdit(power_settings.get("login_password", ""))
        self.login_password_input.setPlaceholderText("自动登录密码")
        self.login_password_input.setEchoMode(QLineEdit.Password)
        self.login_password_input.editingFinished.connect(self._save_power_settings)
        row_pwd.addWidget(pwd_label)
        row_pwd.addWidget(self.login_password_input)
        panel_layout.addLayout(row_pwd)

        v_title = QLabel("虚拟显示器（息屏识图）")
        v_title.setStyleSheet("font-size: 17px; color: #222; font-weight: 700;")
        panel_layout.addWidget(v_title)

        self.virtual_display_auto_prepare_checkbox = QCheckBox("每次执行 OpenCV 前自动准备虚拟显示器")
        self.virtual_display_auto_prepare_checkbox.setChecked(
            bool(power_settings.get("virtual_display_auto_prepare", False))
        )
        self.virtual_display_auto_prepare_checkbox.toggled.connect(self._save_power_settings)
        panel_layout.addWidget(self.virtual_display_auto_prepare_checkbox)

        self.virtual_display_auto_install_checkbox = QCheckBox("若未检测到虚拟显示器则自动安装驱动（需管理员）")
        self.virtual_display_auto_install_checkbox.setChecked(
            bool(power_settings.get("virtual_display_auto_install", False))
        )
        self.virtual_display_auto_install_checkbox.toggled.connect(self._save_power_settings)
        panel_layout.addWidget(self.virtual_display_auto_install_checkbox)

        self.virtual_display_prepare_on_start_checkbox = QCheckBox("启动 AutoPlua 时自动准备虚拟显示器")
        self.virtual_display_prepare_on_start_checkbox.setChecked(
            bool(power_settings.get("virtual_display_prepare_on_app_start", True))
        )
        self.virtual_display_prepare_on_start_checkbox.toggled.connect(self._save_power_settings)
        panel_layout.addWidget(self.virtual_display_prepare_on_start_checkbox)

        self.virtual_display_strict_isolation_checkbox = QCheckBox("强制在非主显示器执行自动化（失败则中止）")
        self.virtual_display_strict_isolation_checkbox.setChecked(
            bool(power_settings.get("virtual_display_strict_isolation", True))
        )
        self.virtual_display_strict_isolation_checkbox.toggled.connect(self._save_power_settings)
        panel_layout.addWidget(self.virtual_display_strict_isolation_checkbox)

        inf_row = QHBoxLayout()
        inf_row.setSpacing(8)
        inf_label = QLabel("驱动 INF")
        inf_label.setStyleSheet("font-size: 15px; color: #222;")
        self.virtual_display_inf_input = QLineEdit(power_settings.get("virtual_display_driver_inf", ""))
        self.virtual_display_inf_input.setPlaceholderText("可留空使用项目内置驱动；也可手动指定 INF 覆盖")
        self.virtual_display_inf_input.editingFinished.connect(self._save_power_settings)
        inf_pick_btn = QPushButton("选择INF")
        inf_pick_btn.setCursor(Qt.PointingHandCursor)
        inf_pick_btn.clicked.connect(self._pick_virtual_display_inf)
        inf_install_btn = QPushButton("安装并启用")
        inf_install_btn.setCursor(Qt.PointingHandCursor)
        inf_install_btn.clicked.connect(self._install_virtual_display_driver)
        inf_test_btn = QPushButton("检测状态")
        inf_test_btn.setCursor(Qt.PointingHandCursor)
        inf_test_btn.clicked.connect(self._test_virtual_display_ready)
        inf_row.addWidget(inf_label)
        inf_row.addWidget(self.virtual_display_inf_input)
        inf_row.addWidget(inf_pick_btn)
        inf_row.addWidget(inf_install_btn)
        inf_row.addWidget(inf_test_btn)
        panel_layout.addLayout(inf_row)

        layout.addWidget(panel)

        entry_title = QLabel("电源方式")
        entry_title.setStyleSheet("font-size: 24px; font-weight: 700; color: #333;")
        layout.addWidget(entry_title)

        task_btn = self._build_power_nav_button(
            title="Windows 任务计划",
            description="仅跳转页面，不在此处开发具体任务计划功能",
            on_click=partial(self._open_power_subpage, "task"),
        )
        layout.addWidget(task_btn)

        wol_btn = self._build_power_nav_button(
            title="WOL 方式",
            description="仅跳转页面，不在此处开发具体唤醒功能",
            on_click=partial(self._open_power_subpage, "wol"),
        )
        layout.addWidget(wol_btn)

        tip = QLabel("说明：完整的自动开机需要管理员权限、服务常驻和主板/BIOS 支持；当前版本已提供配置与定时触发能力。")
        tip.setWordWrap(True)
        tip.setStyleSheet("font-size: 13px; color: #666;")
        layout.addWidget(tip)

        self._apply_power_schedule()
        layout.addStretch()
        return page

    def _build_power_nav_button(self, title: str, description: str, on_click) -> QPushButton:
        button = QPushButton(f"{title}\n{description}")
        button.setCursor(Qt.PointingHandCursor)
        button.setMinimumHeight(86)
        button.setStyleSheet(
            """
            QPushButton {
                text-align: left;
                padding: 14px 18px;
                border: 1px solid #dedede;
                border-radius: 10px;
                background: #ffffff;
                color: #242424;
                font-size: 16px;
            }
            QPushButton:hover {
                border-color: #b8d8ff;
                background: #f7fbff;
            }
            """
        )
        button.clicked.connect(on_click)
        return button

    def _build_power_placeholder_page(self, title_text: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        top_row = QHBoxLayout()
        back_btn = QPushButton("返回")
        back_btn.setCursor(Qt.PointingHandCursor)
        back_btn.setFixedSize(92, 38)
        back_btn.setStyleSheet(
            "QPushButton { border: 1px solid #cfcfcf; border-radius: 8px; background: #fff; font-size: 15px; }"
            "QPushButton:hover { background: #f5f5f5; }"
        )
        back_btn.clicked.connect(self._back_to_power_home)
        top_row.addWidget(back_btn)
        top_row.addStretch()
        layout.addLayout(top_row)

        title = QLabel(title_text)
        title.setStyleSheet("font-size: 30px; font-weight: 700; color: #222;")
        layout.addWidget(title)

        desc = QLabel("当前仅提供页面跳转占位，不包含具体功能开发。")
        desc.setStyleSheet("font-size: 18px; color: #4d4d4d;")
        layout.addWidget(desc)
        layout.addStretch()
        return page

    def _create_power_icon(self, size: int = 24) -> QIcon:
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing, True)
        color = QColor("#1f2937")

        pen = QPen(color)
        pen.setWidth(max(2, size // 10))
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)

        margin = max(3, size // 8)
        circle_size = size - margin * 2
        painter.drawArc(margin, margin, circle_size, circle_size, 40 * 16, 280 * 16)

        x = size // 2
        top = margin - 1
        bottom = size // 2 + max(1, size // 14)
        painter.drawLine(x, top, x, bottom)

        painter.end()
        return QIcon(pixmap)

    def _build_start_page(self) -> QWidget:
        page = QWidget()
        page.setStyleSheet("background-color: #f9f9f9;")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        header_layout = QHBoxLayout()
        start_label = QLabel("开始")
        start_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #333;")
        header_layout.addWidget(start_label)

        header_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        self.launch_btn = QPushButton("启动")
        self.launch_btn.setCursor(Qt.PointingHandCursor)
        self.launch_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 30px;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """)
        self.launch_btn.clicked.connect(self._toggle_project_runtime)
        header_layout.addWidget(self.launch_btn)
        self._update_launch_button_state()
        layout.addLayout(header_layout)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("background-color: #ccc; max-height: 1px;")
        layout.addWidget(line)

        loop_label = QLabel("每日循环列表")
        loop_label.setStyleSheet("font-size: 44px; font-weight: bold; color: #313131;")
        layout.addWidget(loop_label)

        self.loop_list = QListWidget()
        self.loop_list.setSelectionMode(QListWidget.NoSelection)
        self.loop_list.setSpacing(12)
        self.loop_list.setStyleSheet("""
            QListWidget {
                border: none;
                background-color: transparent;
                outline: 0;
            }
            QListWidget::item {
                background-color: white;
                border-radius: 10px;
                border: 1px solid #e0e0e0;
            }
        """)
        layout.addWidget(self.loop_list)

        self.start_add_btn = QPushButton("+", page)
        self.start_add_btn.setFixedSize(60, 60)
        self.start_add_btn.setCursor(Qt.PointingHandCursor)
        self.start_add_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border-radius: 30px;
                font-size: 30px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """)
        self.start_add_btn.clicked.connect(self._add_exe_file)
        self.start_add_btn.raise_()
        QTimer.singleShot(0, self._position_start_add_button)
        return page

    def _position_start_add_button(self) -> None:
        if not hasattr(self, "start_add_btn") or not hasattr(self, "start_page"):
            return
        page = self.start_page
        margin_right = 28
        margin_bottom = 22
        x = max(0, page.width() - self.start_add_btn.width() - margin_right)
        y = max(0, page.height() - self.start_add_btn.height() - margin_bottom)
        self.start_add_btn.move(x, y)
        self.start_add_btn.raise_()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._position_start_add_button()

    def _build_log_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(16)

        title = QLabel("日志")
        title.setStyleSheet("font-size: 34px; font-weight: 700; color: #222;")
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #ccc; max-height: 1px;")
        layout.addWidget(line)

        app_label = QLabel("AutoPlua 日志")
        app_label.setStyleSheet("font-size: 18px; font-weight: 700; color: #333;")
        layout.addWidget(app_label)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet(
            "QTextEdit { background: #fff; border: 1px solid #d9d9d9; border-radius: 10px; font-size: 14px; }"
        )
        self.log_text.verticalScrollBar().valueChanged.connect(self._on_app_log_scroll_changed)
        layout.addWidget(self.log_text)

        program_label = QLabel("被启动程序日志（含 Python 终端输出）")
        program_label.setStyleSheet("font-size: 18px; font-weight: 700; color: #333;")
        layout.addWidget(program_label)

        self.program_log_text = QTextEdit()
        self.program_log_text.setReadOnly(True)
        self.program_log_text.setStyleSheet(
            "QTextEdit { background: #fff; border: 1px solid #d9d9d9; border-radius: 10px; font-size: 14px; }"
        )
        self.program_log_text.verticalScrollBar().valueChanged.connect(self._on_program_log_scroll_changed)
        layout.addWidget(self.program_log_text)

        row = QHBoxLayout()
        refresh_btn = QPushButton("刷新")
        refresh_btn.setCursor(Qt.PointingHandCursor)
        refresh_btn.clicked.connect(self._refresh_log_page)
        clear_btn = QPushButton("清空")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self._clear_logs)
        row.addWidget(refresh_btn)
        row.addWidget(clear_btn)
        row.addStretch()
        layout.addLayout(row)

        return page

    def _build_about_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(14)

        title = QLabel("关于")
        title.setStyleSheet("font-size: 34px; font-weight: 700; color: #222;")
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background-color: #ccc; max-height: 1px;")
        layout.addWidget(line)

        intro = QLabel("AutoPlua 是一个 Windows 自动化控制工具。")
        intro.setStyleSheet("font-size: 18px; color: #333;")
        layout.addWidget(intro)

        feature = QLabel("功能模块：程序启动/停止、定时调度、电源控制、日志管理")
        feature.setStyleSheet("font-size: 16px; color: #4a4a4a;")
        layout.addWidget(feature)

        version = QLabel("版本：MVP")
        version.setStyleSheet("font-size: 16px; color: #4a4a4a;")
        layout.addWidget(version)

        layout.addStretch()
        return page

    def _switch_page(self, index: int, checked: bool = False) -> None:
        for i, btn in enumerate(self.nav_buttons):
            btn.setChecked(i == index)
        self.stacked_widget.setCurrentIndex(index)
        if index == 2:
            self._refresh_log_page()

    def _toggle_sidebar(self) -> None:
        self.is_sidebar_collapsed = not self.is_sidebar_collapsed
        if self.is_sidebar_collapsed:
            self.sidebarWidget.setFixedWidth(64)
            for btn in self.nav_buttons:
                btn.setText("")
                btn.setStyleSheet(SIDEBAR_COLLAPSED_STYLE)
        else:
            self.sidebarWidget.setFixedWidth(160)
            for i, btn in enumerate(self.nav_buttons):
                btn.setText(self.nav_items_text[i])
                btn.setStyleSheet(SIDEBAR_EXPANDED_STYLE)

    def _on_power_enabled_toggled(self, enabled: bool) -> None:
        self.config["power_enabled"] = enabled
        save_config(self.config)
        self._apply_power_schedule()

    def _open_power_subpage(self, page_type: str) -> None:
        if page_type == "task":
            self.power_stack.setCurrentWidget(self.power_task_page)
        elif page_type == "wol":
            self.power_stack.setCurrentWidget(self.power_wol_page)

    def _back_to_power_home(self) -> None:
        self.power_stack.setCurrentWidget(self.power_home_page)

    def _add_exe_file(self) -> None:
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "选择可执行文件",
            "",
            "Executable Files (*.exe);;All Files (*)"
        )
        if not filepath:
            return

        exe_name = os.path.basename(filepath)
        entry = {
            "name": exe_name,
            "command": filepath,
            "args": [],
            "launch_args_raw": "",
            "input_mode": "foreground",
            "target_window_title": "",
            "opencv_step_retry_seconds": 30,
            "cwd": str(Path(filepath).parent),
            "enabled": True,
            "start_time": "00:00",
            "end_time": "00:00",
            "time_points": [{"start": "00:00", "end": "00:00"}],
            "opencv_flow": {"nodes": [], "edges": []},
        }
        self.program_entries.append(entry)
        self._refresh_program_list()
        self._save_programs_to_config()
        self._append_log(f"已添加程序：{exe_name}")

    def _add_program_item(self, entry: dict) -> None:
        item_widget = DailyLoopItemWidget(
            entry=entry,
            on_toggle_enabled=partial(self._toggle_program_enabled, entry),
            on_config_clicked=partial(self._open_program_config, entry),
            on_start_clicked=partial(self._start_single_program, entry),
            on_stop_clicked=partial(self._stop_single_program, entry),
            on_remove_clicked=partial(self._remove_program, entry),
        )
        list_item = QListWidgetItem(self.loop_list)
        list_item.setSizeHint(QSize(0, 122))
        self.loop_list.addItem(list_item)
        self.loop_list.setItemWidget(list_item, item_widget)

    def _toggle_program_enabled(self, entry: dict, enabled: bool) -> None:
        if entry in self.program_entries:
            entry["enabled"] = enabled
            self._save_programs_to_config()

    def _open_program_config(self, entry: dict) -> None:
        if entry not in self.program_entries:
            return
        dialog = ProgramConfigDialog(entry=entry, parent=self)
        if dialog.exec() != QDialog.Accepted or not dialog.result_data:
            return

        entry["launch_args_raw"] = dialog.result_data.get("launch_args_raw", "")
        entry["args"] = dialog.result_data.get("args", [])
        entry["input_mode"] = dialog.result_data.get("input_mode", "foreground")
        entry["target_window_title"] = dialog.result_data.get("target_window_title", "")
        entry["opencv_step_retry_seconds"] = int(dialog.result_data.get("opencv_step_retry_seconds", 30))
        time_points = dialog.result_data.get("time_points", [])
        if isinstance(time_points, list) and time_points:
            entry["time_points"] = time_points
            first = time_points[0] if isinstance(time_points[0], dict) else {}
            entry["start_time"] = str(first.get("start", entry.get("start_time", "00:00")))
            entry["end_time"] = str(first.get("end", entry.get("end_time", "00:00")))
        entry["opencv_flow"] = dialog.result_data.get("opencv_flow", {"nodes": [], "edges": []})
        self._refresh_program_list()
        self._save_programs_to_config()
        self._append_log(f"已更新配置：{entry.get('name', '')}")

    def _load_programs_from_config(self) -> None:
        saved_programs = self.config.get("programs", [])
        if not isinstance(saved_programs, list):
            return

        self.program_entries = []
        self.loop_list.clear()

        for raw in saved_programs:
            if not isinstance(raw, dict):
                continue
            command = raw.get("command", "")
            name = raw.get("name") or os.path.basename(command)
            loaded_points = raw.get("time_points", [])
            if not isinstance(loaded_points, list) or not loaded_points:
                loaded_points = [{
                    "start": raw.get("start_time", "00:00"),
                    "end": raw.get("end_time", "00:00"),
                }]
            entry = {
                "name": name,
                "command": command,
                "args": raw.get("args", []),
                "launch_args_raw": raw.get("launch_args_raw", ""),
                "input_mode": raw.get("input_mode", "background_window_message"),
                "target_window_title": raw.get("target_window_title", Path(name).stem),
                "opencv_step_retry_seconds": int(raw.get("opencv_step_retry_seconds", 30) or 30),
                "cwd": raw.get("cwd") or (str(Path(command).parent) if command else ""),
                "enabled": bool(raw.get("enabled", True)),
                "start_time": raw.get("start_time", "00:00"),
                "end_time": raw.get("end_time", "00:00"),
                "time_points": loaded_points,
                "opencv_flow": raw.get("opencv_flow", {"nodes": [], "edges": []}),
            }
            self.program_entries.append(entry)

        self._refresh_program_list()
        self._save_programs_to_config()

    def _refresh_program_list(self) -> None:
        self.loop_list.clear()
        for entry in self.program_entries:
            self._add_program_item(entry)

    def _start_single_program(self, entry: dict, trigger: str = "manual") -> None:
        command = entry.get("command", "")
        if not command:
            return
        try:
            resolved_args = self._resolve_program_args(entry)
            managed = ManagedProgram(
                name=entry.get("name") or os.path.basename(command),
                command=command,
                args=resolved_args,
                cwd=entry.get("cwd") or None,
            )
            self.process_service.start(managed)
            prefix = "定时启动" if trigger == "schedule" else "手动启动"
            self._append_log(f"{prefix}：{managed.name}")
            self._run_post_launch_flow(entry, managed.name)
        except Exception as exc:
            prefix = "定时启动失败" if trigger == "schedule" else "手动启动失败"
            self._append_log(f"{prefix}：{entry.get('name', command)} - {exc}")

    def _stop_single_program(self, entry: dict, trigger: str = "manual") -> None:
        name = entry.get("name") or os.path.basename(entry.get("command", ""))
        command = entry.get("command", "")
        if not name:
            return
        stopped = self.process_service.stop(name, command=command)
        if stopped:
            prefix = "定时结束" if trigger == "schedule" else "手动结束"
            self._append_log(f"{prefix}：{name}")
        else:
            self._append_log(f"未找到运行中的程序：{name}")

    def _remove_program(self, entry: dict) -> None:
        name = entry.get("name") or os.path.basename(entry.get("command", ""))
        result = QMessageBox.question(
            self,
            "确认移除",
            f"确定要移除程序：{name} ？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if result != QMessageBox.Yes:
            return

        if entry in self.program_entries:
            self.program_entries.remove(entry)
            self._refresh_program_list()
            self._save_programs_to_config()
            self._append_log(f"已移除程序：{name}")

    def _save_programs_to_config(self) -> None:
        self.config["programs"] = self.program_entries
        save_config(self.config)

    def _toggle_project_runtime(self) -> None:
        enabled_targets = [entry for entry in self.program_entries if entry.get("enabled")]
        if not self._project_runtime_active and not enabled_targets:
            QMessageBox.information(self, "提示", "请先勾选至少一个项目，再启动运行状态。")
            return

        if self._project_runtime_active:
            self._stop_project_runtime()
        else:
            self._start_project_runtime()

    def _start_project_runtime(self) -> None:
        self._project_runtime_active = True
        self._project_runtime_marks.clear()
        self._update_launch_button_state()
        self._append_log("项目调度运行已开启，任务将按配置时间点执行。")

        try:
            self.scheduler_service.add_interval_job(
                job_id=self._program_runtime_job_id,
                seconds=20,
                func=self._program_runtime_tick,
            )
        except Exception as exc:
            self._project_runtime_active = False
            self._update_launch_button_state()
            self._append_log(f"项目调度运行开启失败：{exc}")

    def _stop_project_runtime(self) -> None:
        self._project_runtime_active = False
        self._update_launch_button_state()
        try:
            self.scheduler_service.remove_job(self._program_runtime_job_id)
        except Exception:
            pass
        self._append_log("项目调度运行已停止。")

    def _update_launch_button_state(self) -> None:
        if self._project_runtime_active:
            self.launch_btn.setText("运行中")
            self.launch_btn.setStyleSheet(
                "QPushButton { background-color: #16a34a; color: white; border: none; border-radius: 6px;"
                " padding: 10px 30px; font-size: 16px; font-weight: bold; }"
                "QPushButton:hover { background-color: #15803d; }"
            )
            return

        self.launch_btn.setText("启动")
        self.launch_btn.setStyleSheet(
            "QPushButton { background-color: #0078d7; color: white; border: none; border-radius: 6px;"
            " padding: 10px 30px; font-size: 16px; font-weight: bold; }"
            "QPushButton:hover { background-color: #005a9e; }"
        )

    def _program_runtime_tick(self) -> None:
        if not self._project_runtime_active:
            return

        now = datetime.now()
        current_hhmm = now.strftime("%H:%M")
        day_key = now.strftime("%Y-%m-%d")

        for entry in self.program_entries:
            if not entry.get("enabled"):
                continue

            name = entry.get("name") or os.path.basename(entry.get("command", ""))
            if not name:
                continue

            time_points = entry.get("time_points", [])
            if not isinstance(time_points, list) or not time_points:
                time_points = [{
                    "start": entry.get("start_time", ""),
                    "end": entry.get("end_time", ""),
                }]

            for idx, point in enumerate(time_points):
                if not isinstance(point, dict):
                    continue

                start_hhmm = str(point.get("start", "")).strip()
                end_hhmm = str(point.get("end", "")).strip()

                if start_hhmm and current_hhmm == start_hhmm:
                    start_mark = f"start|{name}|{idx}|{start_hhmm}"
                    if self._project_runtime_marks.get(start_mark) != day_key:
                        self._project_runtime_marks[start_mark] = day_key
                        self.schedule_start_signal.emit(entry)

                if end_hhmm and current_hhmm == end_hhmm:
                    end_mark = f"end|{name}|{idx}|{end_hhmm}"
                    if self._project_runtime_marks.get(end_mark) != day_key:
                        self._project_runtime_marks[end_mark] = day_key
                        self.schedule_stop_signal.emit(entry)

    @staticmethod
    def _resolve_program_args(entry: dict) -> list[str]:
        raw = str(entry.get("launch_args_raw", "")).strip()
        if raw:
            try:
                return shlex.split(raw, posix=False)
            except ValueError:
                return entry.get("args", []) if isinstance(entry.get("args", []), list) else []

        args = entry.get("args", [])
        return args if isinstance(args, list) else []

    def _run_post_launch_flow(self, entry: dict, program_name: str) -> None:
        launch_args = self._resolve_program_args(entry)
        if launch_args:
            self._append_log(f"{program_name} 使用启动参数：{' '.join(launch_args)}")

        if not self._prepare_virtual_display_for_flow(program_name):
            return

        flow = entry.get("opencv_flow", {})
        if not isinstance(flow, dict) or not flow.get("nodes"):
            self._append_log(f"{program_name} 未配置 OpenCV 流程")
            return

        nodes_count = len(flow.get("nodes", [])) if isinstance(flow.get("nodes", []), list) else 0
        self._append_log(f"{program_name} 开始执行 OpenCV 流程，节点数：{nodes_count}")

        input_mode = str(entry.get("input_mode", "foreground")).strip() or "foreground"
        settings = self._get_power_settings()
        strict_isolation = bool(settings.get("virtual_display_strict_isolation", True))
        isolation_display_ready = self.virtual_display_service.has_non_primary_monitor()
        if strict_isolation and isolation_display_ready and input_mode != "background_window_message":
            input_mode = "background_window_message"
            self._append_log(
                f"{program_name} 已启用强制隔离执行，输入模式自动切换为后台窗口消息。"
            )
        elif strict_isolation and not isolation_display_ready:
            self._append_log(
                f"{program_name} 未检测到可用虚拟显示器，已回退到原执行逻辑。"
            )

        success, message = self.opencv_flow_service.run_flow(
            flow,
            timeout_seconds=180,
            default_wait_seconds=2,
            startup_wait_seconds=3,
            step_retry_seconds=max(5, int(entry.get("opencv_step_retry_seconds", 30) or 30)),
            execution_options={
                "input_mode": input_mode,
                "target_window_title": entry.get("target_window_title", ""),
                "target_pid": self.process_service.get_running_pid(program_name),
                "target_process_name": program_name,
            },
        )
        if success:
            if message != "ok":
                self._append_log(f"{program_name} OpenCV 流程执行完成：{message}")
            else:
                self._append_log(f"{program_name} OpenCV 流程执行完成")
            return

        if message == "missing-dependency-pyautogui":
            self._append_log(
                f"{program_name} OpenCV 流程失败：缺少依赖 pyautogui，请执行 pip install pyautogui"
            )
            return

        if message == "screen-capture-unavailable-possibly-screen-off-or-locked":
            self._append_log(
                f"{program_name} OpenCV 流程失败：当前截图源不可用（常见于息屏/锁屏）。"
                "请接入虚拟显示器驱动，或保持目标桌面持续渲染。"
            )
            return

        if message == "target-window-not-found":
            self._append_log(
                f"{program_name} OpenCV 流程失败：未找到目标窗口标题。"
                "请在程序配置中填写与系统窗口标题完全一致的文本。"
            )
            return

        if message.startswith("template-image-not-found:"):
            missing = message.split(":", 1)[1]
            self._append_log(f"{program_name} OpenCV 流程失败：模板图片不存在 -> {missing}")
            return

        if message.endswith("click-target-not-found"):
            self._append_log(
                f"{program_name} OpenCV 流程失败：未识别到点击目标。"
                "请确认目标窗口标题、模板图片路径、以及虚拟显示器是否已启用扩展显示。"
            )
            return

        if "click-target-not-found-score-" in message:
            score = message.split("click-target-not-found-score-", 1)[1]
            self._append_log(
                f"{program_name} OpenCV 流程失败：未识别到点击目标（最高相似度={score}）。"
                "后台模式下建议模板相似度至少 0.80；请重截模板并确保窗口状态一致。"
            )
            return

        if message == "background-target-window-not-found":
            self._append_log(
                f"{program_name} OpenCV 流程失败：后台目标窗口未找到。"
                "请确认窗口标题，或等待程序主窗口完全创建后再执行。"
            )
            return

        if "source=window-minimized" in message or "window-minimized" in message:
            self._append_log(
                f"{program_name} OpenCV 流程失败：目标窗口处于最小化状态，后台截图不可用。"
                "请保持窗口非最小化（可被遮挡/切到后台），或放到虚拟显示器上运行。"
            )
            return

        self._append_log(f"{program_name} OpenCV 流程失败：{message}")

    def _append_log(self, message: str) -> None:
        self.app_log_signal.emit(message)

    def _append_program_log(self, message: str) -> None:
        self.program_log_signal.emit(message)

    def _handle_append_log(self, message: str) -> None:
        self.logger.info(message)
        self.runtime_logs.append(message)
        if hasattr(self, "log_text"):
            self.log_text.append(message)
            if self._app_log_follow_tail:
                self._scroll_textedit_to_bottom(self.log_text)

    def _handle_append_program_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted = f"{timestamp} {message}"
        self.program_runtime_logs.append(formatted)

        try:
            self._program_log_file_path().parent.mkdir(parents=True, exist_ok=True)
            with self._program_log_file_path().open("a", encoding="utf-8") as f:
                f.write(formatted + "\n")
        except OSError:
            pass

        if hasattr(self, "program_log_text"):
            self.program_log_text.append(formatted)
            if self._program_log_follow_tail:
                self._scroll_textedit_to_bottom(self.program_log_text)

    def _handle_schedule_start(self, entry: dict) -> None:
        self._start_single_program(entry, trigger="schedule")

    def _handle_schedule_stop(self, entry: dict) -> None:
        self._stop_single_program(entry, trigger="schedule")

    def _on_program_output(self, program_name: str, line: str) -> None:
        self._append_program_log(f"[{program_name}] {line}")

    def _on_program_exit(self, program_name: str, code: int | None) -> None:
        code_text = "unknown" if code is None else str(code)
        self._append_program_log(f"[{program_name}] [进程退出] code={code_text}")

    @staticmethod
    def _program_log_file_path() -> Path:
        appdata = os.getenv("APPDATA")
        base = Path(appdata) if appdata else (Path.home() / ".autoplua")
        return base / "AutoPlua" / "logs" / "program.log"

    def _refresh_log_page(self) -> None:
        if not hasattr(self, "log_text"):
            return

        lines = []
        lines.extend(self.runtime_logs)

        try:
            log_path = Path(os.getenv("APPDATA", "")) / "AutoPlua" / "logs" / "app.log"
            if log_path.exists():
                file_lines = log_path.read_text(encoding="utf-8", errors="ignore").splitlines()
                lines.extend(file_lines[-200:])
        except OSError:
            pass

        if not lines:
            self.log_text.setPlainText("暂无日志")
            return

        seen = set()
        unique_lines = []
        for line in lines:
            if line not in seen:
                unique_lines.append(line)
                seen.add(line)
        self.log_text.setPlainText("\n".join(unique_lines[-300:]))
        if self._app_log_follow_tail:
            self._scroll_textedit_to_bottom(self.log_text)

        if hasattr(self, "program_log_text"):
            program_lines = []
            program_lines.extend(self.program_runtime_logs)
            try:
                p_log = self._program_log_file_path()
                if p_log.exists():
                    file_lines = p_log.read_text(encoding="utf-8", errors="ignore").splitlines()
                    program_lines.extend(file_lines[-400:])
            except OSError:
                pass

            if not program_lines:
                self.program_log_text.setPlainText("暂无程序日志")
            else:
                seen_program = set()
                unique_program = []
                for line in program_lines:
                    if line not in seen_program:
                        unique_program.append(line)
                        seen_program.add(line)
                self.program_log_text.setPlainText("\n".join(unique_program[-600:]))
                if self._program_log_follow_tail:
                    self._scroll_textedit_to_bottom(self.program_log_text)

    def _clear_logs(self) -> None:
        self.runtime_logs.clear()
        self.program_runtime_logs.clear()
        if hasattr(self, "log_text"):
            self.log_text.clear()
            self._app_log_follow_tail = True
        if hasattr(self, "program_log_text"):
            self.program_log_text.clear()
            self._program_log_follow_tail = True
        try:
            p_log = self._program_log_file_path()
            if p_log.exists():
                p_log.unlink()
        except OSError:
            pass

    def _on_app_log_scroll_changed(self, value: int) -> None:
        if not hasattr(self, "log_text"):
            return
        bar = self.log_text.verticalScrollBar()
        self._app_log_follow_tail = value >= (bar.maximum() - 4)

    def _on_program_log_scroll_changed(self, value: int) -> None:
        if not hasattr(self, "program_log_text"):
            return
        bar = self.program_log_text.verticalScrollBar()
        self._program_log_follow_tail = value >= (bar.maximum() - 4)

    @staticmethod
    def _scroll_textedit_to_bottom(editor: QTextEdit) -> None:
        bar = editor.verticalScrollBar()
        bar.setValue(bar.maximum())

    def _get_power_settings(self) -> dict:
        raw = self.config.get("power_settings", {})
        if not isinstance(raw, dict):
            raw = {}
        defaults = {
            "boot_frequency": "每天",
            "boot_time": "06:30",
            "shutdown_frequency": "每天",
            "shutdown_time": "23:00",
            "shutdown_action": "关机",
            "login_user": "",
            "login_domain": "",
            "login_password": "",
            "wake_mode": "Windows任务计划",
            "wol_mac": "",
            "wol_host": "255.255.255.255",
            "virtual_display_auto_prepare": False,
            "virtual_display_auto_install": False,
            "virtual_display_prepare_on_app_start": True,
            "virtual_display_strict_isolation": True,
            "virtual_display_driver_inf": "",
        }
        return {**defaults, **raw}

    @staticmethod
    def _parse_hhmm(value: str, fallback: QTime) -> QTime:
        parsed = QTime.fromString(value, "HH:mm")
        return parsed if parsed.isValid() else fallback

    def _save_power_settings(self) -> None:
        if not hasattr(self, "boot_freq_combo"):
            return

        settings = {
            "boot_frequency": self.boot_freq_combo.currentText(),
            "boot_time": self.boot_time_edit.time().toString("HH:mm"),
            "shutdown_frequency": self.shutdown_freq_combo.currentText(),
            "shutdown_time": self.shutdown_time_edit.time().toString("HH:mm"),
            "shutdown_action": self.shutdown_action_combo.currentText(),
            "login_user": self.login_user_input.text().strip(),
            "login_password": self.login_password_input.text(),
            "virtual_display_auto_prepare": bool(self.virtual_display_auto_prepare_checkbox.isChecked()),
            "virtual_display_auto_install": bool(self.virtual_display_auto_install_checkbox.isChecked()),
            "virtual_display_prepare_on_app_start": bool(self.virtual_display_prepare_on_start_checkbox.isChecked()),
            "virtual_display_strict_isolation": bool(self.virtual_display_strict_isolation_checkbox.isChecked()),
            "virtual_display_driver_inf": self.virtual_display_inf_input.text().strip(),
        }
        self.config["power_settings"] = settings
        save_config(self.config)
        self._append_log("电源自动化设置已保存")
        self._apply_power_schedule()

    def _apply_power_schedule(self) -> None:
        try:
            if self.config.get("power_enabled", False):
                self.scheduler_service.add_interval_job(
                    job_id="power_automation_tick",
                    seconds=30,
                    func=self._power_automation_tick,
                )
                self._last_power_tick_at = datetime.now() - timedelta(seconds=35)
                self._schedule_next_wake_if_possible()
                self._append_log("电源自动化调度已启用")
            else:
                self.scheduler_service.remove_job("power_automation_tick")
                self.power_service.cancel_wake_timer()
                self._scheduled_wake_marker = ""
                self._append_log("电源自动化调度已停用")
        except Exception:
            pass

    def _power_automation_tick(self) -> None:
        if not self.config.get("power_enabled", False):
            return

        settings = self._get_power_settings()
        now = datetime.now()
        window_start = self._last_power_tick_at or (now - timedelta(seconds=35))
        self._last_power_tick_at = now

        boot_hit, boot_target = self._time_hit_in_window(
            settings.get("boot_frequency", "每天"),
            settings.get("boot_time", "06:30"),
            window_start,
            now,
        )
        if boot_hit and boot_target is not None:
            day_key = boot_target.strftime("%Y-%m-%d")
            if self._power_state.get("boot") != day_key:
                self._power_state["boot"] = day_key
                self._execute_boot_login(settings)

        shutdown_hit, shutdown_target = self._time_hit_in_window(
            settings.get("shutdown_frequency", "每天"),
            settings.get("shutdown_time", "23:00"),
            window_start,
            now,
        )
        if shutdown_hit and shutdown_target is not None:
            day_key = shutdown_target.strftime("%Y-%m-%d")
            if self._power_state.get("shutdown") != day_key:
                self._power_state["shutdown"] = day_key
                self._execute_shutdown_action(settings.get("shutdown_action", "关机"))

        self._schedule_next_wake_if_possible(now=now)

    @staticmethod
    def _freq_matches(freq: str, now: datetime) -> bool:
        weekday = now.weekday()
        if freq == "每天":
            return True
        if freq == "工作日":
            return weekday < 5
        if freq == "周末":
            return weekday >= 5
        return True

    def _execute_boot_login(self, settings: dict) -> None:
        self._append_log("触发定时开机登录任务")

        user = settings.get("login_user", "")
        password = settings.get("login_password", "")
        if user and password:
            self._configure_auto_login_registry(user, password)

    def _execute_shutdown_action(self, action: str) -> None:
        self._append_log(f"触发定时电源动作：{action}")
        if action == "关机":
            self.power_service.shutdown()
        elif action == "注销":
            subprocess.run(["shutdown", "/l"], check=False)
        elif action == "重启":
            self.power_service.restart()
        elif action == "睡眠":
            self.power_service.sleep()
        elif action == "锁屏":
            self.power_service.lock()

    def _configure_auto_login_registry(self, user: str, password: str) -> None:
        domain = os.getenv("COMPUTERNAME", "")
        pairs = {
            "AutoAdminLogon": "1",
            "DefaultUserName": user,
            "DefaultPassword": password,
        }
        if domain:
            pairs["DefaultDomainName"] = domain

        for key, value in pairs.items():
            subprocess.run(
                [
                    "reg",
                    "add",
                    r"HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon",
                    "/v",
                    key,
                    "/t",
                    "REG_SZ",
                    "/d",
                    value,
                    "/f",
                ],
                check=False,
                capture_output=True,
            )
        self._append_log("已尝试写入自动登录注册表（需管理员权限）")

    @staticmethod
    def _time_hit_in_window(freq: str, hhmm: str, window_start: datetime, window_end: datetime) -> tuple[bool, datetime | None]:
        try:
            target_time = datetime.strptime(hhmm.strip(), "%H:%M").time()
        except ValueError:
            return False, None

        start = min(window_start, window_end)
        end = max(window_start, window_end)
        days = (end.date() - start.date()).days

        for delta in range(days + 1):
            day_dt = start + timedelta(days=delta)
            candidate = datetime.combine(day_dt.date(), target_time)
            if not (start < candidate <= end):
                continue
            if MainWindow._freq_matches(freq, candidate):
                return True, candidate

        return False, None

    def _next_occurrence(self, freq: str, hhmm: str, now: datetime) -> datetime | None:
        try:
            target_time = datetime.strptime(hhmm.strip(), "%H:%M").time()
        except ValueError:
            return None

        for offset in range(0, 8):
            day = (now + timedelta(days=offset)).date()
            candidate = datetime.combine(day, target_time)
            if candidate <= now:
                continue
            if self._freq_matches(freq, candidate):
                return candidate
        return None

    def _schedule_next_wake_if_possible(self, now: datetime | None = None) -> None:
        if not self.config.get("power_enabled", False):
            return

        settings = self._get_power_settings()
        now = now or datetime.now()
        next_boot = self._next_occurrence(
            settings.get("boot_frequency", "每天"),
            settings.get("boot_time", "06:30"),
            now,
        )
        if next_boot is None:
            return

        wake_target = next_boot - timedelta(minutes=1)
        if wake_target <= now:
            wake_target = now + timedelta(seconds=15)

        marker = wake_target.strftime("%Y-%m-%d %H:%M")
        if marker == self._scheduled_wake_marker:
            return

        if self.power_service.schedule_wake(wake_target):
            self._scheduled_wake_marker = marker
            self._append_log(f"已安排下一次系统唤醒：{wake_target.strftime('%Y-%m-%d %H:%M')}")
        else:
            self._append_log("系统唤醒定时器设置失败（可能缺少权限或系统策略不支持）")

    def _send_wol_packet(self, mac: str, host: str = "255.255.255.255", port: int = 9) -> None:
        cleaned = mac.replace(":", "").replace("-", "").strip()
        if len(cleaned) != 12:
            self._append_log("WOL 发送失败：MAC 地址格式不正确")
            return

        payload = b"FF" * 6 + (cleaned.encode("ascii") * 16)
        packet = struct.pack("B" * 6, *[0xFF] * 6) + bytes.fromhex(cleaned) * 16
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
                sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
                sock.sendto(packet, (host, port))
            self._append_log(f"已发送 WOL 魔术包到 {host}:{port}")
        except OSError as exc:
            self._append_log(f"WOL 发送失败：{exc}")

    def _test_wol(self) -> None:
        self._save_power_settings()
        settings = self._get_power_settings()
        self._send_wol_packet(settings.get("wol_mac", ""), settings.get("wol_host", "255.255.255.255"))

    def _pick_virtual_display_inf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择虚拟显示器驱动 INF",
            "",
            "INF Files (*.inf);;All Files (*)",
        )
        if not path:
            return
        self.virtual_display_inf_input.setText(path)
        self._save_power_settings()

    def _install_virtual_display_driver(self) -> None:
        self._save_power_settings()
        settings = self._get_power_settings()
        inf = settings.get("virtual_display_driver_inf", "")
        ok, message = self.virtual_display_service.install_driver_from_inf(inf)
        if not ok:
            if message == "admin-required":
                self._append_log("虚拟显示器安装失败：需要管理员权限运行 AutoPlua。")
            elif message == "invalid-inf-path":
                self._append_log("虚拟显示器安装失败：INF 路径无效。")
            elif message == "embedded-driver-not-found":
                self._append_log(
                    "虚拟显示器安装失败：未找到项目内置驱动。"
                    "请将驱动 INF 放到 drivers/virtual_display 目录，或手动选择 INF。"
                )
            elif message == "driver-package-added-but-device-not-created":
                self._append_log(
                    "虚拟显示器安装失败：驱动包已导入系统，但未创建设备实例。"
                    "当前仅用 pnputil 安装不会弹安装窗口；若设备管理器中没有 'Virtual Display Driver/MttVDD'，"
                    "需使用驱动作者提供的控制程序或脚本创建 Root\\MttVDD 设备。"
                )
            elif message.startswith("virtual-driver-device-create-failed:"):
                detail = message.split(":", 1)[1]
                self._append_log(
                    "虚拟显示器安装失败：自动创建设备实例失败。"
                    f"系统错误细节={detail}。"
                    "请以管理员权限运行，并确认系统未拦截未签名/不受信任驱动。"
                )
            else:
                self._append_log(f"虚拟显示器安装失败：{message}")
            return

        ok_extend, msg_extend = self.virtual_display_service.enable_extend_mode()
        if ok_extend:
            self._append_log("虚拟显示器驱动安装成功，已切换为扩展显示模式。")
            return
        self._append_log(f"虚拟显示器驱动已安装，但扩展显示失败：{msg_extend}")

    def _test_virtual_display_ready(self) -> None:
        present = self.virtual_display_service.is_virtual_display_present()
        if present:
            ok_extend, msg = self.virtual_display_service.enable_extend_mode()
            if ok_extend:
                self._append_log("虚拟显示器检测通过，扩展显示已启用。")
            else:
                self._append_log(f"检测到虚拟显示器，但扩展显示启用失败：{msg}")
        else:
            self._append_log("未检测到虚拟显示器。可先选择 INF 后点击“安装并启用”。")

    def _prepare_virtual_display_for_flow(self, program_name: str) -> bool:
        settings = self._get_power_settings()
        strict_isolation = bool(settings.get("virtual_display_strict_isolation", True))
        if not bool(settings.get("virtual_display_auto_prepare", False)):
            if strict_isolation and not self.virtual_display_service.has_non_primary_monitor():
                self._append_log(
                    f"{program_name} 已开启强制隔离，但当前无可用非主显示器。"
                    "检测到未接入虚拟显示器，已回退到原执行逻辑。"
                )
                return True
            return True

        inf = str(settings.get("virtual_display_driver_inf", "")).strip()
        auto_install = bool(settings.get("virtual_display_auto_install", False))
        ok, message = self.virtual_display_service.ensure_automation_display_ready(
            inf_path=inf,
            auto_install=auto_install,
        )
        if ok:
            if message == "installed-and-ready":
                self._append_log(f"{program_name} 已自动安装并启用虚拟显示器。")
            else:
                self._append_log(f"{program_name} 虚拟显示器就绪。")
            return True

        if message == "virtual-display-not-present":
            self._append_log(
                f"{program_name} 未检测到虚拟显示器，且未开启自动安装。"
                "已回退到原执行逻辑。"
                "如需息屏识图，请开启自动安装；INF 可留空使用项目内置驱动。"
            )
            return True

        if message in {
            "virtual-display-not-detected",
            "driver-package-added-but-device-not-created",
        } or message.startswith("virtual-driver-device-create-failed:"):
            self._append_log(
                f"{program_name} 虚拟显示器未成功安装，已回退到原执行逻辑：{message}"
            )
            return True

        if message == "virtual-display-present-but-not-extended":
            self._append_log(
                f"{program_name} 检测到虚拟显示驱动，但未形成扩展桌面。"
                "请检查显示设置中的多显示器模式是否为“扩展这些显示器”。"
            )
            return not strict_isolation

        self._append_log(f"{program_name} 虚拟显示器准备失败：{message}")
        return not strict_isolation

    def _bootstrap_virtual_display_on_app_start(self) -> None:
        settings = self._get_power_settings()
        if not bool(settings.get("virtual_display_prepare_on_app_start", True)):
            return

        inf = str(settings.get("virtual_display_driver_inf", "")).strip()
        auto_install = bool(settings.get("virtual_display_auto_install", False))
        ok, message = self.virtual_display_service.ensure_automation_display_ready(
            inf_path=inf,
            auto_install=auto_install,
        )
        if ok:
            self._append_log("应用启动时已完成虚拟显示器链路准备。")
            return

        if message == "virtual-display-not-present":
            self._append_log(
                "应用启动时未检测到虚拟显示器，且未开启自动安装。"
                "后续可在电源页开启自动安装（INF 可留空使用项目内置驱动）。"
            )
            return

        self._append_log(f"应用启动时虚拟显示器准备失败：{message}")

