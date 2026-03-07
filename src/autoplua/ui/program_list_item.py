from __future__ import annotations

import os

from PySide6.QtCore import Qt, QSize, QFileInfo
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QSpacerItem,
    QVBoxLayout,
    QWidget,
    QFileIconProvider,
)


class DailyLoopItemWidget(QWidget):
    def __init__(
        self,
        entry: dict,
        on_toggle_enabled,
        on_config_clicked,
        on_start_clicked,
        on_stop_clicked,
        on_remove_clicked,
        parent=None,
    ):
        super().__init__(parent)
        self.entry = entry
        self.filepath = entry.get("command", "")
        self.filename = entry.get("name") or os.path.basename(self.filepath)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(14)

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(bool(entry.get("enabled", True)))
        self.checkbox.setStyleSheet("QCheckBox::indicator { width: 22px; height: 22px; }")
        layout.addWidget(self.checkbox)

        icon_box = QFrame()
        icon_box.setFixedSize(64, 64)
        icon_box.setStyleSheet(
            "QFrame { background: #f7f7f7; border: 1px solid #d8d8d8; border-radius: 16px; }"
        )
        icon_box_layout = QVBoxLayout(icon_box)
        icon_box_layout.setContentsMargins(10, 10, 10, 10)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(44, 44)
        self.icon_label.setAlignment(Qt.AlignCenter)

        file_info = QFileInfo(self.filepath)
        icon_provider = QFileIconProvider()
        icon = icon_provider.icon(file_info)
        if icon.isNull():
            icon = self.style().standardIcon(self.style().SP_FileIcon)
        pixmap = icon.pixmap(QSize(44, 44))
        if pixmap.isNull():
            fallback = QIcon.fromTheme("application-x-ms-dos-executable")
            pixmap = fallback.pixmap(QSize(44, 44)) if not fallback.isNull() else QIcon().pixmap(QSize(44, 44))
        self.icon_label.setPixmap(pixmap)

        icon_box_layout.addWidget(self.icon_label)
        layout.addWidget(icon_box)

        center_layout = QVBoxLayout()
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(6)

        self.name_label = QLabel(self.filename)
        self.name_label.setStyleSheet("font-size: 18px; font-weight: 700; color: #1e1e1e;")
        center_layout.addWidget(self.name_label)

        self.config_btn = QPushButton("配置")
        self.config_btn.setCursor(Qt.PointingHandCursor)
        self.config_btn.setFixedSize(74, 36)
        self.config_btn.setStyleSheet(
            """
            QPushButton {
                border: 1px solid #2b2b2b;
                border-radius: 10px;
                background-color: #ffffff;
                color: #222;
                font-size: 16px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #f5f5f5;
            }
            """
        )

        self.start_btn = QPushButton("启动")
        self.start_btn.setCursor(Qt.PointingHandCursor)
        self.start_btn.setFixedSize(74, 36)
        self.start_btn.setStyleSheet(
            """
            QPushButton {
                border: 1px solid #0b6ecf;
                border-radius: 10px;
                background-color: #ffffff;
                color: #0b6ecf;
                font-size: 15px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #eff7ff;
            }
            """
        )

        self.stop_btn = QPushButton("结束")
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.setFixedSize(74, 36)
        self.stop_btn.setStyleSheet(
            """
            QPushButton {
                border: 1px solid #b45309;
                border-radius: 10px;
                background-color: #ffffff;
                color: #b45309;
                font-size: 15px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #fff7ed;
            }
            """
        )

        self.remove_btn = QPushButton("移除")
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.setFixedSize(74, 36)
        self.remove_btn.setStyleSheet(
            """
            QPushButton {
                border: 1px solid #dc2626;
                border-radius: 10px;
                background-color: #ffffff;
                color: #dc2626;
                font-size: 15px;
                font-weight: 700;
            }
            QPushButton:hover {
                background-color: #fff1f2;
            }
            """
        )

        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 0, 0, 0)
        action_row.setSpacing(8)
        action_row.addWidget(self.config_btn)
        action_row.addWidget(self.start_btn)
        action_row.addWidget(self.stop_btn)
        action_row.addWidget(self.remove_btn)
        action_row.addStretch()

        center_layout.addLayout(action_row)
        center_layout.addStretch()
        layout.addLayout(center_layout)

        layout.addSpacerItem(QSpacerItem(24, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        time_points = entry.get("time_points", []) if isinstance(entry.get("time_points", []), list) else []
        first_point = time_points[0] if time_points and isinstance(time_points[0], dict) else {}
        start_value = first_point.get("start") or entry.get("start_time", "00:00")
        end_value = first_point.get("end") or entry.get("end_time", "00:00")
        multi_suffix = f" (+{len(time_points) - 1})" if len(time_points) > 1 else ""

        self.start_time_label = QLabel(f"启动时间 {start_value}{multi_suffix}")
        self.start_time_label.setStyleSheet("color: #4f4f4f; font-size: 18px; font-weight: 600;")

        self.end_time_label = QLabel(f"结束时间 {end_value}{multi_suffix}")
        self.end_time_label.setStyleSheet("color: #4f4f4f; font-size: 18px; font-weight: 600;")

        right_layout.addStretch()
        right_layout.addWidget(self.start_time_label, 0, Qt.AlignRight)
        right_layout.addWidget(self.end_time_label, 0, Qt.AlignRight)
        right_layout.addStretch()

        layout.addLayout(right_layout)
        self.setMinimumHeight(96)

        self.checkbox.toggled.connect(on_toggle_enabled)
        self.config_btn.clicked.connect(on_config_clicked)
        self.start_btn.clicked.connect(on_start_clicked)
        self.stop_btn.clicked.connect(on_stop_clicked)
        self.remove_btn.clicked.connect(on_remove_clicked)
