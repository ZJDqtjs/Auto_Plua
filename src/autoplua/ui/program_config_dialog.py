from __future__ import annotations

import shlex

from PySide6.QtCore import Qt, QTime
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTimeEdit,
    QVBoxLayout,
)

from autoplua.ui.workflow_editor import (
    CLICK_MODULES,
    FLOW_MODULE_TITLES,
    NO_CONFIG_MODULES,
    ModulePaletteButton,
    NodeParamDialog,
    WorkflowCanvas,
)


class ProgramConfigDialog(QDialog):
    def __init__(self, entry: dict, parent=None):
        super().__init__(parent)
        self.entry = dict(entry)
        self.result_data: dict | None = None

        self.setWindowTitle(f"{self.entry.get('name', '')} 程序配置")
        self.resize(1080, 760)
        self.setStyleSheet(
            "QDialog { background: #f5f7fb; }"
            "QLabel { color: #1f2937; }"
            "QPushButton { border: 1px solid #c9d6e4; border-radius: 8px; background: #ffffff; padding: 5px 10px; }"
            "QPushButton:hover { border-color: #0b6ecf; background: #f3f8ff; }"
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(12)
        self._time_rows: list[dict] = []

        title = QLabel(f"{self.entry.get('name', '程序')} 程序配置")
        title.setStyleSheet("font-size: 42px; font-weight: 700; color: #1f2937;")
        root.addWidget(title)

        time_frame = QFrame()
        time_frame.setStyleSheet("QFrame { border: 1px solid #d8e1eb; border-radius: 12px; background: #fff; }")
        time_layout = QVBoxLayout(time_frame)
        time_layout.setContentsMargins(14, 12, 14, 12)
        time_layout.setSpacing(8)

        time_header = QHBoxLayout()
        time_label = QLabel("自动启动/结束时间")
        time_label.setStyleSheet("font-size: 20px; font-weight: 700; color: #1f2937;")
        time_header.addWidget(time_label)
        time_header.addStretch()

        add_time_btn = QPushButton("+ 新增")
        add_time_btn.setCursor(Qt.PointingHandCursor)
        add_time_btn.clicked.connect(self._add_time_row)
        time_header.addWidget(add_time_btn)
        time_layout.addLayout(time_header)

        self.time_rows_layout = QVBoxLayout()
        self.time_rows_layout.setSpacing(6)
        time_layout.addLayout(self.time_rows_layout)

        time_tip = QLabel("可配置多个启动/结束时间点，列表页展示第一组时间。")
        time_tip.setStyleSheet("font-size: 13px; color: #64748b;")
        time_layout.addWidget(time_tip)
        root.addWidget(time_frame)

        self._load_initial_time_rows()

        args_frame = QFrame()
        args_frame.setStyleSheet("QFrame { border: 1px solid #d8e1eb; border-radius: 12px; background: #fff; }")
        args_layout = QVBoxLayout(args_frame)
        args_layout.setContentsMargins(14, 12, 14, 12)
        args_layout.setSpacing(8)

        args_label = QLabel("使用启动参数(如果有):")
        args_label.setStyleSheet("font-size: 20px; font-weight: 700; color: #1f2937;")
        args_layout.addWidget(args_label)

        self.args_input = QLineEdit(self.entry.get("launch_args_raw", ""))
        self.args_input.setPlaceholderText('例如: --mode fast --user "demo"')
        self.args_input.setStyleSheet(
            "QLineEdit { border: 1px solid #c9d6e4; border-radius: 8px; padding: 9px 12px; font-size: 15px; }"
        )
        args_layout.addWidget(self.args_input)

        args_tip = QLabel("若填写启动参数，可直接一步完成启动；OpenCV 流程作为兜底配置使用。")
        args_tip.setStyleSheet("font-size: 14px; color: #4b5563;")
        args_layout.addWidget(args_tip)

        mode_row = QHBoxLayout()
        mode_row.setSpacing(8)
        mode_label = QLabel("输入模式")
        mode_label.setStyleSheet("font-size: 14px; color: #334155;")
        self.input_mode_combo = QComboBox()
        self.input_mode_combo.addItem("前台模拟（兼容高，会占用鼠标键盘）", "foreground")
        self.input_mode_combo.addItem("后台窗口消息（不抢鼠标键盘）", "background_window_message")

        saved_mode = str(self.entry.get("input_mode", "foreground")).strip() or "foreground"
        idx = self.input_mode_combo.findData(saved_mode)
        self.input_mode_combo.setCurrentIndex(idx if idx >= 0 else 0)

        mode_row.addWidget(mode_label)
        mode_row.addWidget(self.input_mode_combo)
        mode_row.addStretch()
        args_layout.addLayout(mode_row)

        target_row = QHBoxLayout()
        target_row.setSpacing(8)
        target_label = QLabel("目标窗口标题")
        target_label.setStyleSheet("font-size: 14px; color: #334155;")
        self.target_window_input = QLineEdit(self.entry.get("target_window_title", ""))
        self.target_window_input.setPlaceholderText("仅后台窗口消息模式必填，例如：碧蓝航线")
        target_row.addWidget(target_label)
        target_row.addWidget(self.target_window_input)
        args_layout.addLayout(target_row)
        root.addWidget(args_frame)

        module_box = QFrame()
        module_box.setStyleSheet("QFrame { border: 1px solid #d8e1eb; border-radius: 12px; background: #fff; }")
        module_layout = QVBoxLayout(module_box)
        module_layout.setContentsMargins(14, 12, 14, 14)
        module_layout.setSpacing(10)

        module_title = QLabel("OpenCV识图点击")
        module_title.setStyleSheet("font-size: 20px; font-weight: 700; color: #1f2937;")
        module_layout.addWidget(module_title)

        chip_row = QHBoxLayout()
        chip_row.setSpacing(10)
        for module_type in ["start", "left_click", "right_click", "scroll", "wait", "text_input", "enter", "end"]:
            chip = ModulePaletteButton(module_type=module_type, text=FLOW_MODULE_TITLES[module_type])
            chip_row.addWidget(chip)
        chip_row.addStretch()
        module_layout.addLayout(chip_row)

        controls = QHBoxLayout()
        controls.setSpacing(10)

        self.connect_btn = QPushButton("连线模式")
        self.connect_btn.setCheckable(True)
        self.connect_btn.setCursor(Qt.PointingHandCursor)
        self.connect_btn.toggled.connect(self._toggle_connect_mode)
        self.connect_btn.setStyleSheet(
            "QPushButton { border: 1px solid #c9d6e4; border-radius: 8px; padding: 6px 10px; font-size: 14px; }"
            "QPushButton:checked { background: #e5f1ff; border-color: #0b6ecf; color: #0b6ecf; font-weight: 700; }"
        )
        controls.addWidget(self.connect_btn)

        remove_node_btn = QPushButton("删除选中模块")
        remove_node_btn.clicked.connect(self._remove_selected_node)
        controls.addWidget(remove_node_btn)

        remove_edge_btn = QPushButton("清除选中模块连线")
        remove_edge_btn.clicked.connect(self._remove_selected_connections)
        controls.addWidget(remove_edge_btn)

        clear_btn = QPushButton("清空流程")
        clear_btn.clicked.connect(self._clear_workflow)
        controls.addWidget(clear_btn)
        controls.addStretch()
        module_layout.addLayout(controls)

        self.canvas = WorkflowCanvas()
        self.canvas.node_edit_requested.connect(self._edit_node_params)
        self.canvas.error_raised.connect(lambda msg: QMessageBox.warning(self, "连线限制", msg))
        module_layout.addWidget(self.canvas)

        self.paste_shortcut = QShortcut(QKeySequence.Paste, self)
        self.paste_shortcut.activated.connect(self._paste_to_selected_node)

        guide = QLabel(
            "拖拽上方模块到下方画布，打开【连线模式】后按顺序点击两个模块可连线。"
            "双击模块可配置参数（点击支持图片识别或手动坐标，支持 Ctrl+V 粘贴截图并在模块下方预览；无等待模块时默认步骤间隔 2 秒）。"
        )
        guide.setWordWrap(True)
        guide.setStyleSheet("font-size: 14px; color: #4b5563;")
        module_layout.addWidget(guide)

        root.addWidget(module_box)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save_and_accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        flow_payload = self.entry.get("opencv_flow", {})
        if isinstance(flow_payload, dict):
            self.canvas.load_payload(flow_payload)
        self._migrate_legacy_step_timeout_to_start_node()

    def _toggle_connect_mode(self, enabled: bool) -> None:
        self.canvas.set_connect_mode(enabled)
        self.connect_btn.setText("连线模式(开启)" if enabled else "连线模式")

    def _remove_selected_node(self) -> None:
        self.canvas.remove_selected_node()

    def _remove_selected_connections(self) -> None:
        self.canvas.remove_selected_node_connections()

    def _clear_workflow(self) -> None:
        result = QMessageBox.question(
            self,
            "清空确认",
            "确定清空当前流程画布吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if result == QMessageBox.Yes:
            self.canvas.clear_all()

    def _edit_node_params(self, node_id: str) -> None:
        node = self.canvas.nodes.get(node_id)
        if not node:
            return
        if node.module_type in NO_CONFIG_MODULES:
            return
        dialog = NodeParamDialog(module_type=node.module_type, initial=node.params, parent=self)
        if dialog.exec() != QDialog.Accepted:
            return
        self.canvas.set_node_params(node_id, dialog.get_data())

    def _paste_to_selected_node(self) -> None:
        ok = self.canvas.apply_clipboard_image_to_selected_node()
        if not ok:
            return

    def _save_and_accept(self) -> None:
        launch_args_raw = self.args_input.text().strip()
        parsed_args: list[str] = []
        if launch_args_raw:
            try:
                parsed_args = shlex.split(launch_args_raw, posix=False)
            except ValueError as exc:
                QMessageBox.warning(self, "参数错误", f"启动参数解析失败：{exc}")
                return

        input_mode = str(self.input_mode_combo.currentData())
        target_window_title = self.target_window_input.text().strip()
        if input_mode == "background_window_message" and not target_window_title:
            QMessageBox.warning(self, "配置不完整", "后台窗口消息模式需要填写目标窗口标题。")
            return

        flow = self.canvas.to_payload()
        has_flow = bool(flow.get("nodes"))
        if not launch_args_raw and not has_flow:
            QMessageBox.warning(self, "配置不完整", "请至少配置启动参数或 OpenCV 流程中的一种。")
            return

        if has_flow:
            valid, message = self.canvas.validate_workflow()
            if not valid:
                QMessageBox.warning(self, "流程不合法", message)
                return
            node_check_ok, node_error = self._validate_node_params(flow)
            if not node_check_ok:
                QMessageBox.warning(self, "模块参数不完整", node_error)
                return

        startup_timeout_seconds = self._extract_startup_timeout(flow)

        self.result_data = {
            "launch_args_raw": launch_args_raw,
            "args": parsed_args,
            "input_mode": input_mode,
            "target_window_title": target_window_title,
            # Keep legacy key for compatibility; source now comes from start-node config.
            "opencv_step_retry_seconds": startup_timeout_seconds,
            "opencv_flow": flow,
            "time_points": self._collect_time_points(),
        }
        self.accept()

    def _migrate_legacy_step_timeout_to_start_node(self) -> None:
        legacy_timeout = int(self.entry.get("opencv_step_retry_seconds", 30) or 30)
        legacy_timeout = max(5, legacy_timeout)

        for node_id, node in self.canvas.nodes.items():
            if node.module_type != "start":
                continue
            params = dict(node.params)
            params.setdefault("startup_timeout_seconds", legacy_timeout)
            params.setdefault("next_step_delay_seconds", 3)
            self.canvas.set_node_params(node_id, params)
            break

    @staticmethod
    def _extract_startup_timeout(flow: dict) -> int:
        default_timeout = 30
        nodes = flow.get("nodes", []) if isinstance(flow, dict) else []
        if not isinstance(nodes, list):
            return default_timeout
        for node in nodes:
            if not isinstance(node, dict) or str(node.get("module", "")) != "start":
                continue
            params = node.get("params", {}) if isinstance(node.get("params", {}), dict) else {}
            try:
                return max(5, int(params.get("startup_timeout_seconds", default_timeout)))
            except (TypeError, ValueError):
                return default_timeout
        return default_timeout

    def _load_initial_time_rows(self) -> None:
        time_points = self.entry.get("time_points", [])
        if not isinstance(time_points, list) or not time_points:
            time_points = [
                {
                    "start": self.entry.get("start_time", "00:00"),
                    "end": self.entry.get("end_time", "00:00"),
                }
            ]

        for point in time_points:
            if not isinstance(point, dict):
                continue
            self._add_time_row(
                start=str(point.get("start", "00:00")),
                end=str(point.get("end", "00:00")),
            )

    def _add_time_row(self, checked: bool = False, start: str = "00:00", end: str = "00:00") -> None:
        _ = checked
        row_frame = QFrame()
        row_frame.setStyleSheet("QFrame { border: 1px solid #e2e8f0; border-radius: 10px; background: #f8fafc; }")

        row = QHBoxLayout(row_frame)
        row.setContentsMargins(10, 8, 10, 8)
        row.setSpacing(8)

        start_label = QLabel("启动时间")
        start_label.setStyleSheet("font-size: 14px; color: #334155;")
        row.addWidget(start_label)

        start_edit = QTimeEdit()
        start_edit.setDisplayFormat("HH:mm")
        start_edit.setTime(self._parse_hhmm(start, QTime(0, 0)))
        start_edit.setStyleSheet("QTimeEdit { min-width: 90px; }")
        row.addWidget(start_edit)

        end_label = QLabel("结束时间")
        end_label.setStyleSheet("font-size: 14px; color: #334155;")
        row.addWidget(end_label)

        end_edit = QTimeEdit()
        end_edit.setDisplayFormat("HH:mm")
        end_edit.setTime(self._parse_hhmm(end, QTime(0, 0)))
        end_edit.setStyleSheet("QTimeEdit { min-width: 90px; }")
        row.addWidget(end_edit)

        row.addStretch()

        remove_btn = QPushButton("删除")
        remove_btn.setCursor(Qt.PointingHandCursor)
        row.addWidget(remove_btn)

        row_data = {"frame": row_frame, "start": start_edit, "end": end_edit}
        remove_btn.clicked.connect(lambda: self._remove_time_row(row_data))

        self._time_rows.append(row_data)
        self.time_rows_layout.addWidget(row_frame)

    def _remove_time_row(self, row_data: dict) -> None:
        if len(self._time_rows) <= 1:
            QMessageBox.information(self, "提示", "至少保留一组启动/结束时间。")
            return
        frame = row_data.get("frame")
        if frame:
            frame.deleteLater()
        if row_data in self._time_rows:
            self._time_rows.remove(row_data)

    def _collect_time_points(self) -> list[dict]:
        points = []
        for row in self._time_rows:
            start_edit = row.get("start")
            end_edit = row.get("end")
            if not isinstance(start_edit, QTimeEdit) or not isinstance(end_edit, QTimeEdit):
                continue
            points.append(
                {
                    "start": start_edit.time().toString("HH:mm"),
                    "end": end_edit.time().toString("HH:mm"),
                }
            )
        return points

    @staticmethod
    def _parse_hhmm(value: str, fallback: QTime) -> QTime:
        parsed = QTime.fromString(value, "HH:mm")
        return parsed if parsed.isValid() else fallback

    def _validate_node_params(self, flow: dict) -> tuple[bool, str]:
        for node in flow.get("nodes", []):
            module_type = node.get("module")
            params = node.get("params", {}) if isinstance(node.get("params", {}), dict) else {}
            title = node.get("title", FLOW_MODULE_TITLES.get(module_type, module_type))
            if module_type in CLICK_MODULES:
                image_path = str(params.get("image_path", "")).strip()
                x = params.get("x", 0)
                y = params.get("y", 0)
                has_manual_coord = (x, y) != (0, 0)
                if not image_path and not has_manual_coord:
                    return False, f"模块【{title}】需要配置识图图片或手动坐标。"
            elif module_type == "scroll":
                steps = int(params.get("steps", 0))
                if steps == 0:
                    return False, "模块【鼠标滚动】的滚动行数不能为 0。"
            elif module_type == "wait":
                seconds = int(params.get("seconds", 1))
                if seconds < 1:
                    return False, "模块【等待】秒数必须大于等于 1。"
            elif module_type == "text_input":
                if not str(params.get("text", "")).strip():
                    return False, "模块【文本输入】不能为空。"
        return True, ""
