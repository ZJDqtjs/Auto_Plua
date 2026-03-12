from __future__ import annotations

import uuid
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

from PySide6.QtCore import Qt, QMimeData, QPoint, QPointF, Signal
from PySide6.QtGui import QColor, QDrag, QGuiApplication, QImage, QKeyEvent, QKeySequence, QMouseEvent, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
)

FLOW_MODULE_TITLES = {
    "start": "启动",
    "left_click": "左键点击",
    "right_click": "右键点击",
    "scroll": "鼠标滚动",
    "wait": "等待",
    "text_input": "文本输入",
    "enter": "回车",
    "end": "结束",
}

CLICK_MODULES = {"left_click", "right_click"}
NO_CONFIG_MODULES = {"end", "enter"}


def _template_image_dir() -> Path:
    path = Path(__file__).resolve().parents[3] / "template_image"
    path.mkdir(parents=True, exist_ok=True)
    return path


def normalize_image_path(raw_path: str) -> str:
    value = str(raw_path or "").strip()
    if not value:
        return ""
    if value.startswith("file://"):
        parsed = urlparse(value)
        candidate = unquote(parsed.path or "")
        if candidate.startswith("/") and len(candidate) >= 3 and candidate[2] == ":":
            candidate = candidate[1:]
        value = candidate
    return str(Path(value).expanduser())


def save_clipboard_image_to_template() -> str | None:
    clipboard = QGuiApplication.clipboard()
    if clipboard is None:
        return None

    mime_data = clipboard.mimeData()
    if mime_data is None:
        return None

    image: QImage | None = None
    if mime_data.hasImage():
        image = QImage(clipboard.image())
    elif mime_data.hasUrls() and mime_data.urls():
        first = mime_data.urls()[0]
        if first.isLocalFile():
            image = QImage(first.toLocalFile())
    elif mime_data.hasText():
        text_path = normalize_image_path(mime_data.text())
        image = QImage(text_path)

    if image is None or image.isNull():
        return None

    filename = f"template_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}.png"
    target = _template_image_dir() / filename
    if not image.save(str(target), "PNG"):
        return None
    return str(target)


class ModulePaletteButton(QPushButton):
    def __init__(self, module_type: str, text: str, parent=None):
        super().__init__(text, parent)
        self.module_type = module_type
        self._drag_start = QPoint()
        self.setCursor(Qt.OpenHandCursor)
        self.setFixedHeight(34)
        self.setStyleSheet(
            "QPushButton { border: 1px solid #c9d6e4; border-radius: 8px; background: #ffffff;"
            " font-size: 15px; font-weight: 700; color: #1f2937; padding: 4px 12px; }"
            "QPushButton:hover { border-color: #0b6ecf; background: #f3f8ff; }"
        )

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self._drag_start = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if not (event.buttons() & Qt.LeftButton):
            return
        if (event.position().toPoint() - self._drag_start).manhattanLength() < 8:
            return

        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setData("application/x-autoplua-module", self.module_type.encode("utf-8"))
        drag.setMimeData(mime_data)

        pixmap = self.grab()
        drag.setPixmap(pixmap)
        drag.setHotSpot(event.position().toPoint())
        drag.exec(Qt.CopyAction)


class FlowNodeWidget(QFrame):
    clicked = Signal(str)
    double_clicked = Signal(str)
    moved = Signal()

    def __init__(self, node_id: str, module_type: str, title: str, parent=None):
        super().__init__(parent)
        self.node_id = node_id
        self.module_type = module_type
        self.title = title
        self.params: dict = {}
        self._drag_start_global = QPoint()
        self._drag_origin = QPoint()
        self._selected = False

        if self.module_type == "start":
            self.setFixedSize(136, 88)
        elif self.module_type == "end":
            self.setFixedSize(124, 74)
        elif self.module_type == "enter":
            self.setFixedSize(124, 82)
        elif self.module_type in CLICK_MODULES:
            self.setFixedSize(198, 168)
        else:
            self.setFixedSize(144, 88)
        self.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        self.title_label = QLabel(title)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet(
            "font-size: 16px; font-weight: 700; color: #1f2937;"
            "border: 1px solid #b7c9dd; border-radius: 10px; padding: 4px 6px; background: #ffffff;"
        )

        self.sub_label = QLabel("双击配置")
        self.sub_label.setAlignment(Qt.AlignCenter)
        self.sub_label.setStyleSheet(
            "font-size: 13px; color: #4b5563;"
            "border: 1px solid #b7c9dd; border-radius: 9px; padding: 2px 6px; background: #ffffff;"
        )

        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setFixedHeight(96)
        self.preview_label.setStyleSheet("border: 1px dashed #b7c9dd; border-radius: 8px; background: #f8fbff;")
        self.preview_label.hide()

        layout.addWidget(self.title_label)
        if self.module_type not in NO_CONFIG_MODULES:
            if self.module_type in CLICK_MODULES:
                self.sub_label.hide()
            else:
                layout.addWidget(self.sub_label)
        else:
            self.sub_label.hide()
        if self.module_type in CLICK_MODULES:
            layout.addWidget(self.preview_label)
        self._apply_style()

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self._apply_style()

    def set_params(self, params: dict) -> None:
        self.params = dict(params)
        if self.module_type in CLICK_MODULES:
            image_path = normalize_image_path(str(self.params.get("image_path", "")))
            self.params["image_path"] = image_path
        if self.module_type in CLICK_MODULES:
            self._refresh_preview()
        if self.module_type in NO_CONFIG_MODULES:
            return
        if self.module_type not in CLICK_MODULES:
            self.sub_label.setText("双击配置")
        self._apply_style()

    def _refresh_preview(self) -> None:
        image_path = normalize_image_path(str(self.params.get("image_path", "")))
        if not image_path:
            self.preview_label.clear()
            self.preview_label.hide()
            return
        pix = QPixmap(image_path)
        if pix.isNull():
            self.preview_label.clear()
            self.preview_label.hide()
            return
        target_w = max(80, self.preview_label.width() - 8)
        target_h = max(40, self.preview_label.height() - 8)
        scaled = pix.scaled(target_w, target_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.preview_label.setPixmap(scaled)
        self.preview_label.show()

    def _apply_style(self) -> None:
        if self.module_type in NO_CONFIG_MODULES:
            border_color = "transparent"
            bg = "transparent"
        else:
            is_configured = self._is_configured()
            base_color = "#16a34a" if is_configured else "#dc2626"
            border_color = "#0b6ecf" if self._selected else base_color
            bg = "#f3f8ff" if self._selected else "#ffffff"
        self.setStyleSheet(
            "QFrame { border: 2px solid %s; border-radius: 12px; background: %s; }" % (border_color, bg)
        )

    def _is_configured(self) -> bool:
        if self.module_type in NO_CONFIG_MODULES:
            return True
        if self.module_type == "start":
            return True
        if self.module_type in CLICK_MODULES:
            image_path = normalize_image_path(str(self.params.get("image_path", "")))
            has_image = bool(image_path and Path(image_path).exists())
            x = int(self.params.get("x", 0))
            y = int(self.params.get("y", 0))
            return has_image or (x, y) != (0, 0)
        if self.module_type == "scroll":
            return int(self.params.get("steps", 0)) != 0
        if self.module_type == "text_input":
            return bool(str(self.params.get("text", "")).strip())
        if self.module_type == "wait":
            return int(self.params.get("seconds", 0)) >= 1
        return bool(self.params)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self._drag_start_global = event.globalPosition().toPoint()
            self._drag_origin = self.pos()
            self.clicked.emit(self.node_id)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        if not (event.buttons() & Qt.LeftButton):
            return
        parent = self.parentWidget()
        if parent is None:
            return

        delta = event.globalPosition().toPoint() - self._drag_start_global
        next_pos = self._drag_origin + delta

        max_x = max(0, parent.width() - self.width())
        max_y = max(0, parent.height() - self.height())
        clamped = QPoint(min(max(next_pos.x(), 0), max_x), min(max(next_pos.y(), 0), max_y))
        self.move(clamped)
        self.moved.emit()

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.LeftButton:
            self.double_clicked.emit(self.node_id)
        super().mouseDoubleClickEvent(event)


class WorkflowCanvas(QFrame):
    node_edit_requested = Signal(str)
    error_raised = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMinimumHeight(360)
        self.setStyleSheet(
            "QFrame { background: #fbfdff; border: 2px dashed #bfd0e4; border-radius: 16px; }"
        )

        self.nodes: dict[str, FlowNodeWidget] = {}
        self.edges: list[tuple[str, str]] = []
        self.selected_node_id: str | None = None
        self.connect_mode = False
        self._pending_source_id: str | None = None

    def set_connect_mode(self, enabled: bool) -> None:
        self.connect_mode = enabled
        self._pending_source_id = None
        self.update()

    def dragEnterEvent(self, event) -> None:
        mime = event.mimeData()
        if mime.hasFormat("application/x-autoplua-module"):
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event) -> None:
        mime = event.mimeData()
        if not mime.hasFormat("application/x-autoplua-module"):
            event.ignore()
            return
        module_type = bytes(mime.data("application/x-autoplua-module")).decode("utf-8")
        drop_pos = event.position().toPoint()
        self.add_node(module_type=module_type, pos=drop_pos)
        event.acceptProposedAction()

    def add_node(self, module_type: str, pos: QPoint | None = None, params: dict | None = None, node_id: str | None = None) -> str:
        title = FLOW_MODULE_TITLES.get(module_type, module_type)
        node_id = node_id or str(uuid.uuid4())
        node = FlowNodeWidget(node_id=node_id, module_type=module_type, title=title, parent=self)
        node.clicked.connect(self._on_node_clicked)
        node.double_clicked.connect(self.node_edit_requested.emit)
        node.moved.connect(self.update)
        node.set_params(params or {})

        if pos is None:
            pos = QPoint(26 + (len(self.nodes) * 24) % 260, 20 + (len(self.nodes) * 30) % 180)

        max_x = max(0, self.width() - node.width())
        max_y = max(0, self.height() - node.height())
        clamped = QPoint(min(max(pos.x(), 0), max_x), min(max(pos.y(), 0), max_y))

        node.move(clamped)
        node.show()

        self.nodes[node_id] = node
        self._set_selected(node_id)
        self.update()
        return node_id

    def remove_selected_node(self) -> None:
        if not self.selected_node_id:
            return
        node_id = self.selected_node_id
        node = self.nodes.pop(node_id, None)
        if node:
            node.deleteLater()
        self.edges = [(s, t) for s, t in self.edges if s != node_id and t != node_id]
        self.selected_node_id = None
        if self._pending_source_id == node_id:
            self._pending_source_id = None
        self.update()

    def clear_all(self) -> None:
        for node in self.nodes.values():
            node.deleteLater()
        self.nodes.clear()
        self.edges.clear()
        self.selected_node_id = None
        self._pending_source_id = None
        self.update()

    def remove_selected_node_connections(self) -> None:
        if not self.selected_node_id:
            return
        node_id = self.selected_node_id
        self.edges = [(s, t) for s, t in self.edges if s != node_id and t != node_id]
        self._pending_source_id = None
        self.update()

    def set_node_params(self, node_id: str, params: dict) -> None:
        node = self.nodes.get(node_id)
        if not node:
            return
        node.set_params(params)

    def to_payload(self) -> dict:
        nodes_payload = []
        for node_id, node in self.nodes.items():
            p = node.pos()
            params = dict(node.params)
            if node.module_type in CLICK_MODULES:
                params["image_path"] = normalize_image_path(str(params.get("image_path", "")))
            nodes_payload.append(
                {
                    "id": node_id,
                    "module": node.module_type,
                    "title": node.title,
                    "x": int(p.x()),
                    "y": int(p.y()),
                    "params": params,
                }
            )
        return {
            "nodes": nodes_payload,
            "edges": [[s, t] for s, t in self.edges],
        }

    def load_payload(self, payload: dict) -> None:
        self.clear_all()
        if not isinstance(payload, dict):
            return

        for raw in payload.get("nodes", []):
            if not isinstance(raw, dict):
                continue
            node_id = str(raw.get("id") or uuid.uuid4())
            module_type = str(raw.get("module") or "")
            x = int(raw.get("x", 20))
            y = int(raw.get("y", 20))
            params = raw.get("params", {}) if isinstance(raw.get("params", {}), dict) else {}
            if module_type:
                self.add_node(module_type=module_type, pos=QPoint(x, y), params=params, node_id=node_id)

        edges = []
        for edge in payload.get("edges", []):
            if not isinstance(edge, (list, tuple)) or len(edge) != 2:
                continue
            source, target = str(edge[0]), str(edge[1])
            if source in self.nodes and target in self.nodes and source != target:
                edges.append((source, target))
        self.edges = edges
        self.update()

    def validate_workflow(self) -> tuple[bool, str]:
        if not self.nodes:
            return True, ""

        starts = [n for n in self.nodes.values() if n.module_type == "start"]
        ends = [n for n in self.nodes.values() if n.module_type == "end"]
        if len(starts) != 1:
            return False, "流程中必须且只能有一个【启动】模块。"
        if len(ends) != 1:
            return False, "流程中必须且只能有一个【结束】模块。"

        incoming = {node_id: 0 for node_id in self.nodes}
        outgoing = {node_id: 0 for node_id in self.nodes}
        for source, target in self.edges:
            outgoing[source] += 1
            incoming[target] += 1

        start_id = starts[0].node_id
        end_id = ends[0].node_id

        if incoming[start_id] != 0:
            return False, "【启动】模块不能有入线。"
        if outgoing[end_id] != 0:
            return False, "【结束】模块不能有出线。"
        if outgoing[start_id] != 1:
            return False, "【启动】模块必须连接到下一步。"
        if incoming[end_id] != 1:
            return False, "【结束】模块必须由上一步连接。"

        for node_id, node in self.nodes.items():
            if node_id in (start_id, end_id):
                continue
            if incoming[node_id] != 1 or outgoing[node_id] != 1:
                return False, f"模块【{node.title}】需要且只能有一条入线和一条出线。"

        visited = set()
        cursor = start_id
        while cursor not in visited:
            visited.add(cursor)
            next_nodes = [t for s, t in self.edges if s == cursor]
            if not next_nodes:
                break
            cursor = next_nodes[0]

        if cursor != end_id or len(visited) != len(self.nodes):
            return False, "流程必须是从【启动】到【结束】的单链路，不能存在分叉或孤立模块。"

        return True, ""

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        pen = QPen(QColor("#3f3f3f"))
        pen.setWidth(3)
        painter.setPen(pen)

        for source, target in self.edges:
            source_node = self.nodes.get(source)
            target_node = self.nodes.get(target)
            if not source_node or not target_node:
                continue
            start = source_node.geometry().center()
            end = target_node.geometry().center()
            painter.drawLine(QPointF(start), QPointF(end))

        if self.connect_mode and self._pending_source_id in self.nodes:
            source_node = self.nodes[self._pending_source_id]
            center = source_node.geometry().center()
            preview_pen = QPen(QColor("#0b6ecf"))
            preview_pen.setWidth(3)
            preview_pen.setStyle(Qt.DotLine)
            painter.setPen(preview_pen)
            painter.drawEllipse(center, 8, 8)

    def _set_selected(self, node_id: str | None) -> None:
        self.selected_node_id = node_id
        for nid, node in self.nodes.items():
            node.set_selected(nid == node_id)

    def _on_node_clicked(self, node_id: str) -> None:
        self.setFocus(Qt.MouseFocusReason)
        self._set_selected(node_id)

        if not self.connect_mode:
            self.update()
            return

        if self._pending_source_id is None:
            self._pending_source_id = node_id
            self.update()
            return

        source = self._pending_source_id
        target = node_id
        self._pending_source_id = None

        if source == target:
            self.error_raised.emit("不能将模块连接到自身。")
            self.update()
            return

        if any(s == source for s, _ in self.edges):
            self.error_raised.emit("每个模块当前仅允许一条出线。")
            self.update()
            return
        if any(t == target for _, t in self.edges):
            self.error_raised.emit("每个模块当前仅允许一条入线。")
            self.update()
            return

        source_node = self.nodes.get(source)
        target_node = self.nodes.get(target)
        if not source_node or not target_node:
            self.update()
            return
        if source_node.module_type == "end":
            self.error_raised.emit("【结束】模块不能作为连线起点。")
            self.update()
            return
        if target_node.module_type == "start":
            self.error_raised.emit("【启动】模块不能作为连线终点。")
            self.update()
            return

        self.edges.append((source, target))
        self.update()

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.matches(QKeySequence.Paste):
            ok = self.apply_clipboard_image_to_selected_node()
            if not ok:
                self.error_raised.emit("请先选中一个点击类模块（左键点击/右键点击），并确保剪贴板里是图片。")
            event.accept()
            return
        super().keyPressEvent(event)

    def apply_clipboard_image_to_selected_node(self) -> bool:
        if not self.selected_node_id:
            return False
        node = self.nodes.get(self.selected_node_id)
        if not node or node.module_type not in CLICK_MODULES:
            return False

        image_path = save_clipboard_image_to_template()
        if not image_path:
            return False

        params = dict(node.params)
        params["image_path"] = image_path
        self.set_node_params(node.node_id, params)
        return True


class NodeParamDialog(QDialog):
    def __init__(self, module_type: str, initial: dict | None = None, parent=None):
        super().__init__(parent)
        self.module_type = module_type
        self.initial = dict(initial or {})
        self.setWindowTitle(f"模块配置 - {FLOW_MODULE_TITLES.get(module_type, module_type)}")
        self.resize(520, 260)

        layout = QVBoxLayout(self)
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setFormAlignment(Qt.AlignTop)
        form.setSpacing(12)

        self.image_input = QLineEdit(self.initial.get("image_path", ""))
        self.image_input.setPlaceholderText("选择图片或按 Ctrl+V 粘贴截图（自动保存到 template_image）")
        image_row = QHBoxLayout()
        image_row.addWidget(self.image_input)
        browse_btn = QPushButton("选择图片")
        browse_btn.clicked.connect(self._pick_image)
        image_row.addWidget(browse_btn)

        paste_btn = QPushButton("粘贴截图")
        paste_btn.clicked.connect(self._paste_image_from_clipboard)
        image_row.addWidget(paste_btn)

        self.x_input = QSpinBox()
        self.x_input.setRange(-99999, 99999)
        self.x_input.setValue(int(self.initial.get("x", 0)))

        self.y_input = QSpinBox()
        self.y_input.setRange(-99999, 99999)
        self.y_input.setValue(int(self.initial.get("y", 0)))

        self.scroll_input = QSpinBox()
        self.scroll_input.setRange(-5000, 5000)
        self.scroll_input.setValue(int(self.initial.get("steps", 0)))

        self.wait_input = QSpinBox()
        self.wait_input.setRange(1, 30)
        self.wait_input.setValue(int(self.initial.get("seconds", 1)))

        self.text_input = QLineEdit(self.initial.get("text", ""))
        self.text_input.setPlaceholderText("需要输入的文本")

        self.start_timeout_input = QSpinBox()
        self.start_timeout_input.setRange(1, 300)
        self.start_timeout_input.setValue(int(self.initial.get("startup_timeout_seconds", 20)))

        self.start_next_delay_input = QSpinBox()
        self.start_next_delay_input.setRange(0, 30)
        self.start_next_delay_input.setValue(int(self.initial.get("next_step_delay_seconds", 3)))

        if module_type == "start":
            # Ordered top-to-bottom as requested.
            form.addRow("启动最大超时(秒)", self.start_timeout_input)
            form.addRow("下一模块延迟点击(秒)", self.start_next_delay_input)
        elif module_type in CLICK_MODULES:
            form.addRow("识图图片", image_row)
            form.addRow("手动 X 坐标", self.x_input)
            form.addRow("手动 Y 坐标", self.y_input)
        elif module_type == "scroll":
            form.addRow("滚动行数", self.scroll_input)
        elif module_type == "wait":
            form.addRow("等待秒数", self.wait_input)
        elif module_type == "text_input":
            form.addRow("输入文本", self.text_input)
        else:
            note = QLabel("该模块无需额外参数。")
            note.setStyleSheet("font-size: 14px; color: #4b5563;")
            form.addRow("说明", note)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _pick_image(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择识图图片",
            "",
            "Image Files (*.png *.jpg *.jpeg *.bmp);;All Files (*)",
        )
        if file_path:
            self.image_input.setText(normalize_image_path(file_path))

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.matches(QKeySequence.Paste) and self.module_type in CLICK_MODULES:
            self._paste_image_from_clipboard()
            event.accept()
            return
        super().keyPressEvent(event)

    def _paste_image_from_clipboard(self) -> None:
        if self.module_type not in CLICK_MODULES:
            return

        image_path = save_clipboard_image_to_template()
        if not image_path:
            return
        self.image_input.setText(normalize_image_path(image_path))

    def get_data(self) -> dict:
        if self.module_type == "start":
            return {
                "startup_timeout_seconds": self.start_timeout_input.value(),
                "next_step_delay_seconds": self.start_next_delay_input.value(),
            }
        if self.module_type in CLICK_MODULES:
            return {
                "image_path": normalize_image_path(self.image_input.text()),
                "x": self.x_input.value(),
                "y": self.y_input.value(),
            }
        if self.module_type == "scroll":
            return {"steps": self.scroll_input.value()}
        if self.module_type == "wait":
            return {"seconds": self.wait_input.value()}
        if self.module_type == "text_input":
            return {"text": self.text_input.text()}
        return {}
