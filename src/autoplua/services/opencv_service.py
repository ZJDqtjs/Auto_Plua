from __future__ import annotations

import os
import time
import ctypes
from ctypes import wintypes
from typing import Any

try:
    import cv2  # type: ignore
    import numpy as np  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    cv2 = None
    np = None

try:
    import pyautogui  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    pyautogui = None


CLICK_MODULES = {"left_click", "right_click"}

WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
WM_MOUSEWHEEL = 0x020A
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CHAR = 0x0102
VK_RETURN = 0x0D

MK_LBUTTON = 0x0001
MK_RBUTTON = 0x0002
DEFAULT_TEMPLATE_THRESHOLD = 0.80


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class OpenCVFlowService:
    def __init__(self) -> None:
        self._last_match_score: float = 0.0

    def run_flow(
        self,
        flow: dict[str, Any],
        timeout_seconds: int = 120,
        default_wait_seconds: int = 2,
        startup_wait_seconds: int = 3,
        step_retry_seconds: int = 20,
        execution_options: dict[str, Any] | None = None,
    ) -> tuple[bool, str]:
        execution_options = execution_options or {}
        input_mode = str(execution_options.get("input_mode", "foreground")).strip().lower() or "foreground"
        target_window_title = str(execution_options.get("target_window_title", "")).strip()

        if input_mode == "background_window_message" and not target_window_title:
            return False, "invalid-target-window-title"

        if self._requires_screen_capture(flow) and not self._is_capture_available():
            return False, "screen-capture-unavailable-possibly-screen-off-or-locked"

        nodes_chain = self._build_linear_chain(flow)
        if not nodes_chain:
            return True, "empty-flow"

        started_at = time.monotonic()

        if str(nodes_chain[0].get("module", "")) == "start" and startup_wait_seconds > 0:
            if not self._safe_sleep(startup_wait_seconds, started_at, timeout_seconds):
                return False, f"timeout-{timeout_seconds}s"

        for index in range(1, len(nodes_chain)):
            now = time.monotonic()
            if now - started_at > timeout_seconds:
                return False, f"timeout-{timeout_seconds}s"

            node = nodes_chain[index]
            module = str(node.get("module", ""))

            if module == "wait":
                ok, message = self._execute_node(
                    node=node,
                    started_at=started_at,
                    timeout_seconds=timeout_seconds,
                    input_mode=input_mode,
                    target_window_title=target_window_title,
                )
            else:
                ok, message = self._execute_node_with_retry(
                    node=node,
                    started_at=started_at,
                    timeout_seconds=timeout_seconds,
                    step_retry_seconds=step_retry_seconds,
                    input_mode=input_mode,
                    target_window_title=target_window_title,
                )
            if not ok:
                return False, message

            # Default pacing: after an action succeeds, wait briefly before the next step.
            has_next = index < (len(nodes_chain) - 1)
            if has_next and module not in {"wait", "end"}:
                next_module = str(nodes_chain[index + 1].get("module", ""))
                if next_module not in {"wait", "end"}:
                    if not self._safe_sleep(default_wait_seconds, started_at, timeout_seconds):
                        return False, f"timeout-{timeout_seconds}s"

        return True, "ok"

    def _execute_node_with_retry(
        self,
        node: dict[str, Any],
        started_at: float,
        timeout_seconds: int,
        step_retry_seconds: int,
        input_mode: str,
        target_window_title: str,
    ) -> tuple[bool, str]:
        deadline = min(time.monotonic() + max(1, step_retry_seconds), started_at + timeout_seconds)
        last_message = "step-not-ready"

        while time.monotonic() <= deadline:
            ok, message = self._execute_node(
                node=node,
                started_at=started_at,
                timeout_seconds=timeout_seconds,
                input_mode=input_mode,
                target_window_title=target_window_title,
            )
            if ok:
                return True, message

            last_message = message
            if self._is_non_retryable(message):
                return False, message
            time.sleep(0.4)

        return False, f"step-timeout-{step_retry_seconds}s-{last_message}"

    @staticmethod
    def _is_non_retryable(message: str) -> bool:
        return message.startswith(
            (
                "missing-dependency",
                "unsupported-module",
                "invalid-",
                "template-image-not-found",
            )
        ) or message in {
            "empty-text",
        }

    def _execute_node(
        self,
        node: dict[str, Any],
        started_at: float,
        timeout_seconds: int,
        input_mode: str,
        target_window_title: str,
    ) -> tuple[bool, str]:
        module = str(node.get("module", ""))
        params = node.get("params", {}) if isinstance(node.get("params"), dict) else {}

        if module in {"start", "end"}:
            return True, "ok"

        if module == "wait":
            seconds = int(params.get("seconds", 1))
            if seconds < 1:
                seconds = 1
            if not self._safe_sleep(seconds, started_at, timeout_seconds):
                return False, f"timeout-{timeout_seconds}s"
            return True, "ok"

        if pyautogui is None and input_mode == "foreground":
            return False, "missing-dependency-pyautogui"

        if module in CLICK_MODULES:
            return self._execute_click(
                module=module,
                params=params,
                input_mode=input_mode,
                target_window_title=target_window_title,
            )

        if module == "scroll":
            steps = int(params.get("steps", 0))
            if steps == 0:
                return False, "invalid-scroll-steps"
            if input_mode == "background_window_message":
                ok, message = self._send_window_scroll(target_window_title=target_window_title, steps=steps)
                if not ok:
                    return False, message
            else:
                pyautogui.scroll(steps)
            return True, "ok"

        if module == "text_input":
            text = str(params.get("text", ""))
            if not text:
                return False, "empty-text"
            if input_mode == "background_window_message":
                ok, message = self._send_window_text(target_window_title=target_window_title, text=text)
                if not ok:
                    return False, message
            else:
                pyautogui.write(text, interval=0.02)
            return True, "ok"

        if module == "enter":
            if input_mode == "background_window_message":
                ok, message = self._send_window_enter(target_window_title=target_window_title)
                if not ok:
                    return False, message
            else:
                pyautogui.press("enter")
            return True, "ok"

        return False, f"unsupported-module-{module}"

    def _execute_click(
        self,
        module: str,
        params: dict[str, Any],
        input_mode: str,
        target_window_title: str,
    ) -> tuple[bool, str]:
        button = "left" if module == "left_click" else "right"
        image_path = str(params.get("image_path", "")).strip()
        x = int(params.get("x", 0))
        y = int(params.get("y", 0))

        target = None
        self._last_match_score = 0.0
        if image_path:
            if not os.path.exists(image_path):
                return False, f"template-image-not-found:{image_path}"
            threshold = float(params.get("threshold", DEFAULT_TEMPLATE_THRESHOLD))
            target = self._locate_by_template(
                image_path,
                target_window_title=target_window_title,
                threshold=threshold,
            )

        if target is None and (x, y) != (0, 0):
            target = (x, y)

        if target is None:
            return False, f"click-target-not-found-score-{self._last_match_score:.3f}"

        if input_mode == "background_window_message":
            ok, message = self._send_window_click(
                target_window_title=target_window_title,
                x=target[0],
                y=target[1],
                button=button,
            )
            if not ok:
                return False, message
        else:
            pyautogui.click(target[0], target[1], button=button)
        return True, "ok"

    def _is_capture_available(self) -> bool:
        if pyautogui is None or np is None:
            return False
        try:
            frame = pyautogui.screenshot()
            arr = np.array(frame)
            if arr.size == 0:
                return False
            if arr.ndim < 2:
                return False
            # If screenshot call succeeds and shape is valid, regard capture backend as available.
            return arr.shape[0] > 0 and arr.shape[1] > 0
        except Exception:
            return False

    @staticmethod
    def _requires_screen_capture(flow: dict[str, Any]) -> bool:
        nodes = flow.get("nodes", []) if isinstance(flow, dict) else []
        if not isinstance(nodes, list):
            return False

        for node in nodes:
            if not isinstance(node, dict):
                continue
            if str(node.get("module", "")) not in CLICK_MODULES:
                continue
            params = node.get("params", {}) if isinstance(node.get("params"), dict) else {}
            image_path = str(params.get("image_path", "")).strip()
            if image_path:
                return True
        return False

    @staticmethod
    def _window_handle_by_title(target_window_title: str) -> int:
        if not target_window_title:
            return 0
        user32 = ctypes.windll.user32
        exact = int(user32.FindWindowW(None, target_window_title))
        if exact:
            return exact

        matches: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        def enum_proc(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value.strip().lower()
            if target_window_title.lower() in title:
                matches.append(int(hwnd))
            return True

        user32.EnumWindows(enum_proc, 0)
        return matches[0] if matches else 0

    @staticmethod
    def _make_lparam(x: int, y: int) -> int:
        return ((y & 0xFFFF) << 16) | (x & 0xFFFF)

    def _screen_to_client(self, hwnd: int, x: int, y: int) -> tuple[int, int] | None:
        point = _POINT(x=x, y=y)
        ok = ctypes.windll.user32.ScreenToClient(wintypes.HWND(hwnd), ctypes.byref(point))
        if not ok:
            return None
        return int(point.x), int(point.y)

    def _send_window_click(self, target_window_title: str, x: int, y: int, button: str) -> tuple[bool, str]:
        top_hwnd = self._window_handle_by_title(target_window_title)
        if not top_hwnd:
            return False, "target-window-not-found"

        target_hwnd = self._resolve_input_hwnd(top_hwnd=top_hwnd, screen_x=x, screen_y=y)

        client_pos = self._screen_to_client(target_hwnd, x, y)
        if client_pos is None:
            return False, "target-window-coordinate-convert-failed"

        cx, cy = client_pos
        lparam = self._make_lparam(cx, cy)
        user32 = ctypes.windll.user32

        user32.PostMessageW(wintypes.HWND(target_hwnd), WM_MOUSEMOVE, 0, lparam)
        if button == "left":
            user32.PostMessageW(wintypes.HWND(target_hwnd), WM_LBUTTONDOWN, MK_LBUTTON, lparam)
            user32.PostMessageW(wintypes.HWND(target_hwnd), WM_LBUTTONUP, 0, lparam)
        else:
            user32.PostMessageW(wintypes.HWND(target_hwnd), WM_RBUTTONDOWN, MK_RBUTTON, lparam)
            user32.PostMessageW(wintypes.HWND(target_hwnd), WM_RBUTTONUP, 0, lparam)
        return True, "ok"

    def _resolve_input_hwnd(self, top_hwnd: int, screen_x: int, screen_y: int) -> int:
        user32 = ctypes.windll.user32

        point = _POINT(x=screen_x, y=screen_y)
        if not user32.ScreenToClient(wintypes.HWND(top_hwnd), ctypes.byref(point)):
            return top_hwnd

        child = user32.ChildWindowFromPointEx(wintypes.HWND(top_hwnd), point, 0)
        return int(child) if child else top_hwnd

    def _send_window_scroll(self, target_window_title: str, steps: int) -> tuple[bool, str]:
        hwnd = self._window_handle_by_title(target_window_title)
        if not hwnd:
            return False, "target-window-not-found"

        wheel_delta = 120 * steps
        wparam = (wheel_delta & 0xFFFF) << 16
        ctypes.windll.user32.PostMessageW(wintypes.HWND(hwnd), WM_MOUSEWHEEL, wparam, 0)
        return True, "ok"

    def _send_window_text(self, target_window_title: str, text: str) -> tuple[bool, str]:
        hwnd = self._window_handle_by_title(target_window_title)
        if not hwnd:
            return False, "target-window-not-found"

        user32 = ctypes.windll.user32
        for ch in text:
            user32.PostMessageW(wintypes.HWND(hwnd), WM_CHAR, ord(ch), 0)
        return True, "ok"

    def _send_window_enter(self, target_window_title: str) -> tuple[bool, str]:
        hwnd = self._window_handle_by_title(target_window_title)
        if not hwnd:
            return False, "target-window-not-found"

        user32 = ctypes.windll.user32
        user32.PostMessageW(wintypes.HWND(hwnd), WM_KEYDOWN, VK_RETURN, 0)
        user32.PostMessageW(wintypes.HWND(hwnd), WM_KEYUP, VK_RETURN, 0)
        return True, "ok"

    def _locate_by_template(
        self,
        image_path: str,
        target_window_title: str = "",
        threshold: float = DEFAULT_TEMPLATE_THRESHOLD,
    ) -> tuple[int, int] | None:
        if cv2 is None or np is None:
            return None

        screen = None
        offset_x = 0
        offset_y = 0
        used_window_capture = False
        if target_window_title:
            capture_result = self._capture_window_bgr(target_window_title)
            if capture_result is not None:
                screen, offset_x, offset_y = capture_result
            used_window_capture = True

        if screen is None:
            if pyautogui is None:
                return None
            screenshot = pyautogui.screenshot()
            screen = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        template = cv2.imread(image_path)
        if template is None:
            return None

        # Guard OpenCV assertion: source image must be >= template size.
        if screen.shape[0] < template.shape[0] or screen.shape[1] < template.shape[1]:
            if used_window_capture and pyautogui is not None:
                screenshot = pyautogui.screenshot()
                fallback_screen = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                if fallback_screen.shape[0] >= template.shape[0] and fallback_screen.shape[1] >= template.shape[1]:
                    screen = fallback_screen
                    offset_x = 0
                    offset_y = 0
                else:
                    self._last_match_score = 0.0
                    return None
            else:
                self._last_match_score = 0.0
                return None

        max_val, max_loc = self._best_match(screen, template)
        self._last_match_score = max_val
        if max_val < threshold:
            return None

        h, w = template.shape[:2]
        center_x = int(offset_x + max_loc[0] + w / 2)
        center_y = int(offset_y + max_loc[1] + h / 2)
        return center_x, center_y

    @staticmethod
    def _best_match(screen: Any, template: Any) -> tuple[float, tuple[int, int]]:
        if screen.shape[0] < template.shape[0] or screen.shape[1] < template.shape[1]:
            return 0.0, (0, 0)

        result_color = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        _, score_color, _, loc_color = cv2.minMaxLoc(result_color)

        try:
            screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)
            template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
            result_gray = cv2.matchTemplate(screen_gray, template_gray, cv2.TM_CCOEFF_NORMED)
            _, score_gray, _, loc_gray = cv2.minMaxLoc(result_gray)
        except Exception:
            score_gray = -1.0
            loc_gray = (0, 0)

        if score_gray > score_color:
            return float(score_gray), (int(loc_gray[0]), int(loc_gray[1]))
        return float(score_color), (int(loc_color[0]), int(loc_color[1]))

    def _capture_window_bgr(self, target_window_title: str) -> tuple[Any, int, int] | None:
        hwnd = self._window_handle_by_title(target_window_title)
        if not hwnd:
            return None

        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32

        rect = _RECT()
        if not user32.GetWindowRect(wintypes.HWND(hwnd), ctypes.byref(rect)):
            return None

        client_rect = _RECT()
        if not user32.GetClientRect(wintypes.HWND(hwnd), ctypes.byref(client_rect)):
            return None

        client_origin = _POINT(0, 0)
        if not user32.ClientToScreen(wintypes.HWND(hwnd), ctypes.byref(client_origin)):
            return None

        full_width = int(rect.right - rect.left)
        full_height = int(rect.bottom - rect.top)
        if full_width <= 0 or full_height <= 0:
            return None

        client_width = int(client_rect.right - client_rect.left)
        client_height = int(client_rect.bottom - client_rect.top)
        if client_width <= 0 or client_height <= 0:
            return None

        hdc_window = user32.GetWindowDC(wintypes.HWND(hwnd))
        if not hdc_window:
            return None

        hdc_mem = gdi32.CreateCompatibleDC(hdc_window)
        hbm = gdi32.CreateCompatibleBitmap(hdc_window, full_width, full_height)
        if not hdc_mem or not hbm:
            if hbm:
                gdi32.DeleteObject(hbm)
            if hdc_mem:
                gdi32.DeleteDC(hdc_mem)
            user32.ReleaseDC(wintypes.HWND(hwnd), hdc_window)
            return None

        gdi32.SelectObject(hdc_mem, hbm)
        ok = user32.PrintWindow(wintypes.HWND(hwnd), hdc_mem, 0x00000002)
        if not ok:
            ok = user32.PrintWindow(wintypes.HWND(hwnd), hdc_mem, 0)

        if not ok:
            gdi32.DeleteObject(hbm)
            gdi32.DeleteDC(hdc_mem)
            user32.ReleaseDC(wintypes.HWND(hwnd), hdc_window)
            return None

        bmp_size = full_width * full_height * 4
        buffer = ctypes.create_string_buffer(bmp_size)
        copied = gdi32.GetBitmapBits(hbm, bmp_size, buffer)

        gdi32.DeleteObject(hbm)
        gdi32.DeleteDC(hdc_mem)
        user32.ReleaseDC(wintypes.HWND(hwnd), hdc_window)

        if copied <= 0:
            return None

        image = np.frombuffer(buffer, dtype=np.uint8)
        if image.size < bmp_size:
            return None
        image = image.reshape((full_height, full_width, 4))
        bgr = image[:, :, :3]

        crop_left = int(client_origin.x - rect.left)
        crop_top = int(client_origin.y - rect.top)
        crop_right = crop_left + client_width
        crop_bottom = crop_top + client_height

        if 0 <= crop_left < crop_right <= full_width and 0 <= crop_top < crop_bottom <= full_height:
            client_bgr = bgr[crop_top:crop_bottom, crop_left:crop_right]
            return client_bgr, int(client_origin.x), int(client_origin.y)

        # Fallback to full window image if client crop cannot be derived.
        return bgr, int(rect.left), int(rect.top)

    def _safe_sleep(self, seconds: int, started_at: float, timeout_seconds: int) -> bool:
        if seconds <= 0:
            return True

        elapsed = time.monotonic() - started_at
        remaining = timeout_seconds - elapsed
        if remaining <= 0:
            return False

        sleep_seconds = min(seconds, max(0.0, remaining))
        time.sleep(sleep_seconds)
        return (time.monotonic() - started_at) <= timeout_seconds

    def _build_linear_chain(self, flow: dict[str, Any]) -> list[dict[str, Any]]:
        if not isinstance(flow, dict):
            return []

        raw_nodes = flow.get("nodes", [])
        raw_edges = flow.get("edges", [])
        if not isinstance(raw_nodes, list) or not isinstance(raw_edges, list):
            return []

        node_map: dict[str, dict[str, Any]] = {}
        for node in raw_nodes:
            if not isinstance(node, dict):
                continue
            node_id = str(node.get("id", "")).strip()
            if not node_id:
                continue
            node_map[node_id] = node

        next_map: dict[str, str] = {}
        for edge in raw_edges:
            if not isinstance(edge, (list, tuple)) or len(edge) != 2:
                continue
            source = str(edge[0])
            target = str(edge[1])
            if source in node_map and target in node_map:
                next_map[source] = target

        start_id = None
        for node_id, node in node_map.items():
            if str(node.get("module", "")) == "start":
                start_id = node_id
                break
        if not start_id:
            return []

        chain: list[dict[str, Any]] = []
        seen = set()
        cursor = start_id
        while cursor in node_map and cursor not in seen:
            seen.add(cursor)
            chain.append(node_map[cursor])
            if cursor not in next_map:
                break
            cursor = next_map[cursor]

        return chain
