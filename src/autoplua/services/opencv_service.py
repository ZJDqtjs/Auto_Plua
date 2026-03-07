from __future__ import annotations

import os
import time
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


class OpenCVFlowService:
    def run_flow(
        self,
        flow: dict[str, Any],
        timeout_seconds: int = 10,
        default_wait_seconds: int = 1,
    ) -> tuple[bool, str]:
        nodes_chain = self._build_linear_chain(flow)
        if not nodes_chain:
            return True, "empty-flow"

        started_at = time.monotonic()

        for index in range(1, len(nodes_chain)):
            now = time.monotonic()
            if now - started_at > timeout_seconds:
                return False, f"timeout-{timeout_seconds}s"

            prev_node = nodes_chain[index - 1]
            node = nodes_chain[index]
            module = str(node.get("module", ""))

            if prev_node.get("module") != "wait" and module not in {"wait", "end"}:
                if not self._safe_sleep(default_wait_seconds, started_at, timeout_seconds):
                    return False, f"timeout-{timeout_seconds}s"

            ok, message = self._execute_node(node=node, started_at=started_at, timeout_seconds=timeout_seconds)
            if not ok:
                return False, message

        return True, "ok"

    def _execute_node(self, node: dict[str, Any], started_at: float, timeout_seconds: int) -> tuple[bool, str]:
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

        if pyautogui is None:
            return False, "missing-dependency-pyautogui"

        if module in CLICK_MODULES:
            return self._execute_click(module=module, params=params)

        if module == "scroll":
            steps = int(params.get("steps", 0))
            if steps == 0:
                return False, "invalid-scroll-steps"
            pyautogui.scroll(steps)
            return True, "ok"

        if module == "text_input":
            text = str(params.get("text", ""))
            if not text:
                return False, "empty-text"
            pyautogui.write(text, interval=0.02)
            return True, "ok"

        if module == "enter":
            pyautogui.press("enter")
            return True, "ok"

        return False, f"unsupported-module-{module}"

    def _execute_click(self, module: str, params: dict[str, Any]) -> tuple[bool, str]:
        button = "left" if module == "left_click" else "right"
        image_path = str(params.get("image_path", "")).strip()
        x = int(params.get("x", 0))
        y = int(params.get("y", 0))

        target = None
        if image_path:
            target = self._locate_by_template(image_path)

        if target is None and (x, y) != (0, 0):
            target = (x, y)

        if target is None:
            return False, "click-target-not-found"

        pyautogui.click(target[0], target[1], button=button)
        return True, "ok"

    def _locate_by_template(self, image_path: str) -> tuple[int, int] | None:
        if cv2 is None or np is None or pyautogui is None:
            return None
        if not os.path.exists(image_path):
            return None

        screenshot = pyautogui.screenshot()
        screen = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        template = cv2.imread(image_path)
        if template is None:
            return None

        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        if max_val < 0.8:
            return None

        h, w = template.shape[:2]
        center_x = int(max_loc[0] + w / 2)
        center_y = int(max_loc[1] + h / 2)
        return center_x, center_y

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
