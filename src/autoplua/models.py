from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ManagedProgram:
    name: str
    command: str
    args: list[str] = field(default_factory=list)
    cwd: str | None = None


@dataclass
class IntervalSchedule:
    id: str
    action: str
    target: str
    seconds: int
    enabled: bool = True
    extra: dict[str, Any] = field(default_factory=dict)

