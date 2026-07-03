"""
EventManager – a lightweight, pure-Python event bus.
Works without Qt so that backend logic and tests never depend on PySide6.

Supported events
  log       (level: str, message: str)
  progress  (current: int, total: int)
  status    (text: str)
  error     (message: str, exception: Exception | None)
"""
from __future__ import annotations

import sys
from collections import defaultdict
from typing import Any, Callable, Dict, List


class EventManager:
    """Simple publish / subscribe event bus."""

    def __init__(self) -> None:
        self._subscribers: Dict[str, List[Callable[..., Any]]] = defaultdict(list)

    # -- public API --------------------------------------------------------

    def subscribe(self, event: str, callback: Callable[..., Any]) -> None:
        """Register *callback* for *event*."""
        self._subscribers[event].append(callback)

    def emit(self, event: str, *args: Any, **kwargs: Any) -> None:
        """Fire *event*, calling every registered callback."""
        for cb in self._subscribers.get(event, []):
            try:
                cb(*args, **kwargs)
            except Exception as exc:  # noqa: BLE001
                # Never let a subscriber crash the pipeline.
                print(f"[EventManager] subscriber error on '{event}': {exc}",
                      file=sys.stderr)

    # -- convenience shortcuts ---------------------------------------------

    def log(self, level: str, message: str) -> None:
        self.emit("log", level, message)

    def progress(self, current: int, total: int) -> None:
        self.emit("progress", current, total)

    def status(self, text: str) -> None:
        self.emit("status", text)

    def error(self, message: str, exception: Exception | None = None) -> None:
        self.emit("error", message, exception)


# ---------------------------------------------------------------------------
# Default console subscriber – handy for CLI / test runs
# ---------------------------------------------------------------------------

_LOG_ICONS = {
    "INFO": "ℹ️",
    "WARN": "⚠️",
    "WARNING": "⚠️",
    "ERROR": "❌",
    "DEBUG": "🐛",
    "SUCCESS": "✅",
}


def console_subscriber(level: str, message: str) -> None:
    """Print log events to stdout."""
    icon = _LOG_ICONS.get(level.upper(), "📋")
    log_str = f"[{level}] {icon} {message}"
    try:
        print(log_str)
    except Exception:
        try:
            enc = sys.stdout.encoding or 'utf-8'
            print(log_str.encode(enc, errors='replace').decode(enc))
        except Exception:
            pass


def attach_console(em: EventManager) -> None:
    """Attach a minimal console logger to *em*."""
    em.subscribe("log", console_subscriber)
    em.subscribe("error", lambda msg, exc: print(f"[ERROR] {msg}  ({exc})", file=sys.stderr))
    em.subscribe("status", lambda text: print(f"[STATUS] {text}"))
    em.subscribe("progress", lambda cur, tot: print(f"[PROGRESS] {cur}/{tot}"))
