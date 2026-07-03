"""
PromptStore – load / save custom prompts per translation mode.

Prompts are persisted in ``prompt_store.json`` in the configs/ directory.
The pipeline reads the prompt based on ``document_settings.document_type``
from SettingsManager.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Optional

from app.core.path_helpers import get_configs_dir

_DEFAULT_PROMPTS: Dict[str, str] = {
    "SOP": (
        "You are a professional translator specialising in Standard Operating "
        "Procedures. Translate the following text accurately, preserving "
        "technical terms and formatting. Return ONLY the translated text."
    ),
    "QC_REPORT": (
        "You are a professional translator specialising in Quality Control reports. "
        "Translate accurately, keeping all measurements, product codes, and test "
        "results unchanged. Return ONLY the translated text."
    ),
    "NORMAL": (
        "You are a professional translator. Translate the following text "
        "naturally and accurately. Return ONLY the translated text."
    ),
    "General": (
        "You are a professional translator. Translate the following text "
        "naturally and accurately. Return ONLY the translated text."
    ),
    "Technical": (
        "You are a technical translator. Preserve all product codes, model "
        "numbers, and measurements exactly. Translate only the descriptive text. "
        "Return ONLY the translated text."
    ),
}


class PromptStore:
    """
    Manages per-mode translation prompts.

    Parameters
    ----------
    config_dir : Path | str | None
        Directory that contains (or will contain) ``prompt_store.json``.
        Defaults to ``<project_root>/configs``.
    """

    FILENAME = "prompt_store.json"

    def __init__(self, config_dir: Optional[Path | str] = None) -> None:
        if config_dir is None:
            self.config_dir = get_configs_dir()
        else:
            self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self._file = self.config_dir / self.FILENAME
        self._prompts: Dict[str, str] = {}
        self._load()

    # ── public API ───────────────────────────────────────────────────

    def get(self, mode: str) -> str:
        """Return the prompt for *mode*, falling back to the built-in default."""
        return self._prompts.get(mode, _DEFAULT_PROMPTS.get(mode, ""))

    def set(self, mode: str, prompt: str) -> None:
        """Set the prompt for *mode* (in-memory only – call ``save()`` to persist)."""
        self._prompts[mode] = prompt

    def save(self) -> None:
        """Persist current prompts to ``prompt_store.json``."""
        self._file.write_text(
            json.dumps(self._prompts, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def delete(self, mode: str) -> None:
        """Remove a custom prompt for *mode*."""
        self._prompts.pop(mode, None)

    def modes(self) -> list[str]:
        """Return all mode names that have a prompt (custom + defaults)."""
        return sorted(set(list(self._prompts.keys()) + list(_DEFAULT_PROMPTS.keys())))

    def get_all(self) -> Dict[str, str]:
        """Return a merged dict of default + custom prompts."""
        merged = dict(_DEFAULT_PROMPTS)
        merged.update(self._prompts)
        return merged

    # ── internal ─────────────────────────────────────────────────────

    def _load(self) -> None:
        if self._file.exists():
            try:
                data = json.loads(self._file.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    self._prompts = data
            except (json.JSONDecodeError, OSError):
                self._prompts = {}
