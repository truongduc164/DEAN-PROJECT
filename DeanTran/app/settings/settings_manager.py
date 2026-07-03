"""
SettingsManager – JSON-based settings for the entire DeanTran app.

All translation behaviour is controlled by ``app_settings.json``.
API keys are NOT stored here — they remain in the OS keyring via SecureStorage.

Usage::

    from app.settings.settings_manager import settings
    model = settings.get("selected_models.gemini")
    settings.set("language_settings.source_lang", "English")
    settings.save()
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Optional

from app.core.path_helpers import get_configs_dir

logger = logging.getLogger("DeanTran.settings")

_DEFAULT_SETTINGS: dict = {
    "translation_tool": "Gemini",
    "selected_models": {
        "gemini": "gemini-3.1-flash-lite",
        "deepseek": "deepseek-chat",
    },
    "language_settings": {
        "source_lang": "Chinese",
        "target_lang": "Vietnamese",
    },
    "document_settings": {
        "document_type": "SOP",
        "result_mode": "Save As",
        "output_mode": "overwrite",
    },
    "processing_options": {
        "check_glossary_before_ai": True,
        "auto_batch_by_content": False,
        "translate_textboxes": False,
        "translate_excel_textboxes": False,
        "auto_open_file": False,
        "add_chinese_pinyin": False,
    },
    "limits_settings": {
        "max_rows": 8000,
        "max_cols": 50,
        "max_items_per_batch": 200,
        "max_chars_per_batch": 30000,
        "circuit_breaker_failures": 5,
        "retry_limit": 3,
        "api_timeout": 60,
    },
    "speed_settings": {
        "min_interval": 0.2,
    },
    "ocr_settings": {
        "vision_context_enabled": False,
        "image_text_translation_enabled": False,
        "engine": "paddle",
        "google_vision_key": "",
        "google_vision_batch_enabled": True,
        "google_vision_max_images_per_request": 16,
        "google_vision_dedupe_images": True,
        "google_vision_canvas_enabled": False,
        "google_vision_canvas_images_per_canvas": 4,
        "google_vision_canvas_padding": 24,
        "google_vision_canvas_max_width": 8192,
        "google_vision_canvas_max_height": 8192,
        "ocr_processing_mode": "auto",
        "image_ocr_mode": "off",
        "min_ocr_confidence": 50.0,
        "max_blocks_per_image": 50,
        "skip_low_confidence_blocks": True,
        "preserve_special_tokens_in_images": True,
        "render_textbox_overlay": True,
        "image_translation_batch_enabled": True,
        "translate_with_api": True,
        "tesseract_psm": 3,
        "merge_lines": True,
        "preprocess_image": True,
        "ocr_display_mode": "overwrite",
        "ocr_display_container": "textbox",
        "ocr_textbox_mode": "smart_adjust",
        "ocr_textbox_placement_mode": "whitespace",
        "smart_adjust_max_shift_px": 80,
        "whitespace_search_max_shift_px": 180,
    },
    "text_style_settings": {
        "translated_text_format_mode": "keep_original_format",
        "keep_format": True,
        "font_family": "Arial",
        "font_size": 14,
        "font_color": "#000000",
        "bold": False,
        "italic": False,
        "underline": False,
        "style_preset": "readable",
        "title_font_family": "Arial",
        "body_font_family": "Arial",
        "note_font_family": "Arial",
        "font_size_mode": "auto_fit",
        "title_font_size": 24,
        "body_font_size": 14,
        "note_font_size": 10,
        "min_font_size": 10,
        "max_font_size": 48,
        "scale_ratio_percent": 100,
        "title_color": "#000000",
        "body_color": "#333333",
        "note_color": "#666666",
        "background_mode": "semi_transparent",
        "background_color": "#FFFFFF",
        "background_opacity": 180,
        "line_spacing": 1.15,
        "fit_strategy": "preserve_readability",
        "bold_title": True,
        "bold_body": False,
    },
}

# Schema: key_path → expected type (for validation)
_SCHEMA: dict[str, type] = {
    "translation_tool": str,
    "selected_models.gemini": str,
    "selected_models.deepseek": str,
    "language_settings.source_lang": str,
    "language_settings.target_lang": str,
    "document_settings.document_type": str,
    "document_settings.result_mode": str,
    "document_settings.output_mode": str,
    "processing_options.check_glossary_before_ai": bool,
    "processing_options.auto_batch_by_content": bool,
    "processing_options.translate_textboxes": bool,
    "processing_options.translate_excel_textboxes": bool,
    "processing_options.auto_open_file": bool,
    "processing_options.add_chinese_pinyin": bool,
    "limits_settings.max_rows": int,
    "limits_settings.max_cols": int,
    "limits_settings.max_items_per_batch": int,
    "limits_settings.max_chars_per_batch": int,
    "limits_settings.circuit_breaker_failures": int,
    "limits_settings.retry_limit": int,
    "limits_settings.api_timeout": int,
    "speed_settings.min_interval": (int, float),  # type: ignore
    "ocr_settings.image_text_translation_enabled": bool,
    "ocr_settings.engine": str,
    "ocr_settings.google_vision_key": str,
    "ocr_settings.google_vision_batch_enabled": bool,
    "ocr_settings.google_vision_max_images_per_request": int,
    "ocr_settings.google_vision_dedupe_images": bool,
    "ocr_settings.google_vision_canvas_enabled": bool,
    "ocr_settings.google_vision_canvas_images_per_canvas": int,
    "ocr_settings.google_vision_canvas_padding": int,
    "ocr_settings.google_vision_canvas_max_width": int,
    "ocr_settings.google_vision_canvas_max_height": int,
    "ocr_settings.ocr_processing_mode": str,
    "ocr_settings.image_ocr_mode": str,
    "ocr_settings.min_ocr_confidence": float,
    "ocr_settings.max_blocks_per_image": int,
    "ocr_settings.skip_low_confidence_blocks": bool,
    "ocr_settings.preserve_special_tokens_in_images": bool,
    "ocr_settings.render_textbox_overlay": bool,
    "ocr_settings.image_translation_batch_enabled": bool,
    "ocr_settings.translate_with_api": bool,
    "ocr_settings.tesseract_psm": int,
    "ocr_settings.merge_lines": bool,
    "ocr_settings.preprocess_image": bool,
    "ocr_settings.ocr_display_mode": str,
    "ocr_settings.ocr_display_container": str,
    "ocr_settings.ocr_textbox_mode": str,
    "ocr_settings.ocr_textbox_placement_mode": str,
    "ocr_settings.smart_adjust_max_shift_px": int,
    "ocr_settings.whitespace_search_max_shift_px": int,
    "text_style_settings.translated_text_format_mode": str,
    "text_style_settings.keep_format": bool,
    "text_style_settings.font_family": str,
    "text_style_settings.font_size": int,
    "text_style_settings.font_color": str,
    "text_style_settings.bold": bool,
    "text_style_settings.italic": bool,
    "text_style_settings.underline": bool,
    "text_style_settings.style_preset": str,
    "text_style_settings.title_font_family": str,
    "text_style_settings.body_font_family": str,
    "text_style_settings.note_font_family": str,
    "text_style_settings.font_size_mode": str,
    "text_style_settings.title_font_size": int,
    "text_style_settings.body_font_size": int,
    "text_style_settings.note_font_size": int,
    "text_style_settings.min_font_size": int,
    "text_style_settings.max_font_size": int,
    "text_style_settings.scale_ratio_percent": int,
    "text_style_settings.title_color": str,
    "text_style_settings.body_color": str,
    "text_style_settings.note_color": str,
    "text_style_settings.background_mode": str,
    "text_style_settings.background_color": str,
    "text_style_settings.background_opacity": int,
    "text_style_settings.line_spacing": (int, float),
    "text_style_settings.fit_strategy": str,
    "text_style_settings.bold_title": bool,
    "text_style_settings.bold_body": bool,
}


class SettingsManager:
    """
    Centralised settings loaded from / saved to ``app_settings.json``.

    Parameters
    ----------
    config_dir : Path | str | None
        Directory for ``app_settings.json``.
        Defaults to ``<project_root>/configs``.
    """

    FILENAME = "app_settings.json"

    def __init__(self, config_dir: Optional[Path | str] = None) -> None:
        if config_dir is None:
            self.config_dir = get_configs_dir()
        else:
            self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self._file = self.config_dir / self.FILENAME
        self._data: dict = {}
        self._load()

    # ── public API ───────────────────────────────────────────────────

    def get(self, key_path: str, default: Any = None) -> Any:
        """
        Retrieve a setting by dot-separated path.

        Examples::

            settings.get("selected_models.gemini")
            settings.get("limits_settings.max_rows")
        """
        keys = key_path.split(".")
        node: Any = self._data
        for k in keys:
            if isinstance(node, dict) and k in node:
                node = node[k]
            else:
                return default
        return node

    def set(self, key_path: str, value: Any) -> None:
        """
        Set a setting by dot-separated path and auto-save.

        Examples::

            settings.set("selected_models.gemini", "gemini-3.1-flash-lite")
        """
        # Validate type if in schema
        expected = _SCHEMA.get(key_path)
        if expected is not None and not isinstance(value, expected):
            logger.warning(
                "Type mismatch for %s: expected %s, got %s (%r)",
                key_path, expected, type(value).__name__, value,
            )

        keys = key_path.split(".")
        node = self._data
        for k in keys[:-1]:
            if k not in node or not isinstance(node[k], dict):
                node[k] = {}
            node = node[k]
        node[keys[-1]] = value
        logger.debug("settings.set %s = %r", key_path, value)
        self.save()

    def save(self) -> None:
        """Persist current settings to ``app_settings.json``."""
        self._file.write_text(
            json.dumps(self._data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        logger.debug("Settings saved to %s", self._file)

    def get_all(self) -> dict:
        """Return a deep copy of all settings."""
        import copy
        return copy.deepcopy(self._data)

    def reset_defaults(self) -> None:
        """Reset all settings to built-in defaults."""
        import copy
        self._data = copy.deepcopy(_DEFAULT_SETTINGS)
        self.save()

    # ── internal ─────────────────────────────────────────────────────

    def _load(self) -> None:
        if self._file.exists():
            try:
                data = json.loads(self._file.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    self._data = self._merge_defaults(data)
                    logger.info("Settings loaded from %s", self._file)
                    return
            except (json.JSONDecodeError, OSError) as exc:
                logger.warning("Failed to load settings: %s", exc)

        # First run or corrupt file → create defaults
        import copy
        self._data = copy.deepcopy(_DEFAULT_SETTINGS)
        self.save()
        logger.info("Default settings created at %s", self._file)

    def _merge_defaults(self, loaded: dict) -> dict:
        """Merge loaded data with defaults so new keys are always present."""
        import copy
        merged = copy.deepcopy(_DEFAULT_SETTINGS)
        self._deep_update(merged, loaded)
        return merged

    @staticmethod
    def _deep_update(base: dict, override: dict) -> None:
        for k, v in override.items():
            if k in base and isinstance(base[k], dict) and isinstance(v, dict):
                SettingsManager._deep_update(base[k], v)
            else:
                base[k] = v


# ── Module-level singleton ────────────────────────────────────────────
# Import as:  from app.settings.settings_manager import settings
settings = SettingsManager()
