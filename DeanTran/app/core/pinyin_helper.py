from __future__ import annotations

import logging
import re

logger = logging.getLogger("DeanTran.pinyin")

_CHINESE_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
_WARNED_MISSING = False

try:
    from pypinyin import Style, lazy_pinyin

    _PYPINYIN_AVAILABLE = True
except ModuleNotFoundError:
    Style = None
    lazy_pinyin = None
    _PYPINYIN_AVAILABLE = False


def is_available() -> bool:
    return _PYPINYIN_AVAILABLE


def contains_chinese(text: str) -> bool:
    return bool(_CHINESE_RE.search(text or ""))


def get_pinyin_line(text: str) -> str:
    global _WARNED_MISSING

    if not text or not contains_chinese(text):
        return ""

    if not _PYPINYIN_AVAILABLE or lazy_pinyin is None or Style is None:
        if not _WARNED_MISSING:
            logger.warning("pypinyin is not installed; Chinese pinyin output is disabled")
            _WARNED_MISSING = True
        return ""

    rendered_lines: list[str] = []
    for raw_line in text.replace("\r", "").split("\n"):
        if not raw_line.strip():
            rendered_lines.append("")
            continue
        if not contains_chinese(raw_line):
            rendered_lines.append("")
            continue

        tokens = lazy_pinyin(
            raw_line,
            style=Style.TONE,
            neutral_tone_with_five=True,
            strict=False,
            errors=lambda chars: list(chars),
        )
        rendered = " ".join(token for token in tokens if token is not None).strip()
        rendered = re.sub(r"\s+([\u3001\u3002\uff0c\uff01\uff1f\uff1b\uff1a,.!?;:])", r"\1", rendered)
        rendered = re.sub(r"([\uff08(\u3010\[])\s+", r"\1", rendered)
        rendered = re.sub(r"\s+([\uff09)\u3011\]])", r"\1", rendered)
        rendered = re.sub(r"\s{2,}", " ", rendered).strip()
        rendered_lines.append(rendered)

    pinyin = "\n".join(rendered_lines).strip()
    return "" if pinyin == text.strip() else pinyin


def render_text_with_pinyin(text: str) -> str:
    if not text:
        return ""

    rendered_lines: list[str] = []
    for raw_line in text.replace("\r", "").split("\n"):
        rendered_lines.append(raw_line)
        pinyin = get_pinyin_line(raw_line)
        if pinyin:
            rendered_lines.append(pinyin)

    return "\n".join(rendered_lines)
