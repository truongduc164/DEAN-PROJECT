"""
RegexProtection – protect untranslatable tokens before AI, restore after.

Protected patterns (never translated):
  * Product codes   – BTN-8, SMI-13, XDK2512016, YB0807944
  * Ratios          – 5:1:1.2, 7:1:8
  * Pantone codes   – 471C, 560C
  * Units           – 12mm, 3.5cm, 100%, 25°C
  * Pure numbers    – 123, 45.67
  * Date-like       – 2024-01-15, 15/01/2024
  * Serial IDs      – alphanumeric 6+ chars starting with letter(s)

Strategy
--------
``protect(text)`` uses a single combined regex with alternation (priority
left-to-right) so that each character position is matched at most once.
Placeholders use Unicode private-use-area characters (no digits/letters)
to avoid collision with subsequent patterns.
``restore(text, mapping)`` reverses the substitution.
"""
from __future__ import annotations

import re
from typing import Dict, Tuple

# ── Individual pattern strings (order = priority) ────────────────────

_PATTERN_STRINGS = [
    # Date-like: 2024-01-15 or 15/01/2024
    r"\b\d{1,4}[-/]\d{1,2}[-/]\d{1,4}\b",
    # Ratios: 5:1:1.2 or 7:1:8  (at least two colon-separated groups)
    r"\b\d+(?:\.\d+)?(?::\d+(?:\.\d+)?){1,}\b",
    # Product / serial codes: letter(s) + digits or letter-digit combos
    # e.g. BTN-8, SMI-13, XDK2512016, YB0807944
    r"\b[A-Z]{1,5}[-]?\d{1,}[A-Z0-9]*\b",
    # Pantone codes: 3-4 digits followed by a letter e.g. 471C, 560C
    r"\b\d{3,4}[A-Z]\b",
    # Measurements with units: 12mm, 3.5cm, 100%, 25°C, 0.5kg
    r"\b\d+(?:\.\d+)?\s?(?:mm|cm|m|kg|g|l|ml|°C|°F|%)\b",
    # Long alphanumeric serial IDs (6+ chars, mixed letters+digits)
    r"\b(?=[A-Za-z]*\d)(?=\d*[A-Za-z])[A-Za-z0-9]{6,}\b",
    # Pure numbers (integers and decimals)
    r"\b\d+(?:\.\d+)?\b",
]

# Single combined regex – alternation gives left-to-right priority and
# each position is consumed once, so no double-matching of placeholders.
_COMBINED = re.compile("|".join(f"({p})" for p in _PATTERN_STRINGS))

# Placeholders use Unicode private-use-area chars so they can never
# collide with real text or with digit/letter patterns.
_PH_L = "\uE000\uE001"   # 
_PH_R = "\uE002\uE003"   # 


def protect(text: str) -> Tuple[str, Dict[str, str]]:
    """
    Replace protected tokens with placeholders.

    Returns
    -------
    (cleaned_text, mapping)
        *mapping* maps placeholder → original token.
    """
    mapping: Dict[str, str] = {}
    counter = 0

    def _replacer(match: re.Match[str]) -> str:
        nonlocal counter
        token = match.group(0)
        placeholder = f"{_PH_L}{counter}{_PH_R}"
        mapping[placeholder] = token
        counter += 1
        return placeholder

    result = _COMBINED.sub(_replacer, text)
    return result, mapping


def restore(text: str, mapping: Dict[str, str]) -> str:
    """
    Restore placeholders back to original tokens.
    """
    for placeholder, original in mapping.items():
        text = text.replace(placeholder, original)
    return text
