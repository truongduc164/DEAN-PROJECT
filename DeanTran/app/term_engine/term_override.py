"""
TermOverride – apply glossary replacements before and after translation.

Strategy
--------
1. **pre_replace(text)** – Before sending text to the LLM, replace every
   glossary source term with its target translation.  This steers the model
   and avoids it inventing its own translation for known terms.

2. **post_fix(text)** – After receiving the LLM output, scan for any
   remaining occurrences of source terms and force-replace them.  This is
   the "hard constraint" that the glossary always wins.
"""
from __future__ import annotations

import re
from typing import Dict, Optional

from app.term_engine.glossary_loader import GlossaryLoader


class TermOverride:
    """
    Glossary-based term replacement engine.

    Parameters
    ----------
    glossary : GlossaryLoader
        A loaded glossary instance.
    """

    def __init__(self, glossary: GlossaryLoader) -> None:
        self.glossary = glossary
        # Build a compiled regex for all source terms (longest first to avoid
        # partial matches, e.g. "SMT line" before "SMT").
        self._pattern: Optional[re.Pattern[str]] = None
        self._build_pattern()

    # ── public API ───────────────────────────────────────────────────

    def pre_replace(self, text: str) -> str:
        """Replace source-language glossary terms with target translations."""
        if self._pattern is None:
            return text
        return self._pattern.sub(self._replacer, text)

    def post_fix(self, text: str) -> str:
        """
        Force-correct the translated text: if any *source* term still appears
        in the output, replace it with the glossary target.
        """
        if self._pattern is None:
            return text
        return self._pattern.sub(self._replacer, text)

    # ── internals ────────────────────────────────────────────────────

    def _build_pattern(self) -> None:
        entries = self.glossary.entries
        if not entries:
            return
        # Sort by length descending so longer terms match first.
        sorted_keys = sorted(entries.keys(), key=len, reverse=True)
        escaped = [re.escape(k) for k in sorted_keys]
        self._pattern = re.compile("|".join(escaped), flags=re.IGNORECASE)

    def _replacer(self, match: re.Match[str]) -> str:
        """Return the glossary target for the matched source term."""
        matched_text = match.group(0)
        target = self.glossary.lookup(matched_text)
        return target if target is not None else matched_text
