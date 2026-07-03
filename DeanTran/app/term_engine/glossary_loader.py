"""
GlossaryLoader – load a term glossary from JSON or multi-language xlsx.

Supports two Excel formats:

**Legacy (2-column):**
    Column A = source term, Column B = target term.

**Multi-language (3+ columns):**
    Row 1 contains language headers: Chinese, Vietnamese, English, etc.
    The loader picks the correct source/target columns based on the
    ``source_lang`` and ``target_lang`` parameters.

    Example:
        | Chinese | Vietnamese   | English    |
        |---------|--------------|------------|
        | 质量    | Chất lượng   | Quality    |
        | 检查    | Kiểm tra     | Inspection |

    GlossaryLoader("file.xlsx", source_lang="Chinese", target_lang="Vietnamese")
    → builds index:  质量 → Chất lượng,  检查 → Kiểm tra

    GlossaryLoader("file.xlsx", source_lang="English", target_lang="Vietnamese")
    → builds index:  Quality → Chất lượng,  Inspection → Kiểm tra

JSON format::

    {
        "source_term": "target_term",
        "BOM": "Bảng vật tư"
    }
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger("DeanTran.glossary_loader")

# Map various language name forms to a canonical key for header matching.
_LANG_ALIASES: Dict[str, List[str]] = {
    "chinese":    ["chinese", "中文", "cn", "zh", "tiếng trung", "trung"],
    "vietnamese": ["vietnamese", "tiếng việt", "vi", "vn", "việt"],
    "english":    ["english", "en", "tiếng anh", "anh"],
}


def _normalize_lang(name: str) -> str:
    """Return a canonical language key (lowercase) for *name*.

    Uses exact alias match first, then falls back to prefix matching
    to handle common header typos (e.g. 'Englissh', 'Vietnames').
    """
    low = name.strip().lower()
    # Exact alias match
    for canonical, aliases in _LANG_ALIASES.items():
        if low in aliases:
            return canonical
    # Prefix / fuzzy fallback – if the header starts with a known alias
    # or a known alias starts with the header (handles truncated names)
    for canonical, aliases in _LANG_ALIASES.items():
        for alias in aliases:
            if low.startswith(alias) or alias.startswith(low):
                return canonical
    return low


class GlossaryLoader:
    """
    Load and query a glossary dictionary.

    Parameters
    ----------
    path : Path | str | None
        Path to ``glossary.json`` or ``.xlsx``.
    source_lang : str | None
        Source language name (e.g. "Chinese"). Used for multi-column xlsx
        to pick the correct source column. Falls back to column A if not
        matched.
    target_lang : str | None
        Target language name (e.g. "Vietnamese"). Falls back to column B.
    """

    def __init__(
        self,
        path: Optional[Path | str] = None,
        source_lang: Optional[str] = None,
        target_lang: Optional[str] = None,
    ) -> None:
        self._entries: Dict[str, str] = {}
        self._all_columns: Dict[str, List[str]] = {}  # lang_header → [values]
        self._source_lang = source_lang
        self._target_lang = target_lang
        if path is not None:
            self._load(Path(path))

    # ── public API ───────────────────────────────────────────────────

    @property
    def entries(self) -> Dict[str, str]:
        """Raw source→target mapping."""
        return dict(self._entries)

    @property
    def all_columns(self) -> Dict[str, List[str]]:
        """All language columns read from xlsx (header → values list)."""
        return dict(self._all_columns)

    def lookup(self, term: str) -> Optional[str]:
        """Case-insensitive lookup.  Returns target term or None."""
        key = term.strip().lower()
        return self._index.get(key)

    def is_empty(self) -> bool:
        return len(self._entries) == 0

    def get_dict(self) -> Dict[str, str]:
        """Return the full source→target dictionary."""
        return dict(self._entries)

    # ── loading ──────────────────────────────────────────────────────

    def _load(self, path: Path) -> None:
        if not path.exists():
            return
        suffix = path.suffix.lower()
        if suffix == ".json":
            self._load_json(path)
        elif suffix in (".xlsx", ".xls"):
            self._load_xlsx(path)

    def _load_json(self, path: Path) -> None:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                self._entries = {str(k): str(v) for k, v in data.items()}
        except (json.JSONDecodeError, OSError):
            self._entries = {}
        self._build_index()

    def _load_xlsx(self, path: Path) -> None:
        """
        Load glossary from an Excel file.

        Strategy:
        1. Read row 1 as headers.
        2. If source_lang/target_lang are given, try to match them to headers.
        3. If match found → use those columns as source/target.
        4. If no match → fall back to legacy 2-column (A=source, B=target).
        """
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(path), read_only=True)
            ws = wb.active

            # Read all rows
            all_rows: List[Tuple] = []
            for row in ws.iter_rows(values_only=True):
                all_rows.append(row)
            wb.close()

            if not all_rows:
                self._build_index()
                return

            headers = all_rows[0]
            data_rows = all_rows[1:]

            # Try multi-language matching
            src_col, tgt_col = self._find_lang_columns(headers)

            if src_col is not None and tgt_col is not None:
                # Multi-language mode
                logger.info(
                    "Glossary multi-lang: source col=%d (%s) → target col=%d (%s)",
                    src_col, headers[src_col], tgt_col, headers[tgt_col],
                )
                for row in data_rows:
                    if len(row) > max(src_col, tgt_col):
                        s = row[src_col]
                        t = row[tgt_col]
                        if s and t:
                            self._entries[str(s)] = str(t)
            else:
                # Legacy 2-column fallback
                logger.info("Glossary legacy 2-column mode (A→B)")
                for row in data_rows:
                    if len(row) >= 2 and row[0] and row[1]:
                        self._entries[str(row[0])] = str(row[1])

            # Store all columns for viewer
            if headers:
                for col_idx, header in enumerate(headers):
                    if header:
                        col_values = []
                        for row in data_rows:
                            if len(row) > col_idx and row[col_idx]:
                                col_values.append(str(row[col_idx]))
                            else:
                                col_values.append("")
                        self._all_columns[str(header)] = col_values

        except Exception as e:
            logger.warning("Failed to load glossary xlsx: %s", e)
            self._entries = {}

        self._build_index()

    def _find_lang_columns(
        self, headers: Tuple
    ) -> Tuple[Optional[int], Optional[int]]:
        """
        Find column indices for source and target languages in the header row.

        Returns (src_col_index, tgt_col_index) or (None, None) if not found.
        """
        if not self._source_lang or not self._target_lang:
            return None, None

        src_norm = _normalize_lang(self._source_lang)
        tgt_norm = _normalize_lang(self._target_lang)

        src_col = None
        tgt_col = None

        for idx, header in enumerate(headers):
            if header is None:
                continue
            h_norm = _normalize_lang(str(header))
            if h_norm == src_norm:
                src_col = idx
            if h_norm == tgt_norm:
                tgt_col = idx

        if src_col is not None and tgt_col is not None:
            return src_col, tgt_col

        return None, None

    def _build_index(self) -> None:
        """Build a lower-cased lookup index."""
        self._index: Dict[str, str] = {
            k.lower(): v for k, v in self._entries.items()
        }
