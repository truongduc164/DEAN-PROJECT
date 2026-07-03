"""
output_naming – Shared helper for generating unique output file paths.

Used by PowerPoint, Excel, Word, and PDF processors to prevent
overwriting previously translated files.

Example:
    input.pptx → input_Vi.pptx → input_Vi(1).pptx → input_Vi(2).pptx
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger("DeanTran.output_naming")


# ── Language suffix map (single source of truth) ─────────────────────

LANG_SUFFIX = {
    "Vietnamese": "Vi",
    "English": "En",
    "Japanese": "Ja",
    "Chinese": "Zh",
    "Korean": "Ko",
}


def get_lang_suffix(target_lang: str) -> str:
    """Return the short language code for the given target language."""
    return LANG_SUFFIX.get(target_lang, target_lang[:2])


def build_output_path(
    input_path: Path,
    target_lang: str,
    output_dir: Optional[Path] = None,
    extension: Optional[str] = None,
) -> Path:
    """Build the BASE output path (without collision check).

    Parameters
    ----------
    input_path : Path
        Original input file.
    target_lang : str
        Target language name, e.g. "Vietnamese".
    output_dir : Path | None
        Directory for output. Defaults to same directory as input.
    extension : str | None
        Override extension (e.g. ".docx"). Defaults to input's extension.

    Returns
    -------
    Path
        e.g.  ``input_Vi.pptx``
    """
    suffix = get_lang_suffix(target_lang)
    ext = extension or input_path.suffix
    parent = output_dir or input_path.parent
    return parent / f"{input_path.stem}_{suffix}{ext}"


def get_unique_output_path(
    input_path: Path,
    target_lang: str,
    output_dir: Optional[Path] = None,
    extension: Optional[str] = None,
) -> Path:
    """Build an output path that does NOT collide with existing files.

    Behaviour
    ---------
    1. Try  ``input_Vi.pptx``
    2. If it exists → ``input_Vi(1).pptx``
    3. If that exists → ``input_Vi(2).pptx``
    4. … and so on until a free name is found.

    Logs a warning when a collision is detected.

    Parameters
    ----------
    input_path : Path
        Original input file.
    target_lang : str
        Target language name, e.g. "Vietnamese".
    output_dir : Path | None
        Output directory. Defaults to same directory as input.
    extension : str | None
        Override file extension.

    Returns
    -------
    Path
        A path that is guaranteed not to exist at the time of calling.
    """
    base = build_output_path(input_path, target_lang, output_dir, extension)

    if not base.exists():
        logger.info("Output path selected: %s", base.name)
        return base

    # Collision detected — find a free numbered variant
    logger.warning("Collision detected: '%s' already exists", base.name)
    counter = 1
    while True:
        candidate = base.parent / f"{base.stem}({counter}){base.suffix}"
        if not candidate.exists():
            logger.info("Resolved to: %s", candidate.name)
            return candidate
        counter += 1
        if counter > 9999:  # safety valve
            raise RuntimeError(
                f"Cannot find unique name after 9999 attempts: {base}"
            )
