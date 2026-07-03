from __future__ import annotations

import sys
from pathlib import Path
import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.pdf_converter import pdf_to_word, pdf_to_excel, pdf_to_ppt


def test_pdf_to_word(tmp_path: Path):
    pdf_path = Path(r"d:\Tạm\HLM 06 Rev. C\HLM 01 - rev C_Zh_Vi.pdf")
    if not pdf_path.exists():
        pytest.skip("Test PDF file not found")

    docx_path = tmp_path / "output.docx"
    pdf_to_word(pdf_path, docx_path)

    assert docx_path.exists()
    assert docx_path.stat().st_size > 0


def test_pdf_to_excel(tmp_path: Path):
    pdf_path = Path(r"d:\Tạm\HLM 06 Rev. C\HLM 01 - rev C_Zh_Vi.pdf")
    if not pdf_path.exists():
        pytest.skip("Test PDF file not found")

    xlsx_path = tmp_path / "output.xlsx"
    pdf_to_excel(pdf_path, xlsx_path)

    assert xlsx_path.exists()
    assert xlsx_path.stat().st_size > 0


def test_pdf_to_ppt(tmp_path: Path):
    pdf_path = Path(r"d:\Tạm\HLM 06 Rev. C\HLM 01 - rev C_Zh_Vi.pdf")
    if not pdf_path.exists():
        pytest.skip("Test PDF file not found")

    pptx_path = tmp_path / "output.pptx"
    pdf_to_ppt(pdf_path, pptx_path)

    assert pptx_path.exists()
    assert pptx_path.stat().st_size > 0
