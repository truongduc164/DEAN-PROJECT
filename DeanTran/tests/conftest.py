"""
Shared test fixtures – every test file in this directory gets these
automatically via pytest's conftest mechanism.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Ensure project root is importable
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.translators.translator_service import MockTranslator


@pytest.fixture
def mock_translator() -> MockTranslator:
    """Provide a fresh MockTranslator – never requires an API key."""
    return MockTranslator()
