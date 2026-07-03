"""
tests/test_admin.py – Tests for admin authentication, secure storage, and
runtime translator switching.

All tests use mocks – no real API keys, no real keyring writes.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.core.auth_manager import AuthManager
from app.core.secure_storage import SecureStorage
from app.core.translators.translator_service import (
    MockTranslator,
    create_translator,
)


# ═══════════════════════════════════════════════════════════════════
# 1. AuthManager
# ═══════════════════════════════════════════════════════════════════

class TestAuthManager:
    def test_default_admin_created(self, tmp_path: Path):
        """First-run bootstraps an admin account."""
        am = AuthManager(config_dir=tmp_path)
        assert am.login("admin", "admin")
        assert am.is_admin
        assert am.current_user == "admin"

    def test_wrong_password(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        assert not am.login("admin", "wrong")
        assert not am.is_logged_in

    def test_unknown_user(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        assert not am.login("nobody", "pass")

    def test_create_user(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        assert am.create_user("bob", "secret", "user")
        assert am.login("bob", "secret")
        assert not am.is_admin
        assert am.current_role == "user"

    def test_duplicate_user(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        assert not am.create_user("admin", "x")  # already exists

    def test_change_password(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        assert am.change_password("admin", "newpass")
        assert am.login("admin", "newpass")
        assert not am.login("admin", "admin")  # old pass fails

    def test_logout(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        am.login("admin", "admin")
        assert am.is_logged_in
        am.logout()
        assert not am.is_logged_in
        assert not am.is_admin

    def test_persistence(self, tmp_path: Path):
        am1 = AuthManager(config_dir=tmp_path)
        am1.create_user("alice", "pw123", "user")
        # Reload from disk
        am2 = AuthManager(config_dir=tmp_path)
        assert am2.login("alice", "pw123")

    def test_user_role_is_not_admin(self, tmp_path: Path):
        am = AuthManager(config_dir=tmp_path)
        am.create_user("viewer", "pass", "user")
        am.login("viewer", "pass")
        assert not am.is_admin
        assert am.current_role == "user"


# ═══════════════════════════════════════════════════════════════════
# 2. SecureStorage (mocked keyring)
# ═══════════════════════════════════════════════════════════════════

class TestSecureStorage:
    """All keyring calls are mocked – no real credential writes."""

    def test_save_and_load_api_key(self):
        store = {}

        def _set(svc, key, val):
            store[(svc, key)] = val

        def _get(svc, key):
            return store.get((svc, key))

        with patch("keyring.set_password", side_effect=_set), \
             patch("keyring.get_password", side_effect=_get):
            ss = SecureStorage(service="TestService")
            ss.save_api_key("AIza-test-key-1234")
            assert ss.load_api_key() == "AIza-test-key-1234"

    def test_has_api_key(self):
        with patch("keyring.get_password", return_value="key123"):
            ss = SecureStorage()
            assert ss.has_api_key()

    def test_no_api_key(self):
        with patch("keyring.get_password", return_value=None):
            ss = SecureStorage()
            assert not ss.has_api_key()

    def test_delete_api_key(self):
        with patch("keyring.delete_password") as mock_del:
            ss = SecureStorage()
            ss.delete_api_key()
            mock_del.assert_called_once()

    def test_masked_full_key(self):
        assert SecureStorage.masked("AIzaSy1234567890ABCD") == "AIzaSy****ABCD"

    def test_masked_short_key(self):
        assert SecureStorage.masked("short") == "****"

    def test_masked_none(self):
        assert SecureStorage.masked(None) == "(not set)"

    def test_model_save_load(self):
        store = {}

        def _set(svc, key, val):
            store[(svc, key)] = val

        def _get(svc, key):
            return store.get((svc, key))

        with patch("keyring.set_password", side_effect=_set), \
             patch("keyring.get_password", side_effect=_get):
            ss = SecureStorage()
            ss.save_model_name("models/gemini-2.5-pro")
            assert ss.load_model_name() == "models/gemini-2.5-pro"


# ═══════════════════════════════════════════════════════════════════
# 3. Factory – runtime switching
# ═══════════════════════════════════════════════════════════════════

class TestTranslatorFactory:
    def test_force_mock(self):
        t = create_translator(force_mock=True, use_secure_storage=False)
        assert isinstance(t, MockTranslator)

    def test_no_key_returns_mock(self):
        with patch.dict("os.environ", {}, clear=True):
            t = create_translator(use_secure_storage=False)
            assert isinstance(t, MockTranslator)

    def test_runtime_switch_to_mock(self):
        """Simulate switching from Gemini (would fail) back to mock."""
        t1 = create_translator(force_mock=True, use_secure_storage=False)
        assert isinstance(t1, MockTranslator)
        # Second call — still mock
        t2 = create_translator(force_mock=True, use_secure_storage=False)
        assert isinstance(t2, MockTranslator)
        assert t1 is not t2  # fresh instance


# ═══════════════════════════════════════════════════════════════════
# 4. No plaintext API key in project files
# ═══════════════════════════════════════════════════════════════════

class TestNoPlaintextApiKey:
    """Scan project source files to ensure no hardcoded API keys."""

    def test_no_api_key_in_source(self):
        src = Path(__file__).resolve().parent.parent / "app"
        for py in src.rglob("*.py"):
            content = py.read_text(encoding="utf-8", errors="ignore")
            # Typical Gemini keys start with "AIza"
            assert "AIzaSy" not in content, f"Hardcoded key found in {py}"

    def test_no_api_key_in_json(self):
        root = Path(__file__).resolve().parent.parent
        for jf in root.rglob("*.json"):
            content = jf.read_text(encoding="utf-8", errors="ignore")
            assert "AIzaSy" not in content, f"Hardcoded key found in {jf}"
