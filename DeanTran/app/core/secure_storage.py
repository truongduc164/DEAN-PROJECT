"""
SecureStorage – store and retrieve API keys using the OS keyring.

The key is stored under the service name ``DeanTran`` in the operating
system's credential manager (Windows Credential Vault / macOS Keychain /
Linux Secret Service).

**Security guarantees**:
  * API key is NEVER written to a JSON/config file.
  * API key is NEVER logged or printed.
  * The ``masked()`` helper returns ``"gemini-****ABCD"`` for UI display.
"""
from __future__ import annotations

from typing import Optional

import keyring


_SERVICE = "DeanTran"
_GEMINI_KEY = "gemini_api_key"
_DEEPSEEK_KEY = "deepseek_api_key"
_MODEL_KEY = "gemini_model_name"
_DEFAULT_MODEL = "gemini-3.1-flash-lite"


class SecureStorage:
    """
    Thin wrapper around the OS keyring for DeanTran secrets.

    Parameters
    ----------
    service : str
        Keyring service name (default ``DeanTran``).
    """

    def __init__(self, service: str = _SERVICE) -> None:
        self.service = service

    # ── API key ──────────────────────────────────────────────────────

    def save_api_keys(self, api_keys: list[str]) -> None:
        """Persist multiple Gemini API keys to the OS keyring, encoded as JSON."""
        import json
        keyring.set_password(self.service, _GEMINI_KEY, json.dumps(api_keys))

    def load_api_keys(self) -> list[str]:
        """Retrieve the Gemini API keys as a list."""
        import json
        val = keyring.get_password(self.service, _GEMINI_KEY)
        if not val:
            return []
        try:
            parsed = json.loads(val)
            if isinstance(parsed, list):
                return parsed
        except json.JSONDecodeError:
            pass
        # Fallback for single key string
        return [val]

    def delete_api_keys(self) -> None:
        """Remove the stored API keys."""
        try:
            keyring.delete_password(self.service, _GEMINI_KEY)
        except keyring.errors.PasswordDeleteError:
            pass

    def has_api_keys(self) -> bool:
        """Check whether any API keys are stored."""
        return len(self.load_api_keys()) > 0

    # ── Backward compatibility adapters ──────────────────────────────

    def save_api_key(self, api_key: str) -> None:
        self.save_api_keys([api_key])

    def load_api_key(self) -> Optional[str]:
        keys = self.load_api_keys()
        return keys[0] if keys else None

    def delete_api_key(self) -> None:
        self.delete_api_keys()

    def has_api_key(self) -> bool:
        return self.has_api_keys()

    # ── DeepSeek API Key ──────────────────────────────────────────────

    def save_deepseek_key(self, key: str) -> None:
        """Persist DeepSeek API key to the OS keyring."""
        keyring.set_password(self.service, _DEEPSEEK_KEY, key)

    def load_deepseek_key(self) -> str:
        """Retrieve DeepSeek API key from the OS keyring."""
        return keyring.get_password(self.service, _DEEPSEEK_KEY) or ""

    def delete_deepseek_key(self) -> None:
        """Remove the stored DeepSeek API key."""
        try:
            keyring.delete_password(self.service, _DEEPSEEK_KEY)
        except keyring.errors.PasswordDeleteError:
            pass

    # ── Model name ───────────────────────────────────────────────────

    def save_model_name(self, model: str) -> None:
        keyring.set_password(self.service, _MODEL_KEY, model)

    def load_model_name(self) -> str:
        return keyring.get_password(self.service, _MODEL_KEY) or _DEFAULT_MODEL

    # ── Helpers ──────────────────────────────────────────────────────

    @staticmethod
    def masked(api_key: Optional[str]) -> str:
        """Return a masked version of the key for safe UI display."""
        if not api_key:
            return "(not set)"
        if len(api_key) <= 8:
            return "****"
        return api_key[:6] + "****" + api_key[-4:]
