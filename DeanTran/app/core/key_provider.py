"""
KeyProvider – unified Gemini API key loading with pool rotation.

Priority:
1. Keyring (admin stored via SecureStorage)
2. Encrypted pool file (gemini_keys_pool.json.enc or .json)
3. Environment variable GEMINI_API_KEY

Supports key rotation: if a key errors (401/429/503), switch to next.
Dead keys get a 5-minute cooldown before retry.
"""
from __future__ import annotations

import json
import logging
import os
import time
from pathlib import Path
from typing import Optional

from app.core.path_helpers import get_configs_dir

logger = logging.getLogger("DeanTran.key_provider")

_POOL_FILE_ENC = "gemini_keys_pool.json.enc"
_POOL_FILE_PLAIN = "gemini_keys_pool.json"
_DEAD_COOLDOWN = 300  # 5 minutes


class KeyProvider:
    """
    Unified key provider with rotation support.

    Usage::

        kp = KeyProvider()
        key = kp.get_key()       # returns best available key
        kp.report_error(key)     # mark key as dead, rotate
        key2 = kp.get_key()      # returns next key
    """

    def __init__(self, config_dir: Optional[Path | str] = None) -> None:
        if config_dir is None:
            self.config_dir = get_configs_dir()
        else:
            self.config_dir = Path(config_dir)

        self._keys: list[str] = []
        self._source: str = "none"
        self._current_idx: int = 0
        self._dead_keys: dict[str, float] = {}  # key_prefix → expiry timestamp

        self._load_keys()

    @property
    def key_source(self) -> str:
        return self._source

    @property
    def key_loaded(self) -> bool:
        return len(self._keys) > 0

    @property
    def key_count(self) -> int:
        return len(self._keys)

    @property
    def current_index(self) -> int:
        return self._current_idx

    @property
    def current_key(self) -> Optional[str]:
        if not self._keys:
            return None
        return self._keys[self._current_idx]


    def get_key(self) -> Optional[str]:
        """Return the next available (non-dead) key, or None."""
        if not self._keys:
            return None

        now = time.time()
        # Try all keys starting from current index
        for _ in range(len(self._keys)):
            key = self._keys[self._current_idx]
            prefix = key[:6] if len(key) >= 6 else key
            dead_until = self._dead_keys.get(prefix, 0)

            if now >= dead_until:
                # Key is alive
                if prefix in self._dead_keys:
                    del self._dead_keys[prefix]
                    logger.info("Key %s**** recovered from cooldown", prefix)
                return key

            # Key is dead, try next
            self._current_idx = (self._current_idx + 1) % len(self._keys)

        logger.warning("All %d keys are in cooldown!", len(self._keys))
        # Return the first key anyway (cooldown expired by now hopefully)
        return self._keys[0]

    def report_error(self, key: str, error_code: int = 0) -> None:
        """Mark a key as dead and rotate to next."""
        prefix = key[:6] if len(key) >= 6 else key
        self._dead_keys[prefix] = time.time() + _DEAD_COOLDOWN
        logger.warning(
            "Key %s**** marked dead (error=%d), cooldown %ds",
            prefix, error_code, _DEAD_COOLDOWN,
        )
        self._current_idx = (self._current_idx + 1) % len(self._keys)

    def report_success(self, key: str) -> None:
        """Clear dead status for a key on success."""
        prefix = key[:6] if len(key) >= 6 else key
        self._dead_keys.pop(prefix, None)

    def log_status(self) -> None:
        """Log key provider status."""
        alive = self.key_count - len(self._dead_keys)
        if self._keys:
            key = self._keys[self._current_idx]
            prefix = key[:6] if len(key) >= 6 else "***"
        else:
            prefix = "(none)"
        logger.info(
            "key_source=%s key_loaded=%s key_count=%d alive=%d key_prefix=%s****",
            self._source, self.key_loaded, self.key_count, alive, prefix,
        )

    # ── Internal ─────────────────────────────────────────────────────

    def _load_keys(self) -> None:
        """Load keys in priority order: keyring → pool → env."""
        # 1. Keyring
        try:
            from app.core.secure_storage import SecureStorage
            ss = SecureStorage()
            keys = ss.load_api_keys()
            if keys:
                for k in keys:
                    if k.strip() and k.strip() not in self._keys:
                        self._keys.append(k.strip())
                self._source = "keyring"
                logger.info("Loaded %d key(s) from keyring", len(keys))
        except Exception as exc:
            logger.debug("Keyring load failed: %s", exc)

        # 2. Pool file
        pool_keys = self._load_pool()
        if pool_keys:
            for pk in pool_keys:
                if pk not in self._keys:
                    self._keys.append(pk)
            if self._source == "none":
                self._source = "pool"
            elif pool_keys:
                self._source = "keyring+pool"
            logger.info("Loaded %d key(s) from pool file", len(pool_keys))

        # 3. Env var
        env_key = os.environ.get("GEMINI_API_KEY", "").strip()
        if env_key and env_key not in self._keys:
            self._keys.append(env_key)
            if self._source == "none":
                self._source = "env"
            logger.info("Loaded key from environment")

        if not self._keys:
            logger.warning("No API keys found from any source")

    def _load_pool(self) -> list[str]:
        """Load keys from pool file (plain JSON for now)."""
        # Try plain JSON first
        plain = self.config_dir / _POOL_FILE_PLAIN
        if plain.exists():
            try:
                data = json.loads(plain.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return [k.strip() for k in data if isinstance(k, str) and k.strip()]
                if isinstance(data, dict) and "keys" in data:
                    return [k.strip() for k in data["keys"] if isinstance(k, str) and k.strip()]
            except (json.JSONDecodeError, OSError) as exc:
                logger.warning("Pool file parse error: %s", exc)

        # Encrypted pool (placeholder — same JSON format but could be encrypted)
        enc = self.config_dir / _POOL_FILE_ENC
        if enc.exists():
            try:
                data = json.loads(enc.read_text(encoding="utf-8"))
                if isinstance(data, list):
                    return [k.strip() for k in data if isinstance(k, str) and k.strip()]
                if isinstance(data, dict) and "keys" in data:
                    return [k.strip() for k in data["keys"] if isinstance(k, str) and k.strip()]
            except (json.JSONDecodeError, OSError) as exc:
                logger.warning("Encrypted pool parse error: %s", exc)

        return []
