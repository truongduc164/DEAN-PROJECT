"""
AuthManager – role-based authentication with bcrypt-hashed passwords.

Roles
-----
* ``admin`` – full access including API key management.
* ``user``  – translation features only, no API settings.

Credentials are stored in ``configs/users.json`` as::

    {
        "admin": {
            "password_hash": "$2b$12$...",
            "role": "admin"
        }
    }

On first run, if the file does not exist, a default admin account is
created with username ``admin`` and password ``admin``.  The admin
should change the password immediately via the UI.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import bcrypt
import time

from app.core.path_helpers import get_configs_dir


class AuthManager:
    """
    Manage user authentication and role enforcement.

    Parameters
    ----------
    config_dir : Path | str | None
        Directory for ``users.json``.  Defaults to ``<project>/configs``.
    """

    FILENAME = "users.json"
    ROLE_ADMIN = "admin"
    ROLE_OPERATOR = "operator"
    ROLE_USER = "user"
    
    SESSION_TIMEOUT = 1800  # 30 minutes

    def __init__(self, config_dir: Optional[Path | str] = None) -> None:
        if config_dir is None:
            self.config_dir = get_configs_dir()
        else:
            self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self._file = self.config_dir / self.FILENAME
        self._users: dict = {}
        self._current_user: Optional[str] = None
        self._current_role: Optional[str] = None
        self._last_activity: float = 0.0
        self._load()

    # ── public API ───────────────────────────────────────────────────

    def login(self, username: str, password: str) -> bool:
        """
        Authenticate.  Returns True on success and stores current session.
        """
        user_entry = self._users.get(username)
        if user_entry is None:
            return False
        stored_hash = user_entry["password_hash"].encode("utf-8")
        if bcrypt.checkpw(password.encode("utf-8"), stored_hash):
            self._current_user = username
            self._current_role = user_entry.get("role", self.ROLE_USER)
            self.touch_session()
            return True
        return False

    def logout(self) -> None:
        self._current_user = None
        self._current_role = None
        self._last_activity = 0.0

    def touch_session(self) -> None:
        """Update last activity timestamp."""
        self._last_activity = time.time()

    def _check_session(self) -> None:
        """Check if admin/operator session timed out and logout if needed."""
        if self._current_role in (self.ROLE_ADMIN, self.ROLE_OPERATOR):
            if time.time() - self._last_activity > self.SESSION_TIMEOUT:
                self.logout()

    @property
    def is_logged_in(self) -> bool:
        self._check_session()
        return self._current_user is not None

    @property
    def is_admin(self) -> bool:
        self._check_session()
        return self._current_role == self.ROLE_ADMIN
        
    @property
    def is_operator(self) -> bool:
        self._check_session()
        return self._current_role == self.ROLE_OPERATOR

    @property
    def current_user(self) -> Optional[str]:
        self._check_session()
        return self._current_user

    @property
    def current_role(self) -> Optional[str]:
        return self._current_role

    def create_user(
        self, username: str, password: str, role: str = ROLE_USER
    ) -> bool:
        """Create a new user.  Returns False if the username already exists."""
        if username in self._users:
            return False
        hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt())
        self._users[username] = {
            "password_hash": hashed.decode("utf-8"),
            "role": role,
        }
        self._save()
        return True

    def change_password(self, username: str, new_password: str) -> bool:
        """Change password for an existing user."""
        if username not in self._users:
            return False
        hashed = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt())
        self._users[username]["password_hash"] = hashed.decode("utf-8")
        self._save()
        return True

    # ── internal ─────────────────────────────────────────────────────

    def _load(self) -> None:
        if self._file.exists():
            try:
                self._users = json.loads(
                    self._file.read_text(encoding="utf-8")
                )
            except (json.JSONDecodeError, OSError):
                self._users = {}
        if not self._users:
            # Bootstrap default admin account
            self.create_user("admin", "admin", self.ROLE_ADMIN)

    def _save(self) -> None:
        self._file.write_text(
            json.dumps(self._users, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
