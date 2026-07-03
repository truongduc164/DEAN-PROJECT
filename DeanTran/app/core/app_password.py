"""
app_password – quản lý mật khẩu mở ứng dụng.

Mật khẩu mặc định hardcode, nhưng có thể đổi từ xa qua Telegram.
Khi đổi, mật khẩu mới được lưu vào configs/app_password.txt.
Nếu file tồn tại → dùng mật khẩu trong file.
Nếu không → dùng mật khẩu mặc định.
"""
from pathlib import Path

from app.core.path_helpers import get_configs_dir

_DEFAULT_PASSWORD = "@Diatangbotat13"
_PASSWORD_FILE = "app_password.txt"


def _get_file() -> Path:
    return get_configs_dir() / _PASSWORD_FILE


def get_app_password() -> str:
    """Lấy mật khẩu app hiện tại."""
    f = _get_file()
    if f.exists():
        pw = f.read_text(encoding="utf-8").strip()
        if pw:
            return pw
    return _DEFAULT_PASSWORD


def set_app_password(new_password: str) -> None:
    """Đổi mật khẩu app (lưu vào file)."""
    f = _get_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(new_password.strip(), encoding="utf-8")
