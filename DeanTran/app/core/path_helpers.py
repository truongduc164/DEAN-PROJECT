"""
path_helpers.py - Helper functions for resolving paths safely in both
development (script) and production (PyInstaller frozen) environments.
"""
import sys
from pathlib import Path

def get_project_root() -> Path:
    """
    Trả về thư mục gốc của dự án.
    - Khi chạy code Python (chưa build): Là thư mục chứa file run.py
    - Khi chạy file .exe (đã build): Là thư mục chứa file .exe (giúp config nằm cạnh .exe)
    """
    if getattr(sys, 'frozen', False):
        # Đang chạy từ file .exe do PyInstaller build
        return Path(sys.executable).parent
    else:
        # Đang chạy từ mã nguồn (.py)
        # File này nằm ở app/core/path_helpers.py -> thư mục gốc là .parents[2]
        return Path(__file__).resolve().parents[2]

def get_configs_dir() -> Path:
    """Trả về đường dẫn tới thư mục configs/"""
    d = get_project_root() / "configs"
    d.mkdir(parents=True, exist_ok=True)
    return d

def get_logs_dir() -> Path:
    """Trả về đường dẫn tới thư mục logs/"""
    d = get_project_root() / "logs"
    d.mkdir(parents=True, exist_ok=True)
    return d

def get_outputs_dir() -> Path:
    """Trả về đường dẫn tới thư mục outputs/"""
    d = get_project_root() / "outputs"
    d.mkdir(parents=True, exist_ok=True)
    return d
