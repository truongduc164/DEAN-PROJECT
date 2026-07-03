import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _python_candidates() -> list[Path]:
    current = Path(sys.executable).resolve()
    py_root = Path.home() / "AppData" / "Local" / "Programs" / "Python"
    if not py_root.exists():
        return []

    candidates = sorted(py_root.glob("Python*/python.exe"), reverse=True)
    return [candidate.resolve() for candidate in candidates if candidate.resolve() != current]


def _find_python_with_pyside() -> Path | None:
    for candidate in _python_candidates():
        try:
            probe = subprocess.run(
                [str(candidate), "-c", "import PySide6"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
                check=False,
            )
        except (OSError, subprocess.SubprocessError):
            continue

        if probe.returncode == 0:
            return candidate
    return None


def _ensure_qt_runtime() -> None:
    try:
        import PySide6  # noqa: F401
        return
    except ModuleNotFoundError as exc:
        if exc.name != "PySide6":
            raise

    candidate = _find_python_with_pyside()
    if candidate is not None:
        rerun = subprocess.run(
            [str(candidate), str(PROJECT_ROOT / "run.py"), *sys.argv[1:]],
            check=False,
        )
        raise SystemExit(rerun.returncode)

    current = Path(sys.executable)
    raise SystemExit(
        "PySide6 is not installed for the current Python interpreter.\n"
        f"Current interpreter: {current}\n"
        "Install it with:\n"
        f'  "{current}" -m pip install PySide6\n'
        "or run this app with a Python that already has PySide6."
    )


_ensure_qt_runtime()

if "--check-bootstrap" in sys.argv:
    print(sys.executable)
    raise SystemExit(0)

from PySide6.QtWidgets import QApplication, QInputDialog, QMessageBox, QLineEdit

from app.ui.main_window import MainWindow
from app.core.app_password import get_app_password


def _check_app_password(qapp: QApplication) -> bool:
    """Prompt for app password. Returns True if correct."""
    current_pw = get_app_password()
    admin_pw = "@Quanambotat12"
    for attempt in range(3):
        pwd, ok = QInputDialog.getText(
            None,
            "🔒 Nhập mật khẩu",
            f"Nhập mật khẩu để mở ứng dụng (lần {attempt + 1}/3):",
            QLineEdit.Password,
        )
        if not ok:
            return False
        if pwd == current_pw or pwd == admin_pw:
            return True
        remaining = 2 - attempt
        if remaining > 0:
            QMessageBox.warning(
                None, "Sai mật khẩu",
                f"❌ Mật khẩu không đúng! Còn {remaining} lần thử."
            )
    QMessageBox.critical(None, "Bị khóa", "🔒 Đã nhập sai 3 lần. Ứng dụng sẽ đóng.")
    return False

if __name__ == "__main__":
    app = QApplication(sys.argv)

    if not _check_app_password(app):
        sys.exit(1)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
