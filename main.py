import sys
import traceback
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMessageBox

from constants import APP_ICON_PATH
from main_window import MainWindow


def _set_high_dpi_attributes_if_supported() -> None:
    # Qt 6 no longer needs these attributes, but setting them when available
    # keeps compatibility with older Qt builds.
    candidates = (Qt, getattr(Qt, "ApplicationAttribute", None))
    for enum_owner in candidates:
        if enum_owner is None:
            continue
        for name in ("AA_EnableHighDpiScaling", "AA_UseHighDpiPixmaps"):
            attr = getattr(enum_owner, name, None)
            if attr is not None:
                QApplication.setAttribute(attr)


def main() -> int:
    _set_high_dpi_attributes_if_supported()
    app = QApplication(sys.argv)
    app_icon = QIcon(str(APP_ICON_PATH)) if APP_ICON_PATH.exists() else QIcon()
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    try:
        w = MainWindow()
    except Exception:
        error_text = traceback.format_exc()
        log_path = Path("startup_error.log").resolve()
        try:
            log_path.write_text(error_text, encoding="utf-8")
            detail = f"Details were written to:\n{log_path}"
        except Exception:
            detail = "Could not write startup_error.log."
        QMessageBox.critical(
            None,
            "Startup Error",
            "Application failed to start.\n\n"
            f"{detail}",
        )
        return 1
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())

