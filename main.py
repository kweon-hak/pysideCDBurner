import sys

from PySide6.QtCore import Qt, qVersion
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from constants import APP_ICON_PATH
from main_window import MainWindow


def main() -> int:
    if qVersion().startswith("5."):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())


