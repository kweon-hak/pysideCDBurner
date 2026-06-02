from pathlib import Path

# Filesystem masks for IMAPI2
FS_ISO9660 = 1
FS_JOLIET = 2
FS_UDF = 4

# About 화면의 텍스트 정보
APP_TITLE = "PySide CD Burner"
AUTHOR = "KHLEE"
UPDATED = "2026.04.16"
DESCRIPTION = "Simple, fast ISO creation and disc burning tool."

# Application icon location (same folder as main modules by default)
APP_ICON_PATH = Path(__file__).resolve().parent / "app_icon.ico"
