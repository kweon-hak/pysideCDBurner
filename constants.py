from pathlib import Path

# Filesystem masks for IMAPI2
FS_ISO9660 = 1
FS_JOLIET = 2
FS_UDF = 4

# Application window title
APP_TITLE = "PySide CD Burner"

# Application icon location (same folder as main modules by default)
APP_ICON_PATH = Path(__file__).resolve().parent / "app_icon.ico"
