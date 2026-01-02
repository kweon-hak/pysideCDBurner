import os
import time
from pathlib import Path
from collections import deque

from PySide6.QtCore import Qt, QSettings, QTimer, QDir, QThread, Signal, QUrl
from PySide6.QtGui import QAction, QActionGroup, QIcon, QTextCursor, QPixmap, QPainter, QColor, QPen, QShortcut, QKeySequence, QDesktopServices, QPalette
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QFormLayout, QGroupBox, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QHBoxLayout,
    QPushButton, QProgressBar, QTextEdit, QVBoxLayout, QWidget, QComboBox,
    QDialog, QAbstractItemView, QListView, QTreeView, QFileSystemModel,
    QSizePolicy, QRadioButton, QCheckBox, QSplitter, QStyle, QProxyStyle,
    QSpacerItem
)

from constants import APP_ICON_PATH, APP_TITLE, FS_ISO9660, FS_JOLIET, FS_UDF
from utils import sanitize_volume_label, force_dialog_accept_label
from widgets import FileFolderDialog, CustomIconProvider
from imapi import list_imapi_writers
from workers import BurnWorker, SizeWorker, IsoCreateWorker


class DropListWidget(QListWidget):
    """QListWidget with simple file/folder drag-and-drop support."""

    def __init__(self, add_callback, parent=None):
        super().__init__(parent)
        self._add_callback = add_callback
        self._hint_text = "Drag and drop files or folders here, or use Add files/folders."
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            urls = event.mimeData().urls()
            for url in urls:
                path = url.toLocalFile()
                if path:
                    self._add_callback(path)
        else:
            super().dropEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.count() == 0 and self._hint_text:
            painter = QPainter(self.viewport())
            pen = painter.pen()
            pen.setColor(QColor("#808080"))
            painter.setPen(pen)
            font = painter.font()
            font.setItalic(True)
            painter.setFont(font)
            rect = self.viewport().rect().adjusted(8, 8, -8, -8)
            painter.drawText(rect, Qt.AlignCenter | Qt.TextWordWrap, self._hint_text)
            painter.end()


class IconSpacingStyle(QProxyStyle):
    """Custom style to widen icon/text spacing on selected buttons."""
    def __init__(self, base_style: QStyle, spacing: int = 10):
        super().__init__(base_style)
        self._spacing = spacing
        # Some Qt builds lack PM_ButtonIconSpacing; capture if available.
        self._pm_icon_spacing = getattr(QStyle, "PM_ButtonIconSpacing", None)

    def pixelMetric(self, metric, option=None, widget=None):
        if self._pm_icon_spacing is not None and metric == self._pm_icon_spacing:
            return self._spacing
        return super().pixelMetric(metric, option, widget)


class WriterLookupWorker(QThread):
    result = Signal(object, object)  # writers or None, error message or None
    def run(self):
        try:
            writers = list_imapi_writers()
            self.result.emit(writers, None)
        except Exception as e:
            self.result.emit(None, str(e))


class SpeedLookupWorker(QThread):
    result = Signal(str, list)  # uid, entries list[(label, val)]
    error = Signal(str, object)  # uid, error
    def __init__(self, uid: str, parent=None):
        super().__init__(parent)
        self.uid = uid

    def run(self):
        coinit = False
        pythoncom = None
        try:
            import pythoncom as _pythoncom
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            coinit = True
            import win32com.client
            rec = win32com.client.Dispatch("IMAPI2.MsftDiscRecorder2")
            rec.InitializeDiscRecorder(self.uid)
            fmt = win32com.client.Dispatch("IMAPI2.MsftDiscFormat2Data")
            fmt.Recorder = rec
            speeds = set()
            for d in getattr(fmt, "WriteSpeedDescriptors", []):
                try:
                    speeds.add(int(getattr(d, "WriteSpeed", 0)))
                except Exception:
                    continue
            if not speeds:
                try:
                    for s in getattr(fmt, "SupportedWriteSpeeds", []):
                        speeds.add(int(s))
                except Exception:
                    pass
            unique = sorted({s for s in speeds if s > 0}, reverse=True)
            entries = []
            for s in unique:
                label = f"{int(round(s/1385.0))}x (~{s} KB/s)"
                entries.append((label, s))
            self.result.emit(self.uid, entries)
        except Exception as e:
            self.error.emit(self.uid, e)
        finally:
            if coinit and pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


class MediaStatusWorker(QThread):
    result = Signal(str, bool, bool, object)  # uid, blank, supported, capacity_bytes or None
    error = Signal(object)
    def __init__(self, uid: str, parent=None):
        super().__init__(parent)
        self.uid = uid

    def run(self):
        coinit = False
        pythoncom = None
        try:
            import pythoncom as _pythoncom
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            coinit = True
            import win32com.client
            rec = win32com.client.Dispatch("IMAPI2.MsftDiscRecorder2")
            rec.InitializeDiscRecorder(self.uid)
            fmt = win32com.client.Dispatch("IMAPI2.MsftDiscFormat2Data")
            fmt.Recorder = rec
            blank = fmt.MediaHeuristicallyBlank
            supported = fmt.IsCurrentMediaSupported(rec)
            try:
                total_sectors = int(getattr(fmt, "TotalSectorsOnMedia", 0) or 0)
                sector_size = int(getattr(fmt, "SectorSize", 2048) or 2048)
                capacity = max(0, total_sectors) * max(1, sector_size)
            except Exception:
                capacity = None
            self.result.emit(self.uid, bool(blank), bool(supported), capacity)
        except Exception as e:
            self.error.emit(e)
        finally:
            if coinit and pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self._app_icon = QIcon(str(APP_ICON_PATH)) if APP_ICON_PATH.exists() else QIcon()
        self.setWindowIcon(self._app_icon)

        self.worker: BurnWorker | None = None
        self._writer_worker: QThread | None = None
        self.settings = QSettings("PySideCDBurner", "PySideCDBurner")
        self._fs_options = [
            ("ISO9660 + Joliet (default)", FS_ISO9660 | FS_JOLIET),
            ("ISO9660 only", FS_ISO9660),
            ("UDF", FS_UDF),
            ("ISO9660 + Joliet + UDF", FS_ISO9660 | FS_JOLIET | FS_UDF),
        ]
        _default_fs = FS_ISO9660 | FS_JOLIET
        last_dir_add_val = self.settings.value("last_dir_add", str(Path.home()))
        self.last_dir_add = Path(last_dir_add_val) if Path(last_dir_add_val).exists() else Path.home()
        last_dir_iso_in_val = self.settings.value("last_dir_iso_in", str(self.last_dir_add))
        self.last_dir_iso_in = Path(last_dir_iso_in_val) if Path(last_dir_iso_in_val).exists() else self.last_dir_add
        last_dir_iso_out_val = self.settings.value("last_dir_iso_out", str(self.last_dir_add))
        self.last_dir_iso_out = Path(last_dir_iso_out_val) if Path(last_dir_iso_out_val).exists() else self.last_dir_add
        self._media_blank = False
        self._burn_started_at: float | None = None
        self._burning: bool = False
        fs_mask_val = self.settings.value("fs_mask", _default_fs)
        try:
            mask_val = int(fs_mask_val)
        except Exception:
            mask_val = _default_fs
        self._fs_mask = mask_val

        self._path_sizes: dict[str, int] = {}
        self._total_size: int = 0
        self._pending_size: set[str] = set()
        self._size_queue = deque()  # Queue for pending size calculations
        self._current_size_worker: SizeWorker | None = None
        self._pending_close: bool = False
        self._known_writers: set[str] = set()
        self._speed_cache: dict[str, list[tuple[str, int | None]]] = {}
        self._last_media_state: tuple | None = None
        self._media_capacity_bytes: int | None = None
        self._last_media_capacity: int | None = None
        self._last_status_text: str | None = None
        self._last_media_text: str | None = None
        # 용량 비교 시 약간의 여유(최소 32MB 또는 1%)를 두고 판단한다.
        self._capacity_headroom_bytes: int = 32 * 1024 * 1024
        self._iso_path: str | None = None          # existing ISO to burn
        self._iso_size: int | None = None
        self._iso_out_path: str | None = None      # destination for creating ISO
        self._iso_out_seen: set[str] = set()       # paths burned once (prompt on subsequent burns)
        self._active_job: str | None = None        # "burn" or "iso"
        self._media_worker: QThread | None = None
        self._speed_worker: QThread | None = None
        self._media_usage_dirty: bool = False
        self._theme = str(self.settings.value("theme", "light")).lower()
        self._default_palette = QApplication.instance().palette() if QApplication.instance() else None

        self._init_ui()
        self._init_actions()
        self._apply_theme(self._theme)

        self.resize(800, 600)
        if self.settings.value("geometry") is not None:
            self.restoreGeometry(self.settings.value("geometry"))
        self.refresh_writers()
        self._apply_label_rules()
        self._set_button_icons()

        self.media_timer = QTimer(self)
        self.media_timer.setInterval(4000)
        self.media_timer.timeout.connect(lambda: self.update_media_status(check_writers=False))
        self.media_timer.start()
        # 주기적으로 드라이브 목록을 자동 리프레시한다(15초 간격, 비동기 워커 사용).
        self.writer_timer = QTimer(self)
        self.writer_timer.setInterval(15000)
        # 주기 리프레시에서는 로그를 남기지 않는다(로그 스팸 방지).
        self.writer_timer.timeout.connect(lambda: self.refresh_writers(log=False))
        self.writer_timer.start()
        self.elapsed_timer = QTimer(self)
        self.elapsed_timer.setInterval(500)
        self.elapsed_timer.timeout.connect(self._update_elapsed_label)
        self._update_list_buttons_and_burn_state()

    def _init_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(8, 8, 8, 8)
        root_layout.setSpacing(0)

        splitter = QSplitter(Qt.Horizontal, root)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)

        left_widget = QWidget(splitter)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        left_splitter = QSplitter(Qt.Vertical, left_widget)
        left_splitter.setChildrenCollapsible(False)
        left_splitter.setHandleWidth(6)
        left_layout.addWidget(left_splitter)

        right_widget = QWidget(splitter)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        # Files
        g = QGroupBox("Files / Folders to burn")
        gl = QVBoxLayout(g)
        self.list = DropListWidget(self._on_drop_add, self)
        gl.addWidget(self.list, 1)
        self.total_size_label = QLabel("Total size: 0 B")
        self.total_size_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        gl.addWidget(self.total_size_label)
        self.media_usage_label = QLabel("Media usage: --")
        self.media_usage_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        gl.addWidget(self.media_usage_label)
        btn_row = QHBoxLayout()
        self.btn_add_files = QPushButton("Add files/folders")
        self.btn_remove = QPushButton("Remove selected")
        self.btn_clear = QPushButton("Clear")
        btn_row.addWidget(self.btn_add_files)
        btn_row.addStretch(1)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        gl.addLayout(btn_row)
        left_splitter.addWidget(g)

        # Log panel under files
        log_box = QGroupBox("Log")
        log_layout = QVBoxLayout(log_box)
        log_layout.setContentsMargins(8, 8, 8, 8)
        toggle_row = QHBoxLayout()
        self.btn_hide_log = QPushButton("Hide log")
        self.btn_hide_log.setFlat(True)
        self.btn_hide_log.setCursor(Qt.PointingHandCursor)
        self.btn_hide_log.clicked.connect(lambda: self._toggle_log_visibility(False))
        self.btn_hide_log.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.btn_clear_log = QPushButton("Clear log")
        self.btn_clear_log.setFlat(True)
        self.btn_clear_log.setCursor(Qt.PointingHandCursor)
        self.btn_clear_log.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        self.btn_clear_log.clicked.connect(self._clear_log)
        toggle_row.setContentsMargins(0, 0, 0, 4)
        toggle_row.addStretch(1)
        toggle_row.addWidget(self.btn_clear_log, alignment=Qt.AlignRight)
        toggle_row.addWidget(self.btn_hide_log, alignment=Qt.AlignRight)
        log_layout.addLayout(toggle_row)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumWidth(260)
        log_layout.addWidget(self.log, 1)
        self.log_container = log_box
        left_splitter.addWidget(self.log_container)
        left_splitter.setStretchFactor(0, 3)
        left_splitter.setStretchFactor(1, 2)

        # Settings
        g = QGroupBox("Settings")
        f = QFormLayout(g)
        self.volume = QLineEdit("DATA")
        f.addRow("Volume label:", self.volume)
        self.drive = QComboBox()
        self.btn_refresh = QPushButton()
        self.btn_refresh.setToolTip("Refresh drives")
        drow = QHBoxLayout()
        drow.addWidget(self.drive, 1)
        drow.addWidget(self.btn_refresh)
        f.addRow("CD/DVD Writer:", drow)
        self.write_speed = QComboBox()
        self.write_speed.addItem("Max (auto)", None)
        f.addRow("Write speed:", self.write_speed)
        self.fs_display = QLabel(self._fs_label(self._fs_mask))
        f.addRow("Filesystem:", self.fs_display)
        right_layout.addWidget(g)

        # Burn / status
        g = QGroupBox("Burn")
        gl = QVBoxLayout(g)
        gl.setContentsMargins(8, 8, 8, 8)

        action_row = QHBoxLayout()
        action_row.addWidget(QLabel("Action:"))
        self.action_burn_disc = QRadioButton("Burn to disc")
        self.action_create_iso = QRadioButton("Create ISO file")
        self.action_burn_disc.setChecked(True)
        action_row.addWidget(self.action_burn_disc)
        action_row.addWidget(self.action_create_iso)
        action_row.addStretch(1)
        gl.addLayout(action_row)

        iso_input_row = QHBoxLayout()
        self.chk_use_iso_input = QCheckBox("Use existing ISO to burn")
        iso_input_row.addWidget(self.chk_use_iso_input)
        iso_input_row.addStretch(1)
        gl.addLayout(iso_input_row)

        burn_iso_row = QHBoxLayout()
        self.iso_path_edit = QLineEdit()
        self.iso_path_edit.setPlaceholderText("Select ISO to burn")
        self.iso_path_edit.setReadOnly(True)
        self.btn_browse_iso = QPushButton("Browse ISO")
        self.btn_clear_iso = QPushButton("Clear")
        burn_iso_row.addWidget(self.iso_path_edit, 1)
        burn_iso_row.addWidget(self.btn_browse_iso)
        burn_iso_row.addWidget(self.btn_clear_iso)
        gl.addLayout(burn_iso_row)

        self.chk_verify = QCheckBox("Verify after operation")
        gl.addWidget(self.chk_verify)

        iso_out_row = QHBoxLayout()
        self.iso_out_path_edit = QLineEdit()
        self.iso_out_path_edit.setPlaceholderText("Select destination ISO file")
        iso_out_row.addWidget(self.iso_out_path_edit, 1)
        self.btn_browse_iso_out = QPushButton("ISO Output")
        iso_out_row.addWidget(self.btn_browse_iso_out)
        gl.addLayout(iso_out_row)

        gl.addSpacing(6)
        gl.addStretch(1)

        act_row = QHBoxLayout()
        # Add a leading space so icon and text have visible separation.
        self.btn_burn = QPushButton(" Burn")
        self.btn_stop = QPushButton(" Stop")
        for btn in (self.btn_burn, self.btn_stop):
            btn.setMinimumHeight(36)
            # Extra padding further separates icon and label.
            btn.setStyleSheet("padding-left: 14px; padding-right: 12px;")
        act_row.addWidget(self.btn_burn)
        act_row.addWidget(self.btn_stop)
        gl.addLayout(act_row)
        self.btn_stop.setEnabled(False)

        gl.addStretch(1)

        self.status = QLabel("Idle")
        self.status.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%p%")
        self.progress.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        progress_row = QHBoxLayout()
        progress_row.setContentsMargins(0, 0, 0, 0)
        progress_row.addWidget(self.progress)
        gl.addWidget(self.status, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        gl.addLayout(progress_row)
        self.progress_info = QLabel()
        self.progress_info.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        gl.addWidget(self.progress_info, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        self._reset_progress_info_label()
        self._ensure_elapsed_timer_running()

        self.mode_hint = QLabel("Tip: choose your burn mode first, then add files or select an ISO.")
        self.mode_hint.setWordWrap(True)
        self.mode_hint.setStyleSheet("color: #6b6b6b; font-size: 12px;")
        gl.addWidget(self.mode_hint)

        right_layout.addWidget(g, 1)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 3)

        root_layout.addWidget(splitter)

        self.status_label = QLabel("Idle")
        # 약간의 하단 여백을 주고 왼쪽에 배치한다.
        self.status_label.setContentsMargins(4, 0, 4, 2)
        self.media_status = QLabel("Media status: unknown")
        self.media_status.setContentsMargins(4, 0, 4, 0)
        self.media_status.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.media_status.setVisible(False)  # 상태바에서 미디어 상태는 숨긴다(중복 표시 제거)
        self.statusBar().addWidget(self.status_label, 0)
        # 상태바에는 미디어 상태를 추가하지 않는다

    def _init_actions(self):
        menu = self.menuBar().addMenu("File")
        fs_menu = menu.addMenu("Filesystem")
        self._setup_filesystem_actions(fs_menu)
        act_eject = QAction("Eject disc", self)
        act_eject.triggered.connect(self.eject_disc)
        menu.addAction(act_eject)
        self.act_eject = act_eject
        menu.addSeparator()
        act_exit = QAction("Exit", self)
        act_exit.triggered.connect(self.close)
        menu.addAction(act_exit)
        view_menu = self.menuBar().addMenu("View")
        self.act_log = QAction("Show Log Panel", self, checkable=True)
        self.act_log.setChecked(True)
        self.act_log.toggled.connect(lambda checked: self._toggle_log_visibility(checked))
        view_menu.addAction(self.act_log)
        self._toggle_log_visibility(self.act_log.isChecked())
        view_menu.addSeparator()
        theme_group = QActionGroup(self)
        self.act_theme_light = QAction("Light Theme", self, checkable=True)
        self.act_theme_dark = QAction("Dark Theme", self, checkable=True)
        theme_group.addAction(self.act_theme_light)
        theme_group.addAction(self.act_theme_dark)
        self.act_theme_light.setChecked(self._theme != "dark")
        self.act_theme_dark.setChecked(self._theme == "dark")
        self.act_theme_light.triggered.connect(lambda: self._set_theme("light"))
        self.act_theme_dark.triggered.connect(lambda: self._set_theme("dark"))
        view_menu.addAction(self.act_theme_light)
        view_menu.addAction(self.act_theme_dark)
        help_menu = self.menuBar().addMenu("Help")
        act_about = QAction("About", self)
        act_about.triggered.connect(self.show_about)
        help_menu.addAction(act_about)

        self.btn_add_files.clicked.connect(self.add_files)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self._clear_list)
        self.btn_browse_iso.clicked.connect(self._select_iso_file)
        self.btn_clear_iso.clicked.connect(self._clear_iso_file)
        self.btn_browse_iso_out.clicked.connect(self._select_iso_output)
        self.chk_use_iso_input.toggled.connect(self._on_use_iso_toggled)
        self.action_burn_disc.toggled.connect(self._on_action_mode_changed)
        self.list.itemSelectionChanged.connect(self._update_remove_button_state)
        self.btn_refresh.clicked.connect(self.refresh_writers)
        self.btn_burn.clicked.connect(self.start_burn)
        self.btn_stop.clicked.connect(self.stop_burn)
        self.drive.currentIndexChanged.connect(lambda: self.update_media_status(check_writers=False))
        self.volume.textEdited.connect(self._normalize_volume_text)
        self.write_speed.currentIndexChanged.connect(self._on_speed_changed)
        self._on_action_mode_changed(self.action_burn_disc.isChecked())

    def _setup_filesystem_actions(self, menu):
        self.fs_group = QActionGroup(self)
        for text, mask in self._fs_options:
            act = QAction(text, self, checkable=True)
            act.setData(mask)
            self.fs_group.addAction(act)
            menu.addAction(act)
            if mask == self._fs_mask:
                act.setChecked(True)
        if not self.fs_group.checkedAction() and menu.actions():
            menu.actions()[0].setChecked(True)
            self._fs_mask = self._fs_options[0][1]
        self.fs_group.triggered.connect(self._on_fs_selected)

    def _status_colors(self) -> tuple[str, str]:
        base = self.palette().window().color()
        lightness = base.lightness()
        ready = "#5cb3ff" if lightness < 128 else "#0078d7"
        warn = "#ff7b7b" if lightness < 128 else "#d32f2f"
        return ready, warn

    def _update_mode_hint(self):
        if not hasattr(self, "mode_hint"):
            return
        create_iso = self._is_create_iso_mode()
        use_iso = self._use_iso_input() and not create_iso
        if create_iso:
            text = "Create ISO: selected files/folders will be saved into an ISO."
        elif use_iso:
            text = "Burn existing ISO: files list is ignored; the chosen ISO will be burned."
        else:
            text = "Burn files/folders directly to disc. Add files and insert blank media."
        self.mode_hint.setText(text)

    def _on_fs_selected(self, action: QAction):
        try:
            mask = int(action.data())
        except Exception:
            mask = FS_ISO9660 | FS_JOLIET
        self._fs_mask = mask
        self.settings.setValue("fs_mask", mask)
        self._append_log(f"Filesystem -> {action.text()}")
        self._apply_label_rules()
        self._normalize_volume_text()
        self._update_fs_display()
        self._populate_write_speeds()

    def _set_theme(self, theme: str):
        theme = (theme or "light").lower()
        if theme not in ("light", "dark"):
            theme = "light"
        self._theme = theme
        self.settings.setValue("theme", theme)
        self.settings.sync()
        self._apply_theme(theme)

    def _apply_theme(self, theme: str):
        app = QApplication.instance()
        if not app:
            return
        if theme == "dark":
            pal = app.palette()
            pal.setColor(QPalette.ColorRole.Window, QColor(40, 40, 40))
            pal.setColor(QPalette.ColorRole.WindowText, QColor(230, 230, 230))
            pal.setColor(QPalette.ColorRole.Base, QColor(30, 30, 30))
            pal.setColor(QPalette.ColorRole.AlternateBase, QColor(45, 45, 45))
            pal.setColor(QPalette.ColorRole.ToolTipBase, QColor(60, 60, 60))
            pal.setColor(QPalette.ColorRole.ToolTipText, QColor(230, 230, 230))
            pal.setColor(QPalette.ColorRole.Text, QColor(230, 230, 230))
            pal.setColor(QPalette.ColorRole.Button, QColor(55, 55, 55))
            pal.setColor(QPalette.ColorRole.ButtonText, QColor(230, 230, 230))
            pal.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
            pal.setColor(QPalette.ColorRole.Link, QColor("#5cb3ff"))
            pal.setColor(QPalette.ColorRole.Highlight, QColor("#2f6bff"))
            pal.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
            app.setPalette(pal)
            app.setStyleSheet(
                "QGroupBox { color: #e6e6e6; border: 1px solid #4f4f4f; border-radius: 4px; margin-top: 6px; }"
                "QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 4px; }"
                "QToolTip { color: #e6e6e6; background-color: #3c3c3c; border: 1px solid #6c6c6c; }"
                "QMenuBar { background-color: #2d2d2d; color: #e6e6e6; }"
                "QMenuBar::item:selected { background-color: #3d3d3d; }"
                "QMenu { background-color: #2d2d2d; color: #e6e6e6; border: 1px solid #4f4f4f; }"
                "QMenu::item:selected { background-color: #3d3d3d; }"
                "QPushButton { background-color: #3a3a3a; color: #e6e6e6; border: 1px solid #555; border-radius: 3px; padding: 4px 8px; }"
                "QPushButton:hover { background-color: #4a4a4a; }"
                "QPushButton:pressed { background-color: #2f2f2f; }"
                "QPushButton:disabled { color: #888888; border-color: #444; background-color: #2b2b2b; }"
                "QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QListWidget, QTreeView {"
                " background-color: #2f2f2f; color: #e6e6e6; border: 1px solid #555; }"
                "QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled, QComboBox:disabled, QListWidget:disabled, QTreeView:disabled {"
                " color: #888888; border: 1px solid #444; background-color: #262626; }"
                "QComboBox QAbstractItemView { background-color: #2f2f2f; color: #e6e6e6; selection-background-color: #3d3d3d; }"
                "QAbstractItemView { outline: 0; }"
                "QAbstractItemView::item { outline: 0; border: 0; }"
                "QAbstractItemView::item:selected { background-color: #3d3d3d; color: #e6e6e6; }"
                "QAbstractItemView::item:selected:active { outline: 0; }"
                "QAbstractItemView::item:selected:!active { outline: 0; }"
                "QAbstractItemView::item:hover { background-color: #444444; color: #ffffff; }"
            )
        else:
            # Restore original app palette and clear custom styles
            if self._default_palette:
                app.setPalette(self._default_palette)
            else:
                app.setPalette(app.style().standardPalette())
            app.setStyleSheet("")
        # Refresh status colors based on palette
        self._update_status_label_colors()
        self._apply_progress_info_style()

    def _update_status_label_colors(self):
        # Recompute status label colors after theme change
        if hasattr(self, "status_label"):
            ready_color, warn_color = self._status_colors()
            if self.btn_burn.isEnabled():
                self.status_label.setStyleSheet(f"color: {ready_color}")
            else:
                self.status_label.setStyleSheet(f"color: {warn_color}")

    def _apply_progress_info_style(self):
        if hasattr(self, "progress_info"):
            if self._theme == "dark":
                self.progress_info.setStyleSheet("color: #cfd2d6; font-size: 12px;")
            else:
                self.progress_info.setStyleSheet("color: #4a4a4a; font-size: 12px;")

    def _fs_label(self, mask: int) -> str:
        mapping = [(FS_ISO9660, "ISO9660"), (FS_JOLIET, "Joliet"), (FS_UDF, "UDF")]
        parts = [name for bit, name in mapping if mask & bit]
        return " + ".join(parts) if parts else "ISO9660 + Joliet"

    def _label_rules_for_mask(self, mask: int) -> dict:
        has_iso = bool(mask & FS_ISO9660)
        has_udf = bool(mask & FS_UDF)
        if has_iso:
            return {"max_len": 16, "allow_lower": False, "allow_space": False, "allow_hyphen": False, "desc": "ISO9660: A-Z, 0-9, _"}
        if has_udf:
            return {"max_len": 64, "allow_lower": True, "allow_space": True, "allow_hyphen": True, "desc": "UDF: loose"}
        return {"max_len": 16, "allow_lower": False, "allow_space": False, "allow_hyphen": False, "desc": "ISO9660"}

    def _apply_label_rules(self):
        rules = self._label_rules_for_mask(self._fs_mask)
        self.volume.setMaxLength(rules["max_len"])
        self.volume.setPlaceholderText(f"Max {rules['max_len']} chars")
        self._normalize_volume_text()
        self._update_fs_display()

    def _normalize_volume_text(self):
        rules = self._label_rules_for_mask(self._fs_mask)
        orig = self.volume.text()
        normalized = sanitize_volume_label(orig, **{k: v for k, v in rules.items() if k != "desc"}, default_label="")
        if normalized != orig:
            self.volume.blockSignals(True)
            self.volume.setText(normalized)
            self.volume.blockSignals(False)

    def _on_speed_changed(self, _index: int):
        text = self.write_speed.currentText()
        self._append_log(f"Write speed -> {text}")

    def _populate_write_speeds(self):
        self.write_speed.blockSignals(True)
        self.write_speed.clear()
        self.write_speed.addItem("Max (auto)", None)
        uid = self.drive.currentData()
        if uid:
            cached = self._speed_cache.get(uid)
            if cached is not None:
                for label, val in cached:
                    self.write_speed.addItem(label, val)
                self.write_speed.blockSignals(False)
                return
            if self._speed_worker and self._speed_worker.isRunning():
                self.write_speed.blockSignals(False)
                return
            worker = SpeedLookupWorker(uid, self)
            self._speed_worker = worker
            worker.result.connect(self._on_speed_result)
            worker.error.connect(self._on_speed_error)
            worker.finished.connect(lambda: setattr(self, "_speed_worker", None))
            worker.start()
        self.write_speed.blockSignals(False)

    def _on_speed_result(self, uid: str, entries: list):
        # 드라이브가 바뀌었으면 결과 무시
        if uid != self.drive.currentData():
            return
        self._speed_cache[uid] = entries
        self.write_speed.blockSignals(True)
        self.write_speed.clear()
        self.write_speed.addItem("Max (auto)", None)
        for label, val in entries:
            self.write_speed.addItem(label, val)
        self.write_speed.blockSignals(False)

    def _on_speed_error(self, uid: str, _err):
        if uid != self.drive.currentData():
            return
        # 실패하면 캐시 없이 기본 상태 유지
        self._speed_cache.pop(uid, None)

    def _append_log(self, msg: str):
        self.log.append(msg)
        cursor = self.log.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log.setTextCursor(cursor)

    def _clear_log(self):
        if hasattr(self, "log"):
            self.log.clear()

    def _toggle_log_visibility(self, visible: bool):
        if hasattr(self, "log_container"):
            self.log_container.setVisible(visible)
        if hasattr(self, "btn_hide_log"):
            self.btn_hide_log.setVisible(visible)
        if hasattr(self, "act_log") and self.act_log.isChecked() != visible:
            self.act_log.blockSignals(True)
            self.act_log.setChecked(visible)
            self.act_log.blockSignals(False)

    def _is_create_iso_mode(self) -> bool:
        return hasattr(self, "action_create_iso") and self.action_create_iso.isChecked()

    def _use_iso_input(self) -> bool:
        return hasattr(self, "chk_use_iso_input") and self.chk_use_iso_input.isChecked()

    def _on_action_mode_changed(self, checked: bool):
        create_iso = self._is_create_iso_mode()
        ui_enabled = self.btn_refresh.isEnabled() if hasattr(self, "btn_refresh") else True
        # Burn-to-disc ISO input controls only relevant when burning to disc
        self.chk_use_iso_input.setEnabled(ui_enabled and not create_iso)
        use_iso = self._use_iso_input() and not create_iso
        self.iso_path_edit.setEnabled(ui_enabled and use_iso)
        self.btn_browse_iso.setEnabled(ui_enabled and use_iso)
        self.btn_clear_iso.setEnabled(ui_enabled and use_iso and self._iso_path is not None)
        # ISO output only for create ISO mode
        self.iso_out_path_edit.setEnabled(ui_enabled and create_iso)
        self.btn_browse_iso_out.setEnabled(ui_enabled and create_iso)
        self._update_list_buttons_and_burn_state()
        self._update_mode_hint()

    def _on_use_iso_toggled(self, checked: bool):
        self._on_action_mode_changed(True)

    def _select_iso_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select ISO file", str(self.last_dir_iso_in), "ISO files (*.iso);;All files (*)")
        if path:
            self._iso_path = path
            self.iso_path_edit.setText(path)
            try:
                self._iso_size = os.path.getsize(path)
            except OSError:
                self._iso_size = None
            self._set_last_dir_iso_in(Path(path).parent)
        self._update_total_size_label()
        self._update_burn_enabled()
        self._update_mode_hint()

    def _clear_iso_file(self):
        self._iso_path = None
        self._iso_size = None
        self.iso_path_edit.clear()
        self._update_total_size_label()
        self._update_burn_enabled()

    def _select_iso_output(self):
        dlg = QFileDialog(self, "Save ISO file", str(self.last_dir_iso_out))
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilter("ISO files (*.iso)")
        dlg.setDefaultSuffix("iso")
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.selectFile("output.iso")
        icon_provider = CustomIconProvider()
        for model in dlg.findChildren(QFileSystemModel):
            model.setIconProvider(icon_provider)
        if dlg.exec():
            path = dlg.selectedFiles()[0]
            if path and not path.lower().endswith(".iso"):
                path = f"{path}.iso"
            if path:
                self._iso_out_path = path
                self.iso_out_path_edit.setText(path)
                # 새 경로를 선택하면 처음 한번은 확인 없이 진행하도록 사용 횟수 기록을 초기화
                self._iso_out_seen = set()
                self._set_last_dir_iso_out(Path(path).parent)

    def _confirm_overwrite(self, path: str) -> bool:
        ret = QMessageBox.question(
            self,
            "Overwrite",
            f"File already exists:\n{path}\nOverwrite it?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return ret == QMessageBox.Yes

    def show_about(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("About")
        dialog.setWindowIcon(self._app_icon)
        dialog.resize(250, 150)
        layout = QVBoxLayout(dialog)

        header_row = QHBoxLayout()
        header_row.setSpacing(16)
        icon_label = QLabel()
        icon_label.setFixedSize(52, 52)
        if not self._app_icon.isNull():
            icon_label.setPixmap(self._app_icon.pixmap(48, 48))
        header_row.addWidget(icon_label, alignment=Qt.AlignTop)

        text_col = QVBoxLayout()
        text_col.setContentsMargins(12, 0, 0, 0)
        label = QLabel(
            "<p align='left'>PySide CD Burner</p>"
            "<p align='left'>Updated: 2025.12</p>"
            "<p align='left'>Author: KHLEE</p>"
            "<p align='left'>Simple, fast ISO creation and disc burning tool.</p>"
        )
        text_col.addWidget(label)
        text_col.addStretch(1)
        header_row.addLayout(text_col, 1)
        layout.addLayout(header_row)
        layout.addStretch(1)

        btn = QPushButton("OK")
        btn.clicked.connect(dialog.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_row.addWidget(btn)
        layout.addLayout(btn_row)
        dialog.exec()

    def _update_fs_display(self):
        if hasattr(self, "fs_display"):
            self.fs_display.setText(self._fs_label(self._fs_mask))

    def _set_button_icons(self):
        self.btn_add_files.setIcon(QIcon.fromTheme("document-open"))
        self.btn_remove.setIcon(QIcon.fromTheme("list-remove"))
        self.btn_clear.setIcon(QIcon.fromTheme("edit-clear"))
        self._icon_burn_default = QIcon.fromTheme("media-record")
        self._icon_burn_ready = self._make_red_record_icon()
        self.btn_burn.setIcon(self._icon_burn_default)
        self.btn_stop.setIcon(QIcon.fromTheme("process-stop"))
        self.btn_refresh.setIcon(QIcon.fromTheme("view-refresh"))
        # 기존 패딩/스타일을 제거해 기본 정렬을 유지한다.

    def _make_red_record_icon(self, size: int = 24) -> QIcon:
        pix = QPixmap(size, size)
        pix.fill(Qt.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor("#2979ff"))  # bright blue fill
        p.setPen(QPen(QColor("#0d47a1"), 1.2))  # darker edge for crispness
        p.drawEllipse(3, 3, size - 6, size - 6)
        p.end()
        return QIcon(pix)

    def add_files(self):
        dlg = FileFolderDialog(self, "Select files", str(self.last_dir_add))
        dlg.setFileMode(QFileDialog.ExistingFiles)
        dlg.setOptions(dlg.options() | QFileDialog.DontUseNativeDialog | QFileDialog.ReadOnly)
        dlg.setOption(QFileDialog.ShowDirsOnly, False)
        dlg.setFilter(QDir.AllDirs | QDir.Files | QDir.NoDotAndDotDot | QDir.Drives)
        
        icon_provider = CustomIconProvider()
        for model in dlg.findChildren(QFileSystemModel):
            model.setIconProvider(icon_provider)

        selection_views = []
        for cls in (QListView, QTreeView):
            for view in dlg.findChildren(cls):
                view.setSelectionMode(QAbstractItemView.ExtendedSelection)
                selection_views.append(view)

        def _first_visible_view():
            for v in selection_views:
                if v.isVisible():
                    return v
            return selection_views[0] if selection_views else None

        initial_view = _first_visible_view()
        if initial_view:
            initial_view.setFocus(Qt.OtherFocusReason)

        shortcut = QShortcut(QKeySequence.SelectAll, dlg)
        shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        def _select_all():
            view = _first_visible_view()
            if view:
                view.setFocus(Qt.ShortcutFocusReason)
                view.selectAll()
        shortcut.activated.connect(_select_all)
        dlg._select_all_shortcut = shortcut  # keep alive
        
        force_dialog_accept_label(dlg, "Add")

        if dlg.exec():
            paths = dlg.selected_paths()
            for p in paths:
                self._add_path(p)
        self._update_list_buttons_and_burn_state()
        try:
            current_dir = Path(dlg.directory().absolutePath())
            if current_dir.exists():
                self.last_dir_add = current_dir
        except Exception:
            pass

    def _on_drop_add(self, path: str):
        self._add_path(path)
        self._update_list_buttons_and_burn_state()

    def _add_path(self, p: str):
        p = os.path.abspath(p)
        for i in range(self.list.count()):
            if self.list.item(i).text() == p:
                return
        self._path_sizes[p] = 0
        self._pending_size.add(p)
        self.list.addItem(QListWidgetItem(p))
        self._start_size_worker(p)

    def remove_selected(self):
        for it in self.list.selectedItems():
            self.list.takeItem(self.list.row(it))
            path = it.text()
            self._pending_size.discard(path)
            try:
                self._size_queue.remove(path)
            except ValueError:
                pass
            size = self._path_sizes.pop(path, 0)
            self._total_size = max(0, self._total_size - size)
        self._update_list_buttons_and_burn_state()

    def _clear_list(self):
        self.list.clear()
        self._path_sizes.clear()
        self._pending_size.clear()
        self._size_queue.clear()
        self._total_size = 0
        self._update_list_buttons_and_burn_state()

    def _start_size_worker(self, path: str):
        self._size_queue.append(path)
        self._process_next_size_worker()

    def _process_next_size_worker(self):
        if self._current_size_worker is not None or not self._size_queue:
            return
        
        path = self._size_queue.popleft()
        self._current_size_worker = SizeWorker(path, self._compute_path_size, self)
        self._current_size_worker.result.connect(self._on_size_computed)
        self._current_size_worker.finished.connect(self._on_size_worker_finished)
        self._current_size_worker.start()

    def _on_size_worker_finished(self):
        self._current_size_worker = None
        self._process_next_size_worker()

    def _on_size_computed(self, path: str, size: int):
        if path not in self._path_sizes:
            self._pending_size.discard(path)
            return
        prev = self._path_sizes.get(path, 0)
        self._path_sizes[path] = size
        self._total_size += max(0, size) - prev
        self._pending_size.discard(path)
        self._update_total_size_label()
        self._update_burn_enabled()

    def _compute_path_size(self, path: str) -> int:
        try:
            if os.path.isfile(path):
                return os.path.getsize(path)
            if os.path.isdir(path):
                total = 0
                for root, _, files in os.walk(path):
                    for f in files:
                        try:
                            total += os.path.getsize(os.path.join(root, f))
                        except OSError:
                            pass
                return total
        except OSError:
            pass
        return 0

    def _refresh_writers_if_changed(self, auto_select: bool = False):
        try:
            writers = list_imapi_writers()
        except Exception as e:
            self._append_log(f"Writer lookup failed: {e}")
            return None

        new_uids = {w["uid"] for w in writers}
        if new_uids != self._known_writers:
            self._known_writers = new_uids
            self._set_drive_list(writers, auto_select=auto_select)
        return writers

    def refresh_writers(self, log: bool = True):
        if self._writer_worker and self._writer_worker.isRunning():
            return
        if log:
            self._append_log("Refreshing writer list...")
        self.btn_refresh.setEnabled(False)
        self._writer_worker = WriterLookupWorker()
        self._writer_worker.result.connect(self._on_writers_result)
        self._writer_worker.finished.connect(self._on_writer_worker_finished)
        self._writer_worker.start()

    def _on_writer_worker_finished(self):
        self.btn_refresh.setEnabled(True)
        self._writer_worker = None

    def _on_writers_result(self, writers, err):
        if err:
            self._append_log(f"Writer lookup failed: {err}")
            return
        writers = writers or []
        new_uids = {w["uid"] for w in writers}
        if new_uids != self._known_writers:
            self._known_writers = new_uids
            self._set_drive_list(writers, auto_select=True)
        self.update_media_status(check_writers=False)

    def _set_drive_list(self, writers: list[dict], auto_select: bool = False):
        current_uid = self.drive.currentData()
        self.drive.blockSignals(True)
        self.drive.clear()
        # 드라이브 목록이 바뀌면 속도 캐시를 초기화한다.
        self._speed_cache.clear()
        for w in writers:
            self.drive.addItem(w["display"], w["uid"])
        self.drive.blockSignals(False)
        if current_uid:
            idx = self.drive.findData(current_uid)
            if idx >= 0:
                self.drive.setCurrentIndex(idx)
                return
        if auto_select and writers and self.drive.currentIndex() == -1:
            self.drive.setCurrentIndex(0)
        self._populate_write_speeds()

    def start_burn(self):
        if self.worker and self.worker.isRunning():
            return
        create_iso_mode = self._is_create_iso_mode()
        use_iso_input = self._use_iso_input()

        # Gather inputs
        paths = [self.list.item(i).text() for i in range(self.list.count())]
        iso_path = self._iso_path if use_iso_input else None
        uid = self.drive.currentData()

        if create_iso_mode:
            if not paths or self._pending_size:
                return
            out_path = self.iso_out_path_edit.text().strip()
            if not out_path:
                QMessageBox.warning(self, "ISO", "Please choose an ISO output path first.")
                return
            if not Path(out_path).exists():
                self._iso_out_seen.add(out_path)  # mark so next burn prompts
            else:
                if out_path in self._iso_out_seen:
                    if QMessageBox.question(
                        self,
                        "Overwrite",
                        f"File already exists:\n{out_path}\nOverwrite it?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No,
                    ) != QMessageBox.Yes:
                        self._set_ui_enabled(True)
                        self._burning = False
                        self._active_job = None
                        self._update_burn_enabled()
                        return
                else:
                    self._iso_out_seen.add(out_path)
        else:
            if use_iso_input:
                if not iso_path or not os.path.isfile(iso_path):
                    return
            else:
                if not paths or self._pending_size:
                    return
            if not uid:
                return

        self.log.clear()
        self.progress.setValue(0)
        self._reset_progress_info_label()
        self.status.setText("Preparing...")
        self._set_ui_enabled(False)
        self._burn_started_at = time.perf_counter()
        self._burning = True
        if hasattr(self, "elapsed_timer"):
            self.elapsed_timer.start()
        self._update_burn_enabled()

        if create_iso_mode:
            self._active_job = "iso"
            self.worker = IsoCreateWorker(
                self.volume.text(),
                paths,
                self._fs_mask,
                out_path,
                verify=self.chk_verify.isChecked(),
                parent=self,
            )
            self.worker.log.connect(self._append_log)
            self.worker.progress.connect(self.progress.setValue)
            self.worker.status.connect(self._on_status_change)
            self.worker.progress_info.connect(self._on_progress_info)
            self.worker.done.connect(self._on_iso_done)
            self.worker.start()
        else:
            self.worker = BurnWorker(
                uid,
                self.volume.text(),
                paths if not use_iso_input else [],
                self._fs_mask,
                self.write_speed.currentData(),
                iso_path=iso_path,
                verify=self.chk_verify.isChecked(),
                parent=self,
            )
            self._active_job = "burn"
            self.worker.log.connect(self._append_log)
            self.worker.progress.connect(self.progress.setValue)
            self.worker.status.connect(self._on_status_change)
            self.worker.progress_info.connect(self._on_progress_info)
            self.worker.done.connect(self._on_done)
            self.worker.start()

    def stop_burn(self):
        if self.worker and self.worker.isRunning():
            self.worker.request_stop()
            self.status.setText("Stopping...")
            self.btn_stop.setEnabled(False)

    def _on_done(self, ok: bool, msg: str):
        self._active_job = None
        self._burning = False
        self._set_ui_enabled(True)
        self.worker = None
        if self._burn_started_at is not None:
            elapsed = time.perf_counter() - self._burn_started_at
            self._append_log(f"Burn time: {self._format_duration(elapsed)}")
            self._burn_started_at = None
        self._stop_elapsed_timer()
        if self._media_usage_dirty:
            self._update_media_usage_label()
        if ok:
            self.status.setText("Completed")
            self.progress.setValue(100)
            QApplication.beep()
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Question)
            box.setWindowTitle("Done")
            box.setText("Eject disc?")
            box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            box.setDefaultButton(QMessageBox.Yes)
            box.setMinimumWidth(260)  # target a narrower width
            spacer = QSpacerItem(240, 0, QSizePolicy.MinimumExpanding, QSizePolicy.Minimum)
            layout = box.layout()
            layout.addItem(spacer, layout.rowCount(), 0, 1, layout.columnCount())
            if box.exec() == QMessageBox.Yes:
                self.eject_disc()
        else:
            self.status.setText("Failed" if msg != "Stopped by user" else "Stopped")
            if msg != "Stopped by user":
                QMessageBox.critical(self, "Failed", msg)
        # Reset to idle state after showing completion/failure
        self.progress.setValue(0)
        self.status.setText("Idle")
        self._reset_progress_info_label()
        self._update_burn_enabled()

    def eject_disc(self):
        uid = self.drive.currentData()
        if uid:
            coinit = False
            pythoncom = None
            try:
                import pythoncom as _pythoncom
                pythoncom = _pythoncom
                pythoncom.CoInitialize()
                coinit = True
                import win32com.client
                rec = win32com.client.Dispatch("IMAPI2.MsftDiscRecorder2")
                rec.InitializeDiscRecorder(uid)
                rec.EjectMedia()
            except Exception as e:
                self._append_log(f"Eject failed: {e}")
            finally:
                if coinit and pythoncom:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

    def _set_ui_enabled(self, enabled: bool):
        create_iso = self._is_create_iso_mode()
        use_iso_input = self._use_iso_input()
        self.btn_add_files.setEnabled(enabled)
        self.btn_remove.setEnabled(enabled and len(self.list.selectedItems()) > 0)
        self.btn_clear.setEnabled(enabled and self.list.count() > 0)
        self.chk_use_iso_input.setEnabled(enabled and not create_iso)
        self.btn_browse_iso.setEnabled(enabled and use_iso_input and not create_iso)
        self.btn_clear_iso.setEnabled(enabled and use_iso_input and not create_iso and self._iso_path is not None)
        self.iso_path_edit.setEnabled(enabled and use_iso_input and not create_iso)
        self.iso_out_path_edit.setEnabled(enabled and create_iso)
        self.btn_browse_iso_out.setEnabled(enabled and create_iso)
        self.btn_refresh.setEnabled(enabled)
        self.drive.setEnabled(enabled and not create_iso)
        self.volume.setEnabled(enabled)
        self.write_speed.setEnabled(enabled and not create_iso)
        if hasattr(self, "action_burn_disc"):
            self.action_burn_disc.setEnabled(enabled)
        if hasattr(self, "action_create_iso"):
            self.action_create_iso.setEnabled(enabled)
        # Disable verify toggle while a burn/ISO job is running
        self.chk_verify.setEnabled(enabled)
        if hasattr(self, "act_eject"):
            self.act_eject.setEnabled(enabled and not self._burning)
        if enabled:
            self.btn_stop.setEnabled(False)
        else:
            self.btn_burn.setEnabled(False)
            self.btn_stop.setEnabled(True)
        self._update_burn_enabled()

    def _set_media_status(self, text: str):
        # 굽기/ISO 작업 중에는 상태 표시/로그를 건드리지 않아 혼선을 막는다.
        if self._burning:
            return
        if text != self._last_media_text:
            self._append_log(f"Media status -> {text}")
            self._last_media_text = text
        self.media_status.setText(text)

    def update_media_status(self, check_writers: bool = True):
        if check_writers:
            self._refresh_writers_if_changed(auto_select=True)
        uid = self.drive.currentData()
        # 드라이브가 바뀌면 이전 상태 비교를 초기화한다.
        prev_uid = self._last_media_state[0] if self._last_media_state else None
        if prev_uid and prev_uid != uid:
            self._last_media_state = None
            self._speed_cache.pop(prev_uid, None)
        if not uid:
            self._set_media_status("Media: N/A")
            self._media_blank = False
            self._media_capacity_bytes = None
            self._last_media_state = None
            self._update_media_usage_label()
            self._update_burn_enabled()
            return
        if self._media_worker and self._media_worker.isRunning():
            return
        worker = MediaStatusWorker(uid, self)
        self._media_worker = worker
        worker.result.connect(self._on_media_status_result)
        worker.error.connect(self._on_media_status_error)
        worker.finished.connect(lambda: setattr(self, "_media_worker", None))
        worker.start()

    def _on_media_status_result(self, uid: str, blank: bool, supported: bool, capacity: object):
        # 결과가 오래된 드라이브에 대한 것이라면 무시
        if uid != self.drive.currentData():
            return
        self._media_capacity_bytes = capacity if isinstance(capacity, int) else None
        status = "Blank disc" if blank else ("Has data" if supported else "No disc/Unsupported")
        self._set_media_status(f"Media status: {status}")
        self._media_blank = bool(blank and supported)
        state = (uid, bool(blank), bool(supported), self._media_capacity_bytes)
        if state != self._last_media_state:
            self._speed_cache.pop(uid, None)
            self._last_media_state = state
        self._update_media_usage_label()
        self._populate_write_speeds()
        self._update_burn_enabled()

    def _on_media_status_error(self, err):
        self._set_media_status("Media: No disc")
        self._media_blank = False
        self._media_capacity_bytes = None
        self._speed_cache.pop(self.drive.currentData(), None)
        self._update_media_usage_label()
        self._update_burn_enabled()

    def _estimate_image_size(self, size: int) -> int:
        """Estimate on-disc size including ISO9660/Joliet overhead for burn preview."""
        if size <= 0:
            return 0
        # No extra overhead for existing ISO input or ISO creation mode.
        if self._use_iso_input() and not self._is_create_iso_mode():
            return size
        if self._is_create_iso_mode():
            return size
        overhead = max(128 * 1024 * 1024, int(size * 0.07))  # min 128MB or 7%
        return size + overhead

    def _update_media_usage_label(self):
        if self._burning:
            # 굽기 중에는 표시를 유지하고 완료 후에 갱신하도록 표시만 남긴다.
            self._media_usage_dirty = True
            return
        use_iso = self._use_iso_input() and not self._is_create_iso_mode()
        if self._pending_size and not use_iso:
            self.media_usage_label.setText("Media usage: calculating...")
            self.media_usage_label.setStyleSheet("")
            return
        if use_iso:
            if self._iso_path and self._iso_size is None:
                self.media_usage_label.setText("Media usage: calculating...")
                self.media_usage_label.setStyleSheet("")
                return
            current_size = self._iso_size or 0
        elif self._is_create_iso_mode():
            current_size = self._total_size
        else:
            current_size = self._total_size
        cap = self._media_capacity_bytes
        if cap and not self._is_create_iso_mode():
            est_size = self._estimate_image_size(current_size)
            headroom = max(self._capacity_headroom_bytes, int(cap * 0.0025))
            usable = max(0, cap - headroom)
            pct = (est_size / usable) * 100 if usable else 0.0
            over = self._is_over_capacity(est_size)
            self.media_usage_label.setText(
                f"Media: {self._format_size(est_size)} of {self._format_size(usable)} usable "
                f"({pct:.1f}%, total {self._format_size(cap)})"
            )
            _, warn_color = self._status_colors()
            self.media_usage_label.setStyleSheet(f"color: {warn_color};" if over else "")
        else:
            self.media_usage_label.setText(f"Media usage: {self._format_size(current_size)} / {'unknown' if not cap else self._format_size(cap)}")
            self.media_usage_label.setStyleSheet("")
        self._media_usage_dirty = False

    def _format_size(self, size: int) -> str:
        for u in ["B", "KB", "MB", "GB"]:
            if size < 1024: return f"{size:.1f} {u}"
            size /= 1024
        return f"{size:.1f} TB"

    def _format_duration(self, seconds: float) -> str:
        if seconds < 1:
            return f"{seconds*1000:.0f} ms"
        m, s = divmod(seconds, 60)
        h, m = divmod(int(m), 60)
        if h:
            return f"{h}h {m:02d}m {s:04.1f}s"
        return f"{m}m {s:04.1f}s"

    def _format_rate(self, bytes_per_sec: float) -> str:
        val = float(bytes_per_sec)
        units = ["B/s", "KB/s", "MB/s", "GB/s", "TB/s"]
        for u in units:
            if val < 1024 or u == units[-1]:
                return f"{val:.1f} {u}"
            val /= 1024

    def _is_over_capacity(self, size: int) -> bool:
        if not self._media_capacity_bytes:
            return False
        headroom = max(self._capacity_headroom_bytes, int(self._media_capacity_bytes * 0.0025))
        return size > max(0, self._media_capacity_bytes - headroom)

    def _reset_progress_info_label(self):
        if hasattr(self, "progress_info"):
            self.progress_info.setText("Elapsed: --")
        self._stop_elapsed_timer()
        if hasattr(self, "elapsed_timer"):
            self.elapsed_timer.stop()

    def _on_progress_info(self, bytes_per_sec: float, eta_seconds):
        if not hasattr(self, "progress_info"):
            return
        self._update_elapsed_label()

    def _update_elapsed_label(self):
        if not hasattr(self, "progress_info"):
            return
        if self._burn_started_at is None:
            self.progress_info.setText("Elapsed: --")
            self._stop_elapsed_timer()
            return
        elapsed = max(0.0, time.perf_counter() - self._burn_started_at)
        self.progress_info.setText(f"Elapsed: {self._format_duration(elapsed)}")

    def _on_status_change(self, text: str):
        self.status.setText(text)
        self._ensure_elapsed_timer_running()

    def _ensure_elapsed_timer_running(self):
        if hasattr(self, "elapsed_timer") and self._burn_started_at is not None and not self.elapsed_timer.isActive():
            self.elapsed_timer.start()

    def _stop_elapsed_timer(self):
        if hasattr(self, "elapsed_timer"):
            self.elapsed_timer.stop()

    def _update_total_size_label(self):
        if self._use_iso_input() and not self._is_create_iso_mode():
            text = "Total size: --"
            if self._iso_path:
                if self._iso_size is None:
                    text = "Total size: calculating..."
                else:
                    text = f"Total size: {self._format_size(self._iso_size)}"
            self.total_size_label.setText(text)
        else:
            if self._pending_size:
                self.total_size_label.setText("Total size: calculating...")
            else:
                self.total_size_label.setText(f"Total size: {self._format_size(self._total_size)}")
        self._update_media_usage_label()

    def _update_list_buttons_and_burn_state(self):
        self.btn_clear.setEnabled(self.list.count() > 0)
        # disable remove while burning; enable only when UI unlocked
        self.btn_remove.setEnabled(self.btn_add_files.isEnabled() and len(self.list.selectedItems()) > 0)
        self.btn_clear_iso.setEnabled(self.btn_browse_iso.isEnabled() and self._iso_path is not None)
        self._update_total_size_label()
        self._update_burn_enabled()

    def _update_burn_enabled(self):
        if self._burning:
            self.btn_burn.setEnabled(False)
            if hasattr(self, "_icon_burn_ready"):
                self.btn_burn.setIcon(self._icon_burn_default)
            status_text = "Processing..." if self._active_job == "iso" else "Burning..."
            if status_text != self._last_status_text:
                self._append_log(f"Status -> {status_text}")
                self._last_status_text = status_text
            self.status_label.setText(status_text)
            self.status_label.setStyleSheet("color: #0078d7")
            return
        reasons = []
        create_iso = self._is_create_iso_mode()
        use_iso_input = self._use_iso_input() and not create_iso
        if use_iso_input:
            if not self._iso_path or not os.path.isfile(self._iso_path):
                reasons.append("select ISO")
                self._iso_size = None
                current_size = 0
            else:
                try:
                    self._iso_size = os.path.getsize(self._iso_path)
                except OSError:
                    self._iso_size = None
                if self._iso_size is None:
                    reasons.append("calculating sizes")
                current_size = self._iso_size or 0
        else:
            if self.list.count() == 0:
                reasons.append("add files/folders")
            if self._pending_size:
                reasons.append("calculating sizes")
            current_size = self._total_size

        if create_iso:
            if not self.iso_out_path_edit.text().strip():
                reasons.append("select ISO output")
        else:
            if not self.drive.currentData():
                reasons.append("select a drive")
            if not self._media_blank:
                reasons.append("no blank/unsupported disc")
            if (
                self._media_capacity_bytes
                and self._is_over_capacity(self._estimate_image_size(current_size))
            ):
                reasons.append("over capacity")

        ready = len(reasons) == 0
        self.btn_burn.setEnabled(ready)
        if hasattr(self, "_icon_burn_ready"):
            self.btn_burn.setIcon(self._icon_burn_ready if ready else self._icon_burn_default)
        status_text = "Ready" if ready else f"Not Ready: {', '.join(reasons)}"
        if status_text != self._last_status_text:
            self._append_log(f"Status -> {status_text}")
            self._last_status_text = status_text
        self.status_label.setText(status_text)
        ready_color, warn_color = self._status_colors()
        self.status_label.setStyleSheet(f"color: {ready_color if ready else warn_color}")

    def _update_remove_button_state(self):
        # Only enable remove when UI is not locked for burning
        self.btn_remove.setEnabled(self.btn_add_files.isEnabled() and len(self.list.selectedItems()) > 0)

    def _set_last_dir_add(self, path: Path):
        if path.exists():
            self.last_dir_add = path
            self.settings.setValue("last_dir_add", str(path))
            self.settings.sync()

    def _set_last_dir_iso_in(self, path: Path):
        if path.exists():
            self.last_dir_iso_in = path
            self.settings.setValue("last_dir_iso_in", str(path))
            self.settings.sync()

    def _set_last_dir_iso_out(self, path: Path):
        if path.exists():
            self.last_dir_iso_out = path
            self.settings.setValue("last_dir_iso_out", str(path))
            self.settings.sync()

    def _set_last_dir_from_selection(self, sel: str):
        p = Path(sel)
        target = p if p.is_dir() else p.parent
        self._set_last_dir_add(target)

    def _recalculate_total_size(self):
        self._total_size = sum(max(0, s) for s in self._path_sizes.values())

    def create_iso(self):
        if self.worker and self.worker.isRunning():
            return
        if self.list.count() == 0 or self._pending_size:
            return
        out_path = self.iso_out_path_edit.text().strip()
        if not out_path:
            self._append_log("Select ISO output path first.")
            QMessageBox.warning(self, "ISO", "Please choose an ISO output path first.")
            return
        if not Path(out_path).exists():
            self._iso_out_seen.add(out_path)
        else:
            if out_path in self._iso_out_seen:
                if not self._confirm_overwrite(out_path):
                    return
            else:
                self._iso_out_seen.add(out_path)

        paths = [self.list.item(i).text() for i in range(self.list.count())]
        self.log.clear()
        self.progress.setValue(0)
        self._reset_progress_info_label()
        self.status.setText("Preparing ISO...")
        self._set_ui_enabled(False)
        self._burn_started_at = time.perf_counter()
        self._burning = True
        self._active_job = "iso"
        if hasattr(self, "elapsed_timer"):
            self.elapsed_timer.start()
        self._update_burn_enabled()

        self.worker = IsoCreateWorker(
            self.volume.text(),
            paths,
            self._fs_mask,
            out_path,
            verify=self.chk_verify.isChecked(),
            parent=self,
        )
        self.worker.log.connect(self._append_log)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.status.setText)
        self.worker.progress_info.connect(self._on_progress_info)
        self.worker.done.connect(self._on_iso_done)
        self.worker.start()

    def _on_iso_done(self, ok: bool, msg: str):
        self._active_job = None
        self._burning = False
        self._set_ui_enabled(True)
        self.worker = None
        if self._burn_started_at is not None:
            elapsed = time.perf_counter() - self._burn_started_at
            self._append_log(f"ISO job time: {self._format_duration(elapsed)}")
            self._burn_started_at = None
        self._stop_elapsed_timer()
        if self._media_usage_dirty:
            self._update_media_usage_label()
        if ok:
            self.status.setText("ISO created")
            self.progress.setValue(100)
            box = QMessageBox(self)
            box.setWindowTitle("ISO")
            box.setText("ISO creation completed.")
            open_btn = box.addButton("Open folder", QMessageBox.AcceptRole)
            box.addButton(QMessageBox.Ok)
            box.exec()
            if box.clickedButton() == open_btn:
                self._open_iso_output_folder()
        else:
            self.status.setText("Failed")
            QMessageBox.critical(self, "ISO", "ISO creation failed.")
        self.progress.setValue(0)
        self.status.setText("Idle")
        self._reset_progress_info_label()
        self._update_burn_enabled()

    def _open_iso_output_folder(self):
        path = self.iso_out_path_edit.text().strip() or self._iso_out_path
        if not path:
            return
        try:
            target = Path(path).resolve().parent
        except Exception:
            return
        if target.exists():
            try:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
            except Exception:
                pass

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            # 굽기/ISO 작업 중에는 먼저 중단 후 종료되도록 막는다.
            if not self._pending_close:
                ret = QMessageBox.question(
                    self,
                    "Exit",
                    "A burn is in progress. Stop and exit after it halts?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if ret != QMessageBox.Yes:
                    event.ignore()
                    return
                self._pending_close = True
                self._append_log("Stopping burn before exit...")
                try:
                    self.worker.finished.connect(self.close)
                except Exception:
                    pass
                try:
                    self.worker.request_stop()
                except Exception:
                    pass
                try:
                    self.status.setText("Stopping...")
                except Exception:
                    pass
            event.ignore()
            return
        self.settings.setValue("geometry", self.saveGeometry())
        super().closeEvent(event)
