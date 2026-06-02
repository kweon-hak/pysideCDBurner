from PySide6.QtCore import QDir, QUrl, Qt, QEvent, QItemSelection, QItemSelectionModel, QStorageInfo, QTimer, QSize, QPointF
from PySide6.QtGui import QIcon, QKeySequence, QShortcut, QAction, QPixmap, QPainter, QColor, QPen, QPolygonF
from PySide6.QtWidgets import (
    QFileDialog, QDialog, QFileIconProvider, QListView, QTreeView,
    QAbstractItemView, QApplication, QWidget, QLineEdit, QToolButton
)


class FileFolderDialog(QFileDialog):
    """
    Non-native QFileDialog that lets users select files and folders together.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._selected_urls: list[QUrl] = []
        self.setOption(QFileDialog.DontUseNativeDialog, True)
        self.installEventFilter(self)
        self._app = QApplication.instance()
        if self._app:
            self._app.installEventFilter(self)
        self._drive_sidebar_keys: set[str] = set()
        self._drive_refresh_timer = QTimer(self)
        self._drive_refresh_timer.setInterval(1500)
        self._drive_refresh_timer.timeout.connect(self._ensure_drive_sidebar_urls)
        self._view_shortcuts: list[QShortcut] = []
        self._select_all_shortcut = QShortcut(QKeySequence.SelectAll, self)
        self._select_all_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self._select_all_shortcut.activated.connect(self._select_all_in_views)
        self._select_all_action = QAction(self)
        self._select_all_action.setShortcut(QKeySequence.SelectAll)
        self._select_all_action.setShortcutContext(Qt.WindowShortcut)
        self._select_all_action.triggered.connect(self._select_all_in_views)
        self.addAction(self._select_all_action)
        self._ensure_drive_sidebar_urls()
        self._install_view_shortcuts()
        for le in self.findChildren(QLineEdit):
            le.installEventFilter(self)

    def accept(self):
        self._selected_urls = self.selectedUrls()
        QDialog.accept(self)

    def selected_paths(self) -> list[str]:
        paths: list[str] = []
        for url in self._selected_urls:
            p = url.toLocalFile()
            if p and p not in paths:
                paths.append(p)
        return paths

    def _select_all_in_views(self) -> bool:
        targets = []
        primary = self.findChild(QListView, "listView")
        if primary:
            targets.append(primary)
        fallback_tree = self.findChild(QTreeView, "treeView")
        if fallback_tree and fallback_tree not in targets:
            targets.append(fallback_tree)
        # add any other views as backup
        for view in self.findChildren(QAbstractItemView):
            if view not in targets:
                targets.append(view)
        if not targets:
            return False
        for v in targets:
            v.setSelectionMode(QAbstractItemView.ExtendedSelection)
            try:
                v.setSelectionBehavior(QAbstractItemView.SelectRows)
            except Exception:
                pass
        # Always prefer the main file list (primary) even when focus is elsewhere (e.g., left nav list)
        target = primary if primary and primary.isVisible() else next((v for v in targets if v.isVisible()), targets[0])
        target.setFocus(Qt.ShortcutFocusReason)
        sel_model = target.selectionModel()
        target.selectAll()
        # If nothing changed (e.g., SingleSelection quirks), force-select all rows via selection model
        if sel_model and not sel_model.hasSelection():
            model = target.model()
            root = getattr(target, "rootIndex", lambda: None)()
            if model:
                rows = model.rowCount(root)
                if rows > 0:
                    first = model.index(0, 0, root)
                    last = model.index(rows - 1, 0, root)
                    if first.isValid() and last.isValid():
                        sel = QItemSelection(first, last)
                        sel_model.select(sel, QItemSelectionModel.Select | QItemSelectionModel.Rows)
                        sel_model.setCurrentIndex(first, QItemSelectionModel.NoUpdate)
        self._update_filename_edit_from_selection(target)
        return True

    def _install_view_shortcuts(self):
        # Attach Ctrl+A directly to each view so it also works when the view already has focus
        for sc in getattr(self, "_view_shortcuts", []):
            try:
                sc.setParent(None)
            except Exception:
                pass
        self._view_shortcuts = []
        for view in self.findChildren(QAbstractItemView):
            view.setSelectionMode(QAbstractItemView.ExtendedSelection)
            self._attach_selection_listener(view)
            sc = QShortcut(QKeySequence.SelectAll, view)
            sc.setContext(Qt.WidgetShortcut)
            sc.activated.connect(lambda v=view: (v.setFocus(Qt.ShortcutFocusReason), v.selectAll()))
            self._view_shortcuts.append(sc)

    def showEvent(self, event):
        super().showEvent(event)
        self._ensure_drive_sidebar_urls()
        self._drive_refresh_timer.start()
        self._install_view_shortcuts()
        self._apply_theme_tweaks()
        # Ensure late-created line edits also route Ctrl+A through our filter
        for le in self.findChildren(QLineEdit):
            le.installEventFilter(self)

    def _ensure_drive_sidebar_urls(self):
        computer_url = self._computer_sidebar_url()
        drive_urls = [QUrl.fromLocalFile(root_path) for root_path in self._available_drive_roots()]
        drive_keys = {self._sidebar_url_key(url) for url in drive_urls}

        current_urls = list(self.sidebarUrls())
        urls = [computer_url]
        seen = {self._sidebar_url_key(computer_url)}

        current_urls = [url for url in current_urls if self._is_valid_sidebar_url(url)]
        previous_drive_keys = self._drive_sidebar_keys
        current_urls = [url for url in current_urls if self._sidebar_url_key(url) not in previous_drive_keys]
        for url in current_urls:
            key = self._sidebar_url_key(url)
            if key in seen:
                continue
            urls.append(url)
            seen.add(key)

        for url in drive_urls:
            key = self._sidebar_url_key(url)
            if key in seen:
                continue
            urls.append(url)
            seen.add(key)

        if self._same_sidebar_urls(current_urls, urls) and drive_keys == self._drive_sidebar_keys:
            return

        self._drive_sidebar_keys = drive_keys
        self.setSidebarUrls(urls)

    @staticmethod
    def _sidebar_url_key(url: QUrl) -> str:
        local = url.toLocalFile()
        if local:
            return QDir.cleanPath(local).rstrip("\\/").casefold()
        return url.toString().casefold()

    @staticmethod
    def _computer_sidebar_url() -> QUrl:
        return QUrl("file://")

    @classmethod
    def _is_valid_sidebar_url(cls, url: QUrl) -> bool:
        if cls._is_computer_sidebar_url(url):
            return True
        local = url.toLocalFile()
        if not local:
            return False
        path = QDir.cleanPath(local)
        if cls._is_drive_root(path):
            return cls._is_drive_ready(path)
        return True

    @classmethod
    def _is_computer_sidebar_url(cls, url: QUrl) -> bool:
        return cls._sidebar_url_key(url) == cls._sidebar_url_key(cls._computer_sidebar_url())

    @staticmethod
    def _is_drive_root(path: str) -> bool:
        cleaned = QDir.cleanPath(path).rstrip("\\/")
        return len(cleaned) == 2 and cleaned[1] == ":"

    @classmethod
    def _same_sidebar_urls(cls, left: list[QUrl], right: list[QUrl]) -> bool:
        return [cls._sidebar_url_key(url) for url in left] == [cls._sidebar_url_key(url) for url in right]

    @staticmethod
    def _available_drive_roots() -> list[str]:
        roots: list[str] = []

        for storage in QStorageInfo.mountedVolumes():
            try:
                if not storage.isValid() or not storage.isReady():
                    continue
                root = storage.rootPath()
            except Exception:
                continue
            if root and root not in roots:
                roots.append(root)

        for drive in QDir.drives():
            root = drive.absoluteFilePath()
            if not FileFolderDialog._is_drive_ready(root):
                continue
            if root and root not in roots:
                roots.append(root)

        return roots

    @staticmethod
    def _is_drive_ready(root_path: str) -> bool:
        try:
            storage = QStorageInfo(root_path)
            return storage.isValid() and storage.isReady()
        except Exception:
            return False

    def eventFilter(self, obj, event):
        if event.type() in (QEvent.KeyPress, QEvent.ShortcutOverride):
            # Catch Ctrl+A anywhere in the dialog (including filename edits) before Qt handles shortcuts
            try:
                key = event.key()
                mods = int(event.modifiers())
                is_ctrl_a = (key == Qt.Key_A) and (mods & Qt.ControlModifier)
            except Exception:
                is_ctrl_a = False
            if is_ctrl_a or (hasattr(event, "matches") and event.matches(QKeySequence.SelectAll)):
                if self._select_all_in_views():
                    event.accept()
                    return True
        return super().eventFilter(obj, event)

    def _attach_selection_listener(self, view: QAbstractItemView):
        if getattr(view, "_ff_selection_connected", False):
            return
        sel_model = view.selectionModel()
        if not sel_model:
            return
        try:
            sel_model.selectionChanged.connect(lambda sel, desel, v=view: self._update_filename_edit_from_selection(v))
            view._ff_selection_connected = True
        except Exception:
            pass

    def _update_filename_edit_from_selection(self, view: QAbstractItemView):
        sel_model = view.selectionModel()
        if not sel_model:
            return
        indexes = []
        try:
            indexes = sel_model.selectedRows(0)
        except Exception:
            indexes = [idx for idx in sel_model.selectedIndexes() if idx.column() == 0]
        if not indexes:
            return
        model = view.model()
        names = []
        for idx in indexes:
            try:
                if hasattr(model, "fileName"):
                    nm = model.fileName(idx)
                else:
                    nm = None
                if not nm:
                    nm = idx.data(Qt.DisplayRole)
                if nm:
                    names.append(str(nm))
            except Exception:
                try:
                    names.append(str(idx.data()))
                except Exception:
                    continue
        names = [n for n in names if n]
        if not names:
            return
        max_items = 5
        max_chars = 120
        shown = names[:max_items]
        suffix = ""
        if len(names) > max_items:
            suffix = f" ... (+{len(names) - max_items} more)"
        text = " ".join(shown)
        if len(text) + len(suffix) > max_chars:
            text = text[: max(0, max_chars - len(suffix) - 3)] + "..."
        text = f"{text}{suffix}"
        for le in self.findChildren(QLineEdit):
            try:
                le.setText(text)
            except Exception:
                continue

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.SelectAll):
            if self._select_all_in_views():
                event.accept()
                return
        super().keyPressEvent(event)

    def _apply_theme_tweaks(self):
        if self._is_dark_palette():
            self._apply_dark_navigation_icons()

    def _is_dark_palette(self) -> bool:
        try:
            window_color = self.palette().color(self.backgroundRole())
            return window_color.lightness() < 128
        except Exception:
            return False

    def _apply_dark_navigation_icons(self):
        icon_map = {
            "backButton": "left",
            "forwardButton": "right",
            "toParentButton": "up",
        }
        for object_name, direction in icon_map.items():
            button = self.findChild(QToolButton, object_name)
            if not button:
                continue
            button.setIcon(self._make_navigation_icon(direction, button.iconSize()))

    def _make_navigation_icon(self, direction: str, size: QSize) -> QIcon:
        width = max(16, size.width() if size.isValid() else 16)
        height = max(16, size.height() if size.isValid() else 16)

        normal = QPixmap(width, height)
        normal.fill(Qt.transparent)
        disabled = QPixmap(width, height)
        disabled.fill(Qt.transparent)

        self._paint_navigation_arrow(normal, direction, QColor("#f2f4f7"))
        self._paint_navigation_arrow(disabled, direction, QColor("#7c8694"))

        icon = QIcon()
        icon.addPixmap(normal, QIcon.Normal, QIcon.Off)
        icon.addPixmap(disabled, QIcon.Disabled, QIcon.Off)
        return icon

    @staticmethod
    def _paint_navigation_arrow(pixmap: QPixmap, direction: str, color: QColor):
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        pen = QPen(color, 2.2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        painter.setPen(pen)

        w = pixmap.width()
        h = pixmap.height()
        margin = max(3.5, min(w, h) * 0.22)
        mid_x = w / 2.0
        mid_y = h / 2.0

        if direction == "left":
            points = [
                (w - margin, margin),
                (margin, mid_y),
                (w - margin, h - margin),
            ]
        elif direction == "right":
            points = [
                (margin, margin),
                (w - margin, mid_y),
                (margin, h - margin),
            ]
        else:
            points = [
                (margin, h - margin),
                (mid_x, margin),
                (w - margin, h - margin),
            ]

        painter.drawPolyline(QPolygonF([QPointF(x, y) for x, y in points]))
        painter.end()

    def closeEvent(self, event):
        self._drive_refresh_timer.stop()
        try:
            if getattr(self, "_app", None):
                self._app.removeEventFilter(self)
        except Exception:
            pass
        super().closeEvent(event)


class CustomIconProvider(QFileIconProvider):
    """Minimal fallback so home (user) entry shows an icon even if theme lacks one."""

    def __init__(self):
        super().__init__()
        self._home_path = QDir.homePath().rstrip("\\/").lower()
        self._home_icon = QIcon.fromTheme("user-home")
        if self._home_icon.isNull():
            self._home_icon = QIcon.fromTheme("folder")

    def icon(self, info):
        icon = super().icon(info)
        if not icon.isNull():
            return icon
        try:
            path = info.absoluteFilePath().rstrip("\\/").lower()
            if path == self._home_path:
                return self._home_icon
        except Exception:
            pass
        return icon
