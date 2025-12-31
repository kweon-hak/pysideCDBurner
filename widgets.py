from PySide6.QtCore import QDir, QUrl, Qt, QEvent, QItemSelection, QItemSelectionModel
from PySide6.QtGui import QIcon, QKeySequence, QShortcut, QAction
from PySide6.QtWidgets import (
    QFileDialog, QDialog, QFileIconProvider, QListView, QTreeView,
    QAbstractItemView, QApplication, QWidget, QLineEdit
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
        self._view_shortcuts: list[QShortcut] = []
        self._select_all_shortcut = QShortcut(QKeySequence.SelectAll, self)
        self._select_all_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self._select_all_shortcut.activated.connect(self._select_all_in_views)
        self._select_all_action = QAction(self)
        self._select_all_action.setShortcut(QKeySequence.SelectAll)
        self._select_all_action.setShortcutContext(Qt.WindowShortcut)
        self._select_all_action.triggered.connect(self._select_all_in_views)
        self.addAction(self._select_all_action)
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
        self._install_view_shortcuts()
        # Ensure late-created line edits also route Ctrl+A through our filter
        for le in self.findChildren(QLineEdit):
            le.installEventFilter(self)

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

    def closeEvent(self, event):
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
