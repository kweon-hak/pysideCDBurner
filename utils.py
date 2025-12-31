import os
import re
import shutil
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QFileDialog, QDialogButtonBox, QListView, QTreeView


def sanitize_volume_label(
    label: str,
    max_len: int = 16,
    allow_lower: bool = False,
    allow_space: bool = False,
    allow_hyphen: bool = False,
    default_label: str | None = "DATA",
) -> str:
    """
    Normalize volume label with simple rules per filesystem selection.
    """
    label = (label or "").strip()
    if not allow_lower:
        label = label.upper()

    allowed = "A-Z0-9_"
    if allow_lower:
        allowed = "A-Za-z0-9_"
    if allow_hyphen:
        allowed += r"\-"
    if allow_space:
        allowed += " "
        label = re.sub(r"\s+", " ", label)

    label = re.sub(fr"[^{allowed}]", "_", label)
    label = re.sub(r"_+", "_", label).strip(" ")
    label = label[:max_len]
    if not label and default_label is not None:
        return default_label[:max_len]
    return label


def _copy_file_chunked(src: str, dst: str, stop_check=None):
    """Copy file with cancellation support."""
    chunk_size = 1024 * 1024  # 1MB
    with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
        while True:
            if stop_check and stop_check():
                raise RuntimeError("Stopped by user")
            buf = fsrc.read(chunk_size)
            if not buf:
                break
            fdst.write(buf)
    shutil.copystat(src, dst)


def _copy_tree_chunked(src: str, dst: str, ignore=None, stop_check=None):
    """Recursive copytree with cancellation support."""
    os.makedirs(dst, exist_ok=True)
    names = os.listdir(src)
    ignored_names = ignore(src, names) if ignore else set()
    
    for name in names:
        if name in ignored_names:
            continue
        src_name = os.path.join(src, name)
        dst_name = os.path.join(dst, name)
        if stop_check and stop_check():
            raise RuntimeError("Stopped by user")
        
        if os.path.isdir(src_name):
            _copy_tree_chunked(src_name, dst_name, ignore, stop_check)
        else:
            _copy_file_chunked(src_name, dst_name, stop_check)


def safe_copy_into_staging(src_path: str, staging_dir: str, stop_check=None) -> None:
    """
    Copy selected file/folder into staging_dir with clash-safe rename.
    """
    src_path = os.path.abspath(src_path)
    base_name = os.path.basename(src_path.rstrip("\\/")) or "ITEM"
    excluded = {".venv", "venv", "__pycache__", ".mypy_cache", ".git"}
    if base_name in excluded:
        return

    def unique_name(name: str) -> str:
        candidate = name
        n = 1
        while os.path.exists(os.path.join(staging_dir, candidate)):
            candidate = f"{name}_{n}"
            n += 1
        return candidate

    os.makedirs(staging_dir, exist_ok=True)

    if os.path.isdir(src_path):
        dst_name = unique_name(base_name)
        dst_path = os.path.join(staging_dir, dst_name)
        ignore = shutil.ignore_patterns(*excluded)
        _copy_tree_chunked(src_path, dst_path, ignore=ignore, stop_check=stop_check)
    else:
        dst_name = unique_name(base_name)
        dst_path = os.path.join(staging_dir, dst_name)
        if base_name in excluded:
            return
        _copy_file_chunked(src_path, dst_path, stop_check=stop_check)


def force_dialog_accept_label(dlg: QFileDialog, label: str) -> None:
    """
    Force the 'Open' button in a QFileDialog to display specific text (e.g., 'Add').
    """
    btn_box = dlg.findChild(QDialogButtonBox)
    if not btn_box:
        return
    add_btn = btn_box.button(QDialogButtonBox.Open)
    if not add_btn:
        return

    def _update_label():
        QTimer.singleShot(0, lambda: add_btn.setText(label))

    _update_label()
    dlg.currentChanged.connect(lambda _: _update_label())
    dlg.directoryEntered.connect(lambda _: _update_label())
    dlg.filesSelected.connect(lambda _: _update_label())

    for cls in (QListView, QTreeView):
        for view in dlg.findChildren(cls):
            sm = view.selectionModel()
            if sm:
                sm.selectionChanged.connect(lambda *_,: _update_label())
