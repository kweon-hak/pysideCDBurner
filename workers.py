import os
import shutil
import tempfile
import time
from collections import deque
from PySide6.QtCore import QThread, Signal

from constants import FS_ISO9660, FS_JOLIET, FS_UDF
from utils import safe_copy_into_staging
from imapi import DiscFormat2DataEvents


class BurnWorker(QThread):
    log = Signal(str)
    progress = Signal(int)   # 0~100
    status = Signal(str)
    done = Signal(bool, str) # (ok, message)
    progress_info = Signal(float, object)  # (bytes_per_sec, eta_seconds or None)

    def __init__(
        self,
        recorder_uid: str,
        volume_label: str,
        source_paths: list[str],
        file_system_mask: int | None = None,
        write_speed: int | None = None,
        iso_path: str | None = None,
        verify: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self.recorder_uid = recorder_uid
        self.volume_label = volume_label
        self.source_paths = source_paths
        self.file_system_mask = file_system_mask or (FS_ISO9660 | FS_JOLIET)
        self.write_speed = write_speed
        self.iso_path = iso_path
        self.verify = verify
        self._stop_requested = False
        self._fmt = None
        self._speed_history = deque(maxlen=8)
        self._max_iso_file_bytes = 4 * 1024**3  # ~4GB single-file limit for ISO9660/Joliet
        self._last_est_size = 0
        self._fsi_limit_msg = "파일 시스템 제한으로 ISO 크기가 허용치를 초과했습니다. UDF로 설정하거나 용량을 줄이세요."

    def request_stop(self):
        self._stop_requested = True

    def _log(self, s: str):
        self.log.emit(s)

    def run(self):
        tmpdir = None
        coinit = False
        pythoncom = None
        self._total_bytes_est = 0
        try:
            import pythoncom as _pythoncom
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            coinit = True
            import win32com.client

            # 1) initialize recorder
            self._log("Initializing recorder...")
            rec = win32com.client.Dispatch("IMAPI2.MsftDiscRecorder2")
            rec.InitializeDiscRecorder(self.recorder_uid)

            # 2) prepare image stream
            if self.iso_path:
                if not os.path.isfile(self.iso_path):
                    raise RuntimeError("ISO file not found")
                self._log("Opening ISO image...")
                try:
                    self._last_est_size = os.path.getsize(self.iso_path)
                except OSError:
                    self._last_est_size = 0
                image_stream = win32com.client.Dispatch("ADODB.Stream")
                image_stream.Type = 1  # binary
                image_stream.Open()
                image_stream.LoadFromFile(self.iso_path)
            else:
                tmpdir = tempfile.mkdtemp(prefix="pyside_imapi_cd_")
                staging_dir = os.path.join(tmpdir, "staging")
                os.makedirs(staging_dir, exist_ok=True)

                self._log("Copying sources into staging...")
                for p in self.source_paths:
                    if self._stop_requested:
                        raise RuntimeError("Stopped by user")
                    safe_copy_into_staging(p, staging_dir, stop_check=lambda: self._stop_requested)
                total_bytes, largest_file = self._compute_staging_size(staging_dir)
                self._last_est_size = total_bytes
                self._log(f"Estimated content size: {self._format_bytes(total_bytes)} (largest file: {self._format_bytes(largest_file)})")

                self._log("Building file system image...")
                fsi = win32com.client.Dispatch("IMAPI2FS.MsftFileSystemImage")
                try:
                    fsi.ChooseImageDefaults(rec)
                except Exception:
                    pass

                mask = self._effective_mask(largest_file)
                fsi.FileSystemsToCreate = int(mask)
                self._configure_size_limits(fsi, total_bytes)
                fsi.VolumeName = self.volume_label
                try:
                    fsi.Root.AddTree(staging_dir, False)
                except Exception as add_err:
                    if "size larger than the current configured limit" in str(add_err):
                        raise RuntimeError(self._fsi_limit_msg)
                    raise

                result = fsi.CreateResultImage()
                image_stream = result.ImageStream

            # 3) burn
            self._log("Starting burn...")
            self._total_bytes_est = max(0, int(self._last_est_size))
            max_speed_bytes = int(self.write_speed) * 1024 if self.write_speed else None
            start_time = time.perf_counter()
            last_progress = 0.0
            last_time = start_time

            def _emit_progress_info(percent: int, status: str | None = None):
                nonlocal last_progress, last_time
                writing = (status and "writing data" in status.lower()) or status is None
                if not writing:
                    self._speed_history.clear()
                    self.progress_info.emit(0.0, None)
                    return
                if self._total_bytes_est <= 0:
                    self.progress_info.emit(0.0, None)
                    return
                now = time.perf_counter()
                pct = max(0.0, min(100.0, float(percent)))
                delta_pct = pct - last_progress
                if delta_pct > 0:
                    bytes_delta = self._total_bytes_est * (delta_pct / 100.0)
                    elapsed = max(now - last_time, 1e-6)
                    inst_speed = bytes_delta / elapsed
                    self._speed_history.append(inst_speed)
                    last_progress = pct
                    last_time = now
                if self._speed_history:
                    speed = sum(self._speed_history) / len(self._speed_history)
                else:
                    elapsed_total = max(now - start_time, 1e-6)
                    done_bytes = (self._total_bytes_est * pct) / 100.0
                    speed = done_bytes / elapsed_total
                if max_speed_bytes and speed > max_speed_bytes:
                    speed = max_speed_bytes
                done_bytes = (self._total_bytes_est * pct) / 100.0
                remaining = max(0.0, self._total_bytes_est - done_bytes)
                eta = remaining / speed if speed > 0 else None
                self.progress_info.emit(speed, eta)
            fmt = win32com.client.DispatchWithEvents(
                "IMAPI2.MsftDiscFormat2Data", DiscFormat2DataEvents
            )
            self._fmt = fmt
            fmt._emit_log = self._log
            fmt._emit_progress = lambda *a, **kw: (
                self.progress.emit(a[0] if a else 0),
                _emit_progress_info(a[0] if a else 0, a[1] if len(a) > 1 else kw.get("st") if kw else None),
            )
            fmt._emit_status = lambda s: self.status.emit(s)
            fmt._stop_check = lambda: self._stop_requested
            fmt.Recorder = rec
            fmt.ClientName = "PySide IMAPI2 Burner"
            fmt.ForceMediaToBeClosed = True
            try:
                fmt.BurnVerificationLevel = 1 if self.verify else 0  # IMAPI_BURN_VERIFICATION_QUICK
            except Exception:
                pass

            if self.write_speed:
                try:
                    fmt.SetWriteSpeed(int(self.write_speed), 0)
                except Exception:
                    pass

            if not fmt.IsCurrentMediaSupported(rec):
                raise RuntimeError("Media not supported for this recorder.")

            if self.verify:
                self._log("Verification enabled (IMAPI quick verify).")
            fmt.Write(image_stream)
            if self.verify:
                self._log("Verification finished.")

            self.progress.emit(100)
            _emit_progress_info(100)
            self.status.emit("Completed")
            self._log("Burn completed")
            self.done.emit(True, "Completed")

        except Exception as e:
            if self._stop_requested:
                self._log("Stopped by user.")
                self.done.emit(False, "Stopped by user")
            else:
                self._log(f"Error: {e}")
                self.done.emit(False, str(e))
        finally:
            self._fmt = None
            if coinit:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

    def _apply_size_hint(self, fsi, staging_dir: str):
        """
        일부 IMAPI 기본값은 이미지 크기를 미디어 기본값으로 제한한다.
        스테이징된 크기 기준으로 FreeMediaBlocks를 넉넉히 올려 제한 오류를 방지한다.
        """
        try:
            block_size = int(getattr(fsi, "BlockSize", 2048) or 2048)
        except Exception:
            block_size = 2048
        total_bytes = 0
        for root, _, files in os.walk(staging_dir):
            for f in files:
                try:
                    total_bytes += os.path.getsize(os.path.join(root, f))
                except OSError:
                    pass
        est_blocks = (total_bytes + block_size - 1) // block_size
        try:
            fsi.FreeMediaBlocks = max(int(getattr(fsi, "FreeMediaBlocks", 0) or 0), int(est_blocks))
        except Exception:
            pass


    def _configure_size_limits(self, fsi, total_bytes: int):
        try:
            fsi.ImageSizeLimit = 0
        except Exception:
            pass
        try:
            block_size = int(getattr(fsi, "BlockSize", 2048) or 2048)
        except Exception:
            block_size = 2048
        est_blocks = (total_bytes + block_size - 1) // block_size
        try:
            current_blocks = int(getattr(fsi, "FreeMediaBlocks", 0) or 0)
            # generous headroom (x2 + 1MB) to avoid premature limit errors
            target_blocks = max(current_blocks, int(est_blocks * 2 + (1024 * 1024) // block_size))
            fsi.FreeMediaBlocks = target_blocks
        except Exception:
            pass

    def _compute_staging_size(self, staging_dir: str) -> tuple[int, int]:
        total_bytes = 0
        largest = 0
        for root, _, files in os.walk(staging_dir):
            for f in files:
                try:
                    sz = os.path.getsize(os.path.join(root, f))
                    total_bytes += sz
                    largest = max(largest, sz)
                except OSError:
                    pass
        return total_bytes, largest

    def _effective_mask(self, largest_file: int) -> int:
        mask = int(self.file_system_mask or (FS_ISO9660 | FS_JOLIET))
        # If a single file is >= ~4GB, ISO/Joliet로는 불가하므로 UDF로 강제 전환한다.
        if largest_file >= self._max_iso_file_bytes:
            iso_bits = mask & (FS_ISO9660 | FS_JOLIET)
            if iso_bits or not (mask & FS_UDF):
                self._log("대용량 파일 감지: ISO/Joliet로는 불가하여 UDF로 전환합니다.")
                try:
                    self.status.emit("큰 파일 감지: UDF로 전환")
                except Exception:
                    pass
                mask = FS_UDF
        return mask

    @staticmethod
    def _format_bytes(n: int) -> str:
        units = ["B", "KB", "MB", "GB", "TB"]
        val = float(n)
        for u in units:
            if val < 1024 or u == units[-1]:
                return f"{val:.2f} {u}"
            val /= 1024


class SizeWorker(QThread):
    result = Signal(str, object)
    def __init__(self, path: str, compute_func, parent=None):
        super().__init__(parent)
        self.path, self._compute_func = path, compute_func
    def run(self):
        self.result.emit(self.path, int(self._compute_func(self.path) or 0))


class IsoCreateWorker(QThread):
    log = Signal(str)
    progress = Signal(int)
    status = Signal(str)
    done = Signal(bool, str)
    progress_info = Signal(float, object)  # (bytes_per_sec, eta_seconds or None)

    def __init__(
        self,
        volume_label: str,
        source_paths: list[str],
        file_system_mask: int | None,
        output_path: str,
        verify: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self.volume_label = volume_label
        self.source_paths = source_paths
        self.file_system_mask = file_system_mask or (FS_ISO9660 | FS_JOLIET)
        self.output_path = output_path
        self.verify = verify
        self._stop_requested = False
        self._max_iso_file_bytes = 4 * 1024**3
        self._last_est_size = 0
        self._fsi_limit_msg = "파일 시스템 제한으로 ISO 크기가 허용치를 초과했습니다. UDF로 설정하거나 용량을 줄이세요."

    def request_stop(self):
        self._stop_requested = True

    def _log(self, msg: str):
        self.log.emit(msg)

    def run(self):
        tmpdir = None
        coinit = False
        pythoncom = None
        try:
            import pythoncom as _pythoncom
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            coinit = True
            import win32com.client

            tmpdir = tempfile.mkdtemp(prefix="pyside_imapi_iso_")
            staging_dir = os.path.join(tmpdir, "staging")
            os.makedirs(staging_dir, exist_ok=True)

            self._log("Copying sources into staging...")
            for p in self.source_paths:
                if self._stop_requested:
                    raise RuntimeError("Stopped by user")
                safe_copy_into_staging(p, staging_dir, stop_check=lambda: self._stop_requested)
            total_bytes, largest_file = self._compute_staging_size(staging_dir)
            self._last_est_size = total_bytes
            self._log(f"Estimated content size: {self._format_bytes(total_bytes)} (largest file: {self._format_bytes(largest_file)})")
            self.progress.emit(10)
            self.status.emit("Building ISO image...")

            fsi = win32com.client.Dispatch("IMAPI2FS.MsftFileSystemImage")
            try:
                fsi.ChooseImageDefaults(None)
            except Exception:
                pass
            mask = self._effective_mask(largest_file)
            fsi.FileSystemsToCreate = int(mask)
            self._configure_size_limits(fsi, total_bytes)
            fsi.VolumeName = self.volume_label
            try:
                fsi.Root.AddTree(staging_dir, False)
            except Exception as add_err:
                if "size larger than the current configured limit" in str(add_err):
                    raise RuntimeError(self._fsi_limit_msg)
                raise

            result = fsi.CreateResultImage()
            image_stream = result.ImageStream
            try:
                block_size = int(getattr(result, "BlockSize", 2048) or 2048)
                total_blocks = int(getattr(result, "TotalBlocks", 0) or 0)
                expected_size = block_size * total_blocks
            except Exception:
                expected_size = 0
            self._total_bytes_est = max(0, int(expected_size))

            self._log("Writing ISO file...")
            written = 0
            try:
                istream = image_stream.QueryInterface(pythoncom.IID_IStream) if pythoncom else None
            except Exception:
                istream = None
            start_time = time.perf_counter()

            def _emit_progress_info():
                if self._total_bytes_est <= 0:
                    self.progress_info.emit(0.0, None)
                    return
                elapsed = max(time.perf_counter() - start_time, 1e-6)
                done_bytes = max(0.0, min(self._total_bytes_est, float(written)))
                speed = done_bytes / elapsed
                remaining = max(0.0, self._total_bytes_est - done_bytes)
                eta = remaining / speed if speed > 0 else None
                self.progress_info.emit(speed, eta)
            with open(self.output_path, "wb") as f:
                chunk_size = 1024 * 1024
                while True:
                    if self._stop_requested:
                        raise RuntimeError("Stopped by user")
                    if istream is not None:
                        data = istream.Read(chunk_size)
                    else:
                        data = image_stream.Read(chunk_size)
                    if not data:
                        break
                    f.write(data)
                    written += len(data)
                    if expected_size:
                        pct = min(99, int((written / expected_size) * 100))
                        self.progress.emit(pct)
                    _emit_progress_info()
            if expected_size and written != expected_size:
                raise RuntimeError("ISO size mismatch")

            if self.verify:
                self._log("Verifying ISO...")
                self.status.emit("Verifying ISO...")
                actual_size = os.path.getsize(self.output_path)
                if expected_size and actual_size != expected_size:
                    raise RuntimeError("Verify failed: size mismatch")
                self._log("Verification completed.")

            self.progress.emit(100)
            self.status.emit("Completed")
            self._log("ISO created")
            self.done.emit(True, "ISO 생성 완료")

        except Exception as e:
            if self._stop_requested:
                self._log("Stopped by user.")
                self.done.emit(False, "Stopped by user")
            else:
                self._log(f"Error: {e}")
                self.done.emit(False, str(e))
        finally:
            if coinit:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

    def _configure_size_limits(self, fsi, total_bytes: int):
        try:
            fsi.ImageSizeLimit = 0
        except Exception:
            pass
        try:
            block_size = int(getattr(fsi, "BlockSize", 2048) or 2048)
        except Exception:
            block_size = 2048
        est_blocks = (total_bytes + block_size - 1) // block_size
        try:
            current_blocks = int(getattr(fsi, "FreeMediaBlocks", 0) or 0)
            # generous headroom (x2 + 1MB) to avoid premature limit errors
            target_blocks = max(current_blocks, int(est_blocks * 2 + (1024 * 1024) // block_size))
            fsi.FreeMediaBlocks = target_blocks
        except Exception:
            pass

    def _compute_staging_size(self, staging_dir: str) -> tuple[int, int]:
        total_bytes = 0
        largest = 0
        for root, _, files in os.walk(staging_dir):
            for f in files:
                try:
                    sz = os.path.getsize(os.path.join(root, f))
                    total_bytes += sz
                    largest = max(largest, sz)
                except OSError:
                    pass
        return total_bytes, largest

    def _effective_mask(self, largest_file: int) -> int:
        mask = int(self.file_system_mask or (FS_ISO9660 | FS_JOLIET))
        # If a single file is >= ~4GB, ISO/Joliet로는 불가하므로 UDF로 강제 전환한다.
        if largest_file >= self._max_iso_file_bytes:
            iso_bits = mask & (FS_ISO9660 | FS_JOLIET)
            if iso_bits or not (mask & FS_UDF):
                self._log("대용량 파일 감지: ISO/Joliet로는 불가하여 UDF로 전환합니다.")
                try:
                    self.status.emit("큰 파일 감지: UDF로 전환")
                except Exception:
                    pass
                mask = FS_UDF
        return mask

    @staticmethod
    def _format_bytes(n: int) -> str:
        units = ["B", "KB", "MB", "GB", "TB"]
        val = float(n)
        for u in units:
            if val < 1024 or u == units[-1]:
                return f"{val:.2f} {u}"
            val /= 1024
