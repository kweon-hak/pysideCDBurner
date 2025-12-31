try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


def list_imapi_writers():
    """
    Return a list of available IMAPI2 writers.
    Each item: {"display": str, "uid": str, "drives": list[str]}
    """
    if not pythoncom or not win32com:
        raise RuntimeError("pywin32 is required. Install with 'pip install pywin32'.")

    pythoncom.CoInitialize()
    try:
        dm = win32com.client.Dispatch("IMAPI2.MsftDiscMaster2")

        writers = []
        for uid in dm:
            try:
                rec = win32com.client.Dispatch("IMAPI2.MsftDiscRecorder2")
                rec.InitializeDiscRecorder(uid)

                drives = []
                for vp in rec.VolumePathNames:
                    if isinstance(vp, str) and len(vp) >= 2 and vp[1] == ":":
                        drives.append(vp[:2].upper())

                vendor = getattr(rec, "VendorId", "") or ""
                product = getattr(rec, "ProductId", "") or ""
                rev = getattr(rec, "ProductRevision", "") or ""
                name = " ".join([x for x in [vendor.strip(), product.strip(), rev.strip()] if x]).strip()
                drive_str = ", ".join(drives) if drives else "(no letter)"

                display = f"{drive_str}  -  {name or 'IMAPI Recorder'}"
                writers.append({"display": display, "uid": uid, "drives": drives})
            except Exception:
                continue

        return writers
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class DiscFormat2DataEvents:
    """Event sink for IMAPI2.MsftDiscFormat2Data."""

    def __init__(self):
        self._emit_log = None
        self._emit_progress = None
        self._emit_status = None
        self._last_percent = -1
        self._last_action = None
        self._stop_check = None

    def OnUpdate(self, sender, args):
        try:
            if self._stop_check and self._stop_check():
                if self._emit_log:
                    self._emit_log("Cancelling burn...")
                try:
                    sender.CancelWrite()
                except Exception:
                    pass
                return
        except Exception:
            pass
        try:
            action = int(args.CurrentAction)
        except Exception:
            action = -1

        percent = None
        try:
            start = int(args.StartLba)
            last = int(args.LastWrittenLba)
            count = int(args.SectorCount)
            if count > 0:
                percent = int(max(0, min(100, ((last - start) * 100) / count)))
        except Exception:
            percent = None

        action_map = {
            0: "Unknown/Idle", 1: "Validating media", 2: "Formatting media",
            3: "Initializing/hibernating", 4: "Calculating progress", 5: "Writing data",
            6: "Finalizing session", 7: "Completed", 8: "Verifying",
        }
        status = action_map.get(action, f"Action={action}")

        if self._emit_log and action != self._last_action:
            self._emit_log(f"Status: {status}")
            self._last_action = action
        if self._emit_status:
            self._emit_status(status)
        if percent is not None and percent != self._last_percent:
            self._last_percent = percent
            if self._emit_progress:
                try:
                    self._emit_progress(percent, status)
                except TypeError:
                    # Backward compatibility if the callback still expects a single argument
                    self._emit_progress(percent)
