"""Cursed PPTMaker - desktop window for make_ppt.py.

Run:  pythonw app.py   (or double-click "Cursed PPTMaker.bat")

The UI is ui/index.html drawn by pywebview (Edge WebView2, offline). This file is the
bridge: it lists local Ollama models, runs make_ppt.auto/build on a worker thread and
streams progress back to the page.
"""
from __future__ import annotations

import base64
import contextlib
import ctypes
import io
import json
import logging
import os
import shutil
import subprocess
import threading
import time
import traceback
import urllib.request
from pathlib import Path

import webview
from PIL import Image
from webview.dom import DOMEventHandler

import make_ppt as mp

HERE = Path(__file__).resolve().parent
LAST_MODEL = HERE / ".last_model"   # shared with the console launcher
LAST_OUTDIR = HERE / ".last_outdir"  # chosen "Save to" folder; empty/missing = next to the source
NO_WINDOW = 0x08000000               # CREATE_NO_WINDOW: no console flashes under pythonw
SUPPORTED = (mp.IMAGE_EXT | set(mp.LEGACY) |
             {".pdf", ".docx", ".xlsx", ".xlsm", ".csv", ".pptx", ".txt", ".md"})
FILE_TYPES = ("Documents (" + ";".join(f"*{e}" for e in sorted(SUPPORTED)) + ")", "All files (*.*)")


# ---------------------------------------------------------------- system info

def free_memory_gb() -> float:
    """Free RAM + free VRAM + memory held by models Ollama has loaded (it unloads on demand)."""
    class MemStatus(ctypes.Structure):
        _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                    ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                    ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                    ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                    ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
    st = MemStatus()
    st.dwLength = ctypes.sizeof(MemStatus)
    ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st))
    free = st.ullAvailPhys / 2**30
    try:
        out = subprocess.run(["nvidia-smi", "--query-gpu=memory.free", "--format=csv,noheader,nounits"],
                             capture_output=True, text=True, timeout=5, creationflags=NO_WINDOW)
        free += float(out.stdout.split()[0]) / 1024
    except (OSError, ValueError, IndexError, subprocess.TimeoutExpired):
        pass
    try:
        free += sum(m.get("size", 0) for m in ollama_get("/api/ps").get("models", [])) / 2**30
    except OSError:
        pass
    return free


def ollama_get(path: str) -> dict:
    with urllib.request.urlopen(mp.OLLAMA + path, timeout=3) as r:
        return json.loads(r.read())


def ensure_ollama() -> str | None:
    """Start `ollama serve` if it isn't running. Returns an error message or None."""
    try:
        ollama_get("/api/tags")
        return None
    except OSError:
        pass
    exe = shutil.which("ollama")
    if not exe:
        return "Ollama is not running and was not found on PATH."
    subprocess.Popen([exe, "serve"], creationflags=NO_WINDOW,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(30):
        time.sleep(1)
        try:
            ollama_get("/api/tags")
            return None
        except OSError:
            continue
    return "Ollama did not start within 30 seconds."


def free_path(first: Path) -> Path:
    """report.pdf -> 'report - slides.pptx', then '(2)', '(3)'... never overwrites."""
    stem = f"{first.stem} - slides"
    out, n = first.with_name(f"{stem}.pptx"), 2
    while out.exists():
        out, n = first.with_name(f"{stem} ({n}).pptx"), n + 1
    return out


def thumbnails(preview_dir: Path, width: int = 360) -> list[str]:
    thumbs = []
    for png in sorted(preview_dir.glob("slide-*.png")):
        with Image.open(png) as im:
            im = im.convert("RGB")
            im.thumbnail((width, width))
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=82)
        thumbs.append("data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode())
    return thumbs


# ---------------------------------------------------------------- bridge

class UILogHandler(logging.Handler):
    def __init__(self, api: "Api"):
        super().__init__()
        self.api = api

    def emit(self, record):
        self.api._on_line(record.getMessage(), record.levelname.lower())


class UIWriter(io.TextIOBase):
    """Stand-in for stdout while a job runs, so make_ppt's print() lines reach the UI."""
    def __init__(self, api: "Api"):
        self.api = api

    def write(self, s):
        for line in s.splitlines():
            if line.strip():
                self.api._on_line(line.strip(), "info")
        return len(s)


class Api:
    def __init__(self):
        self._window: webview.Window | None = None
        self._busy = False
        log = logging.getLogger("make_ppt")
        log.setLevel(logging.INFO)
        log.addHandler(UILogHandler(self))
        log.propagate = False  # pythonw has no console to fall back to

    # -- helpers ---------------------------------------------------------
    def _push(self, event: dict) -> None:
        if self._window:
            self._window.evaluate_js(f"window.onEvent({json.dumps(event)})")

    def _on_line(self, text: str, level: str) -> None:
        self._push({"type": "log", "text": text, "level": level})
        low = text.lower()
        for key, stage in (("extracting", "read"), ("describing", "look"),
                           ("asking", "write"), ("deck:", "render")):
            if low.startswith(key) or f" {key}" in low[:12]:
                self._push({"type": "stage", "stage": stage, "detail": text})
                break

    # -- called from JavaScript ------------------------------------------
    def list_models(self) -> dict:
        if err := ensure_ollama():
            return {"error": err, "models": []}
        tags = ollama_get("/api/tags").get("models", [])
        free = free_memory_gb()
        models = sorted(({"name": m["name"], "size": round(m["size"] / 2**30, 1),
                          "fits": m["size"] / 2**30 * 1.1 <= free}
                         for m in tags if "cloud" not in m["name"] and "embed" not in m["name"]),
                        key=lambda m: m["size"])
        names = [m["name"] for m in models]
        last = LAST_MODEL.read_text().strip() if LAST_MODEL.exists() else ""
        default = next((n for n in (last, mp.DEFAULT_MODEL) if n in names), names[0] if names else "")
        outdir = LAST_OUTDIR.read_text(encoding="utf-8").strip() if LAST_OUTDIR.exists() else ""
        return {"models": models, "free": round(free), "default": default,
                "outdir": outdir if outdir and Path(outdir).is_dir() else ""}

    def pick_files(self) -> list[str]:
        picked = self._window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=True,
                                                file_types=FILE_TYPES)
        return [p for p in (picked or []) if Path(p).suffix.lower() in SUPPORTED]

    def pick_spec(self) -> str | None:
        picked = self._window.create_file_dialog(webview.FileDialog.OPEN, file_types=("Slide spec (*.json)",))
        return picked[0] if picked else None

    def pick_folder(self) -> str:
        picked = self._window.create_file_dialog(webview.FileDialog.FOLDER)
        return picked[0] if picked else ""

    def set_outdir(self, path: str) -> None:
        LAST_OUTDIR.write_text(path or "", encoding="utf-8")

    def open_path(self, path: str) -> None:
        os.startfile(path)

    def edit_spec(self, path: str) -> None:
        subprocess.Popen(["notepad.exe", path])

    def generate(self, opts: dict) -> bool:
        if self._busy:
            return False
        self._busy = True
        threading.Thread(target=self._job, args=("auto", opts), daemon=True).start()
        return True

    def rebuild(self, spec: str) -> bool:
        if self._busy:
            return False
        self._busy = True
        threading.Thread(target=self._job, args=("build", {"spec": spec}), daemon=True).start()
        return True

    # -- worker ------------------------------------------------------------
    def _job(self, kind: str, opts: dict) -> None:
        start = time.time()
        mp.OCR_USED.clear()
        mp.VISION_USED.clear()
        try:
            with contextlib.redirect_stdout(UIWriter(self)):
                if kind == "auto":
                    files = [Path(f) for f in opts["files"]]
                    LAST_MODEL.write_text(opts["model"])
                    outdir = Path(opts.get("outdir") or files[0].parent)
                    outdir.mkdir(parents=True, exist_ok=True)
                    work = outdir / f"{files[0].stem}_ppt"
                    deck = mp.auto(files, free_path(outdir / files[0].name), opts["model"], int(opts["slides"]),
                                   opts.get("focus", "").strip(), True, opts.get("animate", True),
                                   "auto" if opts.get("vision", True) else None, work)
                    spec = work / "spec.json"
                else:
                    spec = Path(opts["spec"])
                    self._push({"type": "stage", "stage": "render", "detail": "Rebuilding from spec"})
                    deck = mp.build(spec, free_path(spec.parent.with_name(spec.parent.name.removesuffix("_ppt"))))
            from pptx import Presentation
            self._push({"type": "done", "deck": str(deck), "name": deck.name, "folder": str(deck.parent),
                       "spec": str(spec), "made_with": Presentation(deck).core_properties.comments,
                       "seconds": round(time.time() - start), "thumbs": thumbnails(spec.parent / "preview")})
        except SystemExit as e:  # make_ppt reports user-facing problems this way
            self._push({"type": "error", "text": str(e.code or e)})
        except Exception as e:
            self._push({"type": "log", "text": traceback.format_exc(), "level": "error"})
            self._push({"type": "error", "text": f"{type(e).__name__}: {e}"})
        finally:
            self._busy = False


def on_start(window: webview.Window, api: Api) -> None:
    api._window = window

    def on_drop(event):
        paths = [f.get("pywebviewFullPath") for f in event["dataTransfer"].get("files", [])]
        ok = [p for p in paths if p and Path(p).suffix.lower() in SUPPORTED]
        api._push({"type": "files", "paths": ok, "skipped": len(paths) - len(ok)})

    doc = window.dom.document
    doc.events.dragover += DOMEventHandler(lambda e: None, True, True)
    doc.events.drop += DOMEventHandler(on_drop, True, True)


def already_running() -> bool:
    """One window at a time: two WebView2 instances sharing a data folder can freeze blank."""
    ctypes.windll.kernel32.CreateMutexW(None, False, "CursedPPTMaker.SingleInstance")
    if ctypes.windll.kernel32.GetLastError() != 183:  # ERROR_ALREADY_EXISTS
        return False
    ctypes.windll.user32.MessageBoxW(None, "Cursed PPTMaker is already open.", "Cursed PPTMaker", 0x40)
    return True


def main() -> None:
    if already_running():
        return
    api = Api()
    window = webview.create_window("Cursed PPTMaker", str(HERE / "ui" / "index.html"), js_api=api,
                                   width=1200, height=800, min_size=(1000, 700),
                                   background_color="#0d0f13")
    webview.start(on_start, (window, api))


if __name__ == "__main__":
    main()
