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
from families import FAMILIES
import providers

HERE = Path(__file__).resolve().parent
LAST_MODEL = HERE / ".last_model"   # shared with the console launcher
LAST_VISION = HERE / ".last_vision"  # image model: "auto", "" (off) or a model id
LAST_DESIGN = HERE / ".last_design"  # chosen design family key; empty = default style
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
        """Local Ollama models, Ollama cloud models and API-provider models, one list.
        Each: {name, label, kind: local|cloud|api, size, fits, provider}."""
        models, error, free = [], None, 0.0
        if err := ensure_ollama():
            error = err
        else:
            try:
                tags = ollama_get("/api/tags").get("models", [])
                free = free_memory_gb()
                retired = mp.retired_models()
                for m in tags:
                    name, gb = m["name"], m["size"] / 2**30
                    if "embed" in name or name in retired:
                        continue  # embedding models can't write slides; retired cloud models are gone
                    if mp.is_online(name):
                        models.append({"name": name, "label": name, "kind": "cloud", "size": 0,
                                       "fits": True, "provider": "Ollama cloud"})
                    else:
                        models.append({"name": name, "label": name, "kind": "local", "size": round(gb, 1),
                                       "fits": gb * 1.1 <= free, "provider": "this PC"})
            except OSError as e:  # includes TimeoutError while Ollama is busy loading a model
                error = f"Ollama did not answer ({e}). Try again in a moment."
        for prov in providers.load():
            for name in prov.get("models", []):
                models.append({"name": f"api:{prov['id']}:{name}", "label": name, "kind": "api",
                               "size": 0, "fits": True, "provider": prov["name"]})
        order = {"local": 0, "cloud": 1, "api": 2}
        models.sort(key=lambda m: (order[m["kind"]], m["size"], m["label"]))
        # image models: Ollama models that report vision support; API models can't be checked
        vision = [dict(m) for m in models if m["kind"] == "api" or mp.vision_capable(m["name"])]
        last_v = LAST_VISION.read_text(encoding="utf-8").strip() if LAST_VISION.exists() else "auto"
        vision_default = last_v if last_v in ("", "auto") or any(v["name"] == last_v for v in vision) else "auto"
        auto_pick = next((n for n in mp.vision_candidates(None)), "")
        names = [m["name"] for m in models]
        last = LAST_MODEL.read_text().strip() if LAST_MODEL.exists() else ""
        local = [m["name"] for m in models if m["kind"] == "local"]
        default = next((n for n in (last, mp.DEFAULT_MODEL) if n in names),
                       (local or names or [""])[0])
        outdir = LAST_OUTDIR.read_text(encoding="utf-8").strip() if LAST_OUTDIR.exists() else ""
        return {"models": models, "free": round(free), "default": default, "error": error,
                "vision": vision, "vision_default": vision_default, "vision_auto": auto_pick,
                "outdir": outdir if outdir and Path(outdir).is_dir() else ""}

    # -- design families ----------------------------------------------------
    def list_designs(self) -> dict:
        """The design families (families.py) with their preview thumbnails."""
        items = []
        for key, fam in FAMILIES.items():
            thumb = HERE / "ui" / "families" / f"{key}.jpg"
            data = base64.b64encode(thumb.read_bytes()).decode() if thumb.exists() else ""
            items.append({"id": key, "name": fam["name"], "about": fam["about"], "swatches": fam["accents"],
                          "thumb": f"data:image/jpeg;base64,{data}" if data else ""})
        last = LAST_DESIGN.read_text(encoding="utf-8").strip() if LAST_DESIGN.exists() else ""
        return {"designs": items, "default": last if last in FAMILIES else ""}

    def open_designs_folder(self) -> None:
        """The reference pictures the families were designed from (PPT_Designs/)."""
        folder = HERE / "PPT_Designs"
        folder.mkdir(exist_ok=True)
        os.startfile(folder)

    # -- online API providers ------------------------------------------------
    def provider_presets(self) -> dict:
        return {"presets": providers.PRESETS, "saved": providers.load()}

    def provider_models(self, base_url: str, key: str, pid: str = "") -> dict:
        """List a provider's models. Empty key = use the stored key of provider `pid`."""
        try:
            if not key and pid:
                key = providers._get(pid)[1]
            return {"models": providers.list_models(base_url, key)}
        except (ValueError, OSError) as e:
            return {"error": str(e)}

    def save_provider(self, name: str, base_url: str, key: str, model: str) -> dict:
        try:
            pid = providers.slug(name)
            existing = next((p["models"] for p in providers.load() if p["id"] == pid), [])
            providers.save(name, base_url, key, list(dict.fromkeys(existing + [model])))
            return {"ok": True, "model": f"api:{pid}:{model}"}
        except (ValueError, OSError) as e:
            return {"error": str(e)}

    def delete_provider(self, pid: str) -> None:
        providers.delete(pid)

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

    # -- slide editor ------------------------------------------------------
    # The editor works on a copy of spec.json in the page; these calls save it, preview it,
    # export it and ask the model for single slides. Keys starting with "_" are the page's own.

    @staticmethod
    def _spec_thumbs(spec_dir: Path) -> list[str]:
        """One preview per spec slide (a long slide's continuation slides are skipped).
        No map = the spec changed since the last build, so the previews no longer match."""
        try:
            first = json.loads((spec_dir / "slide_map.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return []
        thumbs = thumbnails(spec_dir / "preview")
        return [thumbs[j] if j < len(thumbs) else "" for j in first]

    def load_spec(self, path: str) -> dict:
        spec_path = Path(path)
        try:
            spec = json.loads(spec_path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as e:
            return {"error": f"Could not read {spec_path.name}: {e}"}
        decks = sorted(spec_path.parent.parent.glob(f"{spec_path.parent.name.removesuffix('_ppt')}*.pptx"),
                       key=lambda p: p.stat().st_mtime)
        return {"spec": spec, "thumbs": self._spec_thumbs(spec_path.parent),
                "deck": str(decks[-1]) if decks else "", "families": {k: f["name"] for k, f in FAMILIES.items()}}

    def _save_spec(self, path: str, spec: dict) -> list[str]:
        spec = {**spec, "slides": [{k: v for k, v in s.items() if not k.startswith("_") and v not in ("", [], None)}
                                   for s in spec.get("slides", [])]}
        spec_path = Path(path)
        errors = mp.validate(spec, spec_path.parent)
        text = json.dumps(spec, indent=2, ensure_ascii=False)
        if not spec_path.exists() or spec_path.read_text(encoding="utf-8") != text:
            (spec_path.parent / "slide_map.json").unlink(missing_ok=True)  # previews are now stale
            spec_path.write_text(text, encoding="utf-8")
        return errors

    def _quiet(self, fn):
        """Run make_ppt code from a JS call: stdout goes to the log, sys.exit becomes an error."""
        if self._busy:
            return {"error": "Busy with another job - wait for it to finish."}
        self._busy = True
        mp.STOP.clear()
        mp.SKIP.clear()
        try:
            with contextlib.redirect_stdout(UIWriter(self)):
                return fn()
        except mp.Cancelled:
            return {"error": "Cancelled."}
        except SystemExit as e:
            return {"error": str(e.code or e)}
        except Exception as e:
            self._push({"type": "log", "text": traceback.format_exc(), "level": "error"})
            return {"error": f"{type(e).__name__}: {e}"}
        finally:
            self._busy = False

    def save_spec(self, path: str, spec: dict) -> dict:
        return {"errors": self._save_spec(path, spec)}

    def preview_spec(self, path: str, spec: dict) -> dict:
        """Build a scratch deck from the edited spec and return fresh slide thumbnails."""
        def run():
            if errors := self._save_spec(path, spec):
                return {"error": "Fix these first:\n" + "\n".join(errors)}
            scratch = Path(path).parent / "_editor_preview.pptx"
            mp.build(Path(path), scratch, True)
            scratch.unlink(missing_ok=True)
            return {"thumbs": self._spec_thumbs(Path(path).parent)}
        return self._quiet(run)

    def export_deck(self, path: str, spec: dict, suggested: str) -> dict:
        """Ask where to save, then build the edited deck there."""
        if errors := self._save_spec(path, spec):
            return {"error": "Fix these first:\n" + "\n".join(errors)}
        picked = self._window.create_file_dialog(webview.FileDialog.SAVE, save_filename=suggested,
                                                 directory=str(Path(path).parent.parent),
                                                 file_types=("PowerPoint (*.pptx)",))
        if not picked:
            return {}
        out = Path(picked if isinstance(picked, str) else picked[0])
        out = out.with_suffix(".pptx")

        def run():
            deck = mp.build(Path(path), out, True)
            return {"deck": str(deck), "folder": str(deck.parent), "name": deck.name,
                    "thumbs": self._spec_thumbs(Path(path).parent), "all_thumbs": thumbnails(Path(path).parent / "preview")}
        return self._quiet(run)

    def pick_image(self, path: str) -> str:
        """Choose a picture for a slide; it is copied into the deck's assets folder."""
        picked = self._window.create_file_dialog(webview.FileDialog.OPEN, file_types=(
            "Pictures (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp)",))
        if not picked:
            return ""
        src = Path(picked[0])
        assets = Path(path).parent / "assets"
        assets.mkdir(exist_ok=True)
        dst = assets / src.name
        if dst.resolve() != src.resolve():
            shutil.copy2(src, dst)
        return f"assets/{src.name}"

    def ai_slide(self, path: str, spec: dict, index: int, instruction: str, insert: bool, model: str) -> dict:
        def run():
            slide = mp.ai_slide(model, spec, Path(path).parent, index, instruction.strip(), insert)
            return {"slide": slide}
        return self._quiet(run)

    def generate(self, opts: dict) -> bool:
        if self._busy:
            return False
        self._busy = True
        threading.Thread(target=self._job, args=(opts,), daemon=True).start()
        return True

    def cancel(self) -> None:
        mp.STOP.set()

    def skip(self, stage: str) -> None:
        """look = stop describing pictures, write = stop retrying (keep the best spec so far),
        preview = no slide previews."""
        mp.SKIP.add(stage)

    # -- worker ------------------------------------------------------------
    def _job(self, opts: dict) -> None:
        start = time.time()
        mp.STOP.clear()
        mp.SKIP.clear()
        mp.OCR_USED.clear()
        mp.VISION_USED.clear()
        try:
            with contextlib.redirect_stdout(UIWriter(self)):
                files = [Path(f) for f in opts["files"]]
                LAST_MODEL.write_text(opts["model"])
                LAST_VISION.write_text(opts.get("vision", "auto"), encoding="utf-8")
                LAST_DESIGN.write_text(opts.get("design", ""), encoding="utf-8")
                design = opts.get("design") or None
                outdir = Path(opts.get("outdir") or files[0].parent)
                outdir.mkdir(parents=True, exist_ok=True)
                work = outdir / f"{files[0].stem}_ppt"
                deck = mp.auto(files, free_path(outdir / files[0].name), opts["model"], int(opts["slides"]),
                               opts.get("focus", "").strip(), True, opts.get("animate", True),
                               opts.get("vision", "auto") or None, work, design, opts.get("slide_mode", "exact"))
                spec = work / "spec.json"
            from pptx import Presentation
            self._push({"type": "done", "deck": str(deck), "name": deck.name, "folder": str(deck.parent),
                       "spec": str(spec), "made_with": Presentation(deck).core_properties.comments,
                       "seconds": round(time.time() - start), "thumbs": [] if "preview" in mp.SKIP else thumbnails(spec.parent / "preview")})
        except mp.Cancelled:
            self._push({"type": "cancelled"})
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
