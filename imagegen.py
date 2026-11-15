"""Slide pictures from a local ComfyUI server (https://github.com/comfyanonymous/ComfyUI).

Plain txt2img through ComfyUI's HTTP API. ComfyUI is found at COMFYUI_URL (default
http://127.0.0.1:8188); if nothing answers and COMFYUI_DIR (or a known install folder) holds a
portable install, it is started for the run and stopped again afterwards.
"""
import json
import os
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

URL = os.environ.get("COMFYUI_URL", "http://127.0.0.1:8188").rstrip("/")
INSTALLS = [os.environ.get("COMFYUI_DIR", ""), r"F:\ComfyUI\ComfyUI_windows_portable",
            r"C:\ComfyUI_windows_portable", r"D:\ComfyUI_windows_portable"]
# Checkpoints in order of preference: SDXL-class draws the best slide art, SD 1.5 is the fast fallback.
# size: (w, h) near 16:9, in the model's native pixel budget.
PREFER = [("sd_xl_base", (1344, 768)), ("juggernautxl", (1344, 768)), ("dreamshaper", (768, 432)),
          ("sd15", (768, 432)), ("v1-5", (768, 432))]
STYLE = ("clean modern vector illustration, flat colours, soft gradients, simple shapes, "
         "professional presentation art, centered, plain background")
NEGATIVE = "text, letters, words, numbers, watermark, logo, signature, photo of a person, blurry, deformed, cluttered"
STOP = lambda: False  # replaced by make_ppt: True when the user pressed Cancel or Skip


def _call(path: str, body: dict | None = None, timeout: int = 20):
    req = urllib.request.Request(URL + path, None if body is None else json.dumps(body).encode(),
                                 {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read()
    return json.loads(data) if data[:1] in (b"{", b"[") else data


def running() -> bool:
    try:
        _call("/system_stats", timeout=3)
        return True
    except (OSError, ValueError):
        return False


def install_dir() -> Path | None:
    return next((Path(p) for p in INSTALLS if p and (Path(p) / "ComfyUI" / "main.py").exists()), None)


def start() -> subprocess.Popen | None:
    """Start a portable ComfyUI and wait until it answers. None if it was already running."""
    if running():
        return None
    root = install_dir()
    if not root:
        raise RuntimeError(f"ComfyUI is not running at {URL} and no install was found. Start it, or set "
                           "COMFYUI_DIR to the folder holding run_nvidia_gpu.bat.")
    proc = subprocess.Popen([str(root / "python_embeded" / "python.exe"), "-s", str(root / "ComfyUI" / "main.py"),
                             "--windows-standalone-build"], cwd=root, creationflags=0x08000000,  # no window
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(240):  # a cold start loads torch: up to ~2 minutes
        time.sleep(0.5)
        if STOP():
            proc.terminate()
            raise RuntimeError("stopped")
        if running():
            return proc
        if proc.poll() is not None:
            break
    proc.terminate()
    raise RuntimeError("ComfyUI did not start (run its .bat by hand to see why)")


def pick_checkpoint(want: str = "") -> tuple[str, tuple[int, int]]:
    """(checkpoint file name, image size). `want` = a name the user chose; else the best installed."""
    info = _call("/object_info/CheckpointLoaderSimple")
    names = info["CheckpointLoaderSimple"]["input"]["required"]["ckpt_name"][0]
    if want in names:
        return want, next((s for k, s in PREFER if k in want.lower()), (1024, 576))
    for key, size in PREFER:
        if hit := next((n for n in names if key in n.lower()), None):
            return hit, size
    if not names:
        raise RuntimeError("ComfyUI has no checkpoints installed (models/checkpoints)")
    return names[0], (1024, 576)


def draw(prompt: str, out: Path, ckpt: str, size: tuple[int, int], seed: int | None = None) -> Path:
    """Generate one picture into `out` (png)."""
    w, h = size
    steps = 20 if "turbo" not in ckpt.lower() else 4
    graph = {
        "1": {"class_type": "CheckpointLoaderSimple", "inputs": {"ckpt_name": ckpt}},
        "2": {"class_type": "CLIPTextEncode", "inputs": {"text": f"{prompt}, {STYLE}", "clip": ["1", 1]}},
        "3": {"class_type": "CLIPTextEncode", "inputs": {"text": NEGATIVE, "clip": ["1", 1]}},
        "4": {"class_type": "EmptyLatentImage", "inputs": {"width": w, "height": h, "batch_size": 1}},
        "5": {"class_type": "KSampler", "inputs": {
            "model": ["1", 0], "positive": ["2", 0], "negative": ["3", 0], "latent_image": ["4", 0],
            "seed": seed if seed is not None else uuid.uuid4().int % 2**32, "steps": steps, "cfg": 6.5 if steps > 4 else 1.0,
            "sampler_name": "euler", "scheduler": "normal", "denoise": 1.0}},
        "6": {"class_type": "VAEDecode", "inputs": {"samples": ["5", 0], "vae": ["1", 2]}},
        "7": {"class_type": "SaveImage", "inputs": {"images": ["6", 0], "filename_prefix": "pptmaker"}},
    }
    pid = _call("/prompt", {"prompt": graph, "client_id": uuid.uuid4().hex})["prompt_id"]
    for _ in range(1200):  # 10 minutes
        time.sleep(0.5)
        if STOP():
            try:
                _call("/interrupt", {})
            except OSError:
                pass
            raise RuntimeError("stopped")
        hist = _call(f"/history/{pid}").get(pid)
        if hist and hist.get("outputs"):
            img = hist["outputs"]["7"]["images"][0]
            q = urllib.parse.urlencode({"filename": img["filename"], "subfolder": img["subfolder"], "type": img["type"]})
            out.write_bytes(_call(f"/view?{q}", timeout=60))
            return out
        if hist and hist.get("status", {}).get("status_str") == "error":
            raise RuntimeError("ComfyUI failed to draw the picture (see its console)")
    raise RuntimeError("ComfyUI took too long")


def free_memory() -> None:
    """Hand the GPU back (to Ollama) once the pictures are done."""
    try:
        _call("/free", {"unload_models": True, "free_memory": True})
    except (OSError, ValueError):
        pass
