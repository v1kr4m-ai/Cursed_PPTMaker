"""Online LLM providers for make_ppt: any OpenAI-compatible chat API.

Providers are stored in %APPDATA%\\CursedPPTMaker\\providers.json (outside the project, so
never committed). API keys are encrypted with Windows DPAPI: only this Windows user on this
PC can decrypt them. Keys are never logged or printed.

A provider model is addressed as  api:<provider id>:<model name>,  e.g.
api:openai:gpt-4.1-mini  or  api:openrouter:meta-llama/llama-3.3-70b-instruct
"""
from __future__ import annotations

import base64
import ctypes
import json
import os
import re
import urllib.error
import urllib.request
from ctypes import wintypes
from pathlib import Path

CONFIG = Path(os.environ.get("APPDATA", Path.home())) / "CursedPPTMaker" / "providers.json"

PRESETS = {  # name -> OpenAI-compatible base URL
    "OpenAI": "https://api.openai.com/v1",
    "Google Gemini": "https://generativelanguage.googleapis.com/v1beta/openai",
    "OpenRouter": "https://openrouter.ai/api/v1",
    "Groq": "https://api.groq.com/openai/v1",
    "Anthropic": "https://api.anthropic.com/v1",
}


# ---------------------------------------------------------------- DPAPI

class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, encrypt: bool) -> bytes:
    buf = ctypes.create_string_buffer(data, len(data))
    inp = _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = _Blob()
    crypt32 = ctypes.windll.crypt32
    fn = crypt32.CryptProtectData if encrypt else crypt32.CryptUnprotectData
    if not fn(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("Windows could not " + ("encrypt" if encrypt else "decrypt") + " the API key")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(out.pbData)


# ---------------------------------------------------------------- storage

def load() -> list[dict]:
    """Providers without keys: [{id, name, base_url, models}]."""
    if not CONFIG.exists():
        return []
    try:
        data = json.loads(CONFIG.read_text(encoding="utf-8"))
    except ValueError:
        return []
    return [{k: p[k] for k in ("id", "name", "base_url", "models") if k in p} for p in data]


def _raw() -> list[dict]:
    return json.loads(CONFIG.read_text(encoding="utf-8")) if CONFIG.exists() else []


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "provider"


def save(name: str, base_url: str, key: str, models: list[str]) -> dict:
    """Add or replace a provider. An empty key keeps the stored one."""
    pid = slug(name)
    providers = [p for p in _raw() if p["id"] != pid]
    old = next((p for p in _raw() if p["id"] == pid), {})
    enc = base64.b64encode(_dpapi(key.encode(), True)).decode() if key else old.get("key", "")
    if not enc:
        raise ValueError("An API key is required.")
    entry = {"id": pid, "name": name.strip(), "base_url": base_url.rstrip("/"),
             "models": [m.strip() for m in models if m.strip()], "key": enc}
    providers.append(entry)
    CONFIG.parent.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(providers, indent=2), encoding="utf-8")
    return {k: entry[k] for k in ("id", "name", "base_url", "models")}


def delete(pid: str) -> None:
    CONFIG.write_text(json.dumps([p for p in _raw() if p["id"] != pid], indent=2), encoding="utf-8")


def _get(pid: str) -> tuple[dict, str]:
    p = next((p for p in _raw() if p["id"] == pid), None)
    if not p:
        raise ValueError(f"No API provider named '{pid}'. Add it in the app first.")
    return p, _dpapi(base64.b64decode(p["key"]), False).decode()


def parse(model: str) -> tuple[str, str]:
    """'api:openai:gpt-4.1-mini' -> ('openai', 'gpt-4.1-mini')"""
    _, pid, name = model.split(":", 2)
    return pid, name


def label(model: str) -> str:
    pid, name = parse(model)
    p = next((p for p in load() if p["id"] == pid), {"name": pid})
    return f"{name} via {p['name']}"


# ---------------------------------------------------------------- HTTP

def _request(url: str, key: str, body: dict | None, timeout: int) -> dict:
    headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json",
               "User-Agent": "CursedPPTMaker"}
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data, headers, method="POST" if data else "GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())


def _http_reason(e: urllib.error.HTTPError) -> str:
    try:
        err = json.loads(e.read())
        err = err.get("error", err) if isinstance(err, dict) else err
        msg = err.get("message", err) if isinstance(err, dict) else err
        return str(msg)[:300]
    except Exception:
        return f"HTTP {e.code}"


def list_models(base_url: str, key: str, timeout: int = 20) -> list[str]:
    """Model ids from the provider's /models endpoint (used to fill the picker)."""
    try:
        data = _request(base_url.rstrip("/") + "/models", key, None, timeout)
    except urllib.error.HTTPError as e:
        raise ValueError(("API key rejected: " if e.code in (401, 403) else "") + _http_reason(e)) from None
    except (urllib.error.URLError, TimeoutError) as e:
        raise ValueError(f"Could not reach {base_url} ({e})") from None
    ids = [m.get("id", "") for m in data.get("data", data.get("models", []))]
    return sorted({i.removeprefix("models/") for i in ids if i})


def chat_json(model: str, prompt: str, max_tokens: int, timeout: int) -> str:
    """Send one prompt, ask for a JSON reply, return the reply text. Raises ValueError with
    a readable message on failure."""
    pid, name = parse(model)
    p, key = _get(pid)
    url = p["base_url"] + "/chat/completions"
    body = {"model": name, "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3, "max_tokens": max_tokens, "response_format": {"type": "json_object"}}
    for attempt in range(2):
        try:
            data = _request(url, key, body, timeout)
            break
        except urllib.error.HTTPError as e:
            reason = _http_reason(e)
            if e.code == 400 and attempt == 0 and "response_format" in body:
                body.pop("response_format")  # some providers/models don't support JSON mode
                continue
            if e.code in (401, 403):
                raise ValueError(f"{p['name']} rejected the API key: {reason}") from None
            if e.code == 429:
                raise ValueError(f"{p['name']} rate limit or quota reached: {reason}") from None
            raise ValueError(f"{p['name']} error {e.code}: {reason}") from None
        except TimeoutError:
            raise ValueError(f"{p['name']} did not answer within {timeout // 60} minutes.") from None
        except urllib.error.URLError as e:
            raise ValueError(f"Could not reach {p['name']} - are you online? ({e.reason})") from None
    text = (data.get("choices") or [{}])[0].get("message", {}).get("content") or ""
    fenced = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)  # models without JSON mode often fence it
    return (fenced.group(1) if fenced else text).strip()


def describe_image(model: str, jpeg_b64: str, prompt: str, max_tokens: int, timeout: int) -> str:
    """One picture + prompt to a vision-capable chat model (OpenAI image_url format)."""
    pid, name = parse(model)
    p, key = _get(pid)
    body = {"model": name, "temperature": 0.2, "max_tokens": max_tokens, "messages": [{
        "role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{jpeg_b64}"}}]}]}
    try:
        data = _request(p["base_url"] + "/chat/completions", key, body, timeout)
    except urllib.error.HTTPError as e:
        raise ValueError(f"{p['name']} error {e.code}: {_http_reason(e)}") from None
    except TimeoutError:
        raise ValueError(f"{p['name']} did not answer in time") from None
    except urllib.error.URLError as e:
        raise ValueError(f"could not reach {p['name']} ({e.reason})") from None
    return ((data.get("choices") or [{}])[0].get("message", {}).get("content") or "").strip()
