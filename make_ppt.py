#!/usr/bin/env python
"""make_ppt.py - build PowerPoint decks from documents, fully offline.

Commands:
  extract FILE...                 Dump text, tables and images to Markdown so an LLM can read them.
  build   SPEC.json [-o OUT.pptx] Render a deck from a JSON slide spec (see `schema`).
  auto    FILE... [-o OUT.pptx]   extract -> local Ollama model writes the spec -> build.
  schema                          Print the slide spec format.

Inputs: .pdf (scanned pages are OCR'd), .docx/.doc/.odt/.rtf, .xlsx/.xls/.ods/.csv,
.pptx/.ppt/.odp, images (.png/.jpg/...; OCR'd), .txt/.md.

Every build is rendered to PNG previews through LibreOffice so the result can be checked
without opening PowerPoint. Text is auto-shrunk to fit its box; anything that still does
not fit is reported as a warning.

Images (and pictures inside documents) are described by a local vision model when one is
installed (llama3.2-vision, llava, moondream...), so photos and diagrams get sensible slides.

Env overrides: MAKE_PPT_MODEL, OLLAMA_HOST, SOFFICE_PATH, TESSERACT_PATH.
"""
from __future__ import annotations

import argparse
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SOFFICE = os.environ.get("SOFFICE_PATH", r"C:\Program Files\LibreOffice\program\soffice.exe")
TESSERACT = os.environ.get("TESSERACT_PATH", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
# Language data next to this script (eng, osd, hin). Falls back to Tesseract's own folder.
_local_tessdata = Path(__file__).with_name("tessdata")
TESSDATA = os.environ.get("TESSDATA_DIR") or (str(_local_tessdata) if _local_tessdata.is_dir() else "")
CS_FONT = "Nirmala UI"  # Windows font with Devanagari (and other Indic scripts)
OLLAMA = os.environ.get("OLLAMA_HOST", "http://localhost:11434").rstrip("/")
DEFAULT_MODEL = os.environ.get("MAKE_PPT_MODEL", "qwen3-coder:30b")

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
LEGACY = {".doc": "docx", ".odt": "docx", ".rtf": "docx", ".xls": "xlsx", ".ods": "xlsx",
          ".ppt": "pptx", ".odp": "pptx"}
MAX_TABLE_ROWS = 40          # rows per sheet/table written to the extract
MIN_IMAGE_BYTES = 8_000      # skip icons and spacer images in documents

logging.getLogger("pdfminer").setLevel(logging.ERROR)  # FontBBox noise from pdfplumber
log = logging.getLogger("make_ppt")


# ---------------------------------------------------------------- extraction

def md_table(rows: list[list]) -> str:
    def cell(v) -> str:
        return "" if v is None else str(v).replace("|", "\\|").replace("\n", " ").strip()
    rows = [[cell(c) for c in r] for r in rows if r and any(c not in (None, "") for c in r)]
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(lines)


def img_ref(p: Path) -> str:
    """Markdown image link relative to the work dir, which is where spec.json lives."""
    return f"![image](assets/{p.name})"


_langs: set[str] | None = None
OCR_USED: set[str] = set()     # languages actually used, recorded in the deck's properties
VISION_USED: set[str] = set()  # vision models actually used


def ocr_lang(img: Image.Image, config: str) -> str:
    """Pick the OCR language from the page's script: Devanagari -> Hindi + English."""
    import pytesseract
    global _langs
    if _langs is None:
        try:
            _langs = set(pytesseract.get_languages(config=config))
        except Exception:
            _langs = {"eng"}
    if "hin" not in _langs:
        lang = "eng"
    else:
        try:
            script = pytesseract.image_to_osd(img, config=config, output_type=pytesseract.Output.DICT)["script"]
            lang = "hin+eng" if script == "Devanagari" else "eng"
        except Exception:  # OSD needs enough text; hin+eng still reads English well
            lang = "hin+eng"
    OCR_USED.add(lang)
    return lang


def ocr(img: Image.Image) -> str:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = TESSERACT
    config = f"--tessdata-dir {Path(TESSDATA).as_posix()}" if TESSDATA else ""
    try:
        return pytesseract.image_to_string(img, lang=ocr_lang(img, config), config=config).strip()
    except pytesseract.TesseractNotFoundError:
        log.warning("Tesseract not found at %s - skipping OCR", TESSERACT)
        return ""


def save_image(blob: bytes, assets: Path, name: str) -> Path | None:
    if len(blob) < MIN_IMAGE_BYTES:
        return None
    assets.mkdir(parents=True, exist_ok=True)
    out = assets / name
    out.write_bytes(blob)
    try:
        Image.open(out).verify()
    except Exception:
        out.unlink()  # not a format Pillow/PowerPoint can use (e.g. JBIG2, EMF)
        return None
    return out


def extract_pdf(path: Path, assets: Path) -> str:
    import pdfplumber
    import pypdf
    import pypdfium2
    parts, raster = [], None
    reader = pypdf.PdfReader(path)
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = (page.extract_text() or "").strip()
            label = f"## Page {i + 1}"
            scanned = len(text) < 20  # no text layer
            if scanned:
                raster = raster or pypdfium2.PdfDocument(path)
                text = ocr(raster[i].render(scale=2).to_pil())
                label += " (OCR)"
            parts.append(f"{label}\n\n{text}")
            for t in page.extract_tables():
                if t and max(map(len, t)) < 2:
                    continue  # one-column boxes are code/callouts, already in the text
                if table := md_table(t):
                    parts.append(f"Table on page {i + 1}:\n\n{table}")
            if scanned:
                continue  # its only "image" is the scan of the page itself
            try:
                for j, im in enumerate(reader.pages[i].images):
                    ext = Path(im.name).suffix or ".png"
                    if p := save_image(im.data, assets, f"p{i + 1}_img{j + 1}{ext}"):
                        parts.append(img_ref(p))
            except Exception as e:  # odd image encodings shouldn't kill the text extract
                log.warning("page %d: could not extract images (%s)", i + 1, e)
    if raster:
        raster.close()
    return "\n\n".join(parts)


def extract_docx(path: Path, assets: Path) -> str:
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    doc = docx.Document(path)
    parts = []
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            p = Paragraph(el, doc)
            text = p.text.strip()
            if not text:
                continue
            style = (p.style.name if p.style is not None else "") or ""
            if style.startswith("Heading") and style[-1:].isdigit():
                text = "#" * (int(style[-1]) + 1) + " " + text
            elif style == "Title":
                text = "# " + text
            elif "List" in style:
                text = "- " + text
            parts.append(text)
        elif el.tag == qn("w:tbl"):
            t = Table(el, doc)
            parts.append(md_table([[c.text for c in row.cells] for row in t.rows]))
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            name = Path(rel.target_part.partname).name
            if p := save_image(rel.target_part.blob, assets, name):
                parts.append(img_ref(p))
    return "\n\n".join(parts)


def extract_sheet(path: Path, assets: Path) -> str:
    import pandas as pd
    sheets = ({"data": pd.read_csv(path)} if path.suffix.lower() == ".csv"
              else pd.read_excel(path, sheet_name=None))
    parts = []
    for name, df in sheets.items():
        df = df.dropna(how="all").dropna(axis=1, how="all")
        parts.append(f"## Sheet: {name} ({len(df)} rows x {len(df.columns)} cols)")
        rows = [list(df.columns)] + df.head(MAX_TABLE_ROWS).values.tolist()
        parts.append(md_table(rows))
        if len(df) > MAX_TABLE_ROWS:
            parts.append(f"... {len(df) - MAX_TABLE_ROWS} more rows not shown.")
            numeric = df.select_dtypes("number")
            if not numeric.empty:
                parts.append("Column totals:\n\n" + md_table(
                    [["column", "sum", "mean"]] +
                    [[c, round(numeric[c].sum(), 2), round(numeric[c].mean(), 2)] for c in numeric]))
    return "\n\n".join(parts)


def extract_pptx(path: Path, assets: Path) -> str:
    parts = []
    for i, slide in enumerate(Presentation(path).slides, 1):
        lines = [f"## Slide {i}"]
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                lines.append(shape.text_frame.text.strip())
            elif shape.has_table:
                lines.append(md_table([[c.text for c in r.cells] for r in shape.table.rows]))
            elif shape.shape_type == 13:  # picture
                img = shape.image
                if p := save_image(img.blob, assets, f"s{i}_{shape.shape_id}.{img.ext}"):
                    lines.append(img_ref(p))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            lines.append("Notes: " + slide.notes_slide.notes_text_frame.text.strip())
        parts.append("\n\n".join(lines))
    return "\n\n".join(parts)


def extract_image(path: Path, assets: Path) -> str:
    assets.mkdir(parents=True, exist_ok=True)
    copy = assets / path.name
    shutil.copy2(path, copy)
    with Image.open(path) as img:
        size = f"{img.width}x{img.height}"
        text = ocr(img.convert("RGB"))
    out = f"Image file, {size} px:\n\n{img_ref(copy)}"
    return out + (f"\n\nText found in image (OCR):\n\n{text}" if text else "")


def convert_legacy(path: Path, fmt: str, outdir: Path) -> Path:
    if not Path(SOFFICE).exists():
        sys.exit(f"LibreOffice not found at {SOFFICE}; needed to read {path.suffix} files.")
    subprocess.run([SOFFICE, "--headless", "--convert-to", fmt, "--outdir", str(outdir), str(path)],
                   check=True, capture_output=True, timeout=300)
    out = outdir / f"{path.stem}.{fmt}"
    if not out.exists():
        sys.exit(f"LibreOffice could not convert {path}")
    return out


def extract(path: Path, assets: Path) -> str:
    ext = path.suffix.lower()
    if ext in LEGACY:
        with tempfile.TemporaryDirectory() as tmp:
            return extract(convert_legacy(path, LEGACY[ext], Path(tmp)), assets)
    handlers = {".pdf": extract_pdf, ".docx": extract_docx, ".xlsx": extract_sheet,
                ".xlsm": extract_sheet, ".csv": extract_sheet, ".pptx": extract_pptx}
    if ext in handlers:
        return handlers[ext](path, assets)
    if ext in IMAGE_EXT:
        return extract_image(path, assets)
    if ext in {".txt", ".md"}:
        return path.read_text(encoding="utf-8", errors="replace")
    sys.exit(f"Unsupported file type: {path}")


VISION_PREFS = ("llama3.2-vision", "qwen2.5vl", "gemma3", "minicpm-v", "llava", "moondream")
MAX_DESCRIBED = 12  # images described per run; each takes a few seconds
IMG_REF = re.compile(r"!\[image\]\(assets/([^)]+)\)")


def vision_candidates(preferred: str | None) -> list[str]:
    """Installed offline vision models, best first. Empty if Ollama is down or none installed."""
    try:
        with urllib.request.urlopen(f"{OLLAMA}/api/tags", timeout=5) as r:
            names = [m["name"] for m in json.loads(r.read())["models"]]
    except (urllib.error.URLError, OSError, KeyError, ValueError):
        return []
    names = [n for n in names if "cloud" not in n]
    found = [n for pref in VISION_PREFS for n in names if n.split(":")[0] == pref]
    if preferred:
        found = [preferred] + [n for n in found if n != preferred]
    return found


def describe_image(model: str, path: Path) -> str:
    """Short factual description from a local vision model ('' on failure)."""
    import base64
    import io
    with Image.open(path) as im:
        im = im.convert("RGB")
        im.thumbnail((1024, 1024))  # smaller = much faster, still enough detail
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=85)
    body = json.dumps({
        "model": model, "stream": False, "keep_alive": 0,  # free memory for the slide model
        "prompt": "Describe this image in 2-3 factual sentences for someone building a slide deck: "
                  "what it shows, and any text, numbers or labels visible. Do not guess.",
        "images": [base64.b64encode(buf.getvalue()).decode()],
        "options": {"temperature": 0.2, "num_ctx": 4096}}).encode()
    req = urllib.request.Request(f"{OLLAMA}/api/generate", body, {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=600) as r:
        return json.loads(r.read()).get("response", "").strip()


def add_image_descriptions(md: str, assets: Path, preferred: str | None) -> str:
    """Append a vision-model description after each image reference, so a text-only
    slide model knows what photos and diagrams show."""
    names = list(dict.fromkeys(IMG_REF.findall(md)))[:MAX_DESCRIBED]
    if not names:
        return md
    models = vision_candidates(preferred)
    if not models:
        log.warning("no offline vision model installed (e.g. `ollama pull moondream`); "
                    "images without text will not be described")
        return md
    descriptions = {}
    for name in names:
        while models:
            try:
                log.info("describing %s with %s", name, models[0])
                descriptions[name] = describe_image(models[0], assets / name)
                VISION_USED.add(models[0])
                break
            except urllib.error.HTTPError as e:  # usually "not enough memory": try a smaller model
                log.warning("%s failed (%s); trying the next vision model", models[0], e.code)
                models.pop(0)
            except (urllib.error.URLError, OSError) as e:
                log.warning("could not describe %s (%s)", name, e)
                break
    return IMG_REF.sub(lambda m: m.group(0) + (
        f"\n\nImage description (vision model): {descriptions[m.group(1)]}"
        if descriptions.get(m.group(1)) else ""), md)


def extract_all(files: list[Path], workdir: Path, vision: str | None = "auto") -> Path:
    workdir.mkdir(parents=True, exist_ok=True)
    if missing := [str(f) for f in files if not f.exists()]:
        sys.exit("File not found: " + ", ".join(missing))
    chunks = []
    for f in files:
        log.info("extracting %s", f.name)
        chunks.append(f"# Source: {f.name}\n\n{extract(f, workdir / 'assets')}")
    text = "\n\n".join(chunks)
    if vision:  # "auto" = best installed vision model, None = skip
        text = add_image_descriptions(text, workdir / "assets", None if vision == "auto" else vision)
    md = workdir / "extracted.md"
    md.write_text(text, encoding="utf-8")
    return md


# ---------------------------------------------------------------- spec

SCHEMA = """{
  "title": "Deck title",
  "palette": "violet | midnight | forest | coral | terracotta | ocean | charcoal | teal | berry | cherry",
  "slides": [
    {"type": "title",   "title": "...", "subtitle": "...", "footer": "optional small text"},
    {"type": "section", "title": "...", "subtitle": "optional"},
    {"type": "bullets", "title": "...", "bullets": ["Lead phrase: detail", "..."],
                        "image": "optional path", "callout": "optional highlighted tip"},
    {"type": "cards",   "title": "...", "cards": [{"title": "...", "text": "...", "icon": "1-2 chars, optional"}]},
    {"type": "steps",   "title": "...", "steps": [{"title": "...", "text": "..."}]},
    {"type": "stats",   "title": "...", "stats": [{"value": "42%", "label": "..."}], "text": "optional"},
    {"type": "code",    "title": "...", "code": "multi-line string", "lang": "bash", "note": "optional warning"},
    {"type": "table",   "title": "...", "header": ["A", "B"], "rows": [["1", "2"]]},
    {"type": "chart",   "title": "...", "chart": "column | bar | line | pie",
                        "categories": ["Jan", "Feb"], "series": [{"name": "Sales", "values": [10, 20]}],
                        "text": "optional commentary shown beside the chart"},
    {"type": "image",   "title": "...", "image": "path", "caption": "optional", "bullets": ["optional"]},
    {"type": "closing", "title": "...", "items": ["checklist or key takeaways"]}
  ]
}
Every slide may also have "kicker" (small label above the title) and "notes" (speaker notes).
Top level may set "animate": false to turn off fade transitions and entrance animations."""

REQUIRED = {"title": ["title"], "section": ["title"], "bullets": ["title", "bullets"],
            "cards": ["title", "cards"], "steps": ["title", "steps"], "stats": ["title", "stats"],
            "code": ["title", "code"], "table": ["title", "header", "rows"],
            "chart": ["title", "categories", "series"], "image": ["title", "image"],
            "closing": ["title"]}


def validate(spec: dict, base: Path) -> list[str]:
    if not isinstance(spec, dict) or not isinstance(spec.get("slides"), list) or not spec["slides"]:
        return ['top level must be an object with a non-empty "slides" list']
    errors = []
    for i, s in enumerate(spec["slides"], 1):
        t = s.get("type") if isinstance(s, dict) else None
        if t not in REQUIRED:
            errors.append(f"slide {i}: unknown type {t!r}; use one of {sorted(REQUIRED)}")
            continue
        errors += [f"slide {i} ({t}): missing {k!r}" for k in REQUIRED[t] if not s.get(k)]
        if t == "chart":
            n = len(s.get("categories") or [])
            for ser in s.get("series") or []:
                vals = ser.get("values") or []
                if len(vals) != n or not all(isinstance(v, (int, float)) for v in vals):
                    errors.append(f"slide {i}: series {ser.get('name')!r} needs {n} numeric values")
        for key in ("image",):
            if s.get(key) and not (base / s[key]).exists():
                errors.append(f"slide {i}: image not found: {s[key]}")
    return errors


# ---------------------------------------------------------------- rendering

PALETTES = {  # dark (title/code bg), accent (dominant), accent2 (support)
    "violet": ("161A24", "6E4BF2", "149E8E"), "midnight": ("1E2761", "3D5AFE", "00897B"),
    "forest": ("1F3A2B", "2C7A3F", "B7791F"), "coral": ("2F3C7E", "E8505B", "D69E2E"),
    "terracotta": ("3A2A26", "B85042", "5F8A72"), "ocean": ("21295C", "065A82", "1C7293"),
    "charcoal": ("212121", "36454F", "D35400"), "teal": ("0B2E33", "028090", "02A37F"),
    "berry": ("2E1520", "6D2E46", "A26769"), "cherry": ("1C1F3B", "990011", "2F3C7E"),
}
WHITE, BODY, MUTED = RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0x2A, 0x2F, 0x3A), RGBColor(0x6B, 0x72, 0x80)
AMBER, AMBER_TINT = RGBColor(0xE8, 0x9A, 0x1C), RGBColor(0xFD, 0xF3, 0xE2)
CODE_FG, SOFT = RGBColor(0xE6, 0xE8, 0xEE), RGBColor(0xA0, 0xA6, 0xB4)
SANS, MONO = "Calibri", "Courier New"
W, H = 13.333, 7.5
LEFT, RIGHT, TOP, BOTTOM = 0.6, 12.733, 1.65, 6.9   # content area
CW = RIGHT - LEFT


def set_cs_font(run) -> None:
    """Complex-script font for the run, so Hindi (Devanagari) shows in a font that has it."""
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn("a:cs")):
        rPr.remove(old)
    rPr.append(rPr.makeelement(qn("a:cs"), {"typeface": CS_FONT}))


def rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr)


def mix(c: RGBColor, amount: float) -> RGBColor:
    """Blend colour towards white; amount=0.9 gives a pale tint."""
    return RGBColor(*(round(v + (255 - v) * amount) for v in c))


def text_height(paras: list[str], w: float, size: float, mono: bool, gap: float) -> float:
    cpl = max(1, int(w * 72 / (size * (0.6 if mono else 0.5))))  # chars per line, conservative
    lines = sum(max(1, math.ceil(len(p) / cpl)) for p in paras)
    return lines * size * 1.2 / 72 + len(paras) * gap / 72


class Deck:
    def __init__(self, spec: dict, base: Path):
        self.spec, self.base, self.warnings = spec, base, []
        dark, accent, accent2 = PALETTES.get(spec.get("palette", "violet"), PALETTES["violet"])
        self.INK, self.ACCENT, self.ACCENT2 = rgb(dark), rgb(accent), rgb(accent2)
        self.TINT = mix(self.ACCENT, 0.9)
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        self.slide_no = self.section_no = 0
        self.anim = []

    # primitives ---------------------------------------------------------
    def text(self, slide, x, y, w, h, paras, size=16, color=None, bold=False, mono=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, min_size=10, gap=4, bullets=False):
        """paras: str or list; each item is str or list of (text, {bold, color}) runs."""
        paras = [paras] if isinstance(paras, (str, tuple)) else list(paras)
        plain = ["".join(r[0] for r in p) if isinstance(p, list) else str(p) for p in paras]
        inner_w = w - (0.3 if bullets else 0)
        fit = size
        while fit > min_size and text_height(plain, inner_w, fit, mono, gap) > h:
            fit -= 0.5
        if text_height(plain, inner_w, fit, mono, gap) > h * 1.05:
            self.warnings.append(f"slide {self.slide_no}: text may overflow: {plain[0][:50]!r}")
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.space_after = Pt(gap)
            for chunk, opts in (para if isinstance(para, list) else [(str(para), {})]):
                r = p.add_run()
                r.text = chunk
                f = r.font
                f.name, f.size = MONO if mono else SANS, Pt(fit)
                set_cs_font(r)
                f.bold = opts.get("bold", bold)
                f.color.rgb = opts.get("color", color or BODY)
            if bullets:
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", str(Inches(0.3)))
                pPr.set("indent", str(-Inches(0.3)))
                clr = pPr.makeelement(qn("a:buClr"), {})
                clr.append(clr.makeelement(qn("a:srgbClr"), {"val": str(self.ACCENT)}))
                pPr.append(clr)
                pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "\u25cf"}))
        return tb

    def box(self, slide, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08):
        s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        s.fill.solid()
        s.fill.fore_color.rgb = fill
        s.line.fill.background()
        s.shadow.inherit = False
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            s.adjustments[0] = radius
        return s

    def badge(self, slide, x, y, label, fill=None, d=0.6, size=18):
        c = self.box(slide, x, y, d, d, fill or self.ACCENT, MSO_SHAPE.OVAL)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(label)[:2]
        r.font.name, r.font.size, r.font.bold, r.font.color.rgb = SANS, Pt(size), True, WHITE

    def code_block(self, slide, x, y, w, h, code, lang=""):
        self.box(slide, x, y, w, h, self.INK, radius=0.04)
        top = 0.2
        if lang:
            self.text(slide, x + 0.3, y + 0.18, w - 0.6, 0.3, lang.upper(), size=10,
                      color=mix(self.ACCENT2, 0.3), bold=True)
            top = 0.5
        self.text(slide, x + 0.3, y + top, w - 0.6, h - top - 0.15, code.rstrip().split("\n"),
                  size=14, color=CODE_FG, mono=True, min_size=9, gap=1)

    def warn(self, slide, x, y, w, h, msg):
        self.box(slide, x, y, w, h, AMBER_TINT)
        self.badge(slide, x + 0.25, y + (h - 0.5) / 2, "!", AMBER, d=0.5, size=20)
        self.text(slide, x + 0.95, y + 0.1, w - 1.15, h - 0.2, [lead(msg)], size=15,
                  anchor=MSO_ANCHOR.MIDDLE, min_size=11)

    def picture(self, slide, path, x, y, w, h):
        path = self.base / path
        with Image.open(path) as im:
            iw, ih = im.size
        scale = min(w / iw, h / ih)
        pw, ph = iw * scale, ih * scale
        slide.shapes.add_picture(str(path), Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2),
                                 Inches(pw), Inches(ph))

    def new_slide(self, s, dark=False):
        self.slide_no += 1
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        if dark:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = self.INK
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]
        self.anim.append([slide, [0], False])  # [slide, group start indexes, skip first group]
        return slide

    def mark(self, slide):
        """Start a new animation group: shapes added after this fade in together."""
        self.anim[-1][1].append(len(slide.shapes))

    def header(self, s, title=None):
        slide = self.new_slide(s)
        if s.get("kicker"):
            self.text(slide, LEFT, 0.35, CW, 0.3, s["kicker"].upper(), size=12, color=self.ACCENT, bold=True)
        self.text(slide, LEFT, 0.65, CW, 0.8, title or s["title"], size=34, color=self.INK,
                  bold=True, min_size=24)
        self.mark(slide)
        self.anim[-1][2] = True  # title stays put; only content animates
        return slide

    # slide types --------------------------------------------------------
    def s_title(self, s):
        slide = self.new_slide(s, dark=True)
        self.text(slide, 0.8, 2.0, 11.7, 1.9, s["title"], size=54, color=WHITE, bold=True,
                  anchor=MSO_ANCHOR.BOTTOM, min_size=32)
        self.box(slide, 0.8, 4.15, 1.2, 0.12, self.ACCENT, MSO_SHAPE.RECTANGLE)
        self.mark(slide)
        if s.get("subtitle"):
            self.text(slide, 0.8, 4.5, 11, 1.4, s["subtitle"], size=22, color=CODE_FG, min_size=14)
        self.mark(slide)
        if s.get("footer"):
            self.text(slide, 0.8, 6.5, 11.5, 0.4, s["footer"], size=14, color=SOFT)

    def s_section(self, s):
        slide = self.new_slide(s, dark=True)
        self.section_no += 1
        self.badge(slide, 0.8, 2.3, s.get("icon") or str(self.section_no), self.ACCENT, d=0.9, size=26)
        self.text(slide, 0.8, 3.4, 11.7, 1.2, s["title"], size=44, color=WHITE, bold=True, min_size=28)
        self.mark(slide)
        if s.get("subtitle"):
            self.text(slide, 0.8, 4.6, 11, 1.2, s["subtitle"], size=20, color=CODE_FG, min_size=14)

    def s_bullets(self, s):
        items, has_side = s["bullets"], bool(s.get("image"))
        per = 5 if has_side else 6
        chunks = split_even(items, per)
        for n, chunk in enumerate(chunks):
            slide = self.header(s, s["title"] + (" (cont.)" if n else ""))
            bottom = BOTTOM - (1.15 if s.get("callout") and n == len(chunks) - 1 else 0)
            w = 6.6 if has_side else CW
            self.text(slide, LEFT, TOP + 0.1, w, bottom - TOP - 0.1, [lead(b) for b in chunk],
                      size=20, gap=14, min_size=12, bullets=True)
            self.mark(slide)
            if has_side:
                self.box(slide, 7.6, TOP, 5.13, bottom - TOP, self.TINT, radius=0.04)
                self.picture(slide, s["image"], 7.75, TOP + 0.15, 4.83, bottom - TOP - 0.3)
            self.mark(slide)
            if s.get("callout") and n == len(chunks) - 1:
                self.warn(slide, LEFT, BOTTOM - 0.95, CW, 0.95, s["callout"])

    def s_cards(self, s):
        cards = s["cards"][:6]
        cols = 2 if len(cards) == 4 else min(len(cards), 3)
        rows = math.ceil(len(cards) / cols)
        gap = 0.3
        cw = (CW - gap * (cols - 1)) / cols
        ch = min((BOTTOM - TOP - gap * (rows - 1)) / rows, 3.6 if rows == 1 else 2.2)
        slide = self.header(s)
        for i, c in enumerate(cards):
            self.mark(slide)
            x, y = LEFT + (i % cols) * (cw + gap), TOP + (i // cols) * (ch + gap)
            self.box(slide, x, y, cw, ch, self.TINT)
            icon_col = self.ACCENT if i % 2 == 0 else self.ACCENT2
            if rows == 1:  # tall cards: icon on top
                self.badge(slide, x + 0.3, y + 0.3, c.get("icon") or i + 1, icon_col, d=0.7, size=20)
                self.text(slide, x + 0.3, y + 1.2, cw - 0.6, 0.8, c.get("title", ""), size=20,
                          color=self.INK, bold=True, min_size=14)
                self.text(slide, x + 0.3, y + 2.05, cw - 0.6, ch - 2.3, c.get("text", ""), size=16, min_size=11)
            else:  # short cards: icon at left
                self.badge(slide, x + 0.3, y + 0.3, c.get("icon") or i + 1, icon_col, d=0.6)
                self.text(slide, x + 1.15, y + 0.25, cw - 1.45, 0.5, c.get("title", ""), size=18,
                          color=self.INK, bold=True, min_size=13)
                self.text(slide, x + 1.15, y + 0.8, cw - 1.45, ch - 1.0, c.get("text", ""), size=14, min_size=10)

    def s_steps(self, s):
        steps = s["steps"]
        chunks = split_even(steps, 5)
        num = 0
        for n, chunk in enumerate(chunks):
            slide = self.header(s, s["title"] + (" (cont.)" if n else ""))
            gap = 0.2
            rh = min(1.4, (BOTTOM - TOP - gap * (len(chunk) - 1)) / len(chunk))
            for i, st in enumerate(chunk):
                self.mark(slide)
                num += 1
                y = TOP + i * (rh + gap)
                self.box(slide, LEFT, y, CW, rh, self.TINT)
                self.badge(slide, LEFT + 0.3, y + (rh - 0.6) / 2, num)
                self.text(slide, LEFT + 1.2, y + 0.12, 3.6, rh - 0.24, st.get("title", ""), size=18,
                          color=self.INK, bold=True, anchor=MSO_ANCHOR.MIDDLE, min_size=13)
                self.text(slide, LEFT + 5.0, y + 0.12, CW - 5.3, rh - 0.24, st.get("text", ""), size=15,
                          anchor=MSO_ANCHOR.MIDDLE, min_size=10)

    def s_stats(self, s):
        stats = s["stats"][:4]
        slide = self.header(s)
        gap = 0.3
        cw = (CW - gap * (len(stats) - 1)) / len(stats)
        ch = 2.6 if s.get("text") else BOTTOM - TOP
        for i, st in enumerate(stats):
            self.mark(slide)
            x = LEFT + i * (cw + gap)
            self.box(slide, x, TOP, cw, ch, self.TINT)
            self.text(slide, x + 0.3, TOP + 0.3, cw - 0.6, 1.3, str(st.get("value", "")), size=60,
                      color=self.ACCENT if i % 2 == 0 else self.ACCENT2, bold=True,
                      anchor=MSO_ANCHOR.BOTTOM, min_size=28)
            self.text(slide, x + 0.3, TOP + 1.7, cw - 0.6, ch - 1.9, st.get("label", ""), size=16, min_size=11)
        self.mark(slide)
        if s.get("text"):
            self.text(slide, LEFT, TOP + ch + 0.4, CW, BOTTOM - TOP - ch - 0.4, [lead(s["text"])],
                      size=18, min_size=12)

    def s_code(self, s):
        slide = self.header(s)
        y = TOP
        self.mark(slide)
        if s.get("note"):
            self.warn(slide, LEFT, y, CW, 0.95, s["note"])
            y += 1.2
        lines = s["code"].rstrip().split("\n")
        h = min(BOTTOM - y, 0.85 + text_height(lines, CW - 0.6, 14, True, 1))  # counts wrapped lines
        self.mark(slide)
        self.code_block(slide, LEFT, y, CW, max(h, 1.1), s["code"], s.get("lang", ""))

    def s_table(self, s):
        header, rows = s["header"], s["rows"]
        chunks = split_even(rows, 8) or [[]]
        lens = [max(len(str(r[c])) if c < len(r) else 0 for r in [header] + rows) for c in range(len(header))]
        weights = [min(max(l, 6), 40) for l in lens]
        size = 15 if len(header) <= 3 else 13 if len(header) <= 5 else 11
        for n, chunk in enumerate(chunks):
            slide = self.header(s, s["title"] + (" (cont.)" if n else ""))
            data = [header] + chunk
            row_h = min(0.6, (BOTTOM - TOP) / len(data))
            tbl = slide.shapes.add_table(len(data), len(header), Inches(LEFT), Inches(TOP),
                                         Inches(CW), Inches(row_h * len(data))).table
            for c, wt in enumerate(weights):
                tbl.columns[c].width = Emu(int(Inches(CW) * wt / sum(weights)))
            for r, row in enumerate(data):
                for c in range(len(header)):
                    cell = tbl.cell(r, c)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = self.INK if r == 0 else (self.TINT if r % 2 else WHITE)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    cell.margin_left = cell.margin_right = Inches(0.15)
                    run = cell.text_frame.paragraphs[0].add_run()
                    run.text = str(row[c]) if c < len(row) else ""
                    run.font.size, run.font.name = Pt(size), SANS
                    set_cs_font(run)
                    run.font.bold = r == 0
                    run.font.color.rgb = WHITE if r == 0 else BODY

    def s_chart(self, s):
        slide = self.header(s)
        kind = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                "line": XL_CHART_TYPE.LINE_MARKERS, "pie": XL_CHART_TYPE.PIE}.get(
                    s.get("chart", "column"), XL_CHART_TYPE.COLUMN_CLUSTERED)
        data = CategoryChartData()
        data.categories = [str(c) for c in s["categories"]]
        for ser in s["series"]:
            data.add_series(ser.get("name", ""), ser["values"])
        cw = 8.0 if s.get("text") else CW
        chart = slide.shapes.add_chart(kind, Inches(LEFT), Inches(TOP), Inches(cw),
                                       Inches(BOTTOM - TOP), data).chart
        chart.font.size, chart.font.name = Pt(12), SANS
        palette = [self.ACCENT, self.ACCENT2, self.INK, AMBER, mix(self.ACCENT, 0.45), MUTED]
        plot = chart.plots[0]
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(11)
        pie = kind == XL_CHART_TYPE.PIE
        chart.has_legend = pie or len(s["series"]) > 1
        if chart.has_legend:
            chart.legend.position, chart.legend.include_in_layout = XL_LEGEND_POSITION.BOTTOM, False
        if pie:
            for i, pt in enumerate(plot.series[0].points):
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = palette[i % len(palette)]
        else:
            if hasattr(plot, "gap_width"):
                plot.gap_width = 80
            for i, ser in enumerate(plot.series):
                col = palette[i % len(palette)]
                if kind == XL_CHART_TYPE.LINE_MARKERS:
                    ser.format.line.color.rgb = col
                    ser.format.line.width = Pt(2.5)
                else:
                    ser.format.fill.solid()
                    ser.format.fill.fore_color.rgb = col
            chart.value_axis.has_major_gridlines = True
            chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xE5, 0xE7, 0xEB)
            chart.value_axis.format.line.fill.background()
            if kind == XL_CHART_TYPE.BAR_CLUSTERED:  # PowerPoint draws bar categories bottom-up
                chart.category_axis.reverse_order = True
        self.mark(slide)
        if s.get("text"):
            self.box(slide, 8.9, TOP, RIGHT - 8.9, BOTTOM - TOP, self.TINT)
            self.text(slide, 9.2, TOP + 0.3, RIGHT - 9.5, BOTTOM - TOP - 0.6, [lead(s["text"])],
                      size=17, min_size=11)

    def s_image(self, s):
        slide = self.header(s)
        side = bool(s.get("bullets"))
        w = 7.6 if side else CW
        cap_h = 0.5 if s.get("caption") else 0
        self.box(slide, LEFT, TOP, w, BOTTOM - TOP - cap_h, self.TINT, radius=0.04)
        self.picture(slide, s["image"], LEFT + 0.2, TOP + 0.2, w - 0.4, BOTTOM - TOP - cap_h - 0.4)
        if s.get("caption"):
            self.text(slide, LEFT, BOTTOM - 0.4, w, 0.4, s["caption"], size=12, color=MUTED,
                      align=PP_ALIGN.CENTER)
        self.mark(slide)
        if side:
            self.text(slide, LEFT + w + 0.4, TOP + 0.1, CW - w - 0.4, BOTTOM - TOP - 0.1,
                      [lead(b) for b in s["bullets"]], size=18, gap=12, min_size=11, bullets=True)

    def s_closing(self, s):
        slide = self.new_slide(s, dark=True)
        items = s.get("items", [])[:6]
        self.text(slide, 0.8, 0.7, 11.7, 0.9, s["title"], size=40, color=WHITE, bold=True, min_size=26)
        step = min(0.95, 4.9 / max(len(items), 1))
        for i, item in enumerate(items):
            self.mark(slide)
            y = 1.95 + i * step
            self.badge(slide, 0.8, y, "\u2713", self.ACCENT2, d=0.6, size=20)
            self.text(slide, 1.65, y, 10.8, 0.6, item, size=22, color=WHITE,
                      anchor=MSO_ANCHOR.MIDDLE, min_size=13)

    def build(self, out: Path, animate: bool = True) -> Path:
        for s in self.spec["slides"]:
            getattr(self, "s_" + s["type"])(s)
        if animate and self.spec.get("animate", True):
            for slide, marks, skip in self.anim:
                shapes = list(slide.shapes)
                bounds = list(zip(marks, marks[1:] + [len(shapes)]))[1 if skip else 0:]
                add_animation(slide, [shapes[a:b] for a, b in bounds if b > a])
        self.prs.core_properties.title = self.spec.get("title", "")
        self.prs.core_properties.comments = self.spec.get(
            "made_with", "Built by make_ppt.py (offline) from a hand-written spec; no AI model involved.")
        out.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.prs.save(out)
        except PermissionError:
            sys.exit(f"Cannot write {out} - is it open in PowerPoint?")
        return out


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
FADE_MS, STAGGER_MS = 500, 300


def add_animation(slide, groups: list[list]) -> None:
    """Fade transition, plus each group fading in automatically, one after another.

    python-pptx has no animation API, so this writes the <p:transition> and <p:timing>
    XML that PowerPoint itself produces for an "After Previous" fade entrance.
    """
    from lxml import etree
    sld = slide._element
    for tag in ("transition", "timing"):
        for old in sld.findall(qn(f"p:{tag}")):
            sld.remove(old)
    anchor = sld.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = sld.find(qn("p:cSld"))
    transition = etree.fromstring(
        f'<p:transition xmlns:p="{P_NS}" spd="med"><p:fade/></p:transition>')
    anchor.addnext(transition)
    if not groups:
        return
    ids = iter(range(4, 100_000))
    effects, builds = [], []
    for g, shapes in enumerate(groups):
        inner = []
        for shape in shapes:
            spid = shape.shape_id
            a, b, c = next(ids), next(ids), next(ids)
            node = "afterEffect" if g == 0 and not inner else "withEffect"
            tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            inner.append(
                f'<p:par><p:cTn id="{a}" presetID="10" presetClass="entr" presetSubtype="0" '
                f'fill="hold" grpId="0" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="{g * STAGGER_MS}"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/>'
                f'</p:stCondLst></p:cTn>{tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName>'
                f'</p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{c}" dur="{FADE_MS}"/>'
                f'{tgt}</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
            if shape.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX):
                builds.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        effects += inner
    bld = f'<p:bldLst>{"".join(builds)}</p:bldLst>' if builds else ""
    timing = etree.fromstring(
        f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f'<p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
        f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
        f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par>'
        f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
        f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
        f'</p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'{bld}</p:timing>')
    transition.addnext(timing)


def split_even(items: list, per: int) -> list[list]:
    """Split into the fewest chunks of <= per items, balanced (7 -> 4+3, not 6+1)."""
    if not items:
        return []
    n = math.ceil(len(items) / per)
    size = math.ceil(len(items) / n)
    return [items[i:i + size] for i in range(0, len(items), size)]


def lead(s: str) -> list:
    """'Lead phrase: rest' -> bold lead run + normal run."""
    head, sep, rest = str(s).partition(": ")
    if sep and len(head) <= 45 and rest:
        return [(head + ": ", {"bold": True}), (rest, {})]
    return [(str(s), {})]


def render_previews(deck: Path, outdir: Path) -> list[Path]:
    """PPTX -> PDF (LibreOffice) -> one PNG per slide + a contact sheet."""
    import pypdfium2
    if not Path(SOFFICE).exists():
        log.warning("LibreOffice not found at %s - skipping previews", SOFFICE)
        return []
    if outdir.exists():
        shutil.rmtree(outdir)
    outdir.mkdir(parents=True)
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "deck.pptx"
        shutil.copy2(deck, src)
        subprocess.run([SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, capture_output=True, timeout=300)
        pdf = pypdfium2.PdfDocument(Path(tmp) / "deck.pdf")
        pngs = []
        for i in range(len(pdf)):
            p = outdir / f"slide-{i + 1:02d}.png"
            pdf[i].render(scale=1.2).to_pil().save(p)
            pngs.append(p)
        pdf.close()
    thumbs = [Image.open(p) for p in pngs]
    tw, th = thumbs[0].width // 2, thumbs[0].height // 2
    sheet = Image.new("RGB", (tw * 2, th * math.ceil(len(thumbs) / 2)), "white")
    for i, im in enumerate(thumbs):
        sheet.paste(im.resize((tw, th)), ((i % 2) * tw, (i // 2) * th))
    sheet.save(outdir / "contact-sheet.png")
    return pngs


def build(spec_path: Path, out: Path | None, preview: bool = True, animate: bool = True) -> Path:
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    base = spec_path.parent
    if errors := validate(spec, base):
        sys.exit("Spec errors:\n  " + "\n  ".join(errors))
    out = out or base / f"{safe_name(spec.get('title') or spec_path.stem)}.pptx"
    deck = Deck(spec, base)
    deck.build(out, animate)
    print(f"Deck: {out} ({deck.slide_no} slides)")
    for w in deck.warnings:
        print(f"  warning: {w}")
    if preview:
        pngs = render_previews(out, base / "preview")
        if pngs:
            print(f"Previews: {pngs[0].parent} (contact-sheet.png shows all slides)")
    return out


def safe_name(s: str) -> str:
    return "".join(c for c in s if c not in '<>:"/\\|?*').strip()[:80] or "deck"


# ---------------------------------------------------------------- auto (Ollama)

PROMPT = """You turn source documents into a clear, well-structured slide deck.

Reply with ONLY a JSON object in this format:
{schema}

Rules:
- About {slides} slides. Start with a "title" slide, end with a "closing" slide.
- Use at most one "section" slide per 5 slides, and none in decks under 10 slides.
- Vary slide types. Prefer cards, steps, stats, tables and charts over plain bullets.
- Max 6 bullets per slide, max 20 words each. A bullet may start with a short bold label and a colon, e.g. "Budget: 12 lakh".
- Put the detail that does not fit on the slide into "notes".
- Use only facts from the source. Do not invent numbers.
- If the source contains commands or code, include them in "code" slides, copied exactly.
  Commands are usually the most important content of a technical guide - never drop them.
- Only make a "chart" when the source has a table of numbers. Never invent chart data.
- Only use image paths listed in the source as ![image](path), copied exactly.
- Use the "Image description" lines to write titles, captions and bullets for images.
  When the source is mostly images, give each important image its own slide.
- Write the slides in the same language as the source (Hindi source -> Hindi slides),
  unless the instructions ask for another language. Keep commands and numbers as they are.
- Pick the palette that best suits the topic.
{extra}
SOURCE DOCUMENTS:
{content}"""


def ask_ollama(model: str, prompt: str) -> str:
    # Context sized to the prompt (~3.5 chars/token) + room for the reply, in 8k steps.
    # A fixed 32k window reserves several GB of memory even for a short document.
    num_ctx = min(32768, math.ceil((len(prompt) / 3.5 + 4096) / 8192) * 8192)
    body = json.dumps({"model": model, "stream": False, "format": "json",
                       "messages": [{"role": "user", "content": prompt}],
                       "options": {"temperature": 0.3, "num_ctx": num_ctx}}).encode()
    req = urllib.request.Request(f"{OLLAMA}/api/chat", body, {"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=1800) as r:
            return json.loads(r.read())["message"]["content"]
    except urllib.error.HTTPError as e:  # Ollama is up but refused, e.g. not enough memory
        try:
            reason = json.loads(e.read()).get("error", str(e))
        except Exception:
            reason = str(e)
        hint = (" Close other heavy apps, or pick a smaller model with -m."
                if "memory" in reason else "")
        sys.exit(f"Ollama error: {reason}.{hint}")
    except urllib.error.URLError as e:
        sys.exit(f"Cannot reach Ollama at {OLLAMA} ({e}). Is `ollama serve` running?")


CODE_HINTS = ("$ ", "PS>", ".exe", "--", "sudo ", "pip ", "npm ", "curl ", "git ", "cd ", "./", ".\\")


def content_checks(spec: dict, source: str) -> list[str]:
    """Catch the two mistakes local models make most: invented charts and dropped commands."""
    errors = []
    numbers = set(re.findall(r"\d+(?:\.\d+)?", source.replace(",", "")))
    for i, s in enumerate(spec.get("slides", []), 1):
        if s.get("type") != "chart":
            continue
        vals = [v for ser in s.get("series") or [] for v in ser.get("values") or []]
        missing = [v for v in vals if f"{v:g}" not in numbers]
        if missing or len(set(vals)) < 2:
            errors.append(f"slide {i}: chart values {missing or vals} are not real data from the source; remove this chart or use a table/cards instead")
    code_lines = [l for l in source.splitlines() if any(h in l for h in CODE_HINTS)]
    if len(code_lines) >= 3 and not any(s.get("type") == "code" for s in spec.get("slides", [])):
        errors.append("the source contains commands (e.g. " + repr(code_lines[0].strip()[:60]) + ") but there is no "
                      "\"code\" slide; add code slides with the commands copied exactly")
    slides = spec.get("slides", [])
    n_sec = sum(s.get("type") == "section" for s in slides)
    if n_sec > max(len(slides) // 5, 0 if len(slides) < 10 else 1):
        errors.append(f"{n_sec} section slides is too many for {len(slides)} slides; remove section slides and keep the content slides")
    return errors


CHART_WORDS = re.compile(r"\b(chart|graph|plot|bar|axis|axes)\b", re.I)


def add_missing_images(spec: dict, files: list[Path], source: str) -> list[str]:
    """Make sure every picture the user gave ends up in the deck.

    Small models often skip some. Each skipped input image gets its own image slide
    (before the closing slide), captioned from its vision description. A chart picture
    is not re-added when the model already rebuilt it as a native chart slide.
    """
    slides = spec["slides"]
    used = {Path(s.get("image", "")).name for s in slides}
    has_chart = any(s.get("type") == "chart" for s in slides)
    descriptions = dict(re.findall(
        r"!\[image\]\(assets/([^)]+)\)\s*\n\nImage description \(vision model\): ([^\n]+)", source))
    added = []
    for f in files:
        if f.suffix.lower() not in IMAGE_EXT or f.name in used:
            continue
        desc = descriptions.get(f.name, "")
        if has_chart and CHART_WORDS.search(desc):
            continue
        first = re.split(r"(?<=[.!?])\s", desc, maxsplit=1)[0] if desc else ""
        first = re.sub(r"^(The|This) image (depicts|shows|presents|is|features)\s+", "", first)
        first = first[:1].upper() + first[1:]
        slide = {"type": "image", "title": f.stem.replace("_", " ").replace("-", " ").title(),
                 "image": f"assets/{f.name}", "notes": desc}
        if first:
            slide["caption"] = first[:150]
        at = len(slides) - 1 if slides and slides[-1].get("type") == "closing" else len(slides)
        slides.insert(at, slide)
        added.append(f.name)
    return added


def auto(files: list[Path], out: Path | None, model: str, slides: int, extra: str,
         preview: bool, animate: bool = True, vision: str | None = "auto") -> Path:
    if model.endswith(":cloud") or "-cloud" in model:
        log.warning("%s is an Ollama cloud model - this run will NOT be offline", model)
    first = files[0]
    work = first.parent / f"{first.stem}_ppt"
    md = extract_all(files, work, vision)
    content = md.read_text(encoding="utf-8")
    if len(content) > 60_000:
        log.warning("source is %d chars; truncating to 60000 for the model", len(content))
        content = content[:60_000]
    prompt = PROMPT.format(schema=SCHEMA, slides=slides, content=content,
                           extra=f"- {extra}\n" if extra else "")
    spec_path = work / "spec.json"
    for attempt in range(3):
        log.info("asking %s for a slide spec (attempt %d) - this can take a few minutes", model, attempt + 1)
        reply = ask_ollama(model, prompt)
        try:
            spec = json.loads(reply)
            errors = validate(spec, work) + content_checks(spec, content)
        except json.JSONDecodeError as e:
            spec, errors = None, [f"reply was not valid JSON: {e}"]
        if spec is not None:
            spec_path.write_text(json.dumps(spec, indent=2, ensure_ascii=False), encoding="utf-8")
        if not errors:
            break
        log.warning("spec problems: %s", "; ".join(errors))
        prompt += ("\n\nYour previous reply had these problems:\n- " + "\n- ".join(errors) +
                   "\nReply again with the full corrected JSON.")
    else:
        sys.exit(f"Model could not produce a valid spec. Last attempt saved to {spec_path}; "
                 f"fix it by hand and run: make_ppt.py build \"{spec_path}\"")
    slides = spec["slides"]  # drop section dividers with nothing after them
    spec["slides"] = [s for i, s in enumerate(slides) if not (
        s.get("type") == "section" and
        (i == len(slides) - 1 or slides[i + 1].get("type") in ("section", "closing")))]
    if added := add_missing_images(spec, files, content):
        log.info("added slides for images the model skipped: %s", ", ".join(added))
    spec_path.write_text(json.dumps(spec, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Spec: {spec_path} (edit it and run `build` to tweak the deck)")
    offline = not (model.endswith(":cloud") or "-cloud" in model)
    spec["made_with"] = (
        f"Made {'OFFLINE' if offline else 'with a CLOUD model (NOT offline)'} by make_ppt.py on "
        f"{__import__('datetime').date.today()}. Slide content: {model} (Ollama). "
        f"Image descriptions: {', '.join(sorted(VISION_USED)) or 'none'}. "
        f"OCR: {('Tesseract ' + ', '.join(sorted(OCR_USED))) if OCR_USED else 'not needed'}. "
        "File reading, layout and rendering: local Python + LibreOffice.")
    spec_path.write_text(json.dumps(spec, indent=2, ensure_ascii=False), encoding="utf-8")
    return build(spec_path, out or first.parent / f"{first.stem}.pptx", preview, animate)


# ---------------------------------------------------------------- CLI

def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("extract", help="dump documents to Markdown + images")
    e.add_argument("files", nargs="+", type=Path)
    e.add_argument("-o", "--outdir", type=Path, help="default: <first file>_ppt/ next to it")
    e.add_argument("--vision", default="auto", help="vision model for describing images (default: best installed)")
    e.add_argument("--no-vision", action="store_true", help="do not describe images")
    b = sub.add_parser("build", help="render a JSON spec to .pptx")
    b.add_argument("spec", type=Path)
    b.add_argument("-o", "--out", type=Path)
    b.add_argument("--no-preview", action="store_true")
    b.add_argument("--no-animate", action="store_true", help="no fade transitions/animations")
    a = sub.add_parser("auto", help="documents -> local LLM -> .pptx")
    a.add_argument("files", nargs="+", type=Path)
    a.add_argument("-o", "--out", type=Path)
    a.add_argument("-m", "--model", default=DEFAULT_MODEL)
    a.add_argument("--vision", default="auto", help="vision model for describing images (default: best installed)")
    a.add_argument("--no-vision", action="store_true", help="do not describe images")
    a.add_argument("-n", "--slides", type=int, default=10)
    a.add_argument("-i", "--instructions", default="", help='e.g. "audience: management, focus on costs"')
    a.add_argument("--no-preview", action="store_true")
    a.add_argument("--no-animate", action="store_true", help="no fade transitions/animations")
    sub.add_parser("schema", help="print the slide spec format")
    args = ap.parse_args()

    if args.cmd == "extract":
        md = extract_all(args.files, args.outdir or args.files[0].parent / f"{args.files[0].stem}_ppt",
                         None if args.no_vision else args.vision)
        print(f"Extract: {md}")
    elif args.cmd == "build":
        build(args.spec, args.out, not args.no_preview, not args.no_animate)
    elif args.cmd == "auto":
        auto(args.files, args.out, args.model, args.slides, args.instructions, not args.no_preview,
             not args.no_animate, None if args.no_vision else args.vision)
    else:
        print(SCHEMA)


if __name__ == "__main__":
    main()
