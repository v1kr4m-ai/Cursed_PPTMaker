## 2026-09-29 — Offline PPT toolchain setup
- Installed: pdfplumber, xlrd, matplotlib, pytesseract (pip); LibreOffice 26.8.0 (winget)
- Already present: python-pptx, python-docx, openpyxl, pandas, Pillow, pypdf, pypdfium2, Tesseract 5.4.0
- Pinned list in requirements-ppt.txt
- Smoke test passed: docx/xlsx/PDF/image(OCR) -> PPTX -> rendered PNG, fully offline
- Gotchas: Tesseract + soffice not on PATH, use full paths; Tesseract has eng only
- Next: build actual converter when user supplies files

## 2026-10-03 — First PDF -> PPTX conversion
- Source: Documents\Tech Docs\Tutorial for the Linking of Ollama an....pdf (3 pages)
- Output: Documents\Tech Docs\Ollama to LM Studio Sync Guide.pptx (8 slides, 16:9, speaker notes)
- Built with python-pptx (pptxgenjs not installed globally; stayed on the offline Python stack)
- Rendered via LibreOffice + pdftoppm; fixed code-block sizing after visual QA

## 2026-10-03 — make_ppt.py
- Added make_ppt.py: extract / build / auto / schema. 11 slide types, auto-fit text, PNG previews.
- Tested extract on .doc .xls .png, scanned PDF (OCR) and text PDF; built a 16-slide test deck covering every type.
- Offline auto run with qwen3-coder:30b on the tutorial PDF: ~1.5-2 min, valid deck.
- Local-model failure modes found and guarded in content_checks(): invented chart data,
  dropped commands, too many section slides. Failed checks are fed back to the model (3 attempts).
- Known limit: local model still picks thinner content than a hosted model (e.g. only the
  Windows command, not macOS/Linux). Fix by editing <file>_ppt/spec.json and running build.
- Not done: slide animations (python-pptx has no API; would need raw XML).

## 2026-10-03 — make_ppt.py animations
- Fade transition on every slide + staggered auto fade-in of content groups (cards, steps,
  stats, code, charts, checklist items), 0.5s fade, 0.3s apart. Titles stay static.
- Written as raw <p:transition>/<p:timing> XML (python-pptx has no API). Verified by opening
  in real PowerPoint via COM: no repair prompt, all effects After Previous with correct delays.
- Gotcha: PowerPoint ignores delays on wrapper <p:par> groups; delay must sit on each effect's cTn.
- --no-animate / "animate": false to disable. examples/ holds a demo spec + deck.

## 2026-10-03 — Windows launcher
- "Make PPT.bat" -> make_ppt_launcher.ps1: file picker or drag-drop, starts Ollama if down,
  checks model installed, asks model/slides/focus, writes "<name> - slides.pptx" (never
  overwrites), opens it. Dropping a spec.json rebuilds without the model.
- Desktop shortcut "Make PPT.lnk" created.
- make_ppt.py: num_ctx now sized to prompt (was fixed 32k; saved ~2.4 GB on qwen3-coder:30b);
  Ollama HTTP errors (e.g. not enough memory) now shown with the real reason.
- Memory note: qwen3-coder:30b needs ~16 GB free RAM+VRAM; failed when other apps held memory.
  qwen3:8b fits but took ~6.5 min (thinking model).

## 2026-10-03 — Launcher: model menu + Word/Excel verified
- Word + Excel end to end via launcher (qwen3:8b, 2:45): Excel -> native chart with real values,
  Word -> stat cards + embedded image. Small model still mislabels/invents a little.
- Model menu: numbered list of offline models (cloud/embedding hidden), sizes, free-memory
  estimate (RAM + free VRAM + memory held by models Ollama already has loaded), "may not fit"
  flag; accepts number or name; last pick saved to .last_model (gitignored).

## 2026-10-03 — Images: vision descriptions
- extract/auto describe every image (inputs + images inside docs, max 12) with a local vision
  model via Ollama /api/generate (keep_alive 0 so the slide model gets the memory back).
  Picks llama3.2-vision > qwen2.5vl > gemma3 > minicpm-v > llava > moondream; on a memory
  error falls back to the next one. --vision MODEL / --no-vision.
- Input images the model skips get an auto image slide (caption = first sentence of the
  description); a chart picture is not re-added if a native chart slide exists.
- Also: horizontal bar charts now list categories top-down; empty trailing section slides dropped;
  all input files checked before extraction starts.
- Test: house/flowchart/chart pictures -> 8-slide deck in 2.5 min (qwen3:8b + llama3.2-vision),
  chart values read correctly from the picture.

## 2026-10-03 — Hindi scanned PDFs + provenance stamp
- tessdata/ (eng, osd, hin from tessdata_best, 11.9 MB) next to the script; OCR picks the
  language per page via Tesseract OSD: Devanagari -> hin+eng, else eng.
- Runs get complex-script font Nirmala UI so Devanagari renders; prompt keeps the source language.
- Deck Comments property records models used (slide, vision, OCR) and OFFLINE/CLOUD.
- Scanned pages no longer export their own page scan as an "image".
- Test: scanned Hindi PDF -> Hindi deck in 3:19 (qwen3:8b). OCR misread 12 -> 72 lakh:
  numbers from scans must be checked by a human.

## 2026-10-03 — Moved to its own project
- Moved out of D:\Claude\Claude_Ollama (instruction pack) to
  D:\Claude\In_Progress\03_Windows_Desktop_Apps\Cursed_PPTMaker. requirements-ppt.txt -> requirements.txt.
- Desktop shortcut "Make PPT" repointed. Claude_Ollama README restored to original.

## 2026-10-04 — Desktop app (pywebview)
- app.py + ui/index.html, started by "Cursed PPTMaker.bat" (pythonw, no console). Desktop shortcut
  "Cursed PPTMaker.lnk". Style follows a frosted-glass/neo-tactile reference: light glass panel,
  soft pills, blue glowing primary button/toggles/slider, cyan spinning ring on the active step.
- Features: drag-drop (pywebviewFullPath via Python DOM drop handler) or browse; model menu with
  sizes + fits/may-not-fit; slides slider; focus; animation/vision toggles; Save-to folder
  (remembered in .last_outdir); live 4-step progress from make_ppt log lines; result card with
  Open/Folder/Edit spec/Rebuild, offline stamp, slide thumbnails with lightbox.
- make_ppt.auto gained workdir= so the working folder follows the chosen save folder.
- Gotchas: pywebview introspects every public attribute of js_api - keep window/state private
  (_window) or it recurses into WinForms objects. Two WebView2 windows sharing the data folder
  can freeze blank -> single-instance mutex. PrintWindow can't capture WebView2; use screen capture.
- Tested end to end by driving the real window with evaluate_js (Word+Excel -> deck, rebuild,
  custom save folder). README screenshots of the app not added: captures included personal
  notifications/paths.
