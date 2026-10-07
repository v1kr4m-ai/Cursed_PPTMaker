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

## 2026-10-04 — Fix: crash on malformed model output
- Bug: "TypeError: unsupported operand type(s) for /: 'WindowsPath' and 'dict'" when a local
  model wrote "image" as an object ({"path": ..., "caption": ...}); validate() did base / dict.
- normalize() now runs inside validate(): image object/list -> path string (+caption), single
  bullet string -> list, bullet/item/header/category/cell objects -> text, cards/steps/stats as
  strings -> objects, numeric strings in chart values ("1,200", "15%") -> numbers. Anything still
  wrong becomes a normal validation error sent back to the model instead of a crash.
- Tested with a deliberately messy spec (8 slide types) and the example spec.

## 2026-10-04 — Fix: TimeoutError after a 30-minute hang
- Ollama log showed attempt 1 done in 46 s, attempt 2 running 30m0s until our timeout: the small
  model looped in JSON mode (no reply cap). socket TimeoutError isn't a URLError, so it escaped as
  "TimeoutError: timed out".
- ask_ollama: num_predict = 6144 (MAX_REPLY_TOKENS) + repeat_penalty 1.1; num_ctx sized for it;
  TimeoutError -> friendly message. Vision replies capped at 400 tokens.
- app: model-list call guarded against timeouts while Ollama loads; elapsed-time counter while
  a job runs so long waits don't look frozen.
- Tested: timeout path message; normal llama3.2 run Word+Excel -> 7 slides in ~2 min.

## 2026-10-04 — Online models (optional)
- providers.py: any OpenAI-compatible API (presets: OpenAI, Gemini, OpenRouter, Groq, Anthropic,
  custom). Keys DPAPI-encrypted in %APPDATA%\CursedPPTMaker\providers.json (outside the repo).
  JSON mode with fallback when unsupported; fenced replies unwrapped; readable errors for bad key,
  rate limit, offline. Model id format: api:<provider-id>:<model>.
- make_ppt: ask_llm routes api:* to providers, else Ollama; is_online/model_label; online runs
  warn and the deck stamp says "Made ONLINE ...". Ollama cloud 401 -> "run ollama signin" hint.
- normalize(): missing slide titles get a placeholder (small models drop them via the API route).
- App: model menu grouped (On this PC / Ollama cloud / API providers), amber online markers and
  badge, privacy hint; "+ Add online model..." dialog (preset, key, Load list, Save & use, Remove).
  Console launcher stays offline-only.
- Tested without sending data out: a custom provider pointed at local Ollama's /v1 endpoint ->
  Word+Excel deck in 32 s, stamp "Made ONLINE". Real cloud/API providers not exercised (would send
  data externally / need the user's keys).

## 2026-10-04 — Choosable image (vision) model, incl. online
- UI: "Image model" picker replaces the Describe-images switch: Auto (best local), Off, local
  vision models, Ollama cloud vision models, API models (marked "must accept images").
  Vision support detected via Ollama /api/show "capabilities" (name heuristic fallback).
  Online picks: amber marker, "your pictures will be sent to X", badge combines both models.
  Remembered in .last_vision (gitignored).
- make_ppt: describe_image routes api:* to providers.describe_image (OpenAI image_url data URL);
  fallbacks only ever go to local vision models; text-only replies like "I'm unable to view
  images" are detected (NO_SIGHT) and treated as failure -> fallback. Stamp names the image model
  and says ONLINE if pictures were sent out.
- UI: panel split into fixed frame + scrolling .panel-body (cyan rim no longer scrolls over
  controls); dropdowns open upward / shrink to fit the panel.
- Tested offline: moondream explicit (UI run, stamp + log correct), API route via local Ollama /v1
  with llava, fallback from a text-only API model to llama3.2-vision; menus measured in a hidden
  window (both fully inside the panel). Restored the user's remembered phi4 / Auto afterwards.

## 2026-10-04 — Retired Ollama cloud models
- User hit "glm-5 was retired at 2026-07-15": the local :cloud tag stays in `ollama list` after
  the cloud side retires it. ask_ollama now recognises "retired"/HTTP 410, explains it, suggests
  `ollama rm <model>`, and records the model in .retired_models (gitignored); the app hides those
  from both menus. Verified with glm-5:cloud.

## 2026-10-04 — Fix: PermissionError on the preview folder
- User hit "[WinError 5] Access is denied: ...\<name>_ppt\preview". render_previews did
  rmtree(preview) + mkdir immediately: on Windows the delete can still be pending (or Explorer /
  a viewer holds the folder), so mkdir/rmtree fails.
- Now reuses the folder (mkdir exist_ok) and deletes only old slide-*.png / contact-sheet.png;
  build() treats preview failures as a warning - the deck is already saved.
- Tested: back-to-back rebuilds; a preview PNG held open -> deck saved + clear warning, exit 0.

## 2026-10-04 — README + previews refreshed
- App screenshots (docs/images/app-result.png, app-running.png) rendered off-screen with headless
  Edge from demo copies of ui/index.html with a mocked pywebview API and made-up data, so no
  personal windows, notifications or paths can appear. (Gotchas: plain --headless works; one
  --user-data-dir per run or later runs silently hand off; --force-device-scale-factor broke it.)
- Example slide previews regenerated from the current renderer.
- README: app screenshots, online/image-model features, troubleshooting rows for preview folder,
  retired cloud models, API key/quota, image-model fallback, 30-minute timeout.

## 2026-10-04 — Designs from pictures (PPT_Designs/)
- designs.py: picture -> theme. Background = most common border colour; accents = clusters of
  strongly saturated pixels only (median-cut on vivid pixels), so small coloured icons/bars aren't
  averaged into grey; monochrome designs fall back to their darkest tones. Light grey low-saturation
  canvas -> "soft" (raised cards). Title/closing dark = neutral near-black from the image, else a
  deep shade of the background (darkest pixel is often a photo/shadow detail - was brown).
- make_ppt: spec "theme" overrides palette: slide bg, cards, headings/body text (flip on dark
  themes), tables, chart text/gridlines/series, captions; soft themes add an outerShdw to cards;
  badge text goes dark on light accents. --design for auto/build; apply_design() stores theme.
- App: Design row (Default + thumbnails + "+" opens folder, reloads on window focus), swatch hint,
  remembered in .last_design. PPT_Designs/* git-ignored except its README (third-party images).
- Fixed on the way: "image": null from the model crashed the image safety net; the provenance
  stamp exceeded PowerPoint's 255-char Comments limit with a design name -> compact " | " format
  + hard cap; big-number cards no longer stretch to full height.
- Tested: themes extracted for all 17 designs (swatch sheet checked by eye); example deck in designs
  4/5/10/13/14/17; full app run (hidden window) Word+Excel with design 5. Restored the user's
  remembered phi4 / Auto / Default afterwards.

## 2026-10-04 — Design families (replaces colours-from-pictures)
- User clarified: study the reference designs (shapes, fonts, placement) and build designs
  organically, not copy colours. Removed designs.py and the picture-theme path.
- DESIGN_NOTES.md: per-reference analysis of 17 designs + recurring ideas (numbers as graphics,
  radial structures, geometric containers, soft depth, split compositions, small decorative
  systems, type pairing) -> six families.
- families.py: FAMILIES (palette, fonts, flags, variant choices) + FamilyLayouts mixin with
  covers (t_soft/ribbon/hub/split/wander/band), dividers (sec_bignum/soft/split/wander),
  steps (ribbons/chevrons/hexsteps), cards (diamonds/hexagons/circles), stats (ring gauges),
  light closing. Shapes: BLOCK_ARC (angles = deg*0.6 in python-pptx), DONUT, HEXAGON, DIAMOND,
  CHEVRON/PENTAGON, FLOWCHART_OFFPAGE_CONNECTOR (banner), RIGHT_TRIANGLE (diagonal edge),
  ellipse-cropped picture. Fonts all ship with Windows (Bahnschrift SemiBold Condensed renders
  in LibreOffice too).
- make_ppt: Deck(FamilyLayouts); variant() dispatch per slide type; header decorations per
  family; text(font=), family radius/soft shadows on any raised shape; --design <family>;
  `make_ppt.py families` renders ui/families/*.jpg thumbnails. Deck.lead/split_even exposed
  for the mixin (must be set before main runs).
- App: Design row shows the families (thumbnail = cover + steps-slide inset), hint = name +
  description. Tested: example deck in all 6 families (fixed editorial diagonal orientation and
  soft-cover overlap), full app run in Hexa Hub (hidden window). Restored phi4/Auto/Default.

## 2026-10-04 — Settings drawer (and layout restored)
- Gear button beside the status badge opens a right-side Settings drawer (same frosted style) with
  model, image model, design, save folder, animations and "Open the reference designs folder".
  Main card keeps Sources, a "Using" row of summary chips (click = open drawer at that section),
  Slides, Focus, Generate.
- First attempt moved the brand into a new top bar; user said the previous look was perfect ->
  restored the original composition exactly (brand in the card, badge top-right), gear added
  next to the badge only.
- Bug caught by rendering: updateSummary() ran at start-up and called `on` (a const defined
  later) -> TDZ ReferenceError stopped the whole script (status stuck "Checking Ollama...").
- Tested in a hidden window: chip opens drawer, both dropdowns fit inside it, summary updates,
  Done closes. Remembered choices restored to phi4 / Auto / Default.

## 2026-10-06 — Cancel, Skip, hints, thinking-model fix
- Why only phi4 worked: thinking models (qwen3, deepseek-r1, gpt-oss) spent the reply budget on
  hidden reasoning. Probe on the two-wheeler PDF: gpt-oss:20b used all 6144 tokens thinking ->
  empty reply. Tried `think: false`: qwen3 then ran away (22k-char replies) or wrote placeholder
  slides. Fix kept: thinking stays on, models whose /api/show lists "thinking" get +8192 reply
  tokens, gpt-oss thinks "low". Also: markdown table rows ("|---|") were counted as commands, so
  every table document demanded a code slide and burned retries - fixed. Also content_checks (missing code slide, too
  many sections, invented chart numbers) are now retried but no longer fail the run; only a spec
  that can't render fails.
- Cancel: Generate turns into Cancel while running. make_ppt.STOP / SKIP + interruptible() runs
  slow calls on a helper thread so Cancel/Skip act at once; Ollama calls stream so the
  connection closes and Ollama stops generating. Checkpoints per file and per PDF page.
- Skip buttons on the pictures / write / build steps (look = stop describing, write = keep best
  spec so far, build = no previews).
- Focus field renamed "Hints for the model" (2-line box); the prompt now labels hints as user
  instructions that override the rules.
- REAL root cause (measured with prompt_eval_count): the context window was sized at 3.5
  chars/token, but qwen3 tokenizes the number-heavy price list at 1.83 chars/token -> 17,328
  prompt tokens in a 16,384 window. Ollama dropped the start of the prompt (schema + rules), so
  the model wrote "Slide 1 / content" placeholders. phi4's tokenizer needed 12,681 -> fit, which
  is why only phi4 worked. Fix: CHARS_PER_TOKEN = 1.8, window = prompt + reply budget (4k steps,
  max 32k), and the source is cut to what fits (source_limit) instead of a flat 60k chars.

## 2026-10-07 — Slide editor and slide-count modes
- Slides: Model decides / At least N / About N (segmented control; remembered). Prompt line from
  count_rule(); "at least" adds a retry complaint when the spec is short. CLI: -n 0, --at-least.
- Slide editor (replaces Notepad JSON + Rebuild): list with thumbnails, drag/↑↓ reorder, remove,
  add blank or AI-written slide after the selected one, per-type form, "Rewrite this slide with
  AI" (make_ppt.ai_slide: one slide from outline + source + instruction, validated, 3 tries),
  Refresh previews (scratch build), Export PPT… (save dialog, new file). Closing saves spec.json.
  "Edit a deck made earlier…" opens any spec.json.
- Thumbnail alignment: long slides continue onto extra deck slides, so build() now writes
  slide_map.json (spec slide -> first deck slide); any spec change deletes it, so stale previews
  are never shown against the wrong slide.
- Removed: rebuild() / edit_spec() and the Notepad flow.
- Tested: API (load, reorder/delete, preview 5 s, page-only keys stripped, AI insert with
  llama3.2, validation errors), slide_map on the example deck (16 deck slides -> 15 spec slides),
  headless renders of editor and main panel.
