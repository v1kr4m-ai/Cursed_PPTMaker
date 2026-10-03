<#
.SYNOPSIS
    Friendly front end for make_ppt.py. Started by "Make PPT.bat".
.DESCRIPTION
    - No files given: opens a file picker.
    - Documents given: offline local-model deck (make_ppt.py auto).
    - A spec.json given: rebuilds that deck without the model (make_ppt.py build).
    Starts Ollama if needed, never overwrites an existing deck, opens the result.
#>
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Files)

$Host.UI.RawUI.WindowTitle = "Make PPT (offline)"
$script = Join-Path $PSScriptRoot "make_ppt.py"
$defaultModel = if ($env:MAKE_PPT_MODEL) { $env:MAKE_PPT_MODEL } else { "qwen3-coder:30b" }
$ollamaUrl = if ($env:OLLAMA_HOST) { $env:OLLAMA_HOST.TrimEnd("/") } else { "http://localhost:11434" }
$env:PYTHONUNBUFFERED = "1"   # live progress instead of everything at the end

function Stop-WithMessage([string]$msg) {
    Write-Host "`n$msg" -ForegroundColor Red
    Read-Host "`nPress Enter to close"
    exit 1
}

function Ask([string]$question, [string]$default) {
    $answer = Read-Host "$question [$default]"
    if ([string]::IsNullOrWhiteSpace($answer)) { $default } else { $answer.Trim() }
}

function Get-FreePath([string]$path) {
    # report.pptx -> "report - slides.pptx", then "(2)", "(3)"... so nothing is overwritten
    $dir = Split-Path $path
    $stem = [IO.Path]::GetFileNameWithoutExtension($path) + " - slides"
    $candidate = Join-Path $dir "$stem.pptx"
    $n = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $dir "$stem ($n).pptx"
        $n++
    }
    $candidate
}

Write-Host "=== Make PPT - offline deck builder ===`n" -ForegroundColor Cyan

if (-not (Get-Command python -ErrorAction SilentlyContinue)) { Stop-WithMessage "Python is not on PATH." }
if (-not (Test-Path -LiteralPath $script)) { Stop-WithMessage "make_ppt.py not found next to this launcher." }

# ---- pick files -------------------------------------------------------------
if (-not $Files) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Choose document(s) for the deck - hold Ctrl to pick several"
    $dlg.Multiselect = $true
    $dlg.Filter = "Documents|*.pdf;*.docx;*.doc;*.odt;*.rtf;*.xlsx;*.xls;*.ods;*.csv;*.pptx;*.ppt;*.odp;*.txt;*.md;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp|Slide spec (rebuild)|*.json|All files|*.*"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
    $Files = $dlg.FileNames
}
$Files = @($Files | ForEach-Object { (Resolve-Path -LiteralPath $_ -ErrorAction SilentlyContinue).Path })
if ($Files -contains $null -or -not $Files) { Stop-WithMessage "One of the files could not be found." }

Write-Host "Files:"
$Files | ForEach-Object { Write-Host "  $_" }

# ---- rebuild mode: a spec.json was given ---------------------------------------
if ($Files.Count -eq 1 -and $Files[0] -like "*.json") {
    $spec = $Files[0]
    $first = Split-Path (Split-Path $spec)   # spec lives in <doc>_ppt\, deck goes next to <doc>
    $name = (Split-Path (Split-Path $spec) -Leaf) -replace "_ppt$", ""
    $out = Get-FreePath (Join-Path $first "$name.pptx")
    Write-Host "`nRebuilding from the spec (no model needed)...`n" -ForegroundColor Cyan
    & python $script build $spec -o $out
    if ($LASTEXITCODE -ne 0) { Stop-WithMessage "Rebuild failed - see the messages above." }
    Invoke-Item -LiteralPath $out
    Read-Host "`nDone: $out`nPress Enter to close"
    exit 0
}

# ---- make sure Ollama is up ---------------------------------------------------
function Get-Models {
    try { @((Invoke-RestMethod "$ollamaUrl/api/tags" -TimeoutSec 3).models) } catch { $null }
}
$models = Get-Models
if ($null -eq $models) {
    $ollama = Get-Command ollama -ErrorAction SilentlyContinue
    if (-not $ollama) { Stop-WithMessage "Ollama is not running and 'ollama' is not on PATH." }
    Write-Host "`nStarting Ollama..." -ForegroundColor Yellow
    Start-Process $ollama.Source -ArgumentList "serve" -WindowStyle Hidden
    for ($i = 0; $i -lt 30 -and $null -eq $models; $i++) { Start-Sleep 1; $models = Get-Models }
    if ($null -eq $models) { Stop-WithMessage "Ollama did not start within 30 seconds." }
}
# ---- model menu ---------------------------------------------------------------
# Offline chat models only: cloud models break "offline", embedding models can't write slides.
$local = @($models | Where-Object { $_.name -notmatch "cloud" -and $_.name -notmatch "embed" } |
           Sort-Object size)
if (-not $local) { Stop-WithMessage "No offline Ollama models installed. Pull one, e.g. 'ollama pull qwen3:8b'." }

$freeGB = (Get-CimInstance Win32_OperatingSystem).FreePhysicalMemory / 1MB
try {  # add free GPU memory when an NVIDIA card is present
    $freeGB += [double](nvidia-smi --query-gpu=memory.free --format=csv,noheader,nounits 2>$null |
                        Select-Object -First 1) / 1024
} catch {}
try {  # models Ollama already has loaded get unloaded on demand, so count them as free
    $freeGB += ((Invoke-RestMethod "$ollamaUrl/api/ps" -TimeoutSec 3).models | Measure-Object size -Sum).Sum / 1GB
} catch {}
$width = ($local.name | Measure-Object Length -Maximum).Maximum + 2

$lastFile = Join-Path $PSScriptRoot ".last_model"
$last = if (Test-Path -LiteralPath $lastFile) { (Get-Content -LiteralPath $lastFile -Raw).Trim() } else { "" }
$default = @($last, $defaultModel) | Where-Object { $_ -and $local.name -contains $_ } | Select-Object -First 1
if (-not $default) { $default = $local[0].name }

Write-Host ("`nInstalled offline models (free memory now: about {0:N0} GB):" -f $freeGB)
for ($i = 0; $i -lt $local.Count; $i++) {
    $m = $local[$i]
    $gb = $m.size / 1GB
    $line = "  {0,2}. {1} {2,5:N1} GB" -f ($i + 1), $m.name.PadRight($width), $gb
    if ($m.name -eq $default) { $line += "  <- default" }
    if ($gb * 1.1 -gt $freeGB) {   # rough: Ollama needs a bit more than the file size
        Write-Host "$line  (may not fit right now)" -ForegroundColor DarkYellow
    } else {
        Write-Host $line
    }
}
$model = $null
while (-not $model) {
    $pick = (Read-Host "Pick a number, or press Enter for $default").Trim()
    if (-not $pick) { $model = $default }
    elseif ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $local.Count) { $model = $local[[int]$pick - 1].name }
    elseif ($local.name -contains $pick) { $model = $pick }
    else { Write-Host "  Not in the list - type a number from 1 to $($local.Count)." -ForegroundColor Yellow }
}
Set-Content -LiteralPath $lastFile -Value $model
Write-Host "  Using $model`n"

# ---- other questions (Enter = default) ---------------------------------------
$slides = Ask "How many slides" "10"
if ($slides -notmatch '^\d+$') { $slides = "10" }
$extra = Read-Host "Anything to focus on? e.g. 'audience: management' [none]"

# ---- build ------------------------------------------------------------------
$out = Get-FreePath $Files[0]
$argv = @($script, "auto") + $Files + @("-m", $model, "-n", $slides, "-o", $out)
if ($extra.Trim()) { $argv += @("-i", $extra.Trim()) }

Write-Host "`nWorking - the model can take a few minutes...`n" -ForegroundColor Cyan
$start = Get-Date
& python @argv
if ($LASTEXITCODE -ne 0) { Stop-WithMessage "Something went wrong - see the messages above." }

$work = Join-Path (Split-Path $Files[0]) ([IO.Path]::GetFileNameWithoutExtension($Files[0]) + "_ppt")
Write-Host ("`nDone in {0:mm\:ss}." -f ((Get-Date) - $start)) -ForegroundColor Green
Write-Host "  Deck:     $out"
Write-Host "  Previews: $work\preview"
Write-Host "  To tweak: edit $work\spec.json, then drag it onto 'Make PPT.bat'."
Invoke-Item -LiteralPath $out
Read-Host "`nPress Enter to close"
