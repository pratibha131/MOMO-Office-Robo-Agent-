# Momo setup — installs everything into %LOCALAPPDATA%\Momo (no admin rights needed).
# Heavy runtimes stay OUT of OneDrive; this project folder keeps only source + data.
# Run:  powershell -ExecutionPolicy Bypass -File setup.ps1
param(
  [switch]$SkipNode,
  [switch]$SkipNpm,
  [switch]$SkipPython,
  [switch]$SkipModel
)
$ErrorActionPreference = 'Stop'
$Project  = $PSScriptRoot
$MomoHome = Join-Path $env:LOCALAPPDATA 'Momo'
$NodeDir  = Join-Path $MomoHome 'runtime\node'
$AppDir   = Join-Path $MomoHome 'app'
$ModelDir = Join-Path $MomoHome 'models'
$NodeVer  = 'v22.14.0'

Write-Host "== Momo setup ==" -ForegroundColor Magenta
New-Item -ItemType Directory -Force -Path $MomoHome, $AppDir, $ModelDir | Out-Null
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------- 1. Portable Node.js ----------
if (-not $SkipNode) {
  if (Test-Path (Join-Path $NodeDir 'node.exe')) {
    Write-Host "[1/5] Node already present, skipping" -ForegroundColor Green
  } else {
    Write-Host "[1/5] Downloading portable Node.js $NodeVer ..." -ForegroundColor Cyan
    $zip = Join-Path $env:TEMP "node-$NodeVer.zip"
    Invoke-WebRequest "https://nodejs.org/dist/$NodeVer/node-$NodeVer-win-x64.zip" -OutFile $zip
    Expand-Archive $zip -DestinationPath (Join-Path $MomoHome 'runtime') -Force
    if (Test-Path $NodeDir) { Remove-Item $NodeDir -Recurse -Force }
    Rename-Item (Join-Path $MomoHome "runtime\node-$NodeVer-win-x64") $NodeDir
    Remove-Item $zip -Force
    Write-Host "      Node installed -> $NodeDir" -ForegroundColor Green
  }
}
$env:PATH = "$NodeDir;$env:PATH"

# ---------- 2. Electron + Claude Code engine ----------
if (-not $SkipNpm) {
  Write-Host "[2/5] Installing Electron + Claude engine (this can take a few minutes)..." -ForegroundColor Cyan
  Set-Content -Path (Join-Path $AppDir 'package.json') -Value '{ "name": "momo-runtime", "private": true }' -Encoding utf8
  Push-Location $AppDir
  try {
    & "$NodeDir\npm.cmd" install --no-fund --no-audit electron@latest '@anthropic-ai/claude-code@latest'
    if ($LASTEXITCODE -ne 0) { throw "npm install failed ($LASTEXITCODE)" }
  } finally { Pop-Location }
  Write-Host "      Electron + Claude engine installed" -ForegroundColor Green
}

# ---------- 3. Python voice packages ----------
if (-not $SkipPython) {
  Write-Host "[3/5] Installing Python voice packages (vosk, sounddevice, soundcard)..." -ForegroundColor Cyan
  & pip install --user -r (Join-Path $Project 'python\requirements.txt')
  if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
  Write-Host "      Python packages installed" -ForegroundColor Green
}

# ---------- 4. Vosk speech model (offline, Indian English) ----------
if (-not $SkipModel) {
  $mdl = Join-Path $ModelDir 'vosk-model-small-en-in-0.4'
  if (Test-Path $mdl) {
    Write-Host "[4/5] Speech model already present, skipping" -ForegroundColor Green
  } else {
    Write-Host "[4/5] Downloading speech model (~40 MB)..." -ForegroundColor Cyan
    $zip = Join-Path $env:TEMP 'vosk-model.zip'
    Invoke-WebRequest 'https://alphacephei.com/vosk/models/vosk-model-small-en-in-0.4.zip' -OutFile $zip
    Expand-Archive $zip -DestinationPath $ModelDir -Force
    Remove-Item $zip -Force
    Write-Host "      Model installed -> $mdl" -ForegroundColor Green
  }
}

# ---------- 5. Check Claude login ----------
Write-Host "[5/5] Checking Claude engine login..." -ForegroundColor Cyan
$claudeExe = Join-Path $AppDir 'node_modules\@anthropic-ai\claude-code\bin\claude.exe'
if (Test-Path $claudeExe) {
  $out = & $claudeExe -p "Reply with exactly: OK" --output-format text 2>&1
  if ("$out" -match 'OK') {
    Write-Host "      Claude engine is logged in and working!" -ForegroundColor Green
  } else {
    Write-Host "      Claude engine needs a one-time login. Run this, then follow the prompts:" -ForegroundColor Yellow
    Write-Host "        `"$claudeExe`"" -ForegroundColor Yellow
    Write-Host "      (type /login inside it, finish sign-in, then /exit)" -ForegroundColor Yellow
  }
} else {
  Write-Host "      Claude engine not found - re-run step 2" -ForegroundColor Red
}
# Electron sometimes skips its binary download under npm; fetch it if missing
$edist = Join-Path $AppDir 'node_modules\electron\dist\electron.exe'
if (-not (Test-Path $edist)) {
  Write-Host "      Fetching Electron binary..." -ForegroundColor Cyan
  Push-Location (Join-Path $AppDir 'node_modules\electron')
  & "$NodeDir\node.exe" install.js
  Pop-Location
}

# ---------- Start Menu shortcut with global hotkey (works even when closed) ----------
try {
  $ws = New-Object -ComObject WScript.Shell
  $lnk = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Momo.lnk"
  $s = $ws.CreateShortcut($lnk)
  $s.TargetPath = Join-Path $AppDir "node_modules\electron\dist\electron.exe"
  $s.Arguments = '"' + $Project + '\."'
  $s.WorkingDirectory = $Project
  $ico = Join-Path $Project "app\renderer\momo.ico"
  if (Test-Path $ico) { $s.IconLocation = $ico }
  $s.Hotkey = "Ctrl+Alt+M"
  $s.Description = "Momo - your desktop companion (Ctrl+Alt+M)"
  $s.Save()
  Write-Host "      Ctrl+Alt+M summons Momo from anywhere (Start Menu shortcut)" -ForegroundColor Green
} catch { Write-Host "      shortcut/hotkey setup failed: $_" -ForegroundColor Yellow }

Write-Host ""
Write-Host "Setup complete. Start Momo with:  Start-Momo.bat  (or press Ctrl+Alt+M)" -ForegroundColor Magenta
