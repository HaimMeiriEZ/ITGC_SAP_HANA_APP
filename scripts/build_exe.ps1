#Requires -Version 5.1
# Build Windows onedir package for ITGC SAP HANA APP (PyInstaller).
# Includes knowledge_base + system_settings.json (if present).
# Excludes evidence, compensating controls, findings, and other runtime state.

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $Root "ITGC_SAP_APP.spec"))) {
    $Root = $PSScriptRoot
    if (-not (Test-Path (Join-Path $Root "ITGC_SAP_APP.spec"))) {
        $Root = (Get-Location).Path
    }
}

Set-Location $Root
$Python = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $Python)) {
    $Python = "python"
}

Write-Host "==> Installing build dependency (pyinstaller)..."
& $Python -m pip install -q -r (Join-Path $Root "requirements-build.txt")

Write-Host "==> Cleaning previous build/dist for this app..."
$DistApp = Join-Path $Root "dist\ITGC_SAP_APP"
$BuildApp = Join-Path $Root "build\ITGC_SAP_APP"
if (Test-Path $DistApp) { Remove-Item -Recurse -Force $DistApp }
if (Test-Path $BuildApp) { Remove-Item -Recurse -Force $BuildApp }

Write-Host "==> Running PyInstaller (onedir)..."
& $Python -m PyInstaller --noconfirm --clean (Join-Path $Root "ITGC_SAP_APP.spec")
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE"
}

$DataRoot = Join-Path $DistApp "data"
$OutputDir = Join-Path $DataRoot "output"
$InputDir = Join-Path $DataRoot "input"
$EvidenceDir = Join-Path $DataRoot "evidence"
$CompDir = Join-Path $DataRoot "compensating_controls"
$KbDir = Join-Path $DataRoot "knowledge_base"
New-Item -ItemType Directory -Force -Path $OutputDir, $InputDir, $EvidenceDir, $CompDir, $KbDir | Out-Null

$SrcKb = Join-Path $Root "data\knowledge_base"
if (Test-Path $SrcKb) {
    Copy-Item -Force (Join-Path $SrcKb "controls_catalog.json") (Join-Path $KbDir "controls_catalog.json") -ErrorAction SilentlyContinue
    Copy-Item -Force (Join-Path $SrcKb "field_labels.json") (Join-Path $KbDir "field_labels.json") -ErrorAction SilentlyContinue
}

$SrcSettings = Join-Path $Root "data\output\system_settings.json"
if (Test-Path $SrcSettings) {
    Copy-Item -Force $SrcSettings (Join-Path $OutputDir "system_settings.json")
    Write-Host "==> Included system_settings.json"
} else {
    Write-Host "==> No system_settings.json found - client will start with defaults"
}

$ExePath = Join-Path $DistApp "ITGC_SAP_APP.exe"
Write-Host ""
Write-Host "Build OK."
Write-Host "Package folder: $DistApp"
Write-Host "Run: $ExePath"
