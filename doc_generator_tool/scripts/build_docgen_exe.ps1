# build_docgen_exe.ps1 — Build DocGenerator.exe with PyInstaller
Set-Location -Path $PSScriptRoot\..

$python = "..\.venv\Scripts\python.exe"
$iconScript = "scripts\generate_docgen_icon.py"
$iconFile   = "assets\docgen.ico"
$mainScript = "doc_generator.py"

Write-Host "==> Generating random icon..." -ForegroundColor Cyan
& $python $iconScript
if ($LASTEXITCODE -ne 0) { Write-Error "Icon generation failed"; exit 1 }

Write-Host "==> Building DocGenerator.exe..." -ForegroundColor Cyan
$bastligaFont = "$env:USERPROFILE\OneDrive\Desktop\Bastliga One.otf"

& $python -m PyInstaller `
    --onefile `
    --windowed `
    --name "DocGenerator" `
    --icon $iconFile `
    --paths ".." `
    --add-data "$bastligaFont;." `
    --add-data "AdsMailImporter.gs;." `
    --add-data "..\core;core" `
    --hidden-import "PySide6.QtCore" `
    --hidden-import "PySide6.QtGui" `
    --hidden-import "PySide6.QtWidgets" `
    --hidden-import "PIL._tkinter_finder" `
    --hidden-import "docx" `
    --hidden-import "docx.oxml" `
    --hidden-import "docx.oxml.ns" `
    --hidden-import "requests" `
    --hidden-import "bs4" `
    --hidden-import "openai" `
    --hidden-import "core.site_profiler" `
    --collect-submodules "docx" `
    --collect-submodules "PySide6" `
    --noconfirm `
    $mainScript

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed!"
    exit 1
}

Write-Host ""
Write-Host "==> Done! dist\DocGenerator.exe ready." -ForegroundColor Green
