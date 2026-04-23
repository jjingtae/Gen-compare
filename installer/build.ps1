# Build script for Compare: release binaries + Inno Setup installer.
# Usage (PowerShell):  ./installer/build.ps1
#
# Produces:
#   target/release/compare.exe
#   target/release/compare-gui.exe
#   installer/Output/Compare-Setup-<Version>.exe

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

Write-Host "==> Building CLI + core (release)..." -ForegroundColor Cyan
cargo build --release

Write-Host "==> Building GUI (release)..." -ForegroundColor Cyan
cargo build -p compare-gui --release

$iscc = @(
  'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
  'C:\Program Files\Inno Setup 6\ISCC.exe'
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
  Write-Host "ISCC.exe not found. Install Inno Setup 6 from https://jrsoftware.org/isinfo.php" -ForegroundColor Yellow
  Write-Host "Binaries are built; installer step skipped." -ForegroundColor Yellow
  exit 0
}

Write-Host "==> Running Inno Setup..." -ForegroundColor Cyan
& $iscc 'installer\compare.iss'

Write-Host "==> Done. Installer under installer\Output\" -ForegroundColor Green
Get-ChildItem installer\Output\*.exe | Format-Table Name, Length, LastWriteTime
