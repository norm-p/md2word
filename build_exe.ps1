# build_exe.ps1 — Install deps and build md2word.exe
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $projectDir

Write-Host "=== Installing project dependencies ===" -ForegroundColor Cyan
pip install -e ".[all-ai]" --quiet

Write-Host "=== Installing PyInstaller ===" -ForegroundColor Cyan
pip install pyinstaller --quiet

Write-Host "=== Building md2word.exe ===" -ForegroundColor Cyan
$scriptsDir = python -c "import sysconfig; print(sysconfig.get_path('scripts', 'nt_user'))"
& "$scriptsDir\pyinstaller.exe" md2word.spec --noconfirm --clean

Write-Host ""
if (Test-Path "dist\md2word.exe") {
    $size = (Get-Item "dist\md2word.exe").Length / 1MB
    Write-Host ("SUCCESS: dist\md2word.exe ({0:N1} MB)" -f $size) -ForegroundColor Green
    Write-Host ""
    Write-Host "=== Quick smoke test ===" -ForegroundColor Cyan
    & "dist\md2word.exe" --help
} else {
    Write-Host "FAILED: dist\md2word.exe not found" -ForegroundColor Red
    exit 1
}

Pop-Location
