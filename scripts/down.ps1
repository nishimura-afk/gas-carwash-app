# GAS から取り込む: GAS → ローカル → GitHub
# 必ずプロジェクトルートで .\scripts\down.ps1 で実行するか、scripts 内の down.bat を実行

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "GAS から取得 (down)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/2] Google Apps Script から取得中..." -ForegroundColor Yellow
clasp pull
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ clasp pull 失敗" -ForegroundColor Red
    exit 1
}
Write-Host "✓ 取得完了" -ForegroundColor Green

Write-Host ""
Write-Host "[2/2] GitHub にプッシュ中..." -ForegroundColor Yellow
git add .
$timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
git commit -m "Sync from GAS: $timestamp"
if ($LASTEXITCODE -eq 0) {
    git push origin main
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ GitHub プッシュ完了" -ForegroundColor Green
    } else {
        Write-Host "✗ git push 失敗" -ForegroundColor Red
    }
} else {
    Write-Host "変更がないためコミットをスキップ" -ForegroundColor Gray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "down 完了" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
