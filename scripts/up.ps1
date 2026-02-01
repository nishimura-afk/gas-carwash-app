# GAS（clasp）に送る: ローカル → GitHub → GAS（オプションでデプロイ）
# 必ずプロジェクトルートで .\scripts\up.ps1 で実行するか、scripts 内の up.bat を実行

# デプロイを有効にする場合は $true に変更
$EnableDeploy = $false

# スクリプトがあるフォルダの親 = プロジェクトルート
$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "GAS に送信 (up)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/2] GitHub にプッシュ中..." -ForegroundColor Yellow
git add .
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ git add 失敗" -ForegroundColor Red
    exit 1
}

$timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
git commit -m "Update: $timestamp"
if ($LASTEXITCODE -ne 0) {
    Write-Host "変更がないためコミットをスキップ" -ForegroundColor Gray
} else {
    git push origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host "✗ git push 失敗" -ForegroundColor Red
        exit 1
    }
    Write-Host "✓ GitHub プッシュ完了" -ForegroundColor Green
}

Write-Host ""
Write-Host "[2/2] Google Apps Script に送信中..." -ForegroundColor Yellow
$claspStatus = clasp status 2>&1
if ($LASTEXITCODE -ne 0 -or $claspStatus -match "not logged in" -or $claspStatus -match "Login required") {
    Write-Host "✗ clasp にログインが必要です。 clasp login を実行してください。" -ForegroundColor Red
    exit 1
}

clasp push
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ clasp push 失敗" -ForegroundColor Red
    exit 1
}
Write-Host "✓ GAS 送信完了" -ForegroundColor Green

if ($EnableDeploy) {
    Write-Host ""
    Write-Host "[3/3] デプロイ中..." -ForegroundColor Yellow
    clasp deploy --description "Update: $timestamp"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ デプロイ完了" -ForegroundColor Green
    } else {
        Write-Host "✗ デプロイ失敗" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "up 完了" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
