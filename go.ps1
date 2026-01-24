Write-Host "========================================" -ForegroundColor Cyan
Write-Host "洗車機管理システムをGASに送信" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/2] GitHubにプッシュ中..." -ForegroundColor Yellow
git add .
if ($LASTEXITCODE -eq 0) {
    $timestamp = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    git commit -m "Update: $timestamp"
    if ($LASTEXITCODE -eq 0) {
        git push origin main
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ GitHubプッシュ完了" -ForegroundColor Green
        } else {
            Write-Host "✗ GitHubプッシュ失敗" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "✓ 変更がないためコミットをスキップ" -ForegroundColor Green
    }
} else {
    Write-Host "✗ Git add 失敗" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "[2/2] Google Apps Scriptに送信中..." -ForegroundColor Yellow

# claspログイン状態をチェック
$claspStatus = clasp status 2>&1
if ($LASTEXITCODE -ne 0 -or $claspStatus -match "not logged in" -or $claspStatus -match "Login required") {
    Write-Host "✗ claspにログインが必要です" -ForegroundColor Red
    Write-Host ""
    Write-Host "以下のコマンドでログインしてください:" -ForegroundColor Yellow
    Write-Host "  clasp login" -ForegroundColor White
    Write-Host ""
    Write-Host "ログイン後、再度このスクリプトを実行してください。" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "何かキーを押してください..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

clasp push
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ GAS送信完了" -ForegroundColor Green
} else {
    Write-Host "✗ GAS送信失敗" -ForegroundColor Red
    Write-Host ""
    Write-Host "エラーの原因:"
    Write-Host "- claspのログインが切れている可能性があります" -ForegroundColor Yellow
    Write-Host "- .clasp.jsonのscriptIdが正しくない可能性があります" -ForegroundColor Yellow
    Write-Host "- GASプロジェクトへのアクセス権限がない可能性があります" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "解決方法:" -ForegroundColor Cyan
    Write-Host "1. clasp login でログインし直す" -ForegroundColor White
    Write-Host "2. clasp status でプロジェクトの状態を確認" -ForegroundColor White
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "GASへの送信完了！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "何かキーを押してください..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
