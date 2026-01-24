@echo off
echo ========================================
echo GASから最新版を取得してGitHubにプッシュ
echo ========================================
echo.

echo [1/2] Google Apps Scriptから取得中...
clasp pull

echo.
echo [2/2] GitHubにプッシュ中...
git add .
git commit -m "Sync from GAS: %date% %time%"
git push origin main

echo.
echo ========================================
echo 同期完了！
echo ========================================
pause
