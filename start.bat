@echo off
chcp 65001 > nul
echo.
echo ╔══════════════════════════════════╗
echo ║    Lab52 庫存管理系統            ║
echo ╚══════════════════════════════════╝
echo.

cd /d "%~dp0"

if not exist "node_modules" (
  echo 首次執行，正在安裝套件...
  call npm install
  if errorlevel 1 (
    echo 安裝失敗，請確認 Node.js 已安裝。
    pause
    exit /b 1
  )
)

echo 啟動伺服器中...
timeout /t 2 /nobreak > nul
start "" "http://localhost:3737"
node server.js

pause
