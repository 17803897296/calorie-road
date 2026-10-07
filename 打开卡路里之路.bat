@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo   ==========================================
echo     卡路里之路  正在启动，请稍等 2 秒...
echo     浏览器会自动打开： http://localhost:3001
echo     （关闭这个黑色窗口 = 停止服务）
echo   ==========================================
echo.
timeout /t 2 /nobreak >nul
start "" "http://localhost:3001"
node server.js
