@echo off
chcp 65001 >nul
cd /d "C:\Users\Lenovo\Documents\New OpenCode Project"
echo.
echo  ============================================
echo   卡路里之路 服务器（含 AI 拍照识卡）
echo   电脑访问： http://localhost:3001
echo   手机访问（同一 WiFi）： http://10.4.54.52:3001
echo   关闭本窗口即可停止服务
echo  ============================================
echo.
node server.js
