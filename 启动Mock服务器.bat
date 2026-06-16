@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================================
echo   Mock API 服务器启动中...
echo   地址: http://localhost:8899
echo   按 Ctrl+C 停止，或直接关闭此窗口
echo ========================================================
echo.
.venv\Scripts\python mock_server.py
pause
