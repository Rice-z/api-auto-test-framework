@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================================
echo   一键运行接口自动化测试
echo ========================================================
echo.
.venv\Scripts\python run.py -v -s
echo.
echo 测试完成！报告在 reports 目录下
pause
