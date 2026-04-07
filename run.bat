@echo off
title 38DN Macro Runner (Isolated)
echo.
echo  ========================================
echo   38DN Excel Macro Runner
echo   (Process-isolated - safe for open Excel)
echo  ========================================
echo.
python "%~dp0isolated_runner.py" %*
echo.
pause
