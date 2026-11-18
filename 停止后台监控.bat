@echo off
chcp 65001 >nul
title 停止电子台账后台监控
call "%~dp0system\run_engine.bat" stop
echo.
pause
