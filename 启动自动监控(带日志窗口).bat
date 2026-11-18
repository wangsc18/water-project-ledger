@echo off
chcp 65001 >nul
title 水利工程全流程电子台账 - 实时自动监控窗口
call "%~dp0system\run_engine.bat" watch
pause
