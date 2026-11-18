@echo off
chcp 65001 >nul
title 水利工程全流程电子台账 - 单次扫描与校准
call "%~dp0system\run_engine.bat" scan
echo.
echo [完成] 台账与可视化看板已刷新，按任意键关闭窗口...
pause >nul
