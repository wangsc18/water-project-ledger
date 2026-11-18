@echo off
setlocal
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set "SYS_DIR=%~dp0"
for %%I in ("%SYS_DIR%..") do set "ROOT_DIR=%%~fI"

if exist "%SYS_DIR%runtime\ledger_engine.exe" (
    "%SYS_DIR%runtime\ledger_engine.exe" %* --root "%ROOT_DIR%"
    exit /b %ERRORLEVEL%
)

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    python "%SYS_DIR%main.py" %* --root "%ROOT_DIR%"
    exit /b %ERRORLEVEL%
)

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    py -3 "%SYS_DIR%main.py" %* --root "%ROOT_DIR%"
    exit /b %ERRORLEVEL%
)

echo [错误] 未检测到内置绿色运行库 system\runtime\ledger_engine.exe，也未检测到本机 Python 环境！
echo 若您使用的是源码版，请先安装 Python 3.10+ 或从 GitHub Actions 下载《Windows绿色免安装版.zip》。
pause
exit /b 1
