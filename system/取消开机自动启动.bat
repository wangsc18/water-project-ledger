@echo off
chcp 65001 >nul
set "LNK_PATH=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\WaterProjectLedgerAutoWatch.lnk"
if exist "%LNK_PATH%" (
    del /f /q "%LNK_PATH%"
    echo [成功] 已取消电子台账开机自动启动。
) else (
    echo [提示] 当前未设置开机自动启动，无需取消。
)
pause
