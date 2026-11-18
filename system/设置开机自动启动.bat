@echo off
chcp 65001 >nul
set "SYS_DIR=%~dp0"
for %%I in ("%SYS_DIR%..") do set "ROOT_DIR=%%~fI"
set "VBS_TARGET=%ROOT_DIR%\启动后台自动监控(无黑框).vbs"
set "LNK_PATH=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\WaterProjectLedgerAutoWatch.lnk"

powershell -NoProfile -NonInteractive -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%LNK_PATH%'); $s.TargetPath = '%VBS_TARGET%'; $s.WorkingDirectory = '%ROOT_DIR%'; $s.Description = '水利工程电子台账后台自动归档监控'; $s.Save()"
if %ERRORLEVEL% EQU 0 (
    echo [成功] 已配置 Windows 开机自动静默启动电子台账监控！
    echo 快捷方式位置：%LNK_PATH%
) else (
    echo [失败] 创建开机启动快捷方式未成功，请检查系统权限。
)
pause
