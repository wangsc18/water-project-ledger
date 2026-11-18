# -*- coding: utf-8 -*-
"""
Windows 10/11 右下角原生系统托盘气泡通知模块
使用 System.Windows.Forms.NotifyIcon，无需管理员权限或注册表 AppID，兼容所有 Win10/Win11 版本
"""
from __future__ import annotations

import subprocess
import sys


def send_windows_notification(title: str, message: str, is_warning: bool = False, enabled: bool = True) -> None:
    if not enabled:
        return
    if sys.platform != "win32":
        return

    safe_title = title.replace("'", " ").replace('"', " ")[:60]
    safe_msg = message.replace("'", " ").replace('"', " ")[:200]
    icon_type = "Warning" if is_warning else "Info"

    ps_code = (
        "Add-Type -AssemblyName System.Windows.Forms; "
        "Add-Type -AssemblyName System.Drawing; "
        "$n = New-Object System.Windows.Forms.NotifyIcon; "
        "$n.Icon = [System.Drawing.SystemIcons]::Information; "
        "$n.Visible = $true; "
        f"$n.ShowBalloonTip(5000, '{safe_title}', '{safe_msg}', [System.Windows.Forms.ToolTipIcon]::{icon_type}); "
        "Start-Sleep -Seconds 5; "
        "$n.Dispose();"
    )
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", ps_code],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )
    except Exception:
        pass
