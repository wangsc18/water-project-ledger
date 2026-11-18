@echo off
chcp 65001 >nul
title 安装水利电子台账 Python 依赖库
echo 正在通过清华镜像源安装 Windows 运行依赖（含离线中文 RapidOCR 引擎）...
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
python -m pip install -r "%~dp0requirements-windows.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.
echo [安装完成] 现在可以直接双击根目录的启动脚本使用了！
pause
