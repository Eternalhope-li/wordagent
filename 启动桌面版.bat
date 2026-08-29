@echo off
chcp 65001 >nul
cd /d %~dp0
echo 正在启动 WordAgent 桌面版...
where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw gui.py
) else (
    echo 未找到 pythonw，改用 python 启动（可能带命令行窗口）
    start "" python gui.py
)
