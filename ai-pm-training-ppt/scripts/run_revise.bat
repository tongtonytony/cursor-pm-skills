@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "..\..\..\..\待裁剪PPT资源.pptx" (
    echo 错误：未找到 待裁剪PPT资源.pptx
    echo 请将文件放入项目目录：%CD%\..\..\..\..
    pause
    exit /b 1
)
python revise_training_ppt.py
pause
