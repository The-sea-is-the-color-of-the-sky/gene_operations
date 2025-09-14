@echo off
title 自动编译 EXE
chcp 65001 >nul

:: 配置
set MAIN_FILE=main.py
set ICON_FILE=package\icon.ico

:: 编译
echo [1/2] 开始编译 EXE...
python -m PyInstaller --noconfirm --onefile --windowed --icon="%ICON_FILE%" "%MAIN_FILE%"
if errorlevel 1 (
    echo 编译失败！
    pause
    exit /b 1
)

:: 移动文件
if exist "dist\main.exe" (
    move /Y "dist\main.exe" . >nul
    echo 可执行文件创建成功：main.exe
) else (
    echo 错误：未找到 dist\main.exe
    pause
    exit /b 1
)

:: 清理
echo [2/2] 清理临时文件...
rmdir /S /Q build
rmdir /S /Q dist
del /Q main.spec 2>nul

echo 编译完成！
pause
