@echo off
title 最小化编译 EXE
chcp 65001 >nul

:: 配置
set MAIN_FILE=main.py
set ICON_FILE=package\icon.ico

:: 编译
echo 正在编译...
python -m PyInstaller --onefile --windowed --clean --strip --icon="%ICON_FILE%" ^
    --exclude-module=tkinter.test --exclude-module=unittest --exclude-module=pydoc --exclude-module=distutils ^
    "%MAIN_FILE%"

:: 压缩 (如果有 upx)
if exist "dist\main.exe" upx --best --lzma "dist\main.exe"

:: 移动 & 清理
if exist "dist\main.exe" move /Y "dist\main.exe" . >nul
rmdir /S /Q build dist
del /Q main.spec 2>nul

echo 完成！
pause
