@echo off
title 自动编译 EXE
chcp 65001 >nul

REM ====== 检查 version.txt 是否存在 ======
if not exist "package\version.txt" (
    echo 错误：找不到 package\version.txt！
    pause
    exit /b 1
)

REM ====== 第一步：编译 EXE ======
echo [1/3] 开始编译 EXE...
python -m PyInstaller --noconfirm --onefile --windowed --icon=package\icon.ico ^
--distpath dist --version-file package\version.txt main.py

if not exist "dist\main.exe" (
    echo 编译失败！
    pause
    exit /b 1
)

REM 移动生成的 exe 文件到当前目录
move /Y "dist\main.exe" . >nul
echo 可执行文件创建成功：main.exe

REM ====== 第二步：清理临时文件 ======
echo [2/3] 清理临时文件...
rmdir /S /Q build
rmdir /S /Q dist
del /Q main.spec

REM ====== 第三步：完成 ======
echo [3/3] 编译完成并清理结束！
pause
