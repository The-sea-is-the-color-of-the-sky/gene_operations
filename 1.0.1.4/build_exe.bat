@echo off
title 自动编译 EXE
chcp 65001 >nul

REM ====== 第一步：编译 EXE ======
echo [1/2] 开始编译 EXE...
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

REM 清理临时文件
rmdir /s /q build >nul 2>&1
del /q main.spec >nul 2>&1
rmdir /s /q dist >nul 2>&1

REM ====== 第二步：完成 ======
echo [2/2] 编译完成！（未进行数字签名）
pause
