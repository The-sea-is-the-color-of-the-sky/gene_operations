@echo off
:: 检查是否安装了 pyinstaller
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装 pyinstaller...
    pip install pyinstaller
)

:: 使用 pyinstaller 创建 exe 文件，并设置图标和版本信息
echo 正在创建可执行文件...
pyinstaller --onefile --noconsole --icon=package\icon.ico ^
--version-file=version_info.txt main.py

:: 移动生成的 exe 文件到当前目录
if exist dist\main.exe (
    move dist\main.exe .
    echo 可执行文件创建成功：main.exe
) else (
    echo 创建可执行文件失败！
)

:: 清理生成的临时文件
rmdir /s /q build >nul 2>&1
del main.spec >nul 2>&1
rmdir /s /q dist >nul 2>&1

echo 所有操作完成！
pause