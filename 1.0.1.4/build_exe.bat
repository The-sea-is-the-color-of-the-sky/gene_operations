@echo off
REM 需要先安装 pyinstaller: pip install pyinstaller
REM version.txt 必须与本脚本在同一目录
pyinstaller --noconfirm --onefile --windowed --icon=package\icon.ico --distpath . --version-file package\\version.txt main.py
echo.
echo 打包完成，exe在当前目录下
pause
