from setuptools import setup, find_packages

setup(
    name="gene_tool",
    version="1.1.4",
    description="基因工具",
    author="宋庆海",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "openpyxl"
        # tkinter为标准库，无需安装
    ],
    include_package_data=True,
    entry_points={
        "gui_scripts": [
            "gene_tool = main:main"
        ]
    }
)
# 打包说明:
# 推荐使用 PyInstaller 进行打包:
# 1. 安装 PyInstaller: pip install pyinstaller
# 2. 在主目录下运行如下命令生成单文件可执行程序:
#    pyinstaller -F -w -i package/icon.ico main.py
#    -F: 单文件
#    -w: 无命令行窗口
#    -i: 图标路径
# 3. 打包后在 dist 目录下生成 exe 文件
