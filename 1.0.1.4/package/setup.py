from setuptools import setup, find_packages

setup(
    name="gene_tool",
    version="1.0.1.4",
    description="基因工具",
    author="宋庆海",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "openpyxl",
        "tkinter"
    ],
    include_package_data=True,
    entry_points={
        "gui_scripts": [
            "gene_tool = main:main"
        ]
    }
)
