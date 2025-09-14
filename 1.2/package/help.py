import tkinter as tk
from tkinter import Scrollbar, Text
import os
import tkinter.font as tkFont


class SyntenyGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("帮助说明")
        self.master.geometry("650x500")

        # 设置窗口图标
        icon_file = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_file):
            try:
                self.master.iconbitmap(icon_file)
            except Exception as e:
                print(f"加载图标失败: {e}")

        self.create_help_content()

    def create_help_content(self):
        help_window = self.master
        # 滚动条 + 文本框
        scrollbar = Scrollbar(help_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 设置宋体五号字体 (约12pt)
        font_style = tkFont.Font(family="SimSun", size=14)  # SimSun=宋体

        text = Text(help_window, wrap="word", yscrollcommand=scrollbar.set, font=font_style)
        text.pack(expand=True, fill="both")
        scrollbar.config(command=text.yview)

        # 使用说明内容
        help_content = """
    基因工具 使用说明书

    【一、功能简介】
    本程序用于Excel格式基因表格的批量查询、递归匹配和模糊匹配，
    支持横向、竖向排列，并可对模糊匹配结果进行高亮显示。
    同时提供基因共线性文件解析与可视化功能。

    【二、主要功能】
    1. 文件预处理：
    - 基因ID：用于预处理基因ID格式（功能待完善）。
    - 信息表格：可将原始信息表格转换为标准化格式，以便后续匹配。

    2. 基因匹配工具：
    - 选择“填入表格.xlsx”和“信息表格.xlsx”文件。
    - 支持精确匹配、模糊匹配和递归匹配。
    - 匹配结果会高亮显示，未匹配基因会列出提示。

    3. 可视化：
    - 共线性可视化：读取 .collinearity 文件并绘制基因共线性关系图。
    - 其他可视化：预留功能（开发中）。

    【三、操作步骤】
    1. 进入“文件预处理” → “信息表格”，转换数据格式。
    2. 在“工具” → “基因匹配”中执行基因匹配。
    3. 在“可视化” → “共线性可视化”中查看匹配结果图。

    【四、注意事项】
    - 输入文件需为 Excel 格式（.xlsx）。
    - 输出结果默认保存至程序目录下。
    - 若使用打包后的程序，请将输入文件放在程序同一目录，或指定绝对路径。

    【五、联系与更新】
    如有问题或建议，请联系开发者，或关注版本更新日志。
        """
        text.insert("1.0", help_content)
        text.config(state="disabled")  # 只读