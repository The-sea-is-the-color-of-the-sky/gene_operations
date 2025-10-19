import tkinter as tk
from tkinter import Menu, messagebox, Toplevel, Scrollbar, Text
import os
import sys
import ast
import networkx as nx


# ========== 资源路径函数 ==========
def resource_path(relative_path):
    """获取资源绝对路径，兼容 PyInstaller 打包"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ========== 主窗口 ==========
root = tk.Tk()
root.title("基因工具")
root.geometry("500x350")

# 设置主窗口图标
icon_file = resource_path("D:\a数据库\有意思的东西\基因工具\Version archive\1.2\package\icon.ico")
if os.path.exists(icon_file):
    try:
        root.iconbitmap(icon_file)
    except Exception as e:
        print(f"加载图标失败: {e}")

# ===================== 文件转换 GUI =====================
from package.File_conversion import FileConversionUI

def open_file_conversion():
    new_window = Toplevel(root)
    new_window.title("信息表格转换")
    new_window.geometry("500x200")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    app = FileConversionUI(new_window)

# ===================== 基因匹配 GUI =====================
from package.gene_match_gui import GeneToolApp

def open_gene_match():
    new_window = Toplevel(root)
    new_window.title("基因匹配工具")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    app = GeneToolApp(new_window)

# ===================== 共线性可视化 GUI =====================
from package.Collinearity_Visualization import CV_GUI

def open_synteny():
    new_window = Toplevel(root)
    new_window.title("共线性可视化")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    app = CV_GUI(new_window)

# ===================== 帮助页面 =====================
from package.help import SyntenyGUI

def open_help():
    new_window = Toplevel(root)
    new_window.title("帮助")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    app = SyntenyGUI(new_window)
# ===================== 菜单栏 =====================
menu_bar = Menu(root)
root.config(menu=menu_bar)

# 文件预处理菜单
file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="文件预处理", menu=file_menu)
file_menu.add_command(label="基因ID", command=lambda: messagebox.showinfo("提示", "此功能尚未实现"))
file_menu.add_command(label="信息表格", command=open_file_conversion)
file_menu.add_separator()
file_menu.add_command(label="退出", command=root.quit)

# 工具菜单
edit_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="工具", menu=edit_menu)
edit_menu.add_command(label="基因匹配", command=open_gene_match)

# 可视化二级菜单
visual_menu = Menu(edit_menu, tearoff=0)
edit_menu.add_cascade(label="可视化", menu=visual_menu)
visual_menu.add_command(label="共线性可视化", command=open_synteny)
visual_menu.add_command(label="其他可视化", command=lambda: messagebox.showinfo("提示", "此功能尚未实现"))

# 帮助菜单
help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="帮助", menu=help_menu)
help_menu.add_command(label="使用说明", command=open_help)
help_menu.add_separator()
help_menu.add_command(label="关于", command=lambda: messagebox.showinfo("关于", "gene_operations\nauthor：sea wears sky is color"))

# ===================== 主界面提示 =====================
label = tk.Label(root, text="欢迎使用基因工具！\n请通过菜单选择功能。")
label.pack(expand=True)

# ===================== 启动主循环 =====================
root.mainloop()
