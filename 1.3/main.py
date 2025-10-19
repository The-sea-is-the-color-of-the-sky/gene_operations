import tkinter as tk
from tkinter import Menu, messagebox, Toplevel
import os
import sys
import networkx as nx

# ========== 资源路径函数 ==========
def resource_path(relative_path):
    """获取资源绝对路径，兼容 PyInstaller 打包"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ========== 全局样式 ==========
MENU_BG = "white"               # 菜单默认背景
MENU_FG = "black"               # 菜单默认文字
MENU_ACTIVE_FG = "red"          # 鼠标悬停文字
MENU_ACTIVE_BG = "white"        # 鼠标悬停背景

# ========== 主窗口 ==========
root = tk.Tk()
root.title("基因工具")
root.geometry("500x350")

# 设置主窗口图标
icon_file = os.path.join(os.path.dirname(__file__), "package","image", "icon.ico")
if os.path.exists(icon_file):
    try:
        root.iconbitmap(icon_file)
    except Exception as e:
        print(f"加载图标失败: {e}")

# 添加：将窗口短时置顶的辅助函数（不使用 grab_set，避免主窗口被阻塞）
def bring_to_front(win, duration=200):
    """将窗口提升到最顶层短时间（ms），随后恢复 topmost=False，避免阻塞主窗口"""
    try:
        win.transient(root)  # 保持相对主窗口层级
        win.lift()
        win.attributes("-topmost", True)
        # 经过短时延迟后取消 topmost，使行为自然
        win.after(duration, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

# ===================== 文件转换 GUI =====================
from package.File_conversion import FileConversionUI
def open_file_conversion():
    new_window = Toplevel(root)
    new_window.title("信息文件转换")
    new_window.geometry("500x200")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    new_window.transient(root)
    app = FileConversionUI(new_window)
    # 将新窗口短时置顶（不会阻塞主窗口）
    bring_to_front(new_window)

from package.id_gui import id_UI
def open_id_gui():
    new_window = Toplevel(root)
    new_window.title("ID文件转换")
    new_window.geometry("500x200")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    new_window.transient(root)
    app = id_UI(new_window)
    bring_to_front(new_window)

# ===================== 基因匹配 GUI =====================
from package.gene_match_gui import GeneToolApp
def open_gene_match():
    new_window = Toplevel(root)
    new_window.title("基因匹配工具")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    new_window.transient(root)
    app = GeneToolApp(new_window)
    bring_to_front(new_window)

from package.gene_tool_pro import GeneProApp
def gene_tool_pro():
    new_window = Toplevel(root)
    new_window.title("基因匹配pro")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    new_window.transient(root)
    app = GeneProApp(new_window)
    bring_to_front(new_window)

# ===================== 可视化 GUI =====================
from package.cv_link import CV_LINK_GUI
def open_cv_link():
    new_window = Toplevel(root)
    new_window.title("基因关联可视化")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    new_window.transient(root)
    app = CV_LINK_GUI(new_window)
    bring_to_front(new_window)

# ===================== 帮助页面 =====================
from package.help import SyntenyGUI
def open_help():
    new_window = Toplevel(root)
    new_window.title("帮助")
    if os.path.exists(icon_file):
        new_window.iconbitmap(icon_file)
    new_window.transient(root)
    app = SyntenyGUI(new_window)
    bring_to_front(new_window)

# ===================== 菜单栏 =====================
menu_bar = Menu(root, bg=MENU_BG, fg=MENU_FG, activebackground=MENU_ACTIVE_BG, activeforeground=MENU_ACTIVE_FG)
root.config(menu=menu_bar)

# 文件预处理菜单
file_menu = Menu(menu_bar, tearoff=0, bg=MENU_BG, fg=MENU_FG, activebackground=MENU_ACTIVE_BG, activeforeground=MENU_ACTIVE_FG)
menu_bar.add_cascade(label="文件预处理", menu=file_menu)
file_menu.add_command(label="ID文件转换", command=open_id_gui)
file_menu.add_command(label="信息文件转换", command=open_file_conversion)

# 工具菜单
edit_menu = Menu(menu_bar, tearoff=0, bg=MENU_BG, fg=MENU_FG, activebackground=MENU_ACTIVE_BG, activeforeground=MENU_ACTIVE_FG)
menu_bar.add_cascade(label="工具", menu=edit_menu)
edit_menu.add_command(label="基因匹配", command=open_gene_match)
edit_menu.add_command(label="基因匹配pro", command=gene_tool_pro)
edit_menu.add_separator()
# 可视化二级菜单
visual_menu = Menu(edit_menu, tearoff=0, bg=MENU_BG, fg=MENU_FG, activebackground=MENU_ACTIVE_BG, activeforeground=MENU_ACTIVE_FG)
edit_menu.add_cascade(label="可视化", menu=visual_menu)
visual_menu.add_command(label="基因关联可视化", command=open_cv_link)
visual_menu.add_command(label="其他可视化", command=lambda: messagebox.showinfo("其他可视化", "您居然注意到了这个功能！\n没错，这个功能还在开发。\n敬请期待！"))

# 帮助菜单
help_menu = Menu(menu_bar, tearoff=0, bg=MENU_BG, fg=MENU_FG, activebackground=MENU_ACTIVE_BG, activeforeground=MENU_ACTIVE_FG)
menu_bar.add_cascade(label="帮助", command=open_help)

# 关于菜单
pertain_menu = Menu(menu_bar, tearoff=0, bg=MENU_BG, fg=MENU_FG, activebackground=MENU_ACTIVE_BG, activeforeground=MENU_ACTIVE_FG)
menu_bar.add_cascade(label="关于",  command=lambda: messagebox.showinfo("关于", "gene_operations\nauthor：sea wears sky is color"))


# ===================== 主界面提示 =====================
label = tk.Label(root, text="欢迎使用基因工具！\n请通过菜单选择功能。", font=("微软雅黑", 14))
label.pack(expand=True)

# ===================== 启动主循环 =====================
root.mainloop()

