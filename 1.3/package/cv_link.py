import tkinter as tk
from tkinter import filedialog, ttk
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading
from package.package_tool.cv_tool import plot_chord, plot_heatmap, plot_network


class CV_LINK_GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("基因联系可视化")
        self.root.geometry("1200x650")
        self.root.minsize(800, 500)  # 支持调整窗口大小，设置最小尺寸

        self.path_var = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.x_column = tk.StringVar()
        self.y_column = tk.StringVar()
        self.date_column = tk.StringVar()
        self.status_var = tk.StringVar(value="等待任务开始...")

        self.df = None
        self.figure = None
        self.canvas = None

        self.cv_link_gui()

    def cv_link_gui(self):
        cv_left_panel = ttk.Frame(self.root, width=450, height=500)
        cv_left_panel.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)

        self.cv_visual = ttk.LabelFrame(self.root, text="图形预览", padding=10)
        self.cv_visual.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        cv_left_panel.grid_propagate(False)

        cv_path = ttk.LabelFrame(cv_left_panel, text="文件选择", padding=10)
        cv_path.pack(fill="x", padx=0, pady=(0, 5))

        cv_run = ttk.LabelFrame(cv_left_panel, text="运行", padding=10)
        cv_run.pack(fill="both", expand=True, padx=0, pady=5)

        cv_path.grid_columnconfigure(1, weight=1)

        ttk.Label(cv_path, text="信息文件路径:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(cv_path, textvariable=self.path_var, width=30).grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="we")
        ttk.Button(cv_path, text="选择文件", command=lambda: self.load_file(self.path_var)).grid(row=0, column=1, padx=5, pady=5, sticky="we")

        ttk.Label(cv_path, text="保存图片位置:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(cv_path, textvariable=self.output_dir, width=30).grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="we")
        ttk.Button(cv_path, text="选择位置", command=lambda: self.browse_output(self.output_dir)).grid(row=2, column=1, padx=5, pady=5, sticky="we")

        ttk.Label(cv_path, text="节点1列:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.x_combobox = ttk.Combobox(cv_path, textvariable=self.x_column, state="readonly", width=15)
        self.x_combobox.grid(row=4, column=1, sticky="we")

        ttk.Label(cv_path, text="节点2列:").grid(row=6, column=0, padx=5, pady=5, sticky="w")
        self.y_combobox = ttk.Combobox(cv_path, textvariable=self.y_column, state="readonly", width=15)
        self.y_combobox.grid(row=6, column=1, sticky="we")

        ttk.Label(cv_path, text="数据（权重）列:").grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.date_combobox = ttk.Combobox(cv_path, textvariable=self.date_column, state="readonly", width=15)
        self.date_combobox.grid(row=8, column=1, sticky="we")

        cv_run.grid_columnconfigure((0, 1, 2, 3), weight=1)
        ttk.Label(cv_run, text="图片类型").grid(row=0, column=0, padx=5, pady=5, sticky="we")
        self.function_combo = ttk.Combobox(cv_run, values=["弦图", "关系网络图", "热图"], state="readonly", width=10)
        self.function_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.function_combo.current(0)

        ttk.Button(cv_run, text="生成图形", command=self.run_cv).grid(row=1, column=0, padx=5, pady=5, sticky="we")
        ttk.Button(cv_run, text="保存图形", command=self.save_cv).grid(row=1, column=1, padx=5, pady=5, sticky="we")

        status_label = tk.Label(cv_run, textvariable=self.status_var, anchor="w", wraplength=200, justify="left")
        status_label.grid(row=3, column=0, columnspan=4, sticky="we")

    def load_file(self, path_var):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx;*.xls;*.csv")])
        if not file_path:
            return

        self.path_var.set(file_path)
        try:
            if file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)

            cols = list(self.df.columns)
            cols_with_blank = ["--- 请选择 ---"] + cols

            self.x_combobox["values"] = cols_with_blank
            self.y_combobox["values"] = cols_with_blank
            self.date_combobox["values"] = cols_with_blank
            self.x_combobox.current(0)
            self.y_combobox.current(0)
            self.date_combobox.current(0)

            self.status_var.set("✅ 文件加载成功")
        except Exception as e:
            self.status_var.set(f"❌ 加载文件失败: {e}")
        self.root.update_idletasks()

    def browse_output(self, output_dir):
        path = filedialog.askdirectory(title="选择输出文件位置")
        if path:
            output_dir.set(path)

    def run_cv(self):
        def generate_graph():
            graph_type = self.function_combo.get()
            try:
                if graph_type == "弦图":
                    fig = plot_chord(self, self.x_column.get(), self.y_column.get(), self.show_plot)
                elif graph_type == "关系网络图":
                    fig = plot_network(self, self.x_column.get(), self.y_column.get(), self.date_column.get(), self.show_plot)
                elif graph_type == "热图":
                    fig = plot_heatmap(self, self.x_column.get(), self.y_column.get(), self.date_column.get(), self.show_plot)
                else:
                    self.root.after(0, lambda: self.status_var.set("❌ 未识别的图形类型"))
                    return

                if fig:
                    self.root.after(0, lambda: self.show_plot(fig))
                else:
                    self.root.after(0, lambda: self.status_var.set("❌ 图形生成失败"))
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"❌ 图形生成失败: {e}"))
            self.root.after(0, self.root.update_idletasks)

        if self.df is None:
            self.status_var.set("❌ 请先加载数据文件！")
            return

        if self.x_column.get() in ["--- 请选择 ---", ""] or self.y_column.get() in ["--- 请选择 ---", ""]:
            self.status_var.set("❌ 请选择 节点1列 和 节点2列 作为绘图依据！")
            return

        self.status_var.set("⏳ 正在生成图形...")
        threading.Thread(target=generate_graph, daemon=True).start()

    def show_plot(self, fig):
        self.figure = fig
        if self.canvas:
            self.canvas.get_tk_widget().destroy()
        self.canvas = FigureCanvasTkAgg(fig, master=self.cv_visual)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.status_var.set("✅ 图形生成完成")

    def save_cv(self):
        import re
        if not self.figure:
            self.status_var.set("❌ 请先生成图形！")
            return
        save_dir = self.output_dir.get()
        if not save_dir:
            self.status_var.set("❌ 请先选择保存位置！")
            return

        input_name = os.path.splitext(os.path.basename(self.path_var.get()))[0]
        input_name = re.sub(r'[^\w\-]', '_', input_name)  # 替换非法字符
        plot_type = self.function_combo.get()
        save_path = os.path.join(save_dir, f"{input_name}_{plot_type}.png")

        try:
            self.figure.savefig(save_path, dpi=300, bbox_inches="tight")
            self.status_var.set(f"✅ 图片保存成功: {save_path}")
        except Exception as e:
            self.status_var.set(f"❌ 保存图片失败: {e}")
        self.root.update_idletasks()


if __name__ == "__main__":
    try:
        import networkx as nx
    except ImportError:
        print("NetworkX 库未安装，请运行: pip install networkx")
        exit()

    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    root = tk.Tk()
    app = CV_LINK_GUI(root)
    root.mainloop()