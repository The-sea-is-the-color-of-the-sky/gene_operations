import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import pandas as pd
import package.gene_operations as go
import platform
import subprocess
from datetime import datetime


class GeneToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("基因工具")
        self.root.geometry("750x620")
        self.root.resizable(False, False)

        # ---------------- 变量 ----------------
        self.file_a_path = tk.StringVar()
        self.file_b_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.gene_col = tk.StringVar()
        self.info_a_col = tk.StringVar()
        self.info_b_col = tk.StringVar()

        self.selected_function = None
        self.selected_match_mode = None
        self.fuzzy_var = tk.BooleanVar(value=False)

        # ---------------- GUI ----------------
        self.build_gui()

    def build_gui(self):
        frame_function = ttk.LabelFrame(self.root, text="选择功能", padding=10)
        frame_files = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        frame_columns = ttk.LabelFrame(self.root, text="列名输入", padding=10)
        frame_progress = ttk.LabelFrame(self.root, text="运行进度", padding=10)

        for f in [frame_function, frame_files, frame_columns, frame_progress]:
            f.pack(fill="x", pady=5)

        # 功能选择保持原样
        ttk.Label(frame_function, text="功能类型:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.function_combo = ttk.Combobox(frame_function, values=["基因查询", "基因匹配"], state="readonly", width=20)
        self.function_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.match_mode_label = ttk.Label(frame_function, text="排列方式:")
        self.match_mode_combo = ttk.Combobox(frame_function, values=["横向排列", "竖向排列"], state="readonly", width=18)
        self.fuzzy_check = ttk.Checkbutton(frame_function, text="启用模糊匹配", variable=self.fuzzy_var)

        self.match_mode_label.grid_remove()
        self.match_mode_combo.grid_remove()
        self.fuzzy_check.grid_remove()

        self.function_combo.bind("<<ComboboxSelected>>", self.on_function_select)
        self.match_mode_combo.bind("<<ComboboxSelected>>", self.on_match_mode_select)

        # 文件选择
        ttk.Label(frame_files, text="填入的表格文件路径:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_files, textvariable=self.file_a_path, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame_files, text="选择文件", command=lambda: self.load_file(self.file_a_path, target="a")).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame_files, text="基因信息表格文件路径:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_files, textvariable=self.file_b_path, width=50).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame_files, text="选择文件", command=lambda: self.load_file(self.file_b_path, target="b")).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(frame_files, text="输出文件夹:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame_files, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame_files, text="选择文件夹", command=self.choose_output_dir).grid(row=2, column=2, padx=5, pady=5)

        # 列名输入
        ttk.Label(frame_columns, text="填入表格目标列名:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.gene_combobox = ttk.Combobox(frame_columns, textvariable=self.gene_col, width=40, state="readonly")
        self.gene_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame_columns, text="信息表格的A列名:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.info_a_combobox = ttk.Combobox(frame_columns, textvariable=self.info_a_col, width=40, state="readonly")
        self.info_a_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame_columns, text="信息表格的B列名:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.info_b_combobox = ttk.Combobox(frame_columns, textvariable=self.info_b_col, width=40, state="readonly")
        self.info_b_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # 进度条：主 + 子
        ttk.Label(frame_progress, text="主进度:").pack(anchor="w", padx=5)
        self.main_progress = ttk.Progressbar(frame_progress, orient="horizontal", length=700, mode="determinate", maximum=100)
        self.main_progress.pack(fill="x", padx=5, pady=2)

        ttk.Label(frame_progress, text="子进度:").pack(anchor="w", padx=5)
        self.sub_progress = ttk.Progressbar(frame_progress, orient="horizontal", length=700, mode="determinate", maximum=100)
        self.sub_progress.pack(fill="x", padx=5, pady=2)

        self.status_label = ttk.Label(frame_progress, text="等待运行...", anchor="w")
        self.status_label.pack(fill="x", padx=5, pady=(5,2))

        ttk.Button(frame_progress, text="运行", command=self.run_function).pack(pady=10)

    # ---------------- 事件响应 ----------------
    def on_function_select(self, event):
        self.selected_function = self.function_combo.get()
        if self.selected_function == "基因匹配":
            self.match_mode_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
            self.match_mode_combo.grid(row=0, column=3, padx=5, pady=5, sticky="w")
            self.fuzzy_check.grid(row=0, column=4, padx=5, pady=5, sticky="w")
        else:
            self.match_mode_label.grid_remove()
            self.match_mode_combo.grid_remove()
            self.fuzzy_check.grid_remove()
            self.selected_match_mode = None

    def on_match_mode_select(self, event):
        self.selected_match_mode = self.match_mode_combo.get()

    # ---------------- 文件操作 ----------------
    def load_file(self, var, target):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            var.set(path)
            self.update_columns(path, target)

    def choose_output_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def update_columns(self, filepath, target):
        try:
            df = pd.read_excel(filepath, nrows=0)
            cols = list(df.columns)
            if target == "a":
                # 自动优先选择GeneA
                if "GeneA" in cols:
                    self.gene_col.set("GeneA")
                else:
                    self.gene_col.set(cols[0] if cols else "")
                self.gene_combobox["values"] = cols
            else:
                # 自动优先选择GeneB
                if "GeneB" in cols:
                    self.info_a_col.set("GeneB")
                    self.info_b_col.set(cols[1] if len(cols)>1 else cols[0])
                else:
                    self.info_a_col.set(cols[0] if len(cols)>0 else "")
                    self.info_b_col.set(cols[1] if len(cols)>1 else "")
                self.info_a_combobox["values"] = cols
                self.info_b_combobox["values"] = cols
        except Exception as e:
            messagebox.showerror("错误", f"读取列名失败: {e}")

    # ---------------- 功能运行 ----------------
    def run_function(self):
        if self.selected_function == "基因查询":
            self.run_in_thread(go.gene_correspondence_with_progress, func_name="基因查询")
        elif self.selected_function == "基因匹配":
            is_fuzzy = self.fuzzy_var.get()
            if self.selected_match_mode == "横向排列":
                if is_fuzzy:
                    self.run_in_thread(go.fuzzy_match_with_progress, func_name="基因匹配_横向_模糊")
                else:
                    self.run_in_thread(go.gene_search_with_progress, func_name="基因匹配_横向_精确")
            elif self.selected_match_mode == "竖向排列":
                if is_fuzzy:
                    self.run_in_thread(go.fuzzy_match_with_progress, vertical=True, func_name="基因匹配_竖向_模糊")
                else:
                    self.run_in_thread(go.classify_genes_with_progress, func_name="基因匹配_竖向_精确")
            else:
                messagebox.showerror("错误", "请选择基因匹配的排列方式！")
        else:
            messagebox.showerror("错误", "请选择功能类型！")

    # ---------------- 线程执行 ----------------
    def run_in_thread(self, func, func_name="功能", **extra_kwargs):
        args = (
            self.file_a_path.get(),
            self.file_b_path.get(),
            self.gene_col.get(),
            self.info_a_col.get(),
            self.info_b_col.get(),
        )
        if not all(args) or not self.output_dir.get():
            messagebox.showerror("错误", "请填写所有选项并选择输出文件夹！")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(self.output_dir.get(), f"output_{func_name}_{timestamp}.xlsx")

        def task():
            try:
                self.update_main_progress(0)
                self.update_sub_progress(0)
                self.update_status("正在运行...")
                func(args[0], args[1], output_file, args[2], args[3], args[4],
                     progress_callback=self.update_main_progress,
                     sub_progress_callback=self.update_sub_progress,
                     set_progress_status=self.update_status,
                     **extra_kwargs)
                self.update_main_progress(100)
                self.update_sub_progress(100)
                self.update_status("操作完成")
                open_choice = messagebox.askyesno("成功", f"操作完成，结果已保存到：\n{output_file}\n是否现在打开？")
                if open_choice:
                    self.open_with_default(output_file)
            except Exception as e:
                self.update_status("运行出错")
                messagebox.showerror("错误", f"运行时出现错误：{e}")
            finally:
                self.root.after(1000, lambda: (self.update_main_progress(0), self.update_sub_progress(0), self.update_status("等待运行...")))

        threading.Thread(target=task, daemon=True).start()

    # ---------------- 更新UI ----------------
    def update_main_progress(self, value):
        self.main_progress['value'] = value
        self.root.update_idletasks()

    def update_sub_progress(self, value):
        self.sub_progress['value'] = value
        self.root.update_idletasks()

    def update_status(self, text):
        self.status_label['text'] = text
        self.root.update_idletasks()

    # ---------------- 打开文件 ----------------
    def open_with_default(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", file_path])
            else:
                subprocess.Popen(["xdg-open", file_path])
        except Exception as e:
            messagebox.showwarning("提示", f"无法打开文件: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GeneToolApp(root)
    root.mainloop()
