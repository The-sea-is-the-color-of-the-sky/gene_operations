import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os
from datetime import datetime
from tkinter import ttk
import threading


# ========== 数据解析函数 ==========
def parse_collinearity(file_path, output_folder, self=None):
    """解析文本或表格文件并导出 Excel，新增列名 'id'。

    支持:
    - .txt / .collinearity 等文本文件
    - .xls / .xlsx / .csv 表格文件

    规则:
    - 文本文件会逐行解析，提取词组（按空白分割）。
    - 表格文件会读取所有单元格内容，拆分所有单词或词组为独立的 id。
    """
    if self:
        try:
            self.update_status("正在运行...")
        except Exception:
            pass

    ext = os.path.splitext(file_path)[1].lower()
    data = []

    if ext in [".txt", ".collinearity"]:
        # ====== 文本文件解析 ======
        block_id = "Unassigned"
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                if line.startswith("## Alignment"):
                    m = re.search(r"Alignment\s+(\d+)", line)
                    block_id = m.group(1) if m else block_id
                    continue
                if line.startswith("#"):
                    continue

                # 按空白分割整行的所有词组
                parts = line.split()
                for word in parts:
                    data.append([word])

    elif ext in [".xls", ".xlsx", ".csv"]:
        # ====== 表格文件解析 ======
        if ext == ".csv":
            df = pd.read_csv(file_path, dtype=str, header=None)
        else:
            df = pd.read_excel(file_path, dtype=str, header=None)

        # 遍历所有单元格内容，提取词组
        for row in df.itertuples(index=False):
            for cell in row:
                if pd.isna(cell):
                    continue
                parts = str(cell).strip().split()
                for word in parts:
                    data.append([word])
    else:
        raise ValueError("不支持的文件格式，只能处理 .txt, .xls, .xlsx, .csv")

    if not data:
        raise ValueError("未解析到有效词组")

    # ====== 转为 DataFrame ======
    df_out = pd.DataFrame(data, columns=["id"])

    # ====== 输出 Excel 文件 ======
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = os.path.join(output_folder, f"{base_name}_id_{timestamp}.xlsx")

    df_out.to_excel(output_file, index=False)
    return output_file


# ========== GUI ==========
class id_UI:
    def __init__(self, master):
        self.master = master
        master.title("ID文件转换")
        master.geometry("500x250")
        master.resizable(False, False)

        frame_file = ttk.Frame(master)
        frame_file.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame_file, text="选择文件:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_file = ttk.Entry(frame_file)
        self.entry_file.grid(row=1, column=0, columnspan=2, sticky="we", padx=5)
        ttk.Button(frame_file, text="选择", command=self.browse_file).grid(row=0, column=1, padx=5, pady=5, sticky="e")
        frame_file.columnconfigure(1, weight=1)

        frame_output = ttk.Frame(master)
        frame_output.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame_output, text="输出目录:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entry_output = tk.Entry(frame_output)
        self.entry_output.grid(row=3, column=0, columnspan=2, sticky="we", padx=5)
        ttk.Button(frame_output, text="选择", command=self.browse_output).grid(row=2, column=1, padx=5, pady=5, sticky="e")
        frame_output.columnconfigure(1, weight=1)

        frame_button = ttk.Frame(master)
        frame_button.pack(pady=15)
        ttk.Button(frame_button, text="开始转换", width=10, command=self.start_conversion_thread).grid(row=0, column=0, padx=5, pady=5)
        self.status_label = ttk.Label(frame_button, text="等待运行", wraplength=400, justify="left")
        self.status_label.grid(row=1, column=0, padx=5, pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[
                ("支持的文件", "*.txt *.csv *.xls *.xlsx"),
                ("文本文件", "*.txt"),
                ("Excel文件", "*.xls *.xlsx"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)

    def browse_output(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, folder)

    def start_conversion_thread(self):
        t = threading.Thread(target=self.convert)
        t.daemon = True
        t.start()

    def convert(self):
        file_path = self.entry_file.get()
        output_folder = self.entry_output.get()
        if not file_path or not os.path.exists(file_path):
            self.update_status("错误: 请选择有效的输入文件")
            return
        if not output_folder or not os.path.exists(output_folder):
            self.update_status("错误: 请选择有效的输出目录")
            return
        try:
            output_file = parse_collinearity(file_path, output_folder, self)
            self.update_status(f"转换完成，输出文件:\n{output_file}")
        except Exception as e:
            self.update_status(f"错误: 转换失败 ({e})")

    def update_status(self, text):
        if self.status_label:
            try:
                self.status_label.config(text=text)
                self.master.update_idletasks()
            except Exception:
                self.master.after(0, lambda: self.status_label.config(text=text))


if __name__ == "__main__":
    root = tk.Tk()
    app = id_UI(root)
    root.mainloop()
