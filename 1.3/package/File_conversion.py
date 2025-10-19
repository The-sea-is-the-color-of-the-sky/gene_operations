import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os
from datetime import datetime
from tkinter import ttk
import threading  # ✅ 新增

# ========== 数据解析函数 ==========
def parse_collinearity(file_path, output_folder, self=None):
    """解析 .collinearity 文件并导出 Excel"""
    data = []
    block_id = "Unassigned"
    if self:
        self.update_status("正在运行...")

    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            if line.startswith("## Alignment"):
                match = re.search(r"Alignment\s+(\d+)", line)
                block_id = match.group(1) if match else block_id
                continue

            if line.startswith("#"):
                continue

            m = re.match(r"^(\d+-\s*\d+:)\s+(\S+)\s+(\S+)\s+(\S+)$", line)
            if m:
                block_idx, geneA, geneB, evalue = m.groups()
                data.append([block_idx, geneA, geneB, evalue])
                continue

            parts = line.split()
            if len(parts) == 3:
                geneA, geneB, evalue = parts
                data.append([block_id, geneA, geneB, evalue])
            elif len(parts) == 2:
                geneA, evalue = parts
                data.append([block_id, geneA, "NA", evalue])

    df = pd.DataFrame(data, columns=["Block", "GeneA", "GeneB", "E-value"])

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1].replace(".", "")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = os.path.join(output_folder, f"{base_name}_{ext}_{timestamp}.xlsx")
    df.to_excel(output_file, index=False)
    return output_file


# ========== GUI ==========
class FileConversionUI:
    def __init__(self, master):
        self.master = master
        master.title("信息文件转换")
        master.geometry("500x250")
        master.resizable(False, False)

        # ========== 输入文件 ==========
        frame_file = ttk.Frame(master)
        frame_file.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_file, text="选择文件:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_file = ttk.Entry(frame_file)
        self.entry_file.grid(row=1, column=0, columnspan=2, sticky="we", padx=5)
        ttk.Button(frame_file, text="选择", command=self.browse_file).grid(row=0, column=1, padx=5, pady=5, sticky="e")

        frame_file.columnconfigure(1, weight=1)

        # ========== 输出目录 ==========
        frame_output = ttk.Frame(master)
        frame_output.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_output, text="输出目录:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entry_output = tk.Entry(frame_output)
        self.entry_output.grid(row=3, column=0, columnspan=2, sticky="we", padx=5)
        ttk.Button(frame_output, text="选择", command=self.browse_output).grid(row=2, column=1, padx=5, pady=5, sticky="e")

        frame_output.columnconfigure(1, weight=1)

        # ========== 转换按钮 ==========
        frame_button = ttk.Frame(master)
        frame_button.pack(pady=15)
        ttk.Button(frame_button, text="开始转换", width=10, command=self.start_conversion_thread).grid(row=0, column=0, padx=5, pady=5)
        self.status_label = ttk.Label(
            frame_button,
            text="等待运行",
            wraplength=400,       # 自动换行宽度
            justify="left"        # 左对齐
        )
        self.status_label.grid(row=1, column=0, padx=5, pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 .collinearity 文件",
            filetypes=[("Collinearity files", "*.collinearity"), ("All files", "*.*")]
        )
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)

    def browse_output(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, folder)

    # ===================== 多线程启动 =====================
    def start_conversion_thread(self):
        """使用线程执行耗时任务，保持 GUI 响应"""
        t = threading.Thread(target=self.convert)
        t.daemon = True  # 窗口关闭时线程自动结束
        t.start()

    def convert(self):
        file_path = self.entry_file.get()
        output_folder = self.entry_output.get()
        if not file_path or not os.path.exists(file_path):
            self.update_status("错误,请选择有效的 .collinearity 文件")
            return
        if not output_folder or not os.path.exists(output_folder):
            self.update_status("错误,请选择有效的输出目录")
            return
        try:
            output_file = parse_collinearity(file_path, output_folder, self)
            # 弹窗必须在主线程调用
            self.update_status(f"转换完成 输出文件:\n{output_file}")
        except Exception as e:
            self.update_status(f"错误 转换失败:{e}")

    def update_status(self, text):
        if self.status_label:
            self.status_label.config(text=text)
            self.master.update_idletasks()


# ========== 启动 GUI ==========
if __name__ == "__main__":
    root = tk.Tk()
    app = FileConversionUI(root)
    root.mainloop()
