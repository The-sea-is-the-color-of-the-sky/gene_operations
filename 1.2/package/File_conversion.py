import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os
from datetime import datetime

# ========== 数据解析函数 ==========
def parse_collinearity(file_path, output_folder):
    """解析 .collinearity 文件并导出 Excel"""
    data = []
    block_id = "Unassigned"

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
        master.title("Collinearity 文件转换")
        master.geometry("700x200")
        master.resizable(False, False)

        # ========== 输入文件 ==========
        frame_file = tk.Frame(master)
        frame_file.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_file, text="选择文件:").grid(row=0, column=0, sticky="w")
        self.entry_file = tk.Entry(frame_file)
        self.entry_file.grid(row=0, column=1, sticky="we", padx=5)
        tk.Button(frame_file, text="选择", command=self.browse_file).grid(row=0, column=2)

        frame_file.columnconfigure(1, weight=1)

        # ========== 输出目录 ==========
        frame_output = tk.Frame(master)
        frame_output.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_output, text="输出目录:").grid(row=0, column=0, sticky="w")
        self.entry_output = tk.Entry(frame_output)
        self.entry_output.grid(row=0, column=1, sticky="we", padx=5)
        tk.Button(frame_output, text="选择", command=self.browse_output).grid(row=0, column=2)

        frame_output.columnconfigure(1, weight=1)

        # ========== 转换按钮 ==========
        frame_button = tk.Frame(master)
        frame_button.pack(pady=15)
        tk.Button(frame_button, text="开始转换", width=20, command=self.convert).pack()

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

    def convert(self):
        file_path = self.entry_file.get()
        output_folder = self.entry_output.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("错误", "请选择有效的 .collinearity 文件")
            return
        if not output_folder or not os.path.exists(output_folder):
            messagebox.showerror("错误", "请选择有效的输出目录")
            return
        try:
            output_file = parse_collinearity(file_path, output_folder)
            messagebox.showinfo("完成", f"转换完成！输出文件:\n{output_file}")
        except Exception as e:
            messagebox.showerror("错误", f"转换失败:\n{e}")

# ========== 启动 GUI ==========
if __name__ == "__main__":
    root = tk.Tk()
    app = FileConversionUI(root)
    root.mainloop()
