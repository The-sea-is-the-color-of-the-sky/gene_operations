import os
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
from collections import defaultdict
# 新增：调用外部匹配模块
from package.package_tool import pro_tool


class GeneProApp:
    def __init__(self, root):
        self.root = root
        self.root.title("基因匹配pro")
        self.root.geometry("720x560")

        self.worker_thread = None
        self.stop_flag = False

        self._build_ui()
        

    # ---------- UI 构建 ----------
    def _build_ui(self):
        frame_function = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        frame_files = ttk.LabelFrame(self.root, text="列选择", padding=10)
        frame_columns = ttk.LabelFrame(self.root, text="参数选择", padding=10)
        frame_progress = ttk.LabelFrame(self.root, text="运行状态", padding=10)
        for f in [frame_function, frame_files, frame_columns, frame_progress]:
            f.pack(fill="x", pady=5)



        # === 文件选择 ===
        self._add_file_selector(frame_function, "目标表 A ：", 0, 'A')
        self._add_file_selector(frame_function, "信息表 B ：", 1, 'B')
        self._add_folder_selector(frame_function, "输出文件夹：", 2)

        # === 下拉选择列 ===
        ttk.Label(frame_files, text="A 表目标列：").grid(column=0, row=0, padx=5, pady=5, sticky='e')
        self.cmb_a_col = ttk.Combobox(frame_files, values=[], width=30, state='readonly')
        self.cmb_a_col.grid(column=1, row=0, padx=5, pady=5, sticky='w')

        ttk.Label(frame_files, text="B 表匹配列 (B1)：").grid(column=0, row=1, padx=5, pady=5, sticky='e')
        self.cmb_b1 = ttk.Combobox(frame_files, values=[], width=30, state='readonly')
        self.cmb_b1.grid(column=1, row=1, padx=5, pady=5, sticky='w')

        ttk.Label(frame_files, text="B 表信息列 (B2)：").grid(column=0, row=2, padx=5, pady=5, sticky='e')
        self.cmb_b2 = ttk.Combobox(frame_files, values=[], width=30, state='readonly')
        self.cmb_b2.grid(column=1, row=2, padx=5, pady=5, sticky='w')


        # 新增：token 交集比率下拉（0.5 - 1.0）
        ttk.Label(frame_columns, text="token 交集比率阈值：").grid(column=0, row=0, padx=5, pady=5, sticky='e')
        self.cmb_ratio = ttk.Combobox(frame_columns, values=["0.5","0.6","0.7","0.8","0.9","1.0"],
                                      width=10, state='readonly')
        self.cmb_ratio.set("0.8")
        self.cmb_ratio.grid(column=1, row=0, padx=5, pady=5, sticky='w')

        self.export_other = tk.IntVar(value=1)
        # 保留 checkbutton 引用，_set_ui_state 中需要访问
        self.chk_export_other = ttk.Checkbutton(frame_columns, text="导出 B 表其他列（排除 B1, B2）",
                                                variable=self.export_other)
        # 保持原位稍往下
        self.chk_export_other.grid(column=2, row=0, padx=5, pady=5, sticky='w')
        # === 操作按钮 ===
        btn_frame = ttk.Frame(frame_columns)
        btn_frame.grid(column=0, row=1, columnspan=3, padx=5, pady=5, sticky='w')
        self.btn_start = ttk.Button(btn_frame, text="开始匹配", command=self.start_matching)
        self.btn_start.pack(side='left', padx=6)
        self.btn_stop = ttk.Button(btn_frame, text="停止（安全）",
                                   command=self.stop_matching, state='disabled')
        self.btn_stop.pack(side='left', padx=6)

        # === 进度条与状态（行号向下移动）===
        self._add_progress(frame_progress, "主进度：", 0)
 
        self.status_label = ttk.Label(frame_progress, text="准备就绪。", anchor='w')
        self.status_label.grid(column=1, row=1, columnspan=3, sticky='we')

    # ---------- UI 辅助 ----------
    def _add_file_selector(self, parent, label, row, which):
        pad = {'padx': 6, 'pady': 6}
        ttk.Label(parent, text=label).grid(column=0, row=row, sticky='w', **pad)
        entry = ttk.Entry(parent, width=68)
        entry.grid(column=1, row=row, sticky='w', **pad)
        ttk.Button(parent, text="浏览", command=lambda: self._browse_file(which)).grid(column=2, row=row, **pad)
        if which == 'A':
            self.entry_a = entry
        else:
            self.entry_b = entry
        # 当用户手动粘贴/输入路径并离开 Entry（失去焦点）时，尝试自动读取表头并填充下拉
        entry.bind("<FocusOut>", lambda e, w=which: self._on_entry_focus_out(w))

    def _add_folder_selector(self, parent, label, row):
        pad = {'padx': 6, 'pady': 6}
        ttk.Label(parent, text=label).grid(column=0, row=row, sticky='w', **pad)
        self.entry_out = ttk.Entry(parent, width=68)
        self.entry_out.grid(column=1, row=row, sticky='w', **pad)
        ttk.Button(parent, text="浏览", command=self._browse_output).grid(column=2, row=row, **pad)

    def _add_progress(self, parent, label, row, sub=False):
        pad = {'padx': 6, 'pady': 6}
        ttk.Label(parent, text=label).grid(column=0, row=row, sticky='w', **pad)
        bar = ttk.Progressbar(parent, orient='horizontal', length=500, mode='determinate')
        bar.grid(column=1, row=row, sticky='w', **pad)
        if sub:
            self.progress_sub = bar
        else:
            self.progress_main = bar

    # ---------- 文件操作 ----------
    def _browse_file(self, which):
        ft = [("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        path = filedialog.askopenfilename(title="选择文件", filetypes=ft)
        if not path:
            return
        (self.entry_a if which == 'A' else self.entry_b).delete(0, tk.END)
        (self.entry_a if which == 'A' else self.entry_b).insert(0, path)
        # 浏览后自动尝试读取表头并填充对应下拉，提升体验（线程外、轻量 nrows=0）
        try:
            self._load_cols_for_file(which, path)
        except Exception:
            # 允许失败但不阻塞用户
            pass

    def _browse_output(self):
        """选择输出目录并填充到输出 Entry"""
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            try:
                self.entry_out.delete(0, tk.END)
                self.entry_out.insert(0, path)
            except Exception:
                # 若 UI 控件不存在或其他异常，静默处理以避免阻塞
                pass

    def _load_cols_for_file(self, which, path):
        """智能读取文件表头并填充对应 Combobox（用于浏览时自动加载）"""
        ext = os.path.splitext(path)[1].lower()
        # 仅读取表头（nrows=0），减少 IO 与内存
        if ext in ['.xls', '.xlsx']:
            dfh = pd.read_excel(path, nrows=0)
        else:
            # 尝试 utf-8，再 gbk，兼容常见 CSV 编码
            try:
                dfh = pd.read_csv(path, nrows=0, encoding='utf-8')
            except Exception:
                dfh = pd.read_csv(path, nrows=0, encoding='gbk')
        cols = list(dfh.columns.astype(str))
        if which == 'A':
            self.cmb_a_col['values'] = cols
            if cols:
                self.cmb_a_col.set(cols[0])
        else:
            # B 文件同时填充 B1/B2
            self.cmb_b1['values'] = cols
            self.cmb_b2['values'] = cols
            if cols:
                # 智能默认：若有 >=3 列，取前两列为 B1/B2；否则都用第1列
                if len(cols) >= 3:
                    self.cmb_b1.set(cols[1])
                    self.cmb_b2.set(cols[2])
                else:
                    self.cmb_b1.set(cols[0])
                    self.cmb_b2.set(cols[0])

    # ---------- 加载列 ----------
    def load_columns(self):
        a_path, b_path = self.entry_a.get().strip(), self.entry_b.get().strip()
        if not (a_path and b_path):
            messagebox.showwarning("缺少文件", "请先选择 A 表和 B 表文件")
            return
        try:
            # 使用仅读取表头的助手函数，快速响应
            self._load_cols_for_file('A', a_path)
            self._load_cols_for_file('B', b_path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        self.root.after(0, lambda: self.status_label.config(text="列加载完成。"))

    def _read_table(self, path):
        ext = os.path.splitext(path)[1].lower()
        # 读取完整表用于匹配时使用（保留 dtype=str），对 CSV 做编码兼容处理
        if ext in ['.xls', '.xlsx']:
            df = pd.read_excel(path, dtype=str)
        else:
            try:
                df = pd.read_csv(path, dtype=str, encoding='utf-8')
            except Exception:
                df = pd.read_csv(path, dtype=str, encoding='gbk')
        return df.fillna('').astype(str)

    # 新增：规范化与分词，便于快速匹配与倒排索引
    def _normalize(self, s: str) -> str:
        return s.strip().lower()

    def _tokens(self, s: str):
        # 仅提取字母/数字/下划线序列作为 token
        return re.findall(r'\w+', s.lower())

    # ---------- 启动匹配 ----------
    def start_matching(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("进行中", "已有任务在运行。")
            return

        # 在启动前尝试自动加载列（若下拉尚未填充），以避免用户必须手动点击“加载 A/B 表列”
        a_path = self.entry_a.get().strip()
        b_path = self.entry_b.get().strip()
        try:
            if a_path and (not self.cmb_a_col['values']):
                self._load_cols_for_file('A', a_path)
            if b_path and (not self.cmb_b1['values'] or not self.cmb_b2['values']):
                self._load_cols_for_file('B', b_path)
        except Exception:
            pass

        # 获取阈值，默认已有设置；传入 worker 以线程安全
        try:
            ratio_val = float(self.cmb_ratio.get())
        except Exception:
            ratio_val = 0.6

        params = (
            self.entry_a.get().strip(),
            self.entry_b.get().strip(),
            self.entry_out.get().strip(),
            self.cmb_a_col.get().strip(),
            self.cmb_b1.get().strip(),
            self.cmb_b2.get().strip(),
            bool(self.export_other.get()),
            ratio_val
        )

        if not all(params[:-2]):  # 确保 A/B 文件、输出目录和列已选择（不检查 ratio）
            messagebox.showwarning("缺少参数", "请确保 A/B 文件、输出目录和列已选择。")
            return

        self._set_ui_state(False)
        self.stop_flag = False
        self.progress_main['value'] = 0
        self.status_label.config(text="准备开始匹配...")

        self.worker_thread = threading.Thread(target=self._worker_safe_wrapper, args=params, daemon=True)
        self.worker_thread.start()
        self.btn_stop.config(state='normal')

    def stop_matching(self):
        if messagebox.askyesno("确认停止", "确认安全停止当前任务？"):
            self.stop_flag = True
            self.root.after(0, lambda: self.status_label.config(text="已请求停止，等待当前操作完成..."))
            self.btn_stop.config(state='disabled')

    def _set_ui_state(self, enabled):
        state = 'normal' if enabled else 'disabled'
        widgets = [
            self.entry_a, self.entry_b, self.entry_out,
            self.cmb_a_col, self.cmb_b1, self.cmb_b2,
            self.chk_export_other if hasattr(self, 'chk_export_other') else None,
            self.btn_start
        ]
        for w in widgets:
            if w:
                w.config(state=state)
        if enabled:
            self.btn_stop.config(state='disabled')

    # ---------- 线程安全包装 ----------
    def _worker_safe_wrapper(self, *args):
        try:
            self._worker(*args)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("运行错误", str(e)))
            self.root.after(0, lambda: self.status_label.config(text=f"运行错误：{e}"))
        finally:
            self.root.after(0, lambda: self._set_ui_state(True))
            self.root.after(0, lambda: self.btn_stop.config(state='disabled'))

    # ---------- 主匹配逻辑 ----------
    def _worker(self, a_path, b_path, out_folder, a_col, b1, b2, export_other, threshold_ratio):
        # UI 回调包装：将外部回调通过 root.after 发送到主线程
        def ui_progress(pct):
            try:
                self.root.after(0, lambda: self.progress_main.config(value=int(pct)))
            except Exception:
                pass

        def ui_status(text):
            try:
                self.root.after(0, lambda: self.status_label.config(text=text))
            except Exception:
                pass

        out_file = pro_tool.pro_match(
            a_path, b_path, out_folder, a_col, b1, b2,
            export_other=export_other,
            threshold_ratio=threshold_ratio,
            progress_cb=ui_progress,
            status_cb=ui_status,
            stop_flag_getter=lambda: self.stop_flag
        )

        if out_file:
            self.root.after(0, lambda: messagebox.showinfo("完成", f"输出文件：\n{out_file}"))
        else:
            if self.stop_flag:
                self.root.after(0, lambda: self.status_label.config(text="任务已中止。"))
            else:
                self.root.after(0, lambda: messagebox.showinfo("无结果", "没有生成任何匹配结果。"))

    # 当用户在文件路径 Entry 失去焦点时，尝试自动读取并填充列下拉（容错，不抛错）
    def _on_entry_focus_out(self, which):
        try:
            path = (self.entry_a if which == 'A' else self.entry_b).get().strip()
            if path:
                try:
                    self._load_cols_for_file(which, path)
                except Exception:
                    # 读取表头失败时静默处理
                    pass
        except Exception:
            # 吞掉任何异常，避免阻塞 UI
            pass

# ---------- 主入口 ----------
def main():
    root = tk.Tk()
    GeneProApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
