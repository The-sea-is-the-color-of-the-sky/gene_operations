import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.collections import LineCollection
import matplotlib.cm as cm
import networkx as nx
import threading
import os

class SyntenyGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("共线性可视化工具")
        self.master.geometry("1000x780")

        # 文件路径
        self.file_path = tk.StringVar()
        file_frame = tk.Frame(master)
        file_frame.pack(fill="x", padx=10, pady=5)
        tk.Entry(file_frame, textvariable=self.file_path, width=60).pack(side="left", fill="x", expand=True)
        tk.Button(file_frame, text="选择文件", command=self.load_file).pack(side="left", padx=5)

        # 样式选择
        tk.Label(master, text="选择绘图样式:", font=("Arial", 11)).pack(anchor="w", padx=10, pady=5)
        self.plot_type = ttk.Combobox(master, state="readonly",
                                      values=["带状图 (Ribbons)", "点图 (Dot plot)",
                                              "Block 统计条形图", "网络图 (Network)",
                                              "E-value 热度图"])
        self.plot_type.current(0)
        self.plot_type.pack(anchor="w", padx=10, pady=5)

        # 按钮
        btn_frame = tk.Frame(master)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="绘制", command=self.plot).pack(side="left", padx=5)
        tk.Button(btn_frame, text="保存图片", command=self.save_plot).pack(side="left", padx=5)
        tk.Button(btn_frame, text="清除画布", command=self.clear_canvas).pack(side="left", padx=5)

        # Matplotlib 图像区域
        self.fig, self.ax = plt.subplots(figsize=(7, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # 悬浮显示注释
        self.hover_annotation = self.ax.annotate("", xy=(0,0), xytext=(15,15),
                                                 textcoords="offset points",
                                                 bbox=dict(boxstyle="round", fc="w"),
                                                 arrowprops=dict(arrowstyle="->"))
        self.hover_annotation.set_visible(False)

        self.df = None

    # ===== 文件加载 =====
    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("表格文件", "*.csv *.xlsx")])
        if path:
            try:
                if path.endswith(".csv"):
                    df = pd.read_csv(path)
                else:
                    df = pd.read_excel(path)

                required = {"GeneA", "GeneB"}
                if not required.issubset(df.columns):
                    raise ValueError(f"缺少必要列: {required}")

                self.df = df
                self.file_path.set(path)
                messagebox.showinfo("成功", f"已加载 {path}\n共 {len(df)} 行")
            except Exception as e:
                messagebox.showerror("错误", str(e))

    # ===== 多线程绘图 =====
    def plot(self):
        if self.df is None:
            messagebox.showerror("错误", "请先加载数据文件")
            return

        self.master.config(cursor="wait")
        self.ax.clear()
        self.hover_annotation.set_visible(False)
        self.canvas.draw()
        threading.Thread(target=self._plot_thread, daemon=True).start()

    def _plot_thread(self):
        plot_choice = self.plot_type.get()
        try:
            if plot_choice.startswith("带状图"):
                self.plot_ribbons()
            elif plot_choice.startswith("点图"):
                self.plot_dotplot()
            elif plot_choice.startswith("Block 统计"):
                self.plot_block_stats()
            elif plot_choice.startswith("网络图"):
                self.plot_network()
            elif plot_choice.startswith("E-value"):
                self.plot_evalue_heat()

            self.master.after(0, self.canvas.draw)
        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror("绘图错误", str(e)))
        finally:
            self.master.after(0, lambda: self.master.config(cursor=""))

    # ===== 保存图片 =====
    def save_plot(self):
        if self.df is None:
            messagebox.showerror("错误", "没有可保存的图形")
            return

        base_name = os.path.splitext(os.path.basename(self.file_path.get()))[0]
        plot_type_name = self.plot_type.get().split(" ")[0]
        default_name = f"{base_name}_{plot_type_name}.png"

        filetypes = [("PNG 文件", "*.png"), ("PDF 文件", "*.pdf"), ("SVG 文件", "*.svg")]
        path = filedialog.asksaveasfilename(filetypes=filetypes, defaultextension=".png", initialfile=default_name)
        if path:
            try:
                self.fig.savefig(path, bbox_inches="tight")
                messagebox.showinfo("成功", f"图形已保存到：{path}")
            except Exception as e:
                messagebox.showerror("保存失败", str(e))

    # ===== 清除画布 =====
    def clear_canvas(self):
        self.ax.clear()
        self.hover_annotation.set_visible(False)
        self.canvas.draw()

    # ===== 鼠标悬浮提示 =====
    def enable_hover(self, coords=None, labels=None):
        if coords is None or labels is None:
            return

        def on_move(event):
            if event.inaxes != self.ax:
                self.hover_annotation.set_visible(False)
                self.canvas.draw_idle()
                return

            x, y = event.xdata, event.ydata
            if x is None or y is None:
                self.hover_annotation.set_visible(False)
                self.canvas.draw_idle()
                return

            dist = [(np.hypot(x - cx, y - cy), idx) for idx, (cx, cy) in enumerate(coords)]
            min_dist, min_idx = min(dist, key=lambda t: t[0])
            if min_dist < 0.05 * max(self.ax.get_xlim()[1], self.ax.get_ylim()[1]):
                self.hover_annotation.xy = coords[min_idx]
                self.hover_annotation.set_text(labels[min_idx])
                self.hover_annotation.set_visible(True)
            else:
                self.hover_annotation.set_visible(False)
            self.canvas.draw_idle()

        self.canvas.mpl_connect("motion_notify_event", on_move)

    # ===== 绘图方法 =====
    def plot_ribbons(self):
        df = self.df.copy()
        df["pos_a"] = np.arange(len(df))
        df["pos_b"] = np.arange(len(df))
        df['Block'] = df['Block'].fillna('NA')

        blocks = df['Block'].unique()
        cmap = cm.get_cmap("tab20", len(blocks))

        self.ax.plot([0, 0], [0, len(df)], lw=8, color="lightgray")
        self.ax.plot([1, 1], [0, len(df)], lw=8, color="lightgray")

        for i, block in enumerate(blocks):
            sub = df[df['Block'] == block]
            if len(sub) > 2000:
                sub = sub.iloc[::10]
            coords = np.array([[[0, pa], [1, pb]] for pa, pb in zip(sub['pos_a'], sub['pos_b'])])
            lc = LineCollection(coords, colors=cmap(i), linewidths=0.6, alpha=0.6)
            self.ax.add_collection(lc)

        self.ax.set_xlim(-0.2, 1.2)
        self.ax.set_ylim(0, len(df))
        self.ax.set_xticks([0, 1])
        self.ax.set_xticklabels(["GeneA", "GeneB"], fontsize=12)
        self.ax.set_ylabel("基因序列顺序")
        self.ax.set_title("共线性带状图")
        self.ax.legend(fontsize=8, frameon=False)

        # 悬浮显示
        coords = list(zip(df["pos_a"], df["pos_b"]))
        labels = [f"{a} - {b}" for a, b in zip(df["GeneA"], df["GeneB"])]
        self.enable_hover(coords, labels)

    def plot_dotplot(self):
        df = self.df.copy()
        df["pos_a"] = np.arange(len(df))
        df["pos_b"] = np.arange(len(df))
        if len(df) > 5000:
            df = df.iloc[::10]
        self.ax.scatter(df["pos_a"], df["pos_b"], s=10, c="blue", alpha=0.6)
        self.ax.set_xlabel("GeneA 顺序位置")
        self.ax.set_ylabel("GeneB 顺序位置")
        self.ax.set_title("Dot plot")

        coords = list(zip(df["pos_a"], df["pos_b"]))
        labels = [f"{a} - {b}" for a, b in zip(df["GeneA"], df["GeneB"])]
        self.enable_hover(coords, labels)

    def plot_block_stats(self):
        if "Block" not in self.df.columns:
            raise ValueError("没有 Block 列，无法统计")
        counts = self.df.groupby("Block").size().sort_values(ascending=False)
        counts.plot(kind="bar", ax=self.ax, color="skyblue")
        self.ax.set_ylabel("基因对数目")
        self.ax.set_xlabel("Block")
        self.ax.set_title("每个 Block 的基因数目")

    def plot_network(self):
        G = nx.Graph()
        df_sample = self.df.head(5000)
        for _, row in df_sample.iterrows():
            G.add_node(row["GeneA"], group="A")
            G.add_node(row["GeneB"], group="B")
            G.add_edge(row["GeneA"], row["GeneB"])

        pos = nx.spring_layout(G, k=0.3, iterations=50, seed=42)
        groups = nx.get_node_attributes(G, "group")
        colors = ["red" if groups[n] == "A" else "blue" for n in G.nodes]

        nx.draw(G, pos, node_color=colors, node_size=30,
                edge_color="gray", linewidths=0.2, with_labels=False, ax=self.ax)
        self.ax.set_title(f"基因共线性网络图 ({len(G.nodes)} 节点)")

        coords = [pos[n] for n in G.nodes()]
        labels = [str(n) for n in G.nodes()]
        self.enable_hover(coords, labels)

    def plot_evalue_heat(self):
        if "E-value" not in self.df.columns:
            raise ValueError("没有 E-value 列")
        df = self.df.copy()
        df["pos_a"] = np.arange(len(df))
        df["pos_b"] = np.arange(len(df))
        df["log_e"] = -np.log10(df["E-value"].replace(0, 1e-300).astype(float))

        if len(df) > 5000:
            df = df.iloc[::10]

        sc = self.ax.scatter(df["pos_a"], df["pos_b"], c=df["log_e"],
                             cmap="viridis", s=15, alpha=0.7)
        plt.colorbar(sc, ax=self.ax, label="-log10(E-value)")
        self.ax.set_xlabel("GeneA 顺序")
        self.ax.set_ylabel("GeneB 顺序")
        self.ax.set_title("E-value 热度图")

        coords = list(zip(df["pos_a"], df["pos_b"]))
        labels = [f"{a} - {b}\n-log10(E)={log_e:.2f}" 
                  for a,b,log_e in zip(df["GeneA"], df["GeneB"], df["log_e"])]
        self.enable_hover(coords, labels)


if __name__ == "__main__":
    root = tk.Tk()
    app = SyntenyGUI(root)
    root.mainloop()