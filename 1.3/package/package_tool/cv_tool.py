import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import networkx as nx
from matplotlib.patches import Wedge, PathPatch
from matplotlib.path import Path
import tkinter as tk

# 辅助：从 StringVar 或普通字符串中安全获取值
def _get_val(v):
    try:
        return v.get()
    except Exception:
        return v

# ---------------- 关系图 ----------------
def plot_network(self):
    # 从 self 上安全读取列名（兼容 StringVar 或 直接字符串）
    x_col = _get_val(getattr(self, "x_column", "")) 
    y_col = _get_val(getattr(self, "y_column", "")) 
    weight_col = _get_val(getattr(self, "date_column", ""))

    # 检查 DataFrame
    df = getattr(self, "df", None)
    if df is None or not hasattr(df, "columns"):
        if hasattr(self, "status_var"):
            try:
                self.status_var.set("❌ 请先加载数据（DataFrame 为空）！")
            except Exception:
                pass
        return

    if x_col not in df.columns or y_col not in df.columns:
        if hasattr(self, "status_var"):
            try:
                self.status_var.set("❌ 选中的节点列无效。")
            except Exception:
                pass
        return

    G = nx.Graph()

    # 处理边及权重
    if weight_col not in ["--- 请选择 ---", ""]:
        df_temp = df[[x_col, y_col, weight_col]].copy()
        df_temp[weight_col] = pd.to_numeric(df_temp[weight_col], errors="coerce").fillna(0)
        df_grouped = df_temp.groupby([x_col, y_col])[weight_col].mean().reset_index()
        edges = [(r[x_col], r[y_col], {"weight": r[weight_col]}) for _, r in df_grouped.iterrows()]
    else:
        df_edges = df[[x_col, y_col]].drop_duplicates()
        edges = [(r[x_col], r[y_col]) for _, r in df_edges.iterrows()]

    G.add_edges_from(edges)

    if G.number_of_nodes() == 0:
        if hasattr(self, "status_var"):
            try:
                self.status_var.set("❌ 数据中没有找到任何节点关系。")
            except Exception:
                pass
        return

    np.random.seed(42)

    num_nodes = G.number_of_nodes()
    k_value = 10 / np.sqrt(np.sqrt(num_nodes)) if num_nodes > 1 else 0.5

    # layout 可能耗时，捕获异常降级
    try:
        pos = nx.spring_layout(G, k=k_value, iterations=2000)
    except Exception:
        pos = nx.spring_layout(G, k=0.5, iterations=50)

    # 坐标缩放与调整（对单节点或非常小矩阵做保护）
    try:
        pos_arr = np.array(list(pos.values()))
        if pos_arr.size == 0:
            pos_arr = np.array([[0.0, 0.0]])
        min_xy, max_xy = pos_arr.min(axis=0), pos_arr.max(axis=0)
        scale = 1.4
        dx = max_xy[0] - min_xy[0]
        dy = max_xy[1] - min_xy[1]
        for n in pos:
            pos[n] = (
                scale * (pos[n][0] - min_xy[0]) / (dx + 1e-9),
                scale * (pos[n][1] - min_xy[1]) / (dy + 1e-9),
            )
    except Exception:
        # 退回不缩放的原始 pos
        pass

    for n in pos:
        pos[n] = (pos[n][0] + np.random.uniform(-0.02, 0.02), pos[n][1] + np.random.uniform(-0.02, 0.02))

    # 防止节点重叠（简单平移）
    min_dist = 0.05
    nodes_list = list(pos.keys())
    for i in range(len(nodes_list)):
        for j in range(i + 1, len(nodes_list)):
            ni, nj = nodes_list[i], nodes_list[j]
            dx = pos[nj][0] - pos[ni][0]
            dy = pos[nj][1] - pos[ni][1]
            dist = np.sqrt(dx ** 2 + dy ** 2)
            if dist < min_dist and dist > 0:
                shift = (min_dist - dist) / 2
                pos[ni] = (pos[ni][0] - dx / dist * shift, pos[ni][1] - dy / dist * shift)
                pos[nj] = (pos[nj][0] + dx / dist * shift, pos[nj][1] + dy / dist * shift)

    # 节点大小和颜色
    degrees = dict(G.degree())
    fixed_size = 300
    node_sizes = [fixed_size] * len(degrees)
    node_colors = list(degrees.values())

    # 边宽
    if nx.get_edge_attributes(G, "weight"):
        weights = [G[u][v]["weight"] for u, v in G.edges()]
        max_w = max(weights) if weights else 1
        if max_w == 0:
            max_w = 1
        edge_widths = [0.5 + 3 * (w / max_w) for w in weights]
    else:
        edge_widths = [1.0] * G.number_of_edges()

    cmap = plt.get_cmap("Spectral_r")

    fig, ax = plt.subplots(figsize=(10, 8))
    fig.patch.set_facecolor("#f4f6fa")
    ax.set_facecolor("#f4f6fa")

    nx.draw_networkx_nodes(
        G, pos,
        node_size=[s * 1.8 for s in node_sizes],
        node_color=node_colors,
        cmap=cmap,
        alpha=0.25,
        linewidths=0,
        ax=ax
    )

    nx.draw_networkx_edges(
        G, pos,
        width=edge_widths,
        alpha=0.35,
        edge_color="#7d7d7d",
        style="solid",
        ax=ax
    )

    nx.draw_networkx_nodes(
        G, pos,
        node_size=node_sizes,
        node_color=node_colors,
        cmap=cmap,
        alpha=0.95,
        edgecolors="#eeeeee",
        linewidths=0.8,
        ax=ax
    )

    node_labels = {}
    for i, node in enumerate(G.nodes()):
        if i < 60:
            x, y = pos[node]
            node_labels[node] = node
            pos[node] = (x, y + 0.02)

    nx.draw_networkx_labels(
        G, pos,
        labels=node_labels,
        font_size=9,
        font_family="SimHei",
        font_color="black",
        ax=ax
    )

    edge_weights = nx.get_edge_attributes(G, "weight")
    if edge_weights:
        edge_labels = {k: f"{v:.2f}" for k, v in edge_weights.items()}
        nx.draw_networkx_edge_labels(
            G, pos,
            edge_labels=edge_labels,
            font_color="darkgreen",
            font_size=8,
            rotate=False,
            ax=ax
        )

    plt.axis("off")

    # 色条
    try:
        vmin = min(node_colors) if node_colors else 0
        vmax = max(node_colors) if node_colors else 1
    except Exception:
        vmin, vmax = 0, 1
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=plt.Normalize(vmin=vmin, vmax=vmax))
    sm.set_array([])
    cbar = fig.colorbar(sm, ax=ax, orientation="vertical", shrink=0.75)
    cbar.set_label("连接数", fontdict={"fontname": "SimHei", "fontsize": 14, "color": "#444444"})

    for spine in ax.spines.values():
        spine.set_visible(False)

    plt.tight_layout()

    # 显示并更新状态
    try:
        self.show_plot(fig)
        if hasattr(self, "status_var"):
            self.status_var.set("✅ 图形生成完成")
    except Exception:
        pass

# ---------------- 热图 ----------------
def plot_heatmap(self):
    df = getattr(self, "df", None)
    if df is None:
        try:
            self.status_var.set("❌ 请先选择信息文件！")
        except Exception:
            pass
        return

    x_col = _get_val(getattr(self, "x_column", "")) or df.columns[0]
    y_col = _get_val(getattr(self, "y_column", "")) or (df.columns[1] if len(df.columns) > 1 else df.columns[0])
    date_col = _get_val(getattr(self, "date_column", "")) or (df.columns[2] if len(df.columns) > 2 else df.columns[-1])

    try:
        pivot_table = df.pivot_table(index=y_col, columns=x_col, values=date_col, aggfunc='mean')
    except Exception as e:
        try:
            self.status_var.set(f"❌ 无法生成热图: {e}")
        except Exception:
            pass
        return

    fig, ax = plt.subplots(figsize=(10, 8))
    cax = ax.imshow(pivot_table, aspect='auto', cmap='viridis')
    fig.colorbar(cax, label=date_col)
    ax.set_xticks(range(len(pivot_table.columns)))
    ax.set_xticklabels(pivot_table.columns, rotation=90)
    ax.set_yticks(range(len(pivot_table.index)))
    ax.set_yticklabels(pivot_table.index)
    ax.set_xlabel(_get_val(getattr(self, "x_label", "")) or x_col)
    ax.set_ylabel(_get_val(getattr(self, "y_label", "")) or y_col)
    ax.set_title("Heatmap Visualization")
    fig.tight_layout()
    try:
        self.show_plot(fig)
    except Exception:
        pass

# ---------------- 美观弦图 ----------------
def plot_chord(self):
    df = getattr(self, "df", None)
    if df is None:
        try:
            self.status_var.set("❌ 请先选择信息文件！")
        except Exception:
            pass
        return

    if len(df.columns) < 2:
        try:
            self.status_var.set("❌ 数据列不足两列！")
        except Exception:
            pass
        return

    col1 = _get_val(getattr(self, "x_column", "")) or df.columns[0]
    col2 = _get_val(getattr(self, "y_column", "")) or df.columns[1]
    data = df[[col1, col2]].dropna()
    if col1 == "--- 请选择 ---" or col2 == "--- 请选择 ---":
        try:
            self.status_var.set("❌ 弦图至少需要选择 **第一列 (X)** 和 **第二列 (Y)**！")
        except Exception:
            pass
        return

    nodes = sorted(set(data[col1]).union(set(data[col2])))
    n = len(nodes)
    if n == 0:
        try:
            self.status_var.set("❌ 数据中未发现节点！")
        except Exception:
            pass
        return
    node_index = {v: i for i, v in enumerate(nodes)}
    matrix = np.zeros((n, n))
    for a, b in data.values:
        i, j = node_index[a], node_index[b]
        matrix[i, j] += 1
        matrix[j, i] += 1
    total = np.sum(matrix)
    if total == 0:
        try:
            self.status_var.set("❌ 数据无连接关系，无法绘制弦图。")
        except Exception:
            pass
        return

    pad_deg = 2
    pad = np.deg2rad(pad_deg)
    weights = matrix.sum(axis=1)
    # 防止除以 0
    weights_sum = weights.sum() if weights.sum() > 0 else 1
    angles = (2*np.pi - n*pad) * weights / weights_sum
    starts, ends = np.zeros(n), np.zeros(n)
    current = 0
    for i in range(n):
        starts[i] = current
        ends[i] = current + angles[i]
        current = ends[i] + pad

    fig, ax = plt.subplots(figsize=(8, 8))
    ax.set_xlim(-1.4, 1.4)
    ax.set_ylim(-1.4, 1.4)
    ax.axis("off")

    cmap = plt.get_cmap("Spectral")
    colors = [cmap(i / n) for i in range(n)]
    inner_radius = 0.7

    def angle_to_point(a, r=1):
        return np.array([np.cos(a)*r, np.sin(a)*r])

    for i in range(n):
        wedge = Wedge((0, 0), 1, np.degrees(starts[i]), np.degrees(ends[i]),
                        width=0.15, facecolor=colors[i], edgecolor="white")
        ax.add_patch(wedge)

        mid = 0.5 * (starts[i] + ends[i])
        outer_pt = angle_to_point(mid, 1.05)
        label_pt = angle_to_point(mid, 1.25)

        ax.plot([outer_pt[0], label_pt[0]], [outer_pt[1], label_pt[1]],
                color=colors[i], lw=1.0)

        ha = "left" if np.cos(mid) >= 0 else "right"
        ax.text(label_pt[0], label_pt[1], nodes[i],
                ha=ha, va="center", fontsize=9, color="black")

    max_m = np.max(matrix) if np.max(matrix) > 0 else 1
    for i in range(n):
        for j in range(i+1, n):
            if matrix[i, j] <= 0:
                continue
            v = matrix[i, j]
            a1 = 0.5*(starts[i]+ends[i])
            a2 = 0.5*(starts[j]+ends[j])
            p1 = angle_to_point(a1, inner_radius)
            p2 = angle_to_point(a2, inner_radius)
            verts = [
                (p1[0], p1[1]),
                (p1[0]*0.4, p1[1]*0.4),
                (p2[0]*0.4, p2[1]*0.4),
                (p2[0], p2[1]),
            ]
            codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
            path = Path(verts, codes)
            patch = PathPatch(
                path, facecolor="none",
                edgecolor=colors[i],
                lw=0.8 + 3*(v/max_m),
                alpha=0.4 + 0.6*(v/max_m)
            )
            ax.add_patch(patch)

    ax.set_aspect("equal")
    try:
        self.show_plot(fig)
    except Exception:
        pass