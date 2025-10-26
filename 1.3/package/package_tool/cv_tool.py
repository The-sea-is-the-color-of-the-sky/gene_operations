import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import networkx as nx
from matplotlib.patches import Wedge, PathPatch
from matplotlib.path import Path
import tkinter as tk


def _get_val(v):
    """安全获取 StringVar 或普通字符串的值"""
    try:
        return v.get()
    except Exception:
        return v


def _resolve_args(obj_or_df, x_col, y_col, weight_col, show_plot):
    """统一解析传入参数，支持 GUI 对象或直接传入 DataFrame"""
    df = None
    status_var = None
    if obj_or_df is not None and hasattr(obj_or_df, "df"):
        self_obj = obj_or_df
        df = getattr(self_obj, "df", None)
        x_col = x_col or _get_val(getattr(self_obj, "x_column", "")) or None
        y_col = y_col or _get_val(getattr(self_obj, "y_column", "")) or None
        weight_col = weight_col or _get_val(getattr(self_obj, "date_column", "")) or None
        status_var = getattr(self_obj, "status_var", None)
        if show_plot is None and hasattr(self_obj, "show_plot"):
            show_plot = getattr(self_obj, "show_plot")
    else:
        df = obj_or_df

    # 增强列验证
    if df is not None and hasattr(df, "columns"):
        if x_col and x_col not in df.columns:
            if status_var:
                status_var.set(f"❌ 节点1列 '{x_col}' 无效")
            return None, None, None, None, None, status_var
        if y_col and y_col not in df.columns:
            if status_var:
                status_var.set(f"❌ 节点2列 '{y_col}' 无效")
            return None, None, None, None, None, status_var
        if weight_col and weight_col not in df.columns and weight_col != "--- 请选择 ---":
            if status_var:
                status_var.set(f"❌ 权重列 '{weight_col}' 无效")
            return None, None, None, None, None, status_var
    return df, x_col, y_col, weight_col, show_plot, status_var


def plot_network(obj_or_df=None, x_col=None, y_col=None, weight_col=None, show_plot=None):
    """绘制关系网络图"""
    df, x_col, y_col, weight_col, show_plot, status_var = _resolve_args(obj_or_df, x_col, y_col, weight_col, show_plot)

    if df is None or not hasattr(df, "columns"):
        if status_var is not None:
            status_var.set("❌ 请先加载数据（DataFrame 为空）！")
        return None

    if not x_col:
        x_col = df.columns[0] if len(df.columns) > 0 else None
    if not y_col:
        y_col = df.columns[1] if len(df.columns) > 1 else x_col

    G = nx.Graph()

    if weight_col and weight_col in df.columns and weight_col != "--- 请选择 ---":
        df_temp = df[[x_col, y_col, weight_col]].copy()
        df_temp[weight_col] = pd.to_numeric(df_temp[weight_col], errors="coerce").fillna(0)
        df_grouped = df_temp.groupby([x_col, y_col])[weight_col].mean().reset_index()
        edges = [(r[x_col], r[y_col], {"weight": r[weight_col]}) for _, r in df_grouped.iterrows()]
    else:
        df_edges = df[[x_col, y_col]].drop_duplicates()
        edges = [(r[x_col], r[y_col]) for _, r in df_edges.iterrows()]

    G.add_edges_from(edges)

    if G.number_of_nodes() == 0:
        if status_var is not None:
            status_var.set("❌ 数据中没有找到任何节点关系。")
        return None

    np.random.seed(42)
    num_nodes = G.number_of_nodes()
    k_value = 10 / np.sqrt(np.sqrt(num_nodes)) if num_nodes > 1 else 0.5
    iterations = min(2000, max(50, 1000 // max(1, num_nodes // 10)))

    try:
        pos = nx.spring_layout(G, k=k_value, iterations=iterations)
    except Exception:
        pos = nx.spring_layout(G, k=0.5, iterations=50)

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
        pass

    for n in pos:
        pos[n] = (pos[n][0] + np.random.uniform(-0.02, 0.02), pos[n][1] + np.random.uniform(-0.02, 0.02))

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

    degrees = dict(G.degree())
    fixed_size = 300
    node_sizes = [fixed_size] * len(degrees)
    node_colors = list(degrees.values())

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

    max_labels = min(60, len(G.nodes()))
    node_labels = {node: node for i, node in enumerate(G.nodes()) if i < max_labels}
    for node in node_labels:
        x, y = pos[node]
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

    try:
        if show_plot:
            show_plot(fig)
        if status_var is not None:
            status_var.set("✅ 图形生成完成")
    except Exception:
        pass

    return fig


def plot_heatmap(obj_or_df=None, x_col=None, y_col=None, date_col=None, show_plot=None):
    """绘制热图"""
    df, x_col, y_col, date_col, show_plot, status_var = _resolve_args(obj_or_df, x_col, y_col, date_col, show_plot)

    if df is None:
        if status_var is not None:
            status_var.set("❌ 请先选择信息文件！")
        return None

    x_col = x_col or df.columns[0]
    y_col = y_col or (df.columns[1] if len(df.columns) > 1 else df.columns[0])
    date_col = date_col or (df.columns[2] if len(df.columns) > 2 else df.columns[-1])

    try:
        pivot_table = df.pivot_table(index=y_col, columns=x_col, values=date_col, aggfunc='mean')
    except Exception as e:
        if status_var is not None:
            status_var.set(f"❌ 无法生成热图: {e}")
        return None

    fig, ax = plt.subplots(figsize=(10, 8))
    cax = ax.imshow(pivot_table, aspect='auto', cmap='viridis')
    fig.colorbar(cax, label=date_col)
    ax.set_xticks(range(len(pivot_table.columns)))
    ax.set_xticklabels(pivot_table.columns, rotation=90)
    ax.set_yticks(range(len(pivot_table.index)))
    ax.set_yticklabels(pivot_table.index)
    ax.set_xlabel(x_col)
    ax.set_ylabel(y_col)
    ax.set_title("Heatmap Visualization")
    fig.tight_layout()
    try:
        if show_plot:
            show_plot(fig)
    except Exception:
        pass
    return fig


def plot_chord(obj_or_df=None, col1=None, col2=None, show_plot=None):
    """绘制弦图"""
    df, col1, col2, _unused, show_plot, status_var = _resolve_args(obj_or_df, col1, col2, None, show_plot)

    if df is None:
        if status_var is not None:
            status_var.set("❌ 请先选择信息文件！")
        return None

    if len(df.columns) < 2:
        if status_var is not None:
            status_var.set("❌ 数据列不足两列！")
        return None

    col1 = col1 or df.columns[0]
    col2 = col2 or (df.columns[1] if len(df.columns) > 1 else df.columns[0])

    data = df[[col1, col2]].dropna()
    if col1 == "--- 请选择 ---" or col2 == "--- 请选择 ---":
        if status_var is not None:
            status_var.set("❌ 弦图至少需要选择 **第一列 (X)** 和 **第二列 (Y)**！")
        return None

    nodes = sorted(set(data[col1]).union(set(data[col2])))
    n = len(nodes)
    if n == 0:
        if status_var is not None:
            status_var.set("❌ 数据中未发现节点！")
        return None
    node_index = {v: i for i, v in enumerate(nodes)}
    matrix = np.zeros((n, n))
    for a, b in data.values:
        i, j = node_index[a], node_index[b]
        matrix[i, j] += 1
        matrix[j, i] += 1
    total = np.sum(matrix)
    if total == 0:
        if status_var is not None:
            status_var.set("❌ 数据无连接关系，无法绘制弦图。")
        return None

    pad_deg = 2
    pad = np.deg2rad(pad_deg)
    weights = matrix.sum(axis=1)
    weights_sum = weights.sum() if weights.sum() > 0 else 1
    angles = (2 * np.pi - n * pad) * weights / weights_sum
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
        return np.array([np.cos(a) * r, np.sin(a) * r])

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
        for j in range(i + 1, n):
            if matrix[i, j] <= 0:
                continue
            v = matrix[i, j]
            a1 = 0.5 * (starts[i] + ends[i])
            a2 = 0.5 * (starts[j] + ends[j])
            p1 = angle_to_point(a1, inner_radius)
            p2 = angle_to_point(a2, inner_radius)
            verts = [
                (p1[0], p1[1]),
                (p1[0] * 0.4, p1[1] * 0.4),
                (p2[0] * 0.4, p2[1] * 0.4),
                (p2[0], p2[1]),
            ]
            codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4]
            path = Path(verts, codes)
            patch = PathPatch(
                path, facecolor="none",
                edgecolor=colors[i],
                lw=0.8 + 3 * (v / max_m),
                alpha=0.4 + 0.6 * (v / max_m)
            )
            ax.add_patch(patch)

    ax.set_aspect("equal")
    try:
        if show_plot:
            show_plot(fig)
    except Exception:
        pass
    return fig