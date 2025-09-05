import pandas as pd

def classify_genes_with_progress(file_a, file_b, output_file, gene_column_a, gene_id_column_b, collinear_gene_column_b, progress_callback):
    """基因匹配功能，支持进度更新"""
    # 读取表格 A 和表格 B
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    # 检查列名是否存在
    if gene_column_a not in df_a.columns:
        raise ValueError(f"表格 A 中不存在列名 '{gene_column_a}'")
    if gene_id_column_b not in df_b.columns or collinear_gene_column_b not in df_b.columns:
        raise ValueError(f"表格 B 中不存在列名 '{gene_id_column_b}' 或 '{collinear_gene_column_b}'")

    # 创建一个新 DataFrame 用于存储结果
    result_rows = []

    # 遍历表格 A 的基因列
    total = len(df_a)
    for i, (index, gene) in enumerate(df_a[gene_column_a].items()):
        if pd.isnull(gene):  # 跳过空值
            continue

        # 在表格 B 的 '基因id' 列中查找匹配
        matches_in_gene_id = df_b[df_b[gene_id_column_b] == gene][collinear_gene_column_b].tolist()

        # 在表格 B 的 '共线基因' 列中查找匹配
        matches_in_collinear_gene = df_b[df_b[collinear_gene_column_b] == gene][gene_id_column_b].tolist()

        # 合并匹配结果
        combined_matches = matches_in_gene_id + matches_in_collinear_gene

        # 如果没有匹配结果，存储 "无"
        if not combined_matches:
            result_rows.append({gene_column_a: gene, "匹配结果": "无"})
        else:
            # 如果有多个匹配结果，复制基因并一一对应
            for match in combined_matches:
                result_rows.append({gene_column_a: gene, "匹配结果": match})

        # 更新进度
        progress_callback((i + 1) / total * 100)

    # 将结果转换为 DataFrame
    result_df = pd.DataFrame(result_rows)

    # 保存结果到输出文件
    result_df.to_excel(output_file, index=False, engine="openpyxl")

def gene_correspondence_with_progress(file_a, file_b, output_file, gene_column_a, gene_id_column_b, collinear_gene_column_b, progress_callback):
    """基因对应功能，支持进度更新"""
    # 读取表格 A 和表格 B
    df_a = pd.read_excel(file_a)
    df_b = pd.read_excel(file_b)

    # 检查列名是否存在
    if gene_column_a not in df_a.columns:
        raise ValueError(f"表格 A 中不存在列名 '{gene_column_a}'")
    if gene_id_column_b not in df_b.columns or collinear_gene_column_b not in df_b.columns:
        raise ValueError(f"表格 B 中不存在列名 '{gene_id_column_b}' 或 '{collinear_gene_column_b}'")

    # 遍历表格 A 的基因列
    total = len(df_a)
    for i, (index, gene) in enumerate(df_a[gene_column_a].items()):
        if pd.isnull(gene):  # 跳过空值
            continue

        # 在表格 B 的 '基因id' 列中查找匹配
        matches_in_gene_id = df_b[df_b[gene_id_column_b] == gene][collinear_gene_column_b].tolist()

        # 在表格 B 的 '共线基因' 列中查找匹配
        matches_in_collinear_gene = df_b[df_b[collinear_gene_column_b] == gene][gene_id_column_b].tolist()

        # 合并匹配结果
        combined_matches = matches_in_gene_id + matches_in_collinear_gene

        # 如果没有匹配结果，存储 "无"
        if not combined_matches:
            df_a.at[index, "匹配结果"] = "无"
        else:
            # 如果有多个匹配结果，存储为逗号分隔的字符串
            df_a.at[index, "匹配结果"] = ", ".join(combined_matches)

        # 更新进度
        progress_callback((i + 1) / total * 100)

    # 保存结果到输出文件
    df_a.to_excel(output_file, index=False, engine="openpyxl")
