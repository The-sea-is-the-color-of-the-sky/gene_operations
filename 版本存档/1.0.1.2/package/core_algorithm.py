import pandas as pd

def classify_genes(file_a, file_b, output_file, gene_column_a, gene_id_column_b, collinear_gene_column_b):
    """
    从表格 A 的 '基因' 列查找表格 B 的 '基因id' 和 '共线基因' 列的匹配内容，并归类。
    
    :param file_a: 表格 A 的文件路径
    :param file_b: 表格 B 的文件路径
    :param output_file: 输出文件路径
    :param gene_column_a: 表格 A 中的基因列名
    :param gene_id_column_b: 表格 B 中的基因 ID 列名
    :param collinear_gene_column_b: 表格 B 中的共线基因列名
    """
    try:
        # 读取表格 A 和表格 B
        df_a = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)

        # 创建一个新 DataFrame 用于存储结果
        result_rows = []

        # 遍历表格 A 的基因列
        for index, gene in df_a[gene_column_a].items():  # 将 iteritems() 替换为 items()
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

        # 将结果转换为 DataFrame
        result_df = pd.DataFrame(result_rows)

        # 保存结果到输出文件
        result_df.to_excel(output_file, index=False, engine="openpyxl")
        print(f"匹配完成，结果已保存到 {output_file}")
    except Exception as e:
        print(f"处理文件时出现错误：{str(e)}")

