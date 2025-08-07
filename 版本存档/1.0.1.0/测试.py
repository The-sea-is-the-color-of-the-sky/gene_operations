import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

# 测试文件夹下的“填入表格.xlsx”和“信息表格.xlsx”为测试使用的表格
# 假设文件路径和表名如下，请根据实际情况修改
fillin_path = r'测试\填入表格.xlsx'
info_path = r'测试\信息表格.xlsx'
fillin_sheet = 'Sheet1'  # 填入表的工作表名
info_sheet = 'Sheet1'    # 信息表的工作表名
fillin_col = '基因id'     # 填入表的列名
info_col_a = '基因A'         # 信息表A列列名
info_col_b = '基因B'         # 信息表B列列名

# 读取数据
fillin_df = pd.read_excel(fillin_path, sheet_name=fillin_sheet)
info_df = pd.read_excel(info_path, sheet_name=info_sheet)

result_col = '匹配结果'

def process_row(args):
    row, idx = args
    val = row[fillin_col]
    if pd.isna(val):
        return [{**row, result_col: ''}]
    matches = []
    # 只为前5个线程显示内部进度条
    show_inner_bar = idx < 5
    info_iter = info_df.iterrows()
    if show_inner_bar:
        info_iter = tqdm(info_iter, total=len(info_df), desc=f'查找匹配-{idx}', position=idx+1, leave=False)
    for _, info_row in info_iter:
        if info_row[info_col_a] == val:
            matches.append(info_row[info_col_b])
        if info_row[info_col_b] == val:
            matches.append(info_row[info_col_a])
    if not matches:
        return [{**row, result_col: ''}]
    else:
        result = []
        for i, match in enumerate(matches):
            new_row = row.copy()
            new_row[result_col] = match
            if i > 0:
                new_row[fillin_col] = ''
            result.append(new_row.to_dict())
        return result

rows = []
with ThreadPoolExecutor(max_workers=2) as executor:
    tasks = [(row, idx) for idx, (_, row) in enumerate(fillin_df.iterrows())]
    for res in tqdm(
        executor.map(process_row, tasks),
        total=len(tasks),
        desc='处理填入表',
        position=0,
        leave=True  # 主进度条一直显示
    ):
        rows.extend(res)

result_df = pd.DataFrame(rows, columns=list(fillin_df.columns) + [result_col])
# 保存为新文件
result_df.to_excel(r'测试\填入表格_匹配结果.xlsx', sheet_name=fillin_sheet, index=False)