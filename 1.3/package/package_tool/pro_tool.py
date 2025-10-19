import os
from datetime import datetime
import pandas as pd
import re
from collections import defaultdict

def _normalize(s: str) -> str:
    return s.strip().lower()

def _tokens(s: str):
    return re.findall(r'\w+', s.lower())

def _read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ['.xls', '.xlsx']:
        df = pd.read_excel(path, dtype=str)
    else:
        try:
            df = pd.read_csv(path, dtype=str, encoding='utf-8')
        except Exception:
            df = pd.read_csv(path, dtype=str, encoding='gbk')
    return df.fillna('').astype(str)

def pro_match(a_path, b_path, out_folder, a_col, b1, b2, export_other=True,
              threshold_ratio=0.6, progress_cb=None, status_cb=None, stop_flag_getter=None):
    """
    执行匹配并写出结果文件。
    返回生成的输出文件路径，若无结果或被中断则返回 None。
    progress_cb(percent:int) - 可选，接受 0-100
    status_cb(text:str) - 可选，用于显示状态
    stop_flag_getter() - 可选，返回 True 则立即中断（安全退出，可能已写部分结果）
    """
    if status_cb:
        status_cb("读取文件中...")
    df_a = _read_table(a_path)
    df_b = _read_table(b_path)
    if df_a.empty:
        raise ValueError("A 表为空。")

    if out_folder and not os.path.exists(out_folder):
        os.makedirs(out_folder, exist_ok=True)

    cols_b = list(df_b.columns)
    try:
        pos_b1 = df_b.columns.get_loc(b1)
    except Exception:
        pos_b1 = None
    try:
        pos_b2 = df_b.columns.get_loc(b2)
    except Exception:
        pos_b2 = None

    b_tuples = [tuple(x) for x in df_b.itertuples(index=False, name=None)]
    b_map1 = defaultdict(list)
    b_map2 = defaultdict(list)
    token_index1 = defaultdict(set)
    token_index2 = defaultdict(set)
    b_key_tokens1 = {}
    b_key_tokens2 = {}

    for idx, row in enumerate(b_tuples):
        v1 = _normalize(str(row[pos_b1])) if pos_b1 is not None else ""
        v2 = _normalize(str(row[pos_b2])) if pos_b2 is not None else ""
        if v1:
            b_map1[v1].append(idx)
            tks = set(_tokens(v1))
            b_key_tokens1[v1] = tks
            for t in tks:
                token_index1[t].add(v1)
        if v2:
            b_map2[v2].append(idx)
            tks2 = set(_tokens(v2))
            b_key_tokens2[v2] = tks2
            for t in tks2:
                token_index2[t].add(v2)

    other_b_cols = [c for c in df_b.columns if c not in {b1, b2}]
    results = []
    total = len(df_a)

    try:
        pos_a = df_a.columns.get_loc(a_col)
    except Exception:
        pos_a = None

    progress_step = max(1, total // 100)
    status_step = max(1, progress_step * 5)

    for idx_a, tup in enumerate(df_a.itertuples(index=False, name=None)):
        if stop_flag_getter and stop_flag_getter():
            if status_cb:
                status_cb("任务中止。")
            break

        a_value_raw = str(tup[pos_a]).strip() if pos_a is not None else ""
        a_value = _normalize(a_value_raw)
        a_row_dict = dict(zip(df_a.columns, tup))

        if (idx_a + 1) % progress_step == 0 or idx_a == total - 1:
            p = int((idx_a + 1) / total * 100)
            if progress_cb:
                progress_cb(p)

        matches = []

        if a_value:
            if a_value in b_map1:
                for b_idx in b_map1[a_value]:
                    matches.append(('完全匹配', b1, b_idx, 1.0))
            if a_value in b_map2:
                for b_idx in b_map2[a_value]:
                    matches.append(('完全匹配', b2, b_idx, 1.0))

        if not matches and a_value:
            toks = set(_tokens(a_value))
            cand_keys = set()
            for t in toks:
                if t in token_index1:
                    cand_keys.update(token_index1[t])
                if t in token_index2:
                    cand_keys.update(token_index2[t])

            MAX_CAND = 500
            if len(cand_keys) > MAX_CAND:
                def score_key(k):
                    tk = b_key_tokens1.get(k) or b_key_tokens2.get(k) or set(_tokens(k))
                    return -len(toks & tk)
                cand_keys = set(sorted(cand_keys, key=score_key)[:MAX_CAND])

            for k in cand_keys:
                if not k:
                    continue
                tk = b_key_tokens1.get(k) or b_key_tokens2.get(k) or set(_tokens(k))
                inter = toks & tk
                ratio = (len(inter) / max(1, min(len(toks), len(tk)))) if (toks and tk) else 0.0
                if (k in a_value) or (a_value in k) or (len(inter) > 0 and ratio >= threshold_ratio):
                    if k in b_map1:
                        for b_idx in b_map1.get(k, []):
                            matches.append(('模糊匹配', b1, b_idx, ratio))
                    if k in b_map2:
                        for b_idx in b_map2.get(k, []):
                            matches.append(('模糊匹配', b2, b_idx, ratio))

        if matches:
            for mtype, source_col, b_idx, ratio in matches:
                b_row = b_tuples[b_idx]
                r = a_row_dict.copy()
                r['匹配结果'] = mtype
                # 写匹配原始值与匹配内容
                if source_col == b1 and pos_b1 is not None:
                    r['匹配原始值'] = b_row[pos_b1]
                    r['匹配内容'] = b_row[pos_b2] if pos_b2 is not None else ''
                elif source_col == b2 and pos_b2 is not None:
                    r['匹配原始值'] = b_row[pos_b2]
                    r['匹配内容'] = b_row[pos_b1] if pos_b1 is not None else ''
                else:
                    r['匹配原始值'] = ''
                    r['匹配内容'] = ''
                r['交集率'] = round(float(ratio), 3) if ratio is not None else ''
                if export_other:
                    for c in other_b_cols:
                        try:
                            pos = cols_b.index(c)
                            r[f"B_{c}"] = b_row[pos]
                        except Exception:
                            r[f"B_{c}"] = ''
                results.append(r)
        else:
            r = a_row_dict.copy()
            r['匹配结果'] = '无匹配'
            r['匹配原始值'] = ''
            r['匹配内容'] = ''
            r['交集率'] = ''
            if export_other:
                for c in other_b_cols:
                    r[f"B_{c}"] = ''
            results.append(r)

        if (idx_a + 1) % status_step == 0 or idx_a == total - 1:
            if status_cb:
                status_cb(f"已处理 {idx_a + 1}/{total} 行")

    if not results:
        return None

    df_out = pd.DataFrame(results)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(out_folder, f"pro_{ts}.xlsx")
    with pd.ExcelWriter(out_file, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, index=False, sheet_name='匹配结果')

    if status_cb:
        status_cb("匹配完成。")
    return out_file
