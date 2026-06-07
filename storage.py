# -*- coding: utf-8 -*-
"""ETF 份额历史数据的本地存储模块。

每个 ETF 的数据保存在 data/代码.xlsx，列结构：
    日期, 代码, 名称, 份额, 抓取时间
其中「份额」单位为万份（与深交所基金规模口径一致）。
按「日期」去重（新数据覆盖旧数据），按日期升序保存。
"""

import os
from datetime import datetime

import pandas as pd

# 数据目录（与本文件同级的 data 文件夹）
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

# Excel 列顺序（份额单位：万份）
COLUMNS = ["日期", "代码", "名称", "份额", "抓取时间"]


def _ensure_data_dir():
    if not os.path.isdir(DATA_DIR):
        os.makedirs(DATA_DIR, exist_ok=True)


def get_file_path(code):
    """返回某 ETF 代码对应的 xlsx 文件路径。"""
    return os.path.join(DATA_DIR, "%s.xlsx" % str(code).strip())


def load_df(code):
    """读取某 ETF 的历史数据为 DataFrame，文件不存在则返回空 DataFrame。"""
    path = get_file_path(code)
    if not os.path.isfile(path):
        return pd.DataFrame(columns=COLUMNS)
    try:
        df = pd.read_excel(path, dtype={"日期": str, "代码": str})
    except Exception:
        return pd.DataFrame(columns=COLUMNS)
    # 补齐缺失列，保证列顺序统一
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[COLUMNS]


def save_records(code, records):
    """将若干条记录写入某 ETF 的 xlsx，按日期去重后升序保存。

    参数：
        records: dict 列表，每条至少含 日期/代码/名称/份额/净值
    返回：保存后该文件的总记录数。
    """
    if not records:
        return len(load_df(code))

    _ensure_data_dir()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    new_rows = []
    for rec in records:
        row = dict(rec)
        row.setdefault("代码", str(code))
        row["抓取时间"] = now
        new_rows.append(row)
    new_df = pd.DataFrame(new_rows)

    old_df = load_df(code)
    combined = pd.concat([old_df, new_df], ignore_index=True)
    # 同一日期保留最后一条（即本次新抓取的数据覆盖旧的）
    combined = combined.drop_duplicates(subset=["日期"], keep="last")
    combined = combined.sort_values("日期").reset_index(drop=True)
    combined = combined[COLUMNS]

    combined.to_excel(get_file_path(code), index=False)
    return len(combined)


def load_history(code):
    """读取某 ETF 历史数据，返回按日期升序的 dict 列表，供前端图表/表格使用。"""
    df = load_df(code)
    if df.empty:
        return []
    df = df.sort_values("日期").reset_index(drop=True)
    records = []
    for _, r in df.iterrows():
        records.append(
            {
                "日期": None if pd.isna(r["日期"]) else str(r["日期"]),
                "代码": None if pd.isna(r["代码"]) else str(r["代码"]),
                "名称": None if pd.isna(r["名称"]) else str(r["名称"]),
                "份额": None if pd.isna(r["份额"]) else float(r["份额"]),
            }
        )
    return records
