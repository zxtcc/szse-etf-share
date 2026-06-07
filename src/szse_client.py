# -*- coding: utf-8 -*-
"""深交所 ETF 份额（基金规模）数据抓取模块。

数据来源：深交所「基金规模查询」接口（与官网 ETF 份额页面一致）：
  http://www.szse.cn/api/report/ShowReport/data?SHOWTYPE=JSON
      &CATALOGID=scsj_fund_jjgm&TABKEY=tab1
      &txtDm=<ETF代码>&txtStart=<开始日期>&txtEnd=<结束日期>&jjlb=ETF&random=<随机数>

返回字段（cols）：
  size_date            日期
  fund_code            基金代码
  security_short_name  基金简称（中文）
  current_size         基金规模（单位：万份，带千分位逗号）

要点：
- 必须带 jjlb=ETF，否则接口返回「查询异常」。
- 单次查询日期跨度有上限（约 180 天），超出会返回空，故大区间需切分窗口。
- 每页最多 20 条，需根据 metadata.pagecount 翻页。
- 非交易日：该日无记录（recordcount=0），据此判定。
- 深交所证书链不在 Python 默认 CA bundle，请求使用 verify=False（已抑制告警）。
"""

import time
import random
from datetime import datetime, timedelta

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_URL = "https://www.szse.cn/api/report/ShowReport/data"
CATALOG = "scsj_fund_jjgm"
# 单次查询的最大日期跨度（天）。实测约 180 天可用，留足余量。
MAX_SPAN_DAYS = 180

# 反爬：请求之间的随机休眠区间（秒）
PAGE_SLEEP = (0.4, 1.0)    # 同一窗口翻页之间
WINDOW_SLEEP = (0.8, 1.8)  # 不同窗口之间

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Referer": "https://www.szse.cn/market/fund/volume/etf/index.html",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
}


def _get_json(params, retries=3, timeout=20):
    """统一的 GET JSON 封装，带重试，返回深交所接口的 JSON 数组。"""
    params = dict(params)
    params.setdefault("SHOWTYPE", "JSON")
    params.setdefault("TABKEY", "tab1")
    params.setdefault("CATALOGID", CATALOG)
    params.setdefault("jjlb", "ETF")
    last_err = None
    for attempt in range(retries):
        params["random"] = str(random.random())
        try:
            resp = requests.get(
                BASE_URL,
                params=params,
                headers=HEADERS,
                timeout=timeout,
                verify=False,
            )
            resp.raise_for_status()
            return resp.json()
        except Exception as exc:  # noqa: BLE001 - 统一重试
            last_err = exc
            time.sleep(0.6 * (attempt + 1))
    raise RuntimeError("请求深交所接口失败：%s" % last_err)


def _sleep(span):
    """在给定区间内随机休眠，降低被反爬识别的概率。"""
    time.sleep(random.uniform(span[0], span[1]))


def _parse_size(text):
    """将带千分位逗号的规模字符串（万份）转为 float，无法解析返回 None。"""
    if text is None:
        return None
    cleaned = str(text).replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def fetch_window(code, start, end):
    """抓取一个不超过 MAX_SPAN_DAYS 的日期窗口内的全部记录（自动翻页）。

    返回 dict 列表：{日期, 代码, 名称, 份额}（份额单位：万份）。
    """
    results = []
    pageno = 1
    while True:
        data = _get_json(
            {
                "txtDm": str(code).strip(),
                "txtStart": start,
                "txtEnd": end,
                "tab1PAGENO": pageno,
            }
        )
        block = data[0]
        meta = block.get("metadata", {})
        if block.get("error"):
            # 接口级错误（如参数问题）直接抛出，便于上层感知
            raise RuntimeError("深交所接口返回错误：%s" % block.get("error"))
        rows = block.get("data") or []
        if not rows:
            break
        for r in rows:
            day = str(r.get("size_date", "")).strip()
            if not day:
                continue
            results.append(
                {
                    "日期": day,
                    "代码": str(r.get("fund_code", "")).strip(),
                    "名称": str(r.get("security_short_name", "")).strip(),
                    "份额": _parse_size(r.get("current_size")),
                }
            )
        pagecount = int(meta.get("pagecount", 1) or 1)
        if pageno >= pagecount:
            break
        pageno += 1
        _sleep(PAGE_SLEEP)  # 翻页之间随机休眠
    return results


def split_windows(start, end, max_span=MAX_SPAN_DAYS):
    """将 [start, end] 按 max_span 天切分为若干 (窗口起, 窗口止) 区间。"""
    d0 = datetime.strptime(start, "%Y-%m-%d").date()
    d1 = datetime.strptime(end, "%Y-%m-%d").date()
    windows = []
    cur = d0
    while cur <= d1:
        win_end = min(cur + timedelta(days=max_span - 1), d1)
        windows.append((cur.strftime("%Y-%m-%d"), win_end.strftime("%Y-%m-%d")))
        cur = win_end + timedelta(days=1)
    return windows


def fetch_range(code, start, end):
    """抓取某 ETF 在 [start, end] 区间内的逐日份额（自动切分窗口 + 翻页 + 去重）。

    参数：code 6 位代码；start/end 为 YYYY-MM-DD。
    返回：按日期升序的 dict 列表 {日期, 代码, 名称, 份额}。非交易日不会出现在结果中。
    """
    code = str(code).strip()
    if datetime.strptime(start, "%Y-%m-%d") > datetime.strptime(end, "%Y-%m-%d"):
        return []

    by_date = {}
    windows = split_windows(start, end)
    for i, (ws, we) in enumerate(windows):
        for row in fetch_window(code, ws, we):
            if row["日期"]:
                by_date[row["日期"]] = row
        if i < len(windows) - 1:
            _sleep(WINDOW_SLEEP)  # 窗口之间随机休眠

    return [by_date[k] for k in sorted(by_date)]


def fetch_share(code, date):
    """抓取指定 ETF 在指定单日的份额数据。

    返回 dict {日期, 代码, 名称, 份额}（份额单位：万份），非交易日/无数据返回 None。
    """
    rows = fetch_range(code, date, date)
    return rows[0] if rows else None


if __name__ == "__main__":
    import json

    print("单日 159919 / 2026-06-05:")
    print(json.dumps(fetch_share("159919", "2026-06-05"), ensure_ascii=False, indent=2))
    print("非交易日 2026-06-06(周六):", fetch_share("159919", "2026-06-06"))
    rng = fetch_range("159919", "2026-06-01", "2026-06-05")
    print("区间条数:", len(rng))
    for r in rng:
        print(" ", r)
