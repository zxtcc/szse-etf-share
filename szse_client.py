# -*- coding: utf-8 -*-
"""深交所 ETF 份额数据抓取模块。

数据来源（深交所官网公开接口）：
- 份额数据（用户指定的份额页面）：CATALOGID=1953
  http://www.szse.cn/market/fund/volume/etf/index.html
- ETF 中文简称：CATALOGID=1945（ETF 列表）

注意：深交所站点的 HTTPS 证书链不在 Python 默认 CA bundle 中，故请求时
使用 verify=False，并在模块加载时抑制相关告警。
"""

import re
import time
import random

import requests
import urllib3

# 抑制 verify=False 带来的 InsecureRequestWarning
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 接口基础地址
BASE_URL = "https://www.szse.cn/api/report/ShowReport/data"

# 份额数据接口 CATALOGID（对应用户给定的 ETF 份额页面）
CATALOG_SHARE = "1953"
# ETF 列表接口 CATALOGID（用于补充中文简称）
CATALOG_LIST = "1945"

# 请求头：深交所接口要求带 UA 与 Referer，否则可能拒绝或返回异常
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Referer": "https://www.szse.cn/market/fund/volume/etf/index.html",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
}

# 单元素中文简称缓存，避免同一进程内重复请求 1945 接口
_CN_NAME_CACHE = {}


def _get_json(params, retries=3, timeout=20):
    """统一的 GET JSON 封装，带重试。

    返回解析后的 JSON（深交所接口返回的是一个数组），失败抛出异常。
    """
    params = dict(params)
    params.setdefault("SHOWTYPE", "JSON")
    params.setdefault("TABKEY", "tab1")
    last_err = None
    for attempt in range(retries):
        # 每次带一个随机数，模拟前端行为，避免缓存
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


def _to_int_share(text):
    """将带千分位逗号的份额字符串转为整数，无法解析返回 None。"""
    if text is None:
        return None
    cleaned = str(text).replace(",", "").strip()
    if not cleaned:
        return None
    try:
        # 份额一般为整数，个别可能带小数，统一取浮点再转 int
        return int(float(cleaned))
    except ValueError:
        return None


def _to_float(text):
    """将净值字符串转为浮点，失败返回 None。"""
    if text is None:
        return None
    cleaned = str(text).replace(",", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def fetch_cn_name(code):
    """通过 ETF 列表接口（1945）获取中文简称，失败返回 None。

    1945 接口的 kzjcurl 字段是一段 HTML，中文简称包裹在 <u>...</u> 中。
    """
    code = str(code).strip()
    if code in _CN_NAME_CACHE:
        return _CN_NAME_CACHE[code]

    name = None
    try:
        data = _get_json({"CATALOGID": CATALOG_LIST, "txtQueryKeyAndJC": code})
        rows = data[0].get("data") or []
        for row in rows:
            html = row.get("kzjcurl", "") or ""
            m = re.search(r"<u>(.*?)</u>", html)
            if m:
                name = m.group(1).strip()
                break
    except Exception:
        name = None

    _CN_NAME_CACHE[code] = name
    return name


def fetch_share(code, date, with_cn_name=True):
    """抓取指定 ETF 在指定日期的份额数据。

    参数：
        code: ETF 代码，如 "159150"
        date: 日期字符串，格式 YYYY-MM-DD
        with_cn_name: 是否补充中文简称

    返回：
        成功 -> dict {日期, 代码, 名称, 份额, 净值}
        非交易日 / 无数据 -> None
    """
    code = str(code).strip()
    date = str(date).strip()

    data = _get_json(
        {
            "CATALOGID": CATALOG_SHARE,
            "txtQueryDate": date,
            "txtQueryKeyAndJC": code,
        }
    )

    block = data[0]
    record_count = block.get("metadata", {}).get("recordcount", 0)
    rows = block.get("data") or []
    if not record_count or not rows:
        # 非交易日或该日无该 ETF 数据
        return None

    # 在返回结果中精确匹配代码（接口按代码筛选，一般首条即是）
    row = None
    for r in rows:
        if str(r.get("fund_code", "")).strip() == code:
            row = r
            break
    if row is None:
        row = rows[0]

    name_en = (row.get("security_english_short_name") or "").strip()
    name = name_en
    if with_cn_name:
        cn = fetch_cn_name(code)
        if cn:
            name = cn

    return {
        "日期": date,
        "代码": code,
        "名称": name,
        "份额": _to_int_share(row.get("current_size")),
        "净值": _to_float(row.get("nav_per_share")),
    }


if __name__ == "__main__":
    # 简单自测
    import json

    print("中文名:", fetch_cn_name("159150"))
    print(
        json.dumps(
            fetch_share("159150", "2025-09-04"), ensure_ascii=False, indent=2
        )
    )
    print("周末测试(应为 None):", fetch_share("159150", "2025-09-06"))
