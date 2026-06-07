# -*- coding: utf-8 -*-
"""深交所 ETF 份额查询与历史分析 Web 程序（Flask 入口）。

提供：
- 单日查询：按代码 + 日期抓取深交所份额数据并入库；
- 批量回填：抓取某 ETF 一个日期区间的逐日份额（自动跳过非交易日）；
- 历史分析：读取本地 xlsx，返回逐日份额序列供前端折线图 + 明细表展示。
"""

import re
import time
from datetime import datetime, timedelta

from flask import Flask, jsonify, render_template, request

import szse_client
import storage

app = Flask(__name__)

# 日期格式校验
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# ETF 代码校验（6 位数字）
_CODE_RE = re.compile(r"^\d{6}$")


def _valid_date(s):
    if not s or not _DATE_RE.match(s):
        return False
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def _daterange(start, end):
    """生成 start~end（含端点）的逐日日期字符串。"""
    d0 = datetime.strptime(start, "%Y-%m-%d").date()
    d1 = datetime.strptime(end, "%Y-%m-%d").date()
    cur = d0
    while cur <= d1:
        yield cur.strftime("%Y-%m-%d")
        cur += timedelta(days=1)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/query", methods=["POST"])
def api_query():
    """单日查询：抓取并入库，返回当日份额。"""
    body = request.get_json(silent=True) or {}
    code = str(body.get("code", "")).strip()
    date = str(body.get("date", "")).strip()

    if not _CODE_RE.match(code):
        return jsonify({"ok": False, "msg": "请输入正确的 6 位 ETF 代码"}), 400
    if not _valid_date(date):
        return jsonify({"ok": False, "msg": "请输入正确的日期（YYYY-MM-DD）"}), 400

    try:
        rec = szse_client.fetch_share(code, date)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "msg": "抓取失败：%s" % exc}), 502

    if rec is None:
        return jsonify(
            {"ok": False, "msg": "该日无数据（可能为非交易日或代码不存在）"}
        )

    # 入库（按日期去重）
    storage.save_records(code, [rec])
    return jsonify({"ok": True, "data": rec})


@app.route("/api/backfill", methods=["POST"])
def api_backfill():
    """批量回填：抓取日期区间逐日份额，跳过非交易日。"""
    body = request.get_json(silent=True) or {}
    code = str(body.get("code", "")).strip()
    start = str(body.get("start", "")).strip()
    end = str(body.get("end", "")).strip()

    if not _CODE_RE.match(code):
        return jsonify({"ok": False, "msg": "请输入正确的 6 位 ETF 代码"}), 400
    if not _valid_date(start) or not _valid_date(end):
        return jsonify({"ok": False, "msg": "请输入正确的起止日期"}), 400
    if start > end:
        return jsonify({"ok": False, "msg": "开始日期不能晚于结束日期"}), 400

    dates = list(_daterange(start, end))
    if len(dates) > 400:
        return jsonify({"ok": False, "msg": "区间过大，请控制在 400 天以内"}), 400

    records = []
    skipped = 0
    failed = 0
    for d in dates:
        try:
            rec = szse_client.fetch_share(code, d)
        except Exception:
            failed += 1
            continue
        if rec is None:
            skipped += 1
        else:
            records.append(rec)
        time.sleep(0.3)  # 轻微限速，避免请求过快

    total = storage.save_records(code, records)
    return jsonify(
        {
            "ok": True,
            "msg": "回填完成",
            "成功天数": len(records),
            "跳过天数": skipped,
            "失败天数": failed,
            "文件总记录数": total,
        }
    )


@app.route("/api/history")
def api_history():
    """返回某 ETF 的历史逐日份额序列。"""
    code = str(request.args.get("code", "")).strip()
    if not _CODE_RE.match(code):
        return jsonify({"ok": False, "msg": "请输入正确的 6 位 ETF 代码"}), 400

    records = storage.load_history(code)
    name = ""
    for r in records:
        if r.get("名称"):
            name = r["名称"]
    return jsonify(
        {
            "ok": True,
            "code": code,
            "name": name,
            "count": len(records),
            "data": records,
        }
    )


if __name__ == "__main__":
    import os

    # 端口默认 5000，可用环境变量 PORT 覆盖（macOS 上 5000 常被 AirPlay 占用）
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="127.0.0.1", port=port, debug=True)
