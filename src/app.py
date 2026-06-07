# -*- coding: utf-8 -*-
"""深交所 ETF 份额查询与历史分析 Web 程序（Flask 入口）。

提供：
- 单日查询：按代码 + 日期抓取深交所份额数据并入库；
- 批量回填：抓取某 ETF 一个日期区间的逐日份额（自动跳过非交易日）；
- 历史分析：读取本地 xlsx，返回逐日份额序列供前端折线图 + 明细表展示。
"""

import re
import json
from datetime import datetime, timedelta

from flask import Flask, Response, jsonify, render_template, request

import szse_client
import storage

app = Flask(__name__)

# 日期格式校验
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# ETF 代码校验（6 位数字）
_CODE_RE = re.compile(r"^\d{6}$")
# 默认历史起始日期（与「历史份额分析」板块的默认开始日期一致）
DEFAULT_START_DATE = "2023-01-02"


def _valid_date(s):
    if not s or not _DATE_RE.match(s):
        return False
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return True
    except ValueError:
        return False


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


def _compute_missing(existing, start, end):
    """根据已有日期（升序）计算 [start, end] 内仍需抓取的子区间。

    采用「连续覆盖」假设：把已有数据视为覆盖了 [最早, 最晚] 这段连续区间，
    只抓取落在该区间之外的部分（即向前/向后扩展的增量），从而跳过已抓取的日期。
    若 [start, end] 完全落在已覆盖区间内，则返回空列表（无需请求）。
    """
    if not existing:
        return [(start, end)]

    sd = datetime.strptime(start, "%Y-%m-%d").date()
    ed = datetime.strptime(end, "%Y-%m-%d").date()
    lo = datetime.strptime(existing[0], "%Y-%m-%d").date()
    hi = datetime.strptime(existing[-1], "%Y-%m-%d").date()

    ranges = []
    if sd < lo:  # 向更早方向的缺口
        ranges.append((start, min(ed, lo - timedelta(days=1)).strftime("%Y-%m-%d")))
    if ed > hi:  # 向更晚方向的缺口
        ranges.append((max(sd, hi + timedelta(days=1)).strftime("%Y-%m-%d"), end))
    return ranges


def _backfill_events(code, start, end):
    """执行回填并逐步产出进度事件（dict）。最后写入 xlsx 并产出汇总事件。"""
    existing = storage.existing_dates(code)
    missing = _compute_missing(existing, start, end)

    # 把所有缺口区间切成窗口，得到总窗口数用于进度展示
    windows = []
    for ms, me in missing:
        windows.extend(szse_client.split_windows(ms, me))

    yield {
        "type": "start",
        "已存在天数": len(existing),
        "待抓取窗口数": len(windows),
        "缺口区间": [list(r) for r in missing],
    }

    if not windows:
        total = len(existing)
        yield {
            "type": "done",
            "msg": "该区间数据已全部存在，无需重复请求深交所。",
            "新增交易日数": 0,
            "文件总记录数": total,
        }
        return

    collected = {}
    for idx, (ws, we) in enumerate(windows, start=1):
        rows = szse_client.fetch_window(code, ws, we)
        for r in rows:
            if r["日期"]:
                collected[r["日期"]] = r
        yield {
            "type": "progress",
            "当前窗口": idx,
            "总窗口数": len(windows),
            "窗口区间": [ws, we],
            "本窗口交易日": len(rows),
            "累计新抓取": len(collected),
        }
        if idx < len(windows):
            szse_client._sleep(szse_client.WINDOW_SLEEP)

    records = [collected[k] for k in sorted(collected)]
    total = storage.save_records(code, records)
    yield {
        "type": "done",
        "msg": "回填完成",
        "新增交易日数": len(records),
        "文件总记录数": total,
    }


def _validate_backfill(code, start, end):
    """校验回填参数，返回错误信息字符串；通过则返回 None。"""
    if not _CODE_RE.match(code):
        return "请输入正确的 6 位 ETF 代码"
    if not _valid_date(start) or not _valid_date(end):
        return "请输入正确的起止日期"
    if start > end:
        return "开始日期不能晚于结束日期"
    span_days = (
        datetime.strptime(end, "%Y-%m-%d") - datetime.strptime(start, "%Y-%m-%d")
    ).days
    if span_days > 2200:
        return "区间过大，请控制在约 6 年以内"
    return None


@app.route("/api/backfill_stream")
def api_backfill_stream():
    """SSE 流式回填：实时向前端推送抓取进度。"""
    code = str(request.args.get("code", "")).strip()
    start = str(request.args.get("start", "")).strip()
    end = str(request.args.get("end", "")).strip()

    err = _validate_backfill(code, start, end)

    def gen():
        if err:
            yield "data: %s\n\n" % json.dumps({"type": "error", "msg": err}, ensure_ascii=False)
            return
        try:
            for ev in _backfill_events(code, start, end):
                yield "data: %s\n\n" % json.dumps(ev, ensure_ascii=False)
        except Exception as exc:  # noqa: BLE001
            yield "data: %s\n\n" % json.dumps(
                {"type": "error", "msg": "抓取失败：%s" % exc}, ensure_ascii=False
            )

    headers = {"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    return Response(gen(), mimetype="text/event-stream", headers=headers)


def _run_backfill(code, start, end):
    """消费 _backfill_events，返回最终事件（done/error）的 dict。"""
    final = {}
    for ev in _backfill_events(code, start, end):
        if ev.get("type") in ("done", "error"):
            final = ev
    return final


def _recent_trading_day():
    """最近一个非周末日期（今天若为工作日取今天，否则回退到上一个周五）。"""
    d = datetime.now().date()
    while d.weekday() >= 5:  # 5=周六, 6=周日
        d -= timedelta(days=1)
    return d.strftime("%Y-%m-%d")


@app.route("/api/backfill", methods=["POST"])
def api_backfill():
    """非流式回填（与流式同逻辑：跳过已有日期、切窗抓取），返回最终汇总。"""
    body = request.get_json(silent=True) or {}
    code = str(body.get("code", "")).strip()
    start = str(body.get("start", "")).strip()
    end = str(body.get("end", "")).strip()

    err = _validate_backfill(code, start, end)
    if err:
        return jsonify({"ok": False, "msg": err}), 400

    try:
        final = _run_backfill(code, start, end)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"ok": False, "msg": "抓取失败：%s" % exc}), 502

    if final.get("type") == "error":
        return jsonify({"ok": False, "msg": final.get("msg")}), 502
    return jsonify(
        {
            "ok": True,
            "msg": final.get("msg", "回填完成"),
            "新增交易日数": final.get("新增交易日数", 0),
            "文件总记录数": final.get("文件总记录数", 0),
        }
    )


@app.route("/api/delete", methods=["POST"])
def api_delete():
    """删除选中 ETF 的全部数据（含 data/代码.xlsx 文件）。"""
    body = request.get_json(silent=True) or {}
    codes = body.get("codes") or []
    if not codes:
        return jsonify({"ok": False, "msg": "请先勾选要删除的 ETF"}), 400

    deleted = []
    for code in codes:
        code = str(code).strip()
        if not _CODE_RE.match(code):
            continue
        if storage.delete_code(code):
            deleted.append(code)
    return jsonify(
        {"ok": True, "deleted": deleted, "msg": "已删除 %d 只 ETF 的数据" % len(deleted)}
    )


def _update_events(codes):
    """逐个更新选中 ETF 到最近交易日，产出带 ETF 维度的进度事件。"""
    end = _recent_trading_day()
    total = len(codes)
    results = []
    for idx, code in enumerate(codes, start=1):
        code = str(code).strip()
        yield {"type": "etf_start", "code": code, "etf_index": idx, "etf_total": total}
        if not _CODE_RE.match(code):
            results.append({"代码": code, "ok": False, "msg": "代码无效"})
            yield {"type": "etf_error", "code": code, "msg": "代码无效"}
            continue
        # 按默认区间 [DEFAULT_START_DATE, 最近交易日] 回填，已有日期由回填逻辑自动跳过
        start = DEFAULT_START_DATE
        added, filetotal, err = 0, 0, None
        try:
            for ev in _backfill_events(code, start, end):
                t = ev.get("type")
                if t == "start":
                    yield {
                        "type": "etf_meta", "code": code,
                        "etf_index": idx, "etf_total": total,
                        "total_windows": ev.get("待抓取窗口数", 0),
                    }
                elif t == "progress":
                    yield {
                        "type": "etf_progress", "code": code,
                        "etf_index": idx, "etf_total": total,
                        "win": ev.get("当前窗口"), "total_windows": ev.get("总窗口数"),
                    }
                elif t == "done":
                    added = ev.get("新增交易日数", 0)
                    filetotal = ev.get("文件总记录数", 0)
                elif t == "error":
                    err = ev.get("msg")
        except Exception as exc:  # noqa: BLE001
            err = "抓取失败：%s" % exc

        if err:
            results.append({"代码": code, "ok": False, "msg": err})
            yield {"type": "etf_error", "code": code, "etf_index": idx,
                   "etf_total": total, "msg": err}
        else:
            results.append({"代码": code, "ok": True, "新增交易日数": added,
                            "文件总记录数": filetotal})
            yield {"type": "etf_done", "code": code, "etf_index": idx,
                   "etf_total": total, "新增交易日数": added, "文件总记录数": filetotal}
    yield {"type": "all_done", "更新至": end, "data": results}


@app.route("/api/update_stream")
def api_update_stream():
    """SSE 流式更新：实时推送总进度、当前 ETF 及其窗口进度。"""
    codes_raw = str(request.args.get("codes", "")).strip()
    codes = [c.strip() for c in codes_raw.split(",") if c.strip()]

    def gen():
        if not codes:
            yield "data: %s\n\n" % json.dumps(
                {"type": "error", "msg": "请先勾选要更新的 ETF"}, ensure_ascii=False)
            return
        try:
            for ev in _update_events(codes):
                yield "data: %s\n\n" % json.dumps(ev, ensure_ascii=False)
        except Exception as exc:  # noqa: BLE001
            yield "data: %s\n\n" % json.dumps(
                {"type": "error", "msg": "更新失败：%s" % exc}, ensure_ascii=False)

    headers = {"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    return Response(gen(), mimetype="text/event-stream", headers=headers)


@app.route("/api/update", methods=["POST"])
def api_update():
    """将选中 ETF 的数据增量更新到最近一个交易日（非流式，保留备用）。"""
    body = request.get_json(silent=True) or {}
    codes = body.get("codes") or []
    if not codes:
        return jsonify({"ok": False, "msg": "请先勾选要更新的 ETF"}), 400

    end = _recent_trading_day()
    results = []
    for code in codes:
        code = str(code).strip()
        if not _CODE_RE.match(code):
            results.append({"代码": code, "ok": False, "msg": "代码无效"})
            continue
        # 按默认区间 [DEFAULT_START_DATE, 最近交易日] 回填，已有日期由回填逻辑自动跳过
        try:
            final = _run_backfill(code, DEFAULT_START_DATE, end)
        except Exception as exc:  # noqa: BLE001
            results.append({"代码": code, "ok": False, "msg": "抓取失败：%s" % exc})
            continue
        if final.get("type") == "error":
            results.append({"代码": code, "ok": False, "msg": final.get("msg")})
            continue
        results.append(
            {
                "代码": code,
                "ok": True,
                "新增交易日数": final.get("新增交易日数", 0),
                "文件总记录数": final.get("文件总记录数", 0),
            }
        )
    return jsonify({"ok": True, "更新至": end, "data": results})


@app.route("/api/history")
def api_history():
    """返回某 ETF 的历史逐日份额序列，可选 start/end 按区间过滤。"""
    code = str(request.args.get("code", "")).strip()
    if not _CODE_RE.match(code):
        return jsonify({"ok": False, "msg": "请输入正确的 6 位 ETF 代码"}), 400

    start = str(request.args.get("start", "")).strip()
    end = str(request.args.get("end", "")).strip()
    if start and not _valid_date(start):
        return jsonify({"ok": False, "msg": "开始日期格式不正确（YYYY-MM-DD）"}), 400
    if end and not _valid_date(end):
        return jsonify({"ok": False, "msg": "结束日期格式不正确（YYYY-MM-DD）"}), 400
    if start and end and start > end:
        return jsonify({"ok": False, "msg": "开始日期不能晚于结束日期"}), 400

    records = storage.load_history(code, start or None, end or None)
    name = ""
    for r in records:
        if r.get("名称"):
            name = r["名称"]
    return jsonify(
        {
            "ok": True,
            "code": code,
            "name": name,
            "start": start,
            "end": end,
            "count": len(records),
            "data": records,
        }
    )


@app.route("/api/summary")
def api_summary():
    """返回本地已存储数据的所有 ETF 汇总（代码、名称、数据时间段、记录数）。"""
    return jsonify({"ok": True, "data": storage.list_summary()})


@app.route("/api/order", methods=["POST"])
def api_order():
    """保存汇总表的自定义排列顺序（ETF 代码列表）。"""
    body = request.get_json(silent=True) or {}
    codes = body.get("codes") or []
    storage.save_order(codes)
    return jsonify({"ok": True})


if __name__ == "__main__":
    import os

    # 端口默认 5000，可用环境变量 PORT 覆盖（macOS 上 5000 常被 AirPlay 占用）
    port = int(os.environ.get("PORT", "5000"))
    # 监听地址默认仅本机(127.0.0.1)；如需局域网/外网访问，设 HOST=0.0.0.0
    host = os.environ.get("HOST", "127.0.0.1")
    # 安全：仅本机访问时才开启 debug 调试器（其会暴露任意代码执行风险），
    # 一旦监听到非本机地址则强制关闭 debug。
    debug = host in ("127.0.0.1", "localhost")
    # threaded=True：SSE 流式回填会长时间占用连接，需并发处理其它请求
    app.run(host=host, port=port, debug=debug, threaded=True)
