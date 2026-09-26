"""从最新 Excel 重建预测偏差、价格及其相关关系。"""
import math
import re
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

DATA_DIR = Path(r"E:\数据")
SOURCE_CANDIDATES = [DATA_DIR / "9.1-9.15日前日内实际.xlsx", DATA_DIR / "9.1-9.15.xlsx"]
PRICE = DATA_DIR / "现货价格.xlsx"
DB = Path(__file__).resolve().parent / "data" / "market.sqlite"

SERIES = {
    "系统负荷": ("日系统负荷预测", "日系统负荷预测（日内）", "实际负荷"),
    "省间联络线": ("省间联络线输电曲线预测", "省间联络线输电曲线预测（日内）", "省间联络线实际输电情况"),
    "新能源出力": ("新能源总出力预测", "新能源总出力预测（日内）", "新能源总出力"),
    "水电含抽蓄": ("水电含抽蓄总出力预测", "水电含抽蓄总出力预测（日内）", "水电含抽蓄总出力"),
    "非市场机组": ("非市场机组总出力预测", "非市场机组总出力预测（日内）", "非市场机组总出力"),
    "火电发电空间": ("火电发电空间预测", None, "火电发电空间实际"),
    "风电出力": ("风电有功电力预测", "风电有功电力预测（日内）", "湖南风电发电有功电力"),
    "光伏出力": ("光伏有功电力预测", "光伏有功电力预测（日内）", "湖南光伏发电有功电力"),
}


def as_number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def mean(values):
    values = [v for v in values if v is not None]
    return sum(values) / len(values) if values else None


def excel_date(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()


def pearson(xs, ys):
    pairs = [(x, y) for x, y in zip(xs, ys) if x is not None and y is not None]
    if len(pairs) < 8:
        return None
    ax = mean([x for x, _ in pairs]); ay = mean([y for _, y in pairs])
    sx = sum((x - ax) ** 2 for x, _ in pairs); sy = sum((y - ay) ** 2 for _, y in pairs)
    return sum((x - ax) * (y - ay) for x, y in pairs) / math.sqrt(sx * sy) if sx and sy else None


def locate_series(ws):
    """定位指标名所在单元格，兼容前置表头行和列数变化。"""
    found = {}
    targets = {name for names in SERIES.values() for name in names if name}
    for row in range(1, ws.max_row + 1):
        for col in range(1, min(ws.max_column, 6) + 1):
            value = ws.cell(row, col).value
            if isinstance(value, str) and value.strip() in targets:
                found[value.strip()] = (row, col)
    return found


def quarter_values(ws, location):
    if not location:
        return [None] * 96
    row, label_col = location
    values = [as_number(ws.cell(row, col).value) for col in range(label_col + 1, min(ws.max_column, label_col + 96) + 1)]
    return (values + [None] * 96)[:96]


source = next((p for p in SOURCE_CANDIDATES if p.exists()), None)
if source is None:
    raise SystemExit("未找到运行数据 Excel")
# 普通模式会一次性建立单元格索引，适合本文件需要反复按坐标读取的场景。
# read_only 模式逐格随机访问很慢，15 天数据可能需要数分钟。
wb = openpyxl.load_workbook(source, read_only=False, data_only=True)
pwb = openpyxl.load_workbook(PRICE, read_only=False, data_only=True)
da_ws, rt_ws = pwb["9月日前"], pwb["9月实时"]
price_by_day = {}
for col in range(1, max(da_ws.max_column, rt_ws.max_column) + 1):
    dv = da_ws.cell(26, col).value if col <= da_ws.max_column else None
    if dv is None:
        continue
    trade_date = excel_date(dv)
    price_by_day[trade_date] = {
        "da": [as_number(da_ws.cell(r, col).value) for r in range(1, 25)] if col <= da_ws.max_column else [],
        "rt": [as_number(rt_ws.cell(r, col).value) for r in range(1, 25)] if col <= rt_ws.max_column else [],
    }

day_numbers = sorted({int(m.group(1)) for name in wb.sheetnames if (m := re.fullmatch(r"9\.(\d+)实际", name))})
records, quality, skipped = [], [], []
for day in day_numbers:
    intraday_name, actual_name = f"9.{day}日内", f"9.{day}实际"
    if intraday_name not in wb.sheetnames or actual_name not in wb.sheetnames:
        continue
    intraday_ws, actual_ws = wb[intraday_name], wb[actual_name]
    trade_date = f"2026-09-{day:02d}"
    embedded_dates = []
    for row in range(1, min(actual_ws.max_row, 8) + 1):
        for col in range(1, min(actual_ws.max_column, 6) + 1):
            value = actual_ws.cell(row, col).value
            if isinstance(value, datetime):
                embedded_dates.append(value.date().isoformat())
    embedded_date = embedded_dates[0] if embedded_dates else None
    # 同一批工作簿中，后续日期的工作表可能由前一天复制而来，
    # 内部日期单元格没有同步更新。工作表名明确包含交易日，因此以
    # 工作表名为准继续导入，同时把不一致记录进审计表供人工复核。
    if embedded_date and embedded_date != trade_date:
        skipped.append((trade_date, embedded_date, "内容日期不一致，已按工作表名称日期导入"))
    intraday_rows, actual_rows = locate_series(intraday_ws), locate_series(actual_ws)
    loaded_metrics = 0
    for metric, (da_name, id_name, actual_series) in SERIES.items():
        da_q = quarter_values(actual_ws, actual_rows.get(da_name))
        id_q = quarter_values(intraday_ws, intraday_rows.get(id_name)) if id_name else [None] * 96
        ac_q = quarter_values(actual_ws, actual_rows.get(actual_series))
        if not any(v is not None for v in da_q) or not any(v is not None for v in ac_q):
            continue
        loaded_metrics += 1
        for hour in range(24):
            sl = slice(hour * 4, hour * 4 + 4)
            da, intraday, actual = mean(da_q[sl]), mean(id_q[sl]), mean(ac_q[sl])
            prices = price_by_day.get(trade_date, {})
            da_prices, rt_prices = prices.get("da", []), prices.get("rt", [])
            da_price = da_prices[hour] if len(da_prices) > hour else None
            rt_price = rt_prices[hour] if len(rt_prices) > hour else None
            records.append((trade_date, hour, metric, da, intraday, actual,
                actual - da if actual is not None and da is not None else None,
                actual - intraday if actual is not None and intraday is not None else None,
                abs(actual - da) if actual is not None and da is not None else None,
                abs(actual - intraday) if actual is not None and intraday is not None else None,
                da_price, rt_price, rt_price - da_price if rt_price is not None and da_price is not None else None))
    quality.append((trade_date, loaded_metrics))

conn = sqlite3.connect(DB)
conn.executescript("""
DROP TABLE IF EXISTS forecast_deviation_hourly;
DROP TABLE IF EXISTS deviation_relationship;
DROP TABLE IF EXISTS data_import_audit;
CREATE TABLE forecast_deviation_hourly(
  trade_date TEXT NOT NULL, trade_hour INTEGER NOT NULL, metric_name TEXT NOT NULL,
  day_ahead_value REAL, intraday_value REAL, actual_value REAL,
  day_ahead_error REAL, intraday_error REAL, day_ahead_abs_error REAL, intraday_abs_error REAL,
  day_ahead_price REAL, real_time_price REAL, price_spread REAL,
  PRIMARY KEY(trade_date,trade_hour,metric_name));
CREATE INDEX idx_deviation_metric_date ON forecast_deviation_hourly(metric_name,trade_date,trade_hour);
CREATE TABLE deviation_relationship(
  metric_name TEXT NOT NULL, feature_name TEXT NOT NULL, target_name TEXT NOT NULL,
  correlation REAL, sample_count INTEGER NOT NULL,
  PRIMARY KEY(metric_name,feature_name,target_name));
CREATE TABLE data_import_audit(
  source_file TEXT NOT NULL, source_modified_at TEXT NOT NULL, imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
  first_trade_date TEXT, last_trade_date TEXT, day_count INTEGER, record_count INTEGER, notes TEXT);
""")
conn.executemany("INSERT INTO forecast_deviation_hourly VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", records)
for metric in SERIES:
    subset = [r for r in records if r[2] == metric]
    for flabel, fi in {"实际值": 5, "日前偏差": 6, "日内偏差": 7}.items():
        for tlabel, ti in {"日前价格": 10, "实时价格": 11, "实时-日前价差": 12}.items():
            xs, ys = [r[fi] for r in subset], [r[ti] for r in subset]
            corr = pearson(xs, ys)
            count = sum(x is not None and y is not None for x, y in zip(xs, ys))
            conn.execute("INSERT INTO deviation_relationship VALUES (?,?,?,?,?)", (metric, flabel, tlabel, corr, count))
stat = source.stat()
conn.execute("INSERT INTO data_import_audit(source_file,source_modified_at,first_trade_date,last_trade_date,day_count,record_count,notes) VALUES (?,?,?,?,?,?,?)",
    (source.name, datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"), min((q[0] for q in quality), default=None),
     max((q[0] for q in quality), default=None), len(quality), len(records),
     "; ".join([*(f"{d}:{n}项" for d, n in quality), *(f"跳过{d}(内容日期{e})" for d, e, _ in skipped)])))
conn.execute("PRAGMA optimize")
conn.commit()
print(f"source={source.name} days={len(quality)} records={len(records)} relationships={conn.execute('select count(*) from deviation_relationship').fetchone()[0]}")
print("quality=" + ",".join(f"{d}:{n}" for d, n in quality))
print("skipped=" + ",".join(f"{d}->{e}" for d, e, _ in skipped))
conn.close()

