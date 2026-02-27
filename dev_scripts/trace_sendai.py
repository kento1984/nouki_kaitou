"""sendai_delivery.py の check_sendai_stock_completed トレーススクリプト"""
import os, sys, json, datetime, openpyxl

# nouki_kaitou パッケージの親ディレクトリを追加
_parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_grandparent = os.path.dirname(_parent)
sys.path.insert(0, _grandparent)
sys.path.insert(0, _parent)

FOLDER = r"C:\Users\kento.kashiwabara\Desktop\納期回答書ツールフォルダ（DT)"
PROJECT = r"C:\Users\kento.kashiwabara\nouki_kaitou"

from nouki_kaitou.cache import build_customer_cache, build_pattern_cache
from nouki_kaitou.config import load_branch_settings, load_holidays
from nouki_kaitou.data_loader import load_source_file, get_column_positions
from nouki_kaitou.models import CacheStore, OrderRow
from nouki_kaitou.utils import parse_date, parse_time, normalize_name_for_comparison
from nouki_kaitou.sendai_delivery import check_sendai_stock_completed, _calc_pattern_days
from nouki_kaitou.customer import is_route_delivery, get_customer_pattern, get_customer_delivery_days
from nouki_kaitou.business_days import _is_weekend, is_holiday, add_business_days

# --- Load masters from desktop folder ---
mfg_wb = openpyxl.load_workbook(os.path.join(FOLDER, "メーカー一覧.xlsx"), read_only=True, data_only=True)
cust_wb = openpyxl.load_workbook(os.path.join(FOLDER, "顧客マスター_v2.xlsm"), read_only=True, data_only=True)

sap_path = os.path.join(PROJECT, "26PM.XLS")
data = load_source_file(sap_path)
cols = get_column_positions(data)

branch = load_branch_settings(mfg_wb, data, cols)
holidays = load_holidays(mfg_wb)

cust_days, cust_retention, cust_route, cust_pattern = build_customer_cache(cust_wb)
delivery_patterns = build_pattern_cache(mfg_wb)

cache = CacheStore()
cache.cust_days = cust_days
cache.cust_retention = cust_retention
cache.cust_route = cust_route
cache.cust_pattern = cust_pattern
cache.delivery_patterns = delivery_patterns

out = {
    "branch_name": branch.name,
    "branch_base_center": branch.base_center,
    "sendai_in_name": "仙台" in branch.name,
    "delivery_patterns": {
        k: {
            "cutoff1": list(v.cutoff1),
            "days_before_cutoff1": v.days_before_cutoff1,
            "cutoff2": list(v.cutoff2) if v.cutoff2 else None,
            "days_between": v.days_between_cutoffs,
            "days_after_all": v.days_after_all,
        }
        for k, v in delivery_patterns.items()
    },
    "okayasu_chiba_pattern": cust_pattern.get("岡安産業（株）\u3000千葉営業所", "(not found)"),
}

# --- Column indices ---
c = {name: cols.get(name) for name in [
    "受発注伝票", "明細", "受注先", "出荷先", "伝票タイプ",
    "出荷ステータス", "登録日", "時刻", "指定納期", "受注納期", "保管場所", "出荷先名"
]}

today = datetime.date(2026, 2, 26)
traces = []

for row in data[6:]:
    def cell(col_name):
        idx = c[col_name]
        if idx is not None and idx < len(row):
            return str(row[idx]).strip()
        return ""

    cust = cell("受注先")
    if "岡安" not in cust or "千葉" not in cust:
        continue

    doc_type = cell("伝票タイプ")
    ship_status = cell("出荷ステータス")
    time_val = cell("時刻")
    tp = parse_time(time_val)

    if not (tp and tp[0] >= 15 and doc_type == "【受注】在庫販売" and ship_status == "処理完了"):
        continue

    hour, minute = tp
    order_num = cell("受発注伝票")
    detail = cell("明細")
    reg_date = parse_date(row[c["登録日"]]) if c["登録日"] is not None else None
    spec_date = parse_date(row[c["指定納期"]]) if c["指定納期"] is not None else None
    ship_to = cell("出荷先名")
    storage = cell("保管場所")

    # Build OrderRow
    order_row = OrderRow()
    order_row.order_number = order_num
    order_row.detail_number = detail
    order_row.customer_name = cust
    order_row.ship_to_name = ship_to
    order_row.document_type = doc_type
    order_row.ship_status = ship_status
    order_row.registration_date = reg_date
    order_row.time_value = time_val
    order_row.specified_date = spec_date

    is_rosenbin = is_route_delivery(cust, cache)
    cust_eq_ship = (normalize_name_for_comparison(cust) == normalize_name_for_comparison(ship_to))
    use_ship_rule = (not cust_eq_ship) or is_rosenbin

    t = {"order": order_num, "detail": detail}

    # --- Step-by-step trace ---
    t["S1_branch"] = f"'仙台' in '{branch.name}' -> {('仙台' in branch.name)}"
    t["S2_doc_type"] = doc_type
    t["S3_ship_status"] = ship_status
    t["S4_reg_date"] = str(reg_date)
    t["S5_time"] = f"{hour:02d}:{minute:02d}"

    wkend = _is_weekend(reg_date) if reg_date else None
    hol = is_holiday(reg_date, holidays) if reg_date else None
    t["S5b_weekend_holiday"] = f"weekend={wkend}, holiday={hol}"

    pattern_name = get_customer_pattern(cust, cache)
    t["S6_pattern_name"] = pattern_name if pattern_name else "(empty)"

    pattern = cache.delivery_patterns.get(pattern_name) if pattern_name else None
    t["S7_pattern_found"] = pattern is not None

    if pattern:
        c1 = f"{pattern.cutoff1[0]:02d}:{pattern.cutoff1[1]:02d}"
        c2 = f"{pattern.cutoff2[0]:02d}:{pattern.cutoff2[1]:02d}" if pattern.cutoff2 else "None"
        t["S7_cutoffs"] = f"cutoff1={c1}, cutoff2={c2}"

        is_other = (branch.base_center and storage and storage != branch.base_center)
        t["S8_other_branch_stock"] = f"storage='{storage}' != '{branch.base_center}' -> {is_other}"

        t["S9_ship_rule"] = f"cust_eq_ship={cust_eq_ship}, is_rosenbin={is_rosenbin} -> use_ship_rule={use_ship_rule}"

        # Pattern days calculation
        biz_days = _calc_pattern_days(hour, minute, pattern)
        if (hour, minute) < pattern.cutoff1:
            reason = f"({hour:02d}:{minute:02d}) < cutoff1({c1}) -> days_before_cutoff1={pattern.days_before_cutoff1}"
        elif pattern.cutoff2 and (hour, minute) < pattern.cutoff2:
            reason = f"cutoff1 <= ({hour:02d}:{minute:02d}) < cutoff2({c2}) -> days_between={pattern.days_between_cutoffs}"
        else:
            reason = f"({hour:02d}:{minute:02d}) >= all cutoffs -> days_after_all={pattern.days_after_all}"
        t["S10_biz_days"] = f"{biz_days} ({reason})"

        if biz_days == 0:
            adjusted = reg_date
        else:
            adjusted = add_business_days(reg_date, biz_days, holidays)
        t["S11_adjusted_date"] = str(adjusted)

        del_days = get_customer_delivery_days(cust, cache)
        t["S12_delivery_day_restriction"] = str(del_days) if del_days else "none"

        suffix = "配達予定" if biz_days == 0 else ("配達済み" if adjusted <= today else "配達予定")
        t["S13_expected_suffix"] = suffix

    # Actual function call
    result = check_sendai_stock_completed(order_row, cache, holidays, branch, today, storage, use_ship_rule)
    t["ACTUAL_RESULT"] = result

    traces.append(t)

out["traces"] = traces

outpath = r"C:\Users\kento.kashiwabara\Desktop\sendai_trace.json"
with open(outpath, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print(f"wrote {outpath}")

mfg_wb.close()
cust_wb.close()
