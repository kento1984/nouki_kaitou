"""calculate_delivery_date を完全に通して sendai パスが使われるか確認"""
import os, sys, json, datetime, openpyxl

_parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(_parent))
sys.path.insert(0, _parent)

FOLDER = r"C:\Users\kento.kashiwabara\Desktop\納期回答書ツールフォルダ（DT)"
PROJECT = r"C:\Users\kento.kashiwabara\nouki_kaitou"

from nouki_kaitou.cache import build_all_caches
from nouki_kaitou.config import load_branch_settings, load_holidays
from nouki_kaitou.data_loader import load_source_file, get_column_positions, parse_order_row
from nouki_kaitou.delivery_calc import calculate_delivery_date
from nouki_kaitou.models import BranchSettings, CacheStore
from nouki_kaitou.sendai_delivery import check_sendai_stock_completed
from nouki_kaitou.utils import normalize_name_for_comparison
from nouki_kaitou.customer import is_route_delivery

# --- Load ---
mfg_wb = openpyxl.load_workbook(os.path.join(FOLDER, "メーカー一覧.xlsx"), data_only=True)
cust_wb = openpyxl.load_workbook(os.path.join(FOLDER, "顧客マスター_v2.xlsm"), data_only=True)

sap_path = os.path.join(PROJECT, "26PM.XLS")
data = load_source_file(sap_path)
cols = get_column_positions(data)

branch = load_branch_settings(mfg_wb, data, cols)
holidays = load_holidays(mfg_wb)
cache = build_all_caches(mfg_wb, cust_wb, None, data, cols)

today = datetime.date(2026, 2, 26)
execution_time = datetime.datetime(2026, 2, 26, 16, 0, 0)

out = {
    "branch": branch.name,
    "base_center": branch.base_center,
    "cust_pattern_count": len(cache.cust_pattern),
    "cust_pattern": dict(cache.cust_pattern),
    "delivery_patterns": {k: repr(v) for k, v in cache.delivery_patterns.items()},
}

# --- 岡安千葉の全伝票で calculate_delivery_date を呼ぶ ---
from nouki_kaitou.data_loader import get_data_rows_range, is_data_row

results = []
for i in get_data_rows_range(data, cols):
    if not is_data_row(data, i, cols):
        continue
    row = parse_order_row(data, i, cols)  # (source_data, row_idx, cols)
    if "岡安" not in row.customer_name or "千葉" not in row.customer_name:
        continue
    if row.document_type != "【受注】在庫販売" or row.ship_status != "処理完了":
        continue

    # calculate_delivery_date の完全呼び出し
    full_result = calculate_delivery_date(
        row, cache, holidays, branch, execution_time, today
    )

    # sendai 単独呼び出し（比較用）
    storage = cache.storage.get(row.order_number, "")
    is_rosenbin = is_route_delivery(row.customer_name, cache)
    same_cust = (normalize_name_for_comparison(row.customer_name)
                 == normalize_name_for_comparison(row.ship_to_name))
    use_ship = (not same_cust) or is_rosenbin

    sendai_result = check_sendai_stock_completed(
        row, cache, holidays, branch, today, storage, use_ship
    )

    results.append({
        "order": row.order_number,
        "detail": row.detail_number,
        "time": row.time_value,
        "reg_date": str(row.registration_date),
        "spec_date": str(row.specified_delivery_date),
        "storage": storage,
        "pattern": cache.cust_pattern.get(row.customer_name, ""),
        "full_result": full_result,
        "sendai_only": sendai_result,
        "match": full_result == sendai_result,
    })

out["results"] = results

outpath = r"C:\Users\kento.kashiwabara\Desktop\sendai_trace2.json"
with open(outpath, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print(f"wrote {outpath}")

mfg_wb.close()
cust_wb.close()
