# -*- coding: utf-8 -*-
"""送付履歴のキャッシュあり/なし読み込み時間を計測するベンチマーク"""

import os
import pickle
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from openpyxl import load_workbook

from nouki_kaitou.history import load_delivery_history
from nouki_kaitou.models import CacheStore

tool_folder = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ")
history_path = tool_folder / "送付履歴.xlsx"
cache_file = Path(str(history_path) + ".cache")

print(f"対象ファイル: {history_path}")
print(f"ファイル存在: {history_path.exists()}")
print(f"キャッシュ存在: {cache_file.exists()}")
print()

# ファイルサイズ
if history_path.exists():
    xlsx_size = os.path.getsize(str(history_path))
    print(f"送付履歴.xlsx:       {xlsx_size / 1024:.1f} KB")
if cache_file.exists():
    cache_size = os.path.getsize(str(cache_file))
    print(f"送付履歴.xlsx.cache: {cache_size / 1024:.1f} KB")
print()

# --- キャッシュなし: xlsx読み込み ---
t0 = time.perf_counter()
wb = load_workbook(str(history_path), read_only=True, data_only=True)
t_open = time.perf_counter()
ws_h = wb["送付履歴"]
ws_c = wb["確認中一覧"]
cache_store = CacheStore()
sent = load_delivery_history(ws_h, ws_c, cache_store)
t_load = time.perf_counter()
wb.close()
t_close = time.perf_counter()

open_time = t_open - t0
load_time = t_load - t_open
total_no_cache = t_close - t0

print("=== キャッシュなし（xlsx読み込み） ===")
print(f"  xlsxオープン:          {open_time:.3f}秒")
print(f"  load_delivery_history: {load_time:.3f}秒")
print(f"  合計:                  {total_no_cache:.3f}秒")
print(f"  読み込み件数:          {len(sent)}件")
print()

# --- キャッシュあり: pickle読み込み ---
if cache_file.exists():
    t0 = time.perf_counter()
    with open(cache_file, "rb") as f:
        cached = pickle.load(f)
    t_pickle = time.perf_counter()
    pickle_time = t_pickle - t0
    cached_count = len(cached.get("sent_orders", {}))

    print("=== キャッシュあり（pickle読み込み） ===")
    print(f"  pickle読み込み:        {pickle_time:.3f}秒")
    print(f"  読み込み件数:          {cached_count}件")
    print()
    print("=== 比較 ===")
    print(f"  キャッシュなし: {total_no_cache:.3f}秒")
    print(f"  キャッシュあり: {pickle_time:.3f}秒")
    if pickle_time > 0:
        ratio = total_no_cache / pickle_time
        print(f"  高速化倍率:     {ratio:.1f}x")
    print(f"  短縮時間:       {total_no_cache - pickle_time:.3f}秒")
else:
    print("キャッシュファイルなし（初回実行の計測のみ）")
