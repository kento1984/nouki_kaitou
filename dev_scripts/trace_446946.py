"""GL2C446946 の納期計算トレース"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from nouki_kaitou.data_loader import (
    load_source_file, get_column_positions, parse_order_row, get_data_rows_range, is_data_row
)
from nouki_kaitou.bunno import extract_bunno_info, has_bunno_mitei
from nouki_kaitou.delivery_calc import calculate_delivery_date
from nouki_kaitou.models import CacheStore

TARGET = "GL2C446946"

# 最新のSAPファイルを読み込む
source_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\17PM.xls")
if not source_path.exists():
    source_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\10PM.XLS")
    if not source_path.exists():
        print("ERROR: SAP source file not found")
        sys.exit(1)

print(f"Source: {source_path}")
data = load_source_file(source_path)
cols = get_column_positions(data)
if cols is None:
    print("ERROR: could not parse columns")
    sys.exit(1)

# GL2C446946 を探す
for row_idx in get_data_rows_range(data, cols):
    if not is_data_row(data, row_idx, cols):
        continue
    row = parse_order_row(data, row_idx, cols)
    if row.order_number != TARGET:
        continue

    print(f"\n{'='*70}")
    print(f"注番: {row.order_number}|{row.detail_number}")
    print(f"伝票タイプ: {row.document_type}")
    print(f"受注先: {row.customer_name}")
    print(f"出荷先名: {row.ship_to_name}")
    print(f"品名: {row.product_name}")
    print(f"メーカー: {row.manufacturer_name}")
    print(f"出荷ステータス: {row.ship_status}")
    print(f"保管場所: {row.storage_place}")
    print(f"受注納期: {row.order_delivery_date}")
    print(f"指定納期: {row.specified_delivery_date}")
    print(f"登録日: {row.registration_date}")
    print(f"時刻: {row.time_value}")
    print(f"受注数量: {row.quantity}")
    print(f"品目Group: {row.item_group_code}")
    print(f"拒否理由: [{row.rejection_reason}]")
    print(f"コメント（明細）: [{row.comment_detail}]")
    print(f"コメント（社内）: [{row.comment_internal}]")
    print(f"コメント（社外）: [{row.comment_external}]")
    print()

    # 分納判定
    bunno_info = extract_bunno_info(row.comment_detail)
    print(f"--- 分納判定 ---")
    print(f"extract_bunno_info結果: {bunno_info}")
    if bunno_info:
        print(f"  分納あり: {len(bunno_info)}件")
        for i, b in enumerate(bunno_info):
            print(f"    [{i}] qty={b.quantity}, date={b.date_str}, loc={b.location}")
    else:
        print(f"  分納なし")

    # 欠品判定
    print(f"\n--- 欠品判定 ---")
    print(f"'欠品中' in comment_detail: {'欠品中' in row.comment_detail}")
    print(f"ship_status != '処理完了': {row.ship_status != '処理完了'}")

    # 納期計算
    cache = CacheStore()
    delivery_answer = calculate_delivery_date(row, cache, None, None, None, None)
    print(f"\n--- 納期計算 ---")
    print(f"calculate_delivery_date結果: {delivery_answer}")

    # build_report_row相当のロジック
    if bunno_info and row.ship_status != "処理完了":
        delivery_answer = "分納"
        print(f"→ 分納判定で上書き: {delivery_answer}")

    if ("欠品中" in row.comment_detail
            and row.ship_status != "処理完了"):
        if delivery_answer in ("確認中", "日程調整中"):
            delivery_answer = "欠品中"
            print(f"→ 欠品判定で上書き(確認中→欠品中): {delivery_answer}")
        else:
            delivery_answer = delivery_answer + "（欠品）"
            print(f"→ 欠品判定で(欠品)付加: {delivery_answer}")

    print(f"\n★ 最終 delivery_status: {delivery_answer}")

    # _classify_order相当のロジック
    print(f"\n--- _classify_order相当 ---")
    is_confirming = (
        delivery_answer in ("確認中", "欠品中", "日程調整中")
        or "（欠品）" in delivery_answer
        or "分納" in delivery_answer
    )
    print(f"確認中一覧行き?: {is_confirming}")

    if is_confirming:
        if "分納" in delivery_answer:
            print(f"→ 分納パスへ")
            # has_bunno_miteiは簡易チェック（cache未構築のため）
            if bunno_info:
                for b in bunno_info:
                    print(f"   未定チェック: date_str=[{b.date_str}]")
        elif delivery_answer == "欠品中" or "（欠品）" in delivery_answer:
            ship_status_for_confirm = "欠品中"
            print(f"→ 欠品パスへ: ステータス={ship_status_for_confirm}")
        else:
            ship_status_for_confirm = row.ship_status
            print(f"→ デフォルトパス: ステータス={ship_status_for_confirm}")
    print(f"{'='*70}")
