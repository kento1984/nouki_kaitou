"""SAPデータを検索するスクリプト"""

import sys
import io
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from nouki_kaitou.data_loader import (
    get_column_positions,
    get_data_rows_range,
    is_data_row,
    load_source_file,
    parse_order_row,
)


def main():
    order_no = sys.argv[1] if len(sys.argv) > 1 else "GL2V446963"

    source_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\17PM.xls")

    print(f"検索: {order_no}")
    print(f"ソース: {source_path}")
    print()

    source_data_raw = load_source_file(str(source_path))
    cols = get_column_positions(source_data_raw)

    if cols is None:
        print("ヘッダー行が見つかりません")
        return

    print("=== SAPデータ ===")
    for i in get_data_rows_range(source_data_raw, cols):
        if is_data_row(source_data_raw, i, cols):
            row = parse_order_row(source_data_raw, i, cols)
            if order_no in row.order_number:
                print(f"  注番|明細: {row.order_number}|{row.detail_number}")
                print(f"  顧客: {row.customer_name}")
                print(f"  出荷先: {row.ship_to_name}")
                print(f"  品名: {row.product_name}")
                print(f"  伝票タイプ: {row.document_type}")
                print(f"  出荷ステータス: {row.ship_status}")
                print(f"  保管場所: {row.storage_place}")
                print(f"  指定納期: {row.specified_delivery_date}")
                print(f"  受注納期: {row.order_delivery_date}")
                print(f"  コメント社内: {row.comment_internal[:50] if row.comment_internal else 'なし'}...")
                print()


if __name__ == "__main__":
    main()
