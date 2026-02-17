"""特定注番の日付情報確認"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from nouki_kaitou.data_loader import load_source_file, get_column_positions

def main():
    target_order = "GL2F444945"

    print(f"16AM.xlsから {target_order} の日付情報を取得中...")
    xls_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\16AM.xls")
    source_data = load_source_file(xls_path)
    cols = get_column_positions(source_data)

    if cols is None:
        print("列位置を取得できませんでした")
        return

    # 利用可能な列を表示
    print(f"\n利用可能な列: {list(cols.keys())}")

    col_order = cols.get("受発注伝票", 0)
    col_detail = cols.get("明細", 1)
    col_specified_date = cols.get("指定納期")
    col_order_date = cols.get("受注納期")
    col_register_date = cols.get("登録日")
    col_product = cols.get("品名")
    col_customer = cols.get("受注先")

    print(f"\n列位置:")
    print(f"  指定納期: {col_specified_date}")
    print(f"  受注納期: {col_order_date}")
    print(f"  登録日: {col_register_date}")

    # ヘッダー行を表示（生データ確認用）
    header_row = source_data[4]
    print(f"\nヘッダー行（関連列）:")
    if col_specified_date:
        print(f"  列{col_specified_date}: {header_row[col_specified_date] if col_specified_date < len(header_row) else 'N/A'}")
    if col_order_date:
        print(f"  列{col_order_date}: {header_row[col_order_date] if col_order_date < len(header_row) else 'N/A'}")
    if col_register_date:
        print(f"  列{col_register_date}: {header_row[col_register_date] if col_register_date < len(header_row) else 'N/A'}")

    # 対象注番を検索
    found = False
    for row_idx in range(5, len(source_data)):
        row = source_data[row_idx]
        if len(row) <= col_order:
            continue
        order_num = str(row[col_order] or "").strip()
        if order_num == target_order:
            found = True
            detail_num = str(row[col_detail] if col_detail < len(row) else "").strip()

            # 日付を生データのまま取得
            specified_raw = row[col_specified_date] if col_specified_date and col_specified_date < len(row) else ""
            order_date_raw = row[col_order_date] if col_order_date and col_order_date < len(row) else ""
            register_raw = row[col_register_date] if col_register_date and col_register_date < len(row) else ""
            product = str(row[col_product] if col_product and col_product < len(row) else "").strip()
            customer = str(row[col_customer] if col_customer and col_customer < len(row) else "").strip()

            print(f"\n{'='*80}")
            print(f"注番: {target_order}|{detail_num}")
            print(f"{'='*80}")
            print(f"受注先: {customer}")
            print(f"品名: {product}")
            print(f"\n【日付情報（生データ）】")
            print(f"  指定納期: '{specified_raw}'")
            print(f"  受注納期: '{order_date_raw}'")
            print(f"  登録日: '{register_raw}'")

    if not found:
        print(f"\n{target_order} が見つかりませんでした")

if __name__ == "__main__":
    main()
