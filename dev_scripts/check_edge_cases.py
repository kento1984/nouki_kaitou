"""エッジケース3件の詳細分析

1. 欠品中がいきなり処理完了になった場合
2. 紐付き（直送販売）で納期回答後の送付済み処理
3. 紐付きで受注納期12/31のまま処理完了になった場合
"""

import sys
import io
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from nouki_kaitou.data_loader import load_source_file, get_column_positions

def main():
    print("16AM.xlsを読み込み中...")
    xls_path = Path(r"\\flsv04\316京葉\納期回答書ツールフォルダ\受注一覧\16AM.xls")
    source_data = load_source_file(xls_path)
    cols = get_column_positions(source_data)

    if cols is None:
        print("列位置を取得できませんでした")
        return

    col_order = cols.get("受発注伝票", 0)
    col_detail = cols.get("明細", 1)
    col_customer = cols.get("受注先")
    col_ship_to = cols.get("出荷先名")
    col_product = cols.get("品名")
    col_ship_status = cols.get("出荷ステータス")
    col_comment_detail = cols.get("コメント（明細）")
    col_comment_internal = cols.get("コメント（社内）")
    col_specified = cols.get("指定納期")
    col_order_date = cols.get("受注納期")
    col_doc_type = cols.get("伝票タイプ")
    col_storage = cols.get("保管場所")

    # 対象注番
    target_orders = ["GL2H444845", "GL2C444955"]

    # カテゴリ別に収集
    kepppin_completed = []  # 欠品中+処理完了
    himozuki_orders = []    # 紐付き注文
    himozuki_dec31_completed = []  # 紐付き+12/31+処理完了

    for row_idx in range(5, len(source_data)):
        row = source_data[row_idx]
        if len(row) <= col_order:
            continue

        order_num = str(row[col_order] or "").strip()
        if not order_num:
            continue

        detail_num = str(row[col_detail] if col_detail < len(row) else "").strip()
        customer = str(row[col_customer] if col_customer and col_customer < len(row) else "").strip()
        ship_to = str(row[col_ship_to] if col_ship_to and col_ship_to < len(row) else "").strip()
        product = str(row[col_product] if col_product and col_product < len(row) else "").strip()
        ship_status = str(row[col_ship_status] if col_ship_status and col_ship_status < len(row) else "").strip()
        comment_detail = str(row[col_comment_detail] if col_comment_detail and col_comment_detail < len(row) else "").strip()
        comment_internal = str(row[col_comment_internal] if col_comment_internal and col_comment_internal < len(row) else "").strip()
        specified = str(row[col_specified] if col_specified and col_specified < len(row) else "").strip()
        order_date = str(row[col_order_date] if col_order_date and col_order_date < len(row) else "").strip()
        doc_type = str(row[col_doc_type] if col_doc_type and col_doc_type < len(row) else "").strip()
        storage = str(row[col_storage] if col_storage and col_storage < len(row) else "").strip()

        item = {
            "order": order_num,
            "detail": detail_num,
            "customer": customer[:25],
            "ship_to": ship_to[:25],
            "product": product[:35],
            "ship_status": ship_status,
            "comment_detail": comment_detail[:60],
            "comment_internal": comment_internal[:40],
            "specified": specified,
            "order_date": order_date,
            "doc_type": doc_type,
            "storage": storage,
        }

        # 1. 欠品中+処理完了
        if "欠品" in comment_detail and ship_status == "処理完了":
            kepppin_completed.append(item)

        # 紐付き判定（直送販売 + 保管場所≠転送中（直送用））
        is_himozuki = (
            doc_type == "【受注】直送販売"
            and storage != "転送中（直送用）"
        )

        # 2. 紐付き注文
        if is_himozuki:
            himozuki_orders.append(item)

        # 3. 紐付き+12/31+処理完了
        if is_himozuki and ship_status == "処理完了" and "12/31" in order_date:
            himozuki_dec31_completed.append(item)

    # 結果表示
    print("\n" + "=" * 110)
    print("【1】欠品中のまま処理完了になったケース（コメント明細に「欠品」+出荷ステータス「処理完了」）")
    print("=" * 110)
    print(f"該当件数: {len(kepppin_completed)}件\n")

    # 特定の注番があるか確認
    for target in target_orders:
        matches = [i for i in kepppin_completed if i["order"] == target]
        if matches:
            print(f"★ 指定注番 {target} が見つかりました:")
            for item in matches:
                print(f"   {item['order']}|{item['detail']}")
                print(f"   品名: {item['product']}")
                print(f"   コメント(明細): {item['comment_detail']}")
                print(f"   出荷ステータス: {item['ship_status']}")
                print()

    for i, item in enumerate(kepppin_completed[:10]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   顧客: {item['customer']}")
        print(f"   品名: {item['product']}")
        print(f"   出荷ステータス: {item['ship_status']}")
        print(f"   伝票タイプ: {item['doc_type']}")
        print(f"   コメント(明細): {item['comment_detail']}")
        print(f"   コメント(社内): {item['comment_internal']}")
        print(f"   指定納期: {item['specified']} / 受注納期: {item['order_date']}")
        print()

    if len(kepppin_completed) > 10:
        print(f"... 他 {len(kepppin_completed) - 10}件\n")

    print("\n【処理ロジック解説】")
    print("-" * 80)
    print("report_generator.py の処理:")
    print("1. build_report_row() で欠品判定時:")
    print("   if '欠品中' in row.comment_detail AND row.ship_status != '処理完了':")
    print("       → 欠品判定 (line 165)")
    print("")
    print("2. 処理完了の場合:")
    print("   - 欠品判定がスキップされる")
    print("   - force_delivered = True になる (line 702-706)")
    print("   - delivery_status が '確認中' や '欠品中' でも '納品済み' に上書き (line 456)")
    print("")
    print("★ 結論: 欠品中のまま処理完了 → 「納品済み」として出力される")

    print("\n" + "=" * 110)
    print("【2】紐付き（直送販売）注文")
    print("=" * 110)
    print(f"該当件数: {len(himozuki_orders)}件\n")

    # ステータス別に集計
    status_counts = {}
    for item in himozuki_orders:
        status = item["ship_status"]
        status_counts[status] = status_counts.get(status, 0) + 1

    print("ステータス別件数:")
    for status, count in sorted(status_counts.items(), key=lambda x: -x[1]):
        print(f"  {status}: {count}件")

    print("\n代表例（最大5件）:")
    for i, item in enumerate(himozuki_orders[:5]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   顧客: {item['customer']} → 出荷先: {item['ship_to']}")
        print(f"   品名: {item['product']}")
        print(f"   保管場所: {item['storage']}")
        print(f"   出荷ステータス: {item['ship_status']}")
        print(f"   受注納期: {item['order_date']}")
        print()

    print("\n【処理ロジック解説】")
    print("-" * 80)
    print("紐付き判定 (report_generator.py line 682-688):")
    print("  if doc_type == '【受注】直送販売':")
    print("      if storage != '転送中（直送用）':")
    print("          is_himozuki = True")
    print("")
    print("紐付き注文の特徴:")
    print("  - force_delivered = False を維持 (line 705, 709)")
    print("  - 処理完了でも「納品済み」にならない")
    print("  - 確認中一覧 or 送付履歴で管理される")

    print("\n" + "=" * 110)
    print("【3】紐付き + 受注納期12/31 + 処理完了")
    print("=" * 110)
    print(f"該当件数: {len(himozuki_dec31_completed)}件\n")

    for i, item in enumerate(himozuki_dec31_completed[:10]):
        print(f"{i+1}. {item['order']}|{item['detail']}")
        print(f"   顧客: {item['customer']} → 出荷先: {item['ship_to']}")
        print(f"   品名: {item['product']}")
        print(f"   保管場所: {item['storage']}")
        print(f"   出荷ステータス: {item['ship_status']}")
        print(f"   指定納期: {item['specified']} / 受注納期: {item['order_date']}")
        print(f"   コメント(社内): {item['comment_internal']}")
        print()

    if len(himozuki_dec31_completed) > 10:
        print(f"... 他 {len(himozuki_dec31_completed) - 10}件\n")

    print("\n【処理ロジック解説】")
    print("-" * 80)
    print("紐付き + 12/31 + 処理完了 の場合:")
    print("1. is_himozuki = True")
    print("2. force_delivered は False を維持")
    print("3. 送付履歴にない場合:")
    print("   - 通常の納期計算が適用される")
    print("   - 12/31は「未確定」として扱われる可能性")
    print("4. 送付履歴にある場合:")
    print("   - previous_status に応じて処理")
    print("")
    print("★ 結論: 紐付き注文は処理完了でも「納品済み」強制にならない")
    print("   → 確認中一覧で継続管理 or 元の納期回答のまま送付履歴へ")

if __name__ == "__main__":
    main()
