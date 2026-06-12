"""TWF展示会受注の判定検証スクリプト

最新の10PM.XLS（SAP受注一覧）をスキャンして以下を表示する:
1. TWF検知一覧: コメント（明細）が判定にヒットした明細（=展示会受注扱い）
2. 注番単位の集計: 注番伝播後の対象注番と顧客
3. 近似値一覧: 「TWF」を含むが「TWFNO」ではないコメント
   （「ＴＷＦ特価」等。判定対象外 — 入れ忘れ・表記ゆれの目視確認用）
   参考としてコメント（社外）（社内）のTWF記載も表示する

使い方:
    python -m dev_scripts.scan_twf [10PM.XLSのパス]
    （パス省略時はカレントディレクトリの10PM.XLSを探す）
"""

from __future__ import annotations

import sys
from pathlib import Path

# リポジトリ直下から `python dev_scripts/scan_twf.py` でも動くようにパス追加
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

from nouki_kaitou.data_loader import (
    get_column_positions,
    get_data_rows_range,
    is_data_row,
    load_source_file,
    parse_order_row,
)
from nouki_kaitou.twf import (
    TWF_END_DATE,
    collect_twf_orders,
    is_twf_active,
    is_twf_comment,
    normalize_twf_text,
)


def find_source(arg_path: str | None) -> Path:
    if arg_path:
        p = Path(arg_path)
        if not p.exists():
            print(f"エラー: ファイルが見つかりません: {p}")
            sys.exit(1)
        return p
    candidates = [
        p for p in Path.cwd().iterdir()
        if p.is_file() and p.name.upper() == "10PM.XLS"
    ]
    if not candidates:
        print("エラー: カレントディレクトリに10PM.XLSがありません。"
              "パスを引数で指定してください。")
        sys.exit(1)
    return max(candidates, key=lambda p: p.stat().st_mtime)


def main() -> None:
    source_path = find_source(sys.argv[1] if len(sys.argv) > 1 else None)
    print(f"ソースファイル: {source_path}")
    print(f"TWF期間ゲート: {TWF_END_DATE} まで"
          f"（現在 {'有効' if is_twf_active() else '期限切れ・機能オフ'}）")
    print()

    source_data = load_source_file(source_path)
    result = get_column_positions(source_data)
    if result is None:
        print("エラー: ヘッダー行を検出できませんでした。")
        sys.exit(1)
    cols, header_row_idx = result

    orders = []
    for i in get_data_rows_range(source_data, cols, header_row_idx):
        if is_data_row(source_data, i, cols):
            orders.append(parse_order_row(source_data, i, cols))
    print(f"受注データ: {len(orders)}明細")

    # --- 1. TWF検知一覧（コメント（明細）が判定対象） ---
    twf_orders, detected = collect_twf_orders(orders)
    print()
    print("=" * 60)
    print(f"【1. TWF検知一覧】 {len(detected)}明細が判定にヒット")
    print("=" * 60)
    for onum, dnum, cmt in detected:
        print(f"  {onum}|{dnum}: {cmt}")

    # --- 2. 注番単位の集計（注番伝播後） ---
    print()
    print("=" * 60)
    print(f"【2. 注番単位の集計】 {len(twf_orders)}注番が展示会受注扱い")
    print("=" * 60)
    by_order: dict[str, tuple[str, int, int]] = {}
    for o in orders:
        onum = o.order_number.strip()
        if onum in twf_orders:
            cust, total, hit = by_order.get(onum, (o.customer_name.strip(), 0, 0))
            total += 1
            if is_twf_comment(o.comment_detail):
                hit += 1
            by_order[onum] = (cust, total, hit)
    for onum in sorted(by_order):
        cust, total, hit = by_order[onum]
        mark = "" if hit == total else f"  ※TWF記載 {hit}/{total} 明細（注番伝播で全明細対象）"
        print(f"  {onum} | {cust} | {total}明細{mark}")

    # --- 3. 近似値一覧（TWFを含むがTWFNOではない = 判定対象外） ---
    print()
    print("=" * 60)
    print("【3. 近似値一覧】 TWFを含むがTWFNOではないコメント（判定対象外）")
    print("    ※展示会受注の入れ忘れ・表記ゆれがないか目視確認してください")
    print("=" * 60)
    near_count = 0
    columns = [
        ("コメント（明細）", lambda o: o.comment_detail),
        ("コメント（社外）", lambda o: o.comment_external),
        ("コメント（社内）", lambda o: o.comment_internal),
    ]
    for o in orders:
        for col_name, getter in columns:
            text = getter(o)
            norm = normalize_twf_text(text)
            if "TWF" in norm and "TWFNO" not in norm:
                near_count += 1
                print(f"  [{col_name}] {o.order_number.strip()}|"
                      f"{o.detail_number.strip()} | {o.customer_name.strip()} | "
                      f"{text.strip()}")
    if near_count == 0:
        print("  （該当なし）")

    # 参考: 社外・社内コメントにTWFNO記載がある場合も表示（判定対象外の列）
    print()
    print("【参考】コメント（社外）（社内）のTWFNO記載（判定対象外の列）:")
    ref_count = 0
    for o in orders:
        for col_name, getter in columns[1:]:
            if is_twf_comment(getter(o)):
                ref_count += 1
                print(f"  [{col_name}] {o.order_number.strip()}|"
                      f"{o.detail_number.strip()} | {getter(o).strip()}")
    if ref_count == 0:
        print("  （該当なし）")


if __name__ == "__main__":
    main()
