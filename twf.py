"""東京ウェルディングフェスタ（TWF2026）展示会受注の専用回答書サポート

期間限定機能。TWF_END_DATE を過ぎると is_twf_active() が False を返し、
全機能が自動オフになる（従来動作と完全同一になる）。
展示会受注の納品が完了し運用が終わったら、このモジュールと呼出箇所を
削除してよい。

判定仕様:
- コメント（明細）を NFKC正規化 + 大文字化 + 空白除去 した文字列に
  「TWFNO」が含まれれば展示会受注の明細とみなす
  （例: 「TWFNo.003243　新成（株）」「ＴＷＦ№3243」「twf no 3243」）
- 「ＴＷＦ特価」のような No なしの記載は対象外（展示会前の特価先行受注）
- 判定は注番単位に伝播する: 同一注番の明細が1つでも該当すれば、
  その注番の全明細を展示会受注として扱う（手入力の入れ忘れ対策）
"""

from __future__ import annotations

import datetime
import re
import unicodedata

from nouki_kaitou.models import OrderRow, ReportResult


# ============================================
# 期間ゲート
# ============================================
TWF_END_DATE = datetime.date(2026, 7, 3)
"""TWF専用回答書の生成期限（この日まで生成する）。"""


def is_twf_active(today: datetime.date | None = None) -> bool:
    """TWF専用回答書の生成期間内かどうかを返す。"""
    if today is None:
        today = datetime.date.today()
    return today <= TWF_END_DATE


# ============================================
# TWF記載の判定
# ============================================
def normalize_twf_text(text: str) -> str:
    """TWF判定用にコメント文字列を正規化する。

    NFKC正規化（全角英数→半角、№→No 等）→ 大文字化 → 空白（全半角）除去。
    """
    if not text:
        return ""
    normalized = unicodedata.normalize("NFKC", text).upper()
    return "".join(normalized.split())


def is_twf_comment(comment: str) -> bool:
    """コメント（明細）が展示会受注のTWF記載を含むか判定する。

    正規化後に「TWFNO」を含むか（数字の有無は問わない）。
    「ＴＷＦ特価」等のNoなし記載は対象外。
    """
    return "TWFNO" in normalize_twf_text(comment)


def collect_twf_orders(
    orders: list[OrderRow],
) -> tuple[set[str], list[tuple[str, str, str]]]:
    """全受注からTWF展示会受注の注番セットを構築する。

    判定は注番単位に伝播する（1明細でも該当すれば注番全体が対象）。

    Returns:
        (TWF注番セット, 検知明細リスト[(注番, 明細, コメント原文), ...])
        検知明細リストは目視確認用のコンソール表示に使う。
    """
    order_numbers: set[str] = set()
    detected: list[tuple[str, str, str]] = []

    for row in orders:
        if is_twf_comment(row.comment_detail):
            order_numbers.add(row.order_number.strip())
            detected.append(
                (row.order_number.strip(), row.detail_number.strip(),
                 row.comment_detail.strip())
            )

    return order_numbers, detected


# ============================================
# 備考からのTWF記載除去・整形
# ============================================
# TWF + (空白) + No/№ から行末までを除去する
# （TWF記載には得意先名が続くため行末まで落とす）
_TWF_REMOVE_RE = re.compile(
    r"[TＴtｔ][WＷwｗ][FＦfｆ][ \t　]*(?:[NＮnｎ][OＯoｏ]|№)[^\n]*"
)

# 整形用: No/№ の後の任意ピリオドまで読み、残り（番号+後続テキスト）を捕捉
_TWF_EXTRACT_RE = re.compile(
    r"[TＴtｔ][WＷwｗ][FＦfｆ][ \t　]*(?:[NＮnｎ][OＯoｏ]|№)[.．]?[ \t　]*([^\n]*)"
)


def remove_twf_text(text: str) -> str:
    """備考表示用テキストからTWF記載（行末まで）を除去する。"""
    if not text:
        return text
    return _TWF_REMOVE_RE.sub("", text)


def format_twf_remark(comment: str) -> str:
    """コメント（明細）のTWF記載を備考表示用に整形する（TWF専用回答書のK列用）。

    「TWFNo.003243　新成（株）様」→「No.003243 新成（株）様」
    先頭のTWFを落として「No.」に統一し、連続空白は1つに圧縮する。
    後続テキスト（得意先名・ステータスメモ等）は分類せずそのまま表示。
    TWF記載がなければ空文字を返す。
    """
    if not comment:
        return ""
    m = _TWF_EXTRACT_RE.search(comment)
    if m is None:
        return ""
    rest = re.sub(r"[ \t　]+", " ", m.group(1)).strip()
    return f"No.{rest}" if rest else "No."


def build_twf_remark_map(orders: list[OrderRow]) -> dict[str, str]:
    """注番→整形済みTWF情報のマップを構築する。

    注番内でTWF記載のない明細（入れ忘れ救済分）に、同注番の他明細の
    TWF情報を引き継いで表示するために使う。同一注番に複数のTWF記載が
    ある場合は最初に出現した明細の記載を採用する。
    """
    remark_map: dict[str, str] = {}
    for row in orders:
        order_num = row.order_number.strip()
        if order_num in remark_map:
            continue
        if is_twf_comment(row.comment_detail):
            formatted = format_twf_remark(row.comment_detail)
            if formatted:
                remark_map[order_num] = formatted
    return remark_map


# ============================================
# 文言・表示定数
# ============================================
TWF_REPORT_TITLE = "納期回答書（東京ウェルディングフェスタ ご注文分）"
"""TWF専用回答書のタイトル（通常は「納　期　回　答　書」）。"""

TWF_NOTICE_EXCEL = (
    "※本書は「東京ウェルディングフェスタ」にてご注文いただきました"
    "商品の納期回答書です。納品完了まで最新の状況を毎回ご案内いたします。"
)

TWF_NOTICE_EMAIL = (
    "東京ウェルディングフェスタにてご注文いただきました商品の納期は、"
    "添付の専用回答書にてご案内しております。"
    "専用回答書には納品完了まで、その時点の最新状況を毎回一覧で"
    "記載いたしますので、進捗のご確認にご活用ください。"
)

TWF_FILENAME_TAG = "【展示会】"
"""TWF専用回答書のファイル名タグ（納期回答書【展示会】_顧客名様_...）。"""

TWF_SHEET_PREFIX = "展示会_"
"""TWF専用回答書のシート名プレフィックス。"""


# ============================================
# メール入力の統合（通常 + TWF → 1メール）
# ============================================
def merge_email_input(
    normal: ReportResult | None,
    twf: ReportResult | None = None,
) -> dict:
    """通常回答書とTWF回答書のReportResultを1メール分の入力dictに統合する。

    どちらか一方がNoneでもよい（TWF受注のみの顧客 / TWF受注なしの顧客）。
    送り状・欠品・分納等の情報リストは両方を連結する。

    Returns:
        create_emails() の created_files 要素となるdict
    """
    base = normal or twf
    if base is None:
        raise ValueError("normal と twf の両方が None です")

    results = [r for r in (normal, twf) if r is not None]

    merged: dict = {
        "customer_name": base.customer_name,
        "rep_name": base.rep_name,
        "file_path": normal.file_path if normal else "",
        "attachments": [r.file_path for r in results if r.file_path],
        "stockout_info_list": [
            item for r in results for item in r.stockout_info_list
        ],
        "tracking_info_list": [
            item for r in results for item in r.tracking_info_list
        ],
        "bunno_info_list": [
            item for r in results for item in r.bunno_info_list
        ],
        "bunno_completed_list": [
            item for r in results for item in r.bunno_completed_list
        ],
        "twf_notice": TWF_NOTICE_EMAIL if twf is not None else None,
    }
    return merged
