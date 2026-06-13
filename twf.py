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
from dataclasses import dataclass

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


# ============================================
# TWF記載の構造化パース（専用レイアウト用）
# ============================================
@dataclass
class TwfDetailInfo:
    """TWF記載を分解した結果（TWF専用回答書の列表示用）

    「TWFNo.003243　三友工業様　お持ち帰り」
    → number="003243", customer="三友工業様", memo="お持ち帰り"
    """

    number: str = ""    # 番号（半角化済み。「不明」もあり得る。なければ""）
    customer: str = ""  # お客様名（最初の「様」まで。「様向け」特例あり）
    memo: str = ""      # 備考メモ（お持ち帰り・サービス品・着日等）


# 番号は「数字で始まり、以降は数字か？（全半角）の連続」を1つの番号として取る。
# FAXで読めない桁を担当者が？で埋めて桁数（6桁）を保つ運用に対応
# （例: 「0014？？」＝頭4桁は読めたが残り2桁不明）。？はお客様名ではなく番号の一部。
_TWF_NUM_RE = re.compile(r"^([0-9０-９][0-9０-９？?]*|不明)")


def _split_customer_memo(tail: str) -> tuple[str, str]:
    """後続テキストをお客様名と備考メモに振り分ける。

    ルール（実データ125明細の分析に基づく）:
    - 最初の「様」まで（直後に「向け」が続く場合はそこまで）→ お客様名
    - 残り → 備考メモ
    - 「様」がなければ全文を備考メモへ（社名でも誤分類より安全側に倒す）
    """
    idx = tail.find("様")
    if idx < 0:
        return "", tail
    end = idx + 1
    if tail[end:end + 2] == "向け":
        end += 2
    return tail[:end].strip(), tail[end:].strip()


def parse_twf_comment(comment: str) -> TwfDetailInfo | None:
    """コメント（明細）のTWF記載を構造化する。TWF記載がなければNone。"""
    if not comment:
        return None
    m = _TWF_EXTRACT_RE.search(comment)
    if m is None:
        return None
    body = re.sub(r"[ \t　]+", " ", m.group(1)).strip()
    nm = _TWF_NUM_RE.match(body)
    if nm:
        number = unicodedata.normalize("NFKC", nm.group(1))  # 全角数字→半角
        tail = body[nm.end():].strip()
    else:
        number = ""
        tail = body
    customer, memo = _split_customer_memo(tail)
    return TwfDetailInfo(number=number, customer=customer, memo=memo)


def build_twf_info_map(orders: list[OrderRow]) -> dict[str, TwfDetailInfo]:
    """注番→TWF情報のマップを構築する（入れ忘れ明細への引き継ぎ用）。

    同一注番に複数のTWF記載がある場合は最初に出現した明細を採用する。
    引き継ぐのは番号とお客様名のみ（memoは明細固有のため利用側で捨てる）。
    """
    info_map: dict[str, TwfDetailInfo] = {}
    for row in orders:
        order_num = row.order_number.strip()
        if order_num in info_map:
            continue
        info = parse_twf_comment(row.comment_detail)
        if info is not None:
            info_map[order_num] = info
    return info_map


# ============================================
# 持ち帰り判定（納入先名の「お引き取り」上書き用）
# ============================================
# 実データ12明細の表記: お持ち帰り / お持ち帰り済み / お持ち帰り済 /
# お渡し済み / 6/15or6/16引取りとなります。
_TWF_PICKUP_RE = re.compile(r"持ち?帰|持帰|引き?取|渡し済")

# 純粋な「お持ち帰り」表記（追加情報なし）→ 備考から消してよい
_TWF_PICKUP_ONLY_RE = re.compile(r"^お?持ち?帰り?$")


def is_twf_pickup_memo(memo: str) -> bool:
    """TWFメモが持ち帰り・引取系か判定する（納入先を「お引き取り」に上書き）。"""
    if not memo:
        return False
    return bool(_TWF_PICKUP_RE.search(unicodedata.normalize("NFKC", memo)))


def is_twf_pickup_only_memo(memo: str) -> bool:
    """メモが純粋な「お持ち帰り」表記のみか判定する。

    Trueなら納入先「お引き取り」と完全重複のため備考から省略する。
    「済み」「日程」等の追加情報がある場合はFalse（備考に残す）。
    """
    if not memo:
        return False
    return bool(_TWF_PICKUP_ONLY_RE.match(unicodedata.normalize("NFKC", memo).strip()))


def twf_sort_key(
    number: str, order_number: str, detail_number: str
) -> tuple[int, int, str, int]:
    """TWF専用回答書の行ソートキー。

    TWF No.昇順 → 同一No.内は注番→明細順。番号なし・「不明」は末尾。
    ？を含む番号（例「0014？？」）は ？を0とみなして数値化し、読めている桁の
    位置に並べる（「0014？？」→001400相当）。純粋な数字の順序は不変。
    """
    detail_str = str(detail_number).strip()
    detail = int(detail_str) if detail_str.isdigit() else 0
    core = number.replace("？", "0").replace("?", "0")
    if core.isdigit():
        return (0, int(core), order_number.strip(), detail)
    return (1, 0, order_number.strip(), detail)


# ============================================
# 文言・表示定数
# ============================================
TWF_REPORT_TITLE = "納期回答書（東京ウェルディングフェスタ2026 ご注文分）"
"""TWF専用回答書のタイトル（通常は「納　期　回　答　書」）。"""

TWF_THANKS_EXCEL = (
    "このたびは『東京ウェルディングフェスタ2026』にて"
    "多大なるご尽力を賜り、誠にありがとうございました。"
)
"""ご連絡事項欄の感謝文（回答書がメールから切り離されても感謝が伝わるように）。"""

TWF_NOTICE_EXCEL = (
    "※本書は「東京ウェルディングフェスタ2026」にてご注文いただきました"
    "商品の納期回答書です。納品完了まで最新の状況を毎回ご案内いたします。"
)

# メール本文用。\n は表示時に改行（<br>）に変換される
TWF_NOTICE_EMAIL = (
    "このたびは『東京ウェルディングフェスタ2026』にて、"
    "多大なるご尽力を賜り誠にありがとうございました。\n"
    "展示会でご成約いただきました商品の納期は、"
    "添付の専用回答書にてご案内しております。"
    "納品完了まで、その時点の最新状況を毎回一覧でお知らせいたしますので、"
    "進捗のご確認や、お客様へのご連絡にご活用ください。"
)

TWF_FILENAME_TAG = "【TWF2026】"
"""TWF専用回答書のファイル名タグ（納期回答書【TWF2026】_顧客名様_...）。"""

TWF_SHEET_PREFIX = "TWF2026_"
"""TWF専用回答書のシート名プレフィックス（31文字制限はbuild_sheet_nameが処理）。"""


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
