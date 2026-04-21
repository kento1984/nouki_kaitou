"""汎用ユーティリティ関数

VBAの以下の関数を移植:
- ConvertToHalfWidth (L6883): 全角→半角変換
- ToHalfWidthNum (L6462): 全角→半角数字変換
- ToCircledNumber (L6631): 数字→丸数字変換
- ExtractDateFromString (L5470): 文字列から日付抽出
- HtmlEscape (L7752): → 標準ライブラリ html.escape で代替
- IsFileOpen (L6359): → Python流のファイルロック確認で代替
"""

from __future__ import annotations

import datetime
import os
import re
from pathlib import Path


# ============================================
# VBA: ConvertToHalfWidth (L6883)
# 全角数字・スラッシュ・コロンを半角に変換
# ============================================
def convert_to_half_width(text: str) -> str:
    """全角数字（０-９）・全角スラッシュ（／）・全角コロン（：）を半角に変換する。

    VBA版ではAscWで1文字ずつ処理していたが、Pythonではstr.translateで一括変換。
    """
    if not text:
        return ""

    # 全角数字 → 半角数字
    table = str.maketrans("０１２３４５６７８９／：", "0123456789/:")
    return text.translate(table)


# ============================================
# VBA: ToHalfWidthNum (L6462)
# 全角数字1文字を半角に変換
# ============================================
def to_half_width_num(c: str) -> str:
    """全角数字1文字を半角に変換する。全角でなければそのまま返す。"""
    if len(c) != 1:
        return c
    code = ord(c)
    # 全角数字（０～９）: 0xFF10 ～ 0xFF19
    if 0xFF10 <= code <= 0xFF19:
        return chr(code - 0xFF10 + ord("0"))
    return c


# ============================================
# VBA: IsNumericChar (L6455)
# 半角・全角数字判定 → Python str.isdigit() + 全角チェック
# ============================================
def is_numeric_char(c: str) -> bool:
    """半角・全角数字かどうかを判定する。

    VBA版: (c >= "0" And c <= "9") Or (c >= "０" And c <= "９")
    """
    if len(c) != 1:
        return False
    code = ord(c)
    return (ord("0") <= code <= ord("9")) or (0xFF10 <= code <= 0xFF19)


# ============================================
# VBA: ToCircledNumber (L6631)
# 数字→丸数字変換（①②③…⑩）
# ============================================
_CIRCLED_NUMBERS = ["", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]


def to_circled_number(num: int) -> str:
    """数字を丸数字に変換する。1〜10は丸数字、それ以外は数字文字列を返す。

    VBA版: circled = Array("", "①", "②", ..., "⑩")
    """
    if 1 <= num <= 10:
        return _CIRCLED_NUMBERS[num]
    return str(num)


# ============================================
# VBA: ExtractDateFromString (L5470)
# 文字列から「M月D日」形式の日付を抽出
# ============================================
def extract_date_from_string(text: str) -> datetime.date | None:
    """「○月○日」形式の文字列からdateを抽出する。

    VBA版と同じロジック:
    1. 「月」「日」の位置を探す
    2. 「月」の前の数字列 → 月、「月」と「日」の間の数字列 → 日
    3. 今年の日付を生成。180日以上過去なら翌年にする

    Returns:
        datetime.date or None（抽出失敗時）
    """
    if not text:
        return None

    match = re.search(r"(\d{1,2})月(\d{1,2})日", text)
    if not match:
        return None

    month_num = int(match.group(1))
    day_num = int(match.group(2))

    if not (1 <= month_num <= 12 and 1 <= day_num <= 31):
        return None

    today = datetime.date.today()
    try:
        result_date = datetime.date(today.year, month_num, day_num)
    except ValueError:
        return None

    # 180日以上過去なら翌年の日付とみなす
    if result_date < today and (today - result_date).days > 180:
        try:
            result_date = datetime.date(today.year + 1, month_num, day_num)
        except ValueError:
            return None

    return result_date


# ============================================
# VBA: IsFileOpen (L6359) → Python流で代替
# ============================================
def is_file_open(file_path: str | Path) -> bool:
    """ファイルが他のプロセスで使用中かどうか判定する。

    VBAではFileをOpen/Closeして排他ロックエラーで判定していた。
    Pythonでは書き込みモードでの排他オープンを試みる。
    """
    path = Path(file_path)
    if not path.exists():
        return False

    try:
        # 排他書き込みモードでオープンを試みる
        with open(path, "r+b"):
            pass
        return False
    except (IOError, PermissionError):
        return True


# ============================================
# フォルダ名生成（VBA: 納期回答書作成内のフォルダ作成ロジック）
# ============================================
def get_output_folder(
    tool_folder: str,
    execution_time: datetime.datetime,
) -> Path:
    """出力先フォルダを生成する。

    パス: ツールフォルダ/納期回答書/M月D日(曜日)_①回目/
    同日に複数回実行すると②回目、③回目…とフォルダが増える。
    """
    weekday_names = ["月", "火", "水", "木", "金", "土", "日"]
    weekday_str = weekday_names[execution_time.weekday()]
    date_part = f"{execution_time.month}月{execution_time.day}日({weekday_str})"

    base_dir = Path(tool_folder) / "納期回答書"
    base_dir.mkdir(parents=True, exist_ok=True)

    # 既存フォルダをカウントして回数を決定
    count = 1
    for entry in base_dir.iterdir():
        if entry.is_dir() and entry.name.startswith(date_part):
            count += 1

    folder_name = f"{date_part}_{to_circled_number(count)}回目"
    output_dir = base_dir / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)

    return output_dir


# ============================================
# ファイル名生成
# ============================================
def build_report_filename(
    customer_name: str,
    execution_time: datetime.datetime,
    rep_name: str = "",
    order_numbers: list[str] | None = None,
) -> str:
    """納期回答書のファイル名を生成する。

    通常: 納期回答書_顧客名様_yyyymmdd.xlsx
    担当者分割: 納期回答書_顧客名様_担当者名様_yyyymmdd.xlsx
    注番指定: 納期回答書_顧客名様_注番_yyyymmdd.xlsx（複数なら「複数注番」）
    """
    date_str = execution_time.strftime("%Y%m%d")

    # ファイル名に使えない文字を置換
    safe_customer = _sanitize_filename(customer_name)

    if order_numbers:
        if len(order_numbers) == 1:
            order_part = order_numbers[0]
        else:
            order_part = "複数注番"
        return f"納期回答書_{safe_customer}様_{order_part}_{date_str}.xlsx"

    if rep_name:
        safe_rep = _sanitize_filename(rep_name)
        return f"納期回答書_{safe_customer}様_{safe_rep}様_{date_str}.xlsx"

    return f"納期回答書_{safe_customer}様_{date_str}.xlsx"


def build_sheet_name(customer_name: str, rep_name: str = "") -> str:
    """シート名を生成する（31文字制限）。

    通常: 顧客名様
    担当者分割: 顧客名_担当者名様
    """
    if rep_name:
        name = f"{customer_name}_{rep_name}様"
    else:
        name = f"{customer_name}様"

    # Excelのシート名は31文字制限
    if len(name) > 31:
        name = name[:31]

    return name


def _sanitize_filename(name: str) -> str:
    """ファイル名に使えない文字を置換する。"""
    # Windows禁止文字: \ / : * ? " < > |
    return re.sub(r'[\\/:*?"<>|]', "_", name)


# ============================================
# 日付パース（VBA CDate相当の柔軟なパーサー）
# ============================================
_DATE_FORMATS = [
    "%Y/%m/%d",
    "%Y-%m-%d",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%m/%d/%Y",
]


def parse_date(
    value: object,
    *,
    today: datetime.date | None = None,
) -> datetime.date | None:
    """様々な形式の日付値をdatetime.dateに変換する。

    VBAのCDateは非常に柔軟だが、Pythonでは明示的にパースする必要がある。
    対応形式: datetime, date, str("YYYY/M/D", "YYYY-M-D", "M/D"等)

    月/日のみの場合（例: "1/5", "3/10"）は今年を補完し、
    補完した日付が今日より過去なら翌年にする。
    確認中一覧のJ列（手入力）で年なし入力に対応するための仕様。

    Args:
        value: パース対象の値
        today: 年補完の基準日（省略時は実行日）
    """
    if value is None:
        return None

    if isinstance(value, datetime.datetime):
        return value.date()

    if isinstance(value, datetime.date):
        return value

    if isinstance(value, (int, float)):
        # 0やNaNはNone
        if value == 0 or value != value:  # NaN check
            return None
        # Excelシリアル値（1〜2958465）→ 日付変換
        # Excel基準日: 1899/12/30 + days
        int_val = int(value)
        if 1 <= int_val <= 2958465:
            return datetime.date(1899, 12, 30) + datetime.timedelta(days=int_val)
        return None

    text = str(value).strip()
    if not text:
        return None

    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # "2026/1/5" のような0埋めなしの形式
    match = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", text)
    if match:
        try:
            return datetime.date(
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3)),
            )
        except ValueError:
            pass

    # "1/5", "3/10" のような月/日のみの形式（年を補完）
    match = re.match(r"(\d{1,2})/(\d{1,2})$", text)
    if match:
        base = today or datetime.date.today()
        month = int(match.group(1))
        day = int(match.group(2))
        try:
            result = datetime.date(base.year, month, day)
        except ValueError:
            return None
        if result < base:
            result = result.replace(year=base.year + 1)
        return result

    return None


def parse_time(value: object) -> tuple[int, int] | None:
    """時刻文字列をパースして(時, 分)のタプルを返す。

    VBAでは"8:54:58"形式の文字列をSplit(":")で分割。
    """
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    match = re.match(r"(\d{1,2}):(\d{1,2})", text)
    if match:
        return int(match.group(1)), int(match.group(2))

    return None


def is_december_31(d: datetime.date | None) -> bool:
    """12月31日かどうか判定する。

    VBAではMonth(date) = 12 And Day(date) = 31で判定。
    12/31はSAPの「未定」デフォルト値。
    """
    if d is None:
        return False
    return d.month == 12 and d.day == 31


def format_date_japanese(d: datetime.date) -> str:
    """日付を「M月D日」形式にフォーマットする。

    VBA: Format(d, "m月d日")
    """
    return f"{d.month}月{d.day}日"


# 全角→半角変換テーブル（受注先=出荷先の比較用）
_NORMALIZE_TABLE = str.maketrans(
    "（）「」｛｝＜＞"
    "０１２３４５６７８９"
    "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
    "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
    "－　",
    "()「」{}<>"
    "0123456789"
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    "abcdefghijklmnopqrstuvwxyz"
    "- ",
)


def normalize_name_for_comparison(name: str) -> str:
    """受注先/出荷先の名前を正規化する（比較用）。

    SAPでは同じ顧客でも受注先と出荷先で全角/半角の揺れがある。
    例: 受注先「（有）三橋機工」vs 出荷先「(有)三橋機工」
    """
    return name.strip().translate(_NORMALIZE_TABLE)


def format_quantity(qty: str) -> str:
    """数量文字列から末尾の不要なゼロを除去する。

    SAPの数量は "1.00" のような小数表記で格納されているが、
    表示上は整数なら "1"、小数なら "2.5" のように簡潔にする。

    例: "1.00" → "1", "2.50" → "2.5", "0.50" → "0.5", "3" → "3"
    """
    s = str(qty).strip()
    if not s:
        return s
    try:
        # 数値としてパースし、不要なゼロを除去
        # Decimal で正確に処理（float の丸め誤差を回避）
        from decimal import Decimal, InvalidOperation
        d = Decimal(s)
        # normalize() で "1.00" → "1", "2.50" → "2.5" に変換
        normalized = d.normalize()
        # 指数表記になる場合（例: 1E+2）は固定小数点に戻す
        return f"{normalized:f}" if normalized.as_tuple().exponent > 0 else str(normalized)
    except (InvalidOperation, ValueError):
        # 数値でなければそのまま返す
        return s


def normalize_item_group_code(value: object) -> str:
    """品目GroupCodeを正規化する。

    openpyxlがExcelの数値セルをintで返す場合（75→"75"）と、
    SAPテキストが先頭ゼロ付き文字列で来る場合（"0075"）の
    不一致を解消するための正規化。

    ルール:
    - int型 or 数字のみ文字列 → 4桁ゼロ埋め ("75"→"0075", 75→"0075")
    - 英字を含む → そのまま strip のみ ("A01"→"A01")
    - None / 空文字 → 空文字
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    if s.isdigit():
        return s.zfill(4)
    return s


# ============================================
# 4月SAP切替対応の特別注意文（2026-04-24まで。以降この関数と呼出箇所を削除）
# ============================================
_SPECIAL_NOTICE_EXCEL = (
    "※4月のシステム切替の影響により、納期回答が遅れましたことを"
    "お詫び申し上げます。既にお届け済みの商品につきましても"
    "納期回答が届く場合がございます。何卒ご了承ください。"
)

_SPECIAL_NOTICE_EMAIL = (
    "なお、4月のシステム切替の影響により納期回答が"
    "遅れましたことをお詫び申し上げます。"
    "既にお届け済みの商品につきましても納期回答が届く場合が"
    "ございます。何卒ご了承ください。"
)

_SPECIAL_NOTICE_END_DATE = datetime.date(2026, 4, 24)


def get_special_notice_excel(today: datetime.date | None = None) -> str | None:
    """Excel「ご連絡事項」欄用の特別注意文。期間外ならNone。"""
    if today is None:
        today = datetime.date.today()
    if today <= _SPECIAL_NOTICE_END_DATE:
        return _SPECIAL_NOTICE_EXCEL
    return None


def get_special_notice_email(today: datetime.date | None = None) -> str | None:
    """メール本文用の特別注意文。期間外ならNone。"""
    if today is None:
        today = datetime.date.today()
    if today <= _SPECIAL_NOTICE_END_DATE:
        return _SPECIAL_NOTICE_EMAIL
    return None
