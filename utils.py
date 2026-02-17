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


def parse_date(value: object) -> datetime.date | None:
    """様々な形式の日付値をdatetime.dateに変換する。

    VBAのCDateは非常に柔軟だが、Pythonでは明示的にパースする必要がある。
    対応形式: datetime, date, str("YYYY/M/D", "YYYY-M-D"等)
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
