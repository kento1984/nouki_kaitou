"""分納処理モジュール

VBAの以下の関数を移植:
- ExtractBunnoInfo (L6646): 分納情報抽出（コメント解析）
- CalculateBunnoDate (L6982): 分納日付計算
- HasBunnoMitei (L6914): 分納に未定があるか判定
- HasBunnoKakuninchu (L6947): 計算済み納期に「確認中」があるか判定
- NormalizeBunnoDate (L6856): 分納日付文字列の正規化
- ExtractLocationFromToken (L6779): 括弧内の場所抽出
- SplitQtyAndDate (L7190): 数量と日付の分離
- RemoveBunnoText (L4005): 備考から分納テキスト除去
- StartsWithNumber (L6812): 数字始まり判定
- IsDateToken (L6825): 日付トークン判定
"""

from __future__ import annotations

import datetime
import re
from typing import Optional

from nouki_kaitou.business_days import add_business_days
from nouki_kaitou.confirming import get_confirmed_delivery_date
from nouki_kaitou.models import BunnoEntry, CacheStore, HolidayMap
from nouki_kaitou.utils import convert_to_half_width, format_date_japanese


# ============================================
# VBA: StartsWithNumber (L6812-6820)
# 数字で始まるかチェック
# ============================================
def starts_with_number(text: str) -> bool:
    """文字列が数字で始まるかチェックする。

    VBA版: IsNumeric(Left(text, 1))
    """
    if not text:
        return False
    return text[0].isdigit()


# ============================================
# VBA: IsDateToken (L6825-6851)
# 日付トークンかどうか判定
# ============================================
def is_date_token(token: str) -> bool:
    """トークンが日付（または未定扱い）かどうか判定する。

    以下のパターンをTrue判定:
    - 「未定」
    - 「欠品」「確認中」を含む → 未定扱い
    - スラッシュ含む（12/20形式）
    - 「○月○旬予定」形式
    """
    if not token:
        return False

    # 未定
    if token == "未定":
        return True

    # 欠品中・確認中系は未定扱い
    if "欠品" in token or "確認中" in token:
        return True

    # スラッシュ含む（半角・全角）
    if "/" in token or "／" in token:
        return True

    # ○月○旬予定 形式
    if "月" in token and "予定" in token:
        return True

    return False


# ============================================
# VBA: NormalizeBunnoDate (L6856-6878)
# 分納の日付文字列を正規化
# ============================================
def normalize_bunno_date(date_str: str) -> str:
    """分納の日付文字列を正規化する。

    「12/22出荷」→「12/22」、「欠品中納期確認中」→「未定」

    Args:
        date_str: 日付文字列

    Returns:
        正規化された日付文字列
    """
    result = date_str.strip()

    # 欠品中・確認中系は未定扱い
    if "欠品" in result or "確認中" in result:
        return "未定"

    # 「出荷」「着」等の余分な文字を除去
    result = result.replace("出荷", "")
    result = result.replace("着", "")
    result = result.strip()

    # 「予定」は○旬以外では除去
    if "上旬" not in result and "中旬" not in result and "下旬" not in result:
        result = result.replace("予定", "")
        result = result.strip()

    return result


# ============================================
# VBA: ExtractLocationFromToken (L6779-6807)
# トークンから括弧内の場所を抽出
# ============================================
def extract_location_from_token(token: str) -> tuple[str, str]:
    """トークンから括弧内の場所を抽出し、括弧部分を除去する。

    VBA版はByRefでtokenを変更していた。
    Python版は (修正後token, 場所) のタプルを返す。

    Args:
        token: 入力トークン（例: "700m（滋賀）"）

    Returns:
        (括弧除去後のtoken, 場所文字列)
        場所がなければ (token, "")
    """
    # 全角括弧
    match = re.search(r"（(.+?)）", token)
    if match:
        location = match.group(1)
        modified_token = token[:match.start()]
        return modified_token, location

    # 半角括弧
    match = re.search(r"\((.+?)\)", token)
    if match:
        location = match.group(1)
        modified_token = token[:match.start()]
        return modified_token, location

    return token, ""


# ============================================
# VBA: SplitQtyAndDate (L7190-7218)
# 数量と日付が連結されている場合に分割
# ============================================
_QTY_UNITS = [
    "個", "本", "台", "枚", "セット", "缶", "箱", "袋", "巻", "丁", "組",
    "kg", "ｋｇ", "KG", "m", "ｍ", "M",
]


def split_qty_and_date(token: str) -> tuple[str, str] | None:
    """数量と日付が連結されている場合に分割する。

    例: 「1個12/19」→ ("1個", "12/19")
    分割不要なら None を返す。

    Args:
        token: 入力トークン

    Returns:
        (数量, 日付) のタプル。分割不要ならNone。
    """
    for unit in _QTY_UNITS:
        pos = token.find(unit)
        if pos >= 0:
            unit_end = pos + len(unit)
            after_unit = token[unit_end:]

            # 単位の後に何かあれば分割
            if after_unit:
                # 日付っぽいかチェック（数字で始まるか「未定」）
                if starts_with_number(after_unit) or after_unit == "未定":
                    return token[:unit_end], after_unit

    # 分割不要
    return None


# ============================================
# VBA: ExtractBunnoInfo (L6646-6775)
# 分納情報をコメントから抽出
# ============================================
def extract_bunno_info(comment: str) -> list[BunnoEntry]:
    """コメントから分納情報を抽出する。

    「分納:」または「分納：」の後ろをパースして、
    各分納パートの(数量, 日付, 場所)を返す。

    入力例: 「分納:700m 12/17 滋賀、300m 12/17 東京」
    出力: [BunnoEntry("700m", "12/17", "滋賀"), BunnoEntry("300m", "12/17", "東京")]

    Args:
        comment: コメント文字列

    Returns:
        BunnoEntryのリスト
    """
    result: list[BunnoEntry] = []

    if not comment:
        return result

    # 「分納:」または「分納：」を探す
    start_pos = comment.find("分納:")
    if start_pos < 0:
        start_pos = comment.find("分納：")
    if start_pos < 0:
        return result

    # 分納:以降を取得（「分納:」は3文字）
    bunno_text = comment[start_pos + 3:]

    # 全角数字を半角に変換
    bunno_text = convert_to_half_width(bunno_text)

    # 他のコメントがある場合は分納部分だけ取得（2スペースか改行で区切り）
    for sep in ["  ", "\n"]:
        end_pos = bunno_text.find(sep)
        if end_pos >= 0:
            bunno_text = bunno_text[:end_pos]

    # 「、」をカンマに統一
    bunno_text = bunno_text.replace("、", ",")

    # カンマで分割
    parts = bunno_text.split(",")

    for part in parts:
        part = part.strip()
        if not part:
            continue

        # 全角スペースを半角に
        part = part.replace("\u3000", " ")

        # スペースで分割してトークン化
        tokens = part.split(" ")

        qty = ""
        date_str = ""
        location_str = ""

        for raw_token in tokens:
            token = raw_token.strip()
            if not token:
                continue

            # 括弧内の場所を先に抽出（トークンから括弧部分を除去）
            token, bracket_location = extract_location_from_token(token)
            if bracket_location and not location_str:
                location_str = bracket_location

            # 数量判定（数字で始まる＋単位）
            if not qty and starts_with_number(token):
                # 単位の後に日付が連結されているかチェック
                split_result = split_qty_and_date(token)
                if split_result is not None:
                    qty = split_result[0]
                    if not date_str and split_result[1]:
                        date_str = split_result[1]
                else:
                    qty = token
            # 日付判定
            elif not date_str and is_date_token(token):
                date_str = token
            # それ以外は場所
            elif not location_str and token:
                location_str = token

        # 日付がなければ未定
        if not date_str:
            date_str = "未定"

        # 日付を正規化
        date_str = normalize_bunno_date(date_str)

        # 数値チェック用に単位を除去
        qty_num = qty
        for unit in _QTY_UNITS:
            qty_num = qty_num.replace(unit, "")

        # 数値として妥当かチェック
        if qty_num:
            try:
                float(qty_num)
                result.append(BunnoEntry(
                    quantity=qty,
                    date_str=date_str,
                    location=location_str,
                ))
            except ValueError:
                pass

    return result


# ============================================
# VBA: RemoveBunnoText (L4005-4033)
# 備考から分納テキストを除去
# ============================================
def remove_bunno_text(text: str) -> str:
    """備考テキストから分納テキストを除去する。

    「分納:～」の部分を除去。終端は2スペースか改行。

    Args:
        text: 元テキスト

    Returns:
        分納テキスト除去後の文字列
    """
    if not text:
        return ""

    # 「分納:」または「分納：」を探す
    start_pos = text.find("分納:")
    if start_pos < 0:
        start_pos = text.find("分納：")
    if start_pos < 0:
        return text

    # 分納部分の終了位置を探す
    after_bunno = text[start_pos:]

    end_pos = after_bunno.find("  ")
    if end_pos < 0:
        end_pos = after_bunno.find("\n")

    if end_pos >= 0:
        # 分納部分だけ除去
        bunno_text = after_bunno[:end_pos]
    else:
        # 全部除去
        bunno_text = after_bunno

    return text.replace(bunno_text, "")


# ============================================
# VBA: HasBunnoMitei (L6914-6942)
# 分納に未定があるかチェック
# ============================================
def has_bunno_mitei(
    bunno_info: list[BunnoEntry],
    cache: Optional[CacheStore] = None,
    order_number: str = "",
    detail_number: str = "",
) -> bool:
    """分納に未定（未確定）があるかチェックする。

    日付が「未定」「欠品」「確認中」「予定」を含む場合、
    確認中一覧に確定日があればスキップ、なければTrue。

    Args:
        bunno_info: 分納情報リスト
        cache: キャッシュストア（確認中一覧参照用）
        order_number: 受発注伝票
        detail_number: 明細

    Returns:
        True = 未定の分納がある
    """
    if not bunno_info:
        return False

    # 確認中一覧から確定日を取得
    confirmed_date: Optional[datetime.date] = None
    if cache and order_number and detail_number:
        confirmed_date = get_confirmed_delivery_date(
            order_number, detail_number, cache
        )

    for entry in bunno_info:
        ds = entry.date_str
        if ds == "未定" or "欠品" in ds or "確認中" in ds or "予定" in ds:
            if confirmed_date is not None:
                # 確定済み → 次のentryへ
                continue
            else:
                return True

    return False


# ============================================
# VBA: HasBunnoKakuninchu (L6947-6978)
# 計算済み納期に「確認中」があるかチェック
# ============================================
def has_bunno_kakuninchu(bunno_detail: list[list[str]]) -> bool:
    """計算済み分納詳細に「確認中」があるかチェックする。

    FormatReport内で計算された分納の各行を調べ、
    「確認中」や「○旬予定」（出荷予定・配達予定を除く）が
    あれば未確定として True を返す。

    Args:
        bunno_detail: 計算済み分納詳細リスト。
                      各要素は [qty, date_str, location, calc_date] 形式。

    Returns:
        True = 「確認中」等の未確定がある
    """
    if not bunno_detail:
        return False

    for item in bunno_detail:
        # 計算済み納期（4番目の要素）を見る
        calc_date = ""
        if len(item) >= 4:
            calc_date = str(item[3])

        # 「確認中」があれば未確定
        if calc_date == "確認中":
            return True

        # 「○旬予定」があれば未確定（「出荷予定」「配達予定」は除く）
        if "予定" in calc_date:
            if "出荷" not in calc_date and "配達" not in calc_date:
                return True

    return False


# ============================================
# VBA: CalculateBunnoDate (L6982-7123)
# 分納の日付を納期表示用に変換
# ============================================
def calculate_bunno_date(
    date_str: str,
    is_ship_rule: bool,
    days_to_add: int,
    holidays: HolidayMap | None = None,
    cache: Optional[CacheStore] = None,
    order_number: str = "",
    detail_number: str = "",
    is_rosenbin: bool = False,
    today: Optional[datetime.date] = None,
) -> str:
    """分納の日付を納期表示用に計算する。

    ルール:
    - 未定 → 確認中一覧に確定日があれば計算、なければ「確認中」
    - 予定（○旬予定等） → 確認中一覧チェック後、そのまま返す
    - M/D形式 → 配送日計算して「○月○日出荷/配達 予定/済み」

    isShipRule=True: 直送 → 「出荷予定/済み」
    isShipRule=False, isRosenbin=False: 自社便 → 「配達予定/済み」
    isShipRule=False, isRosenbin=True: 路線便 → 「出荷予定/済み」

    Args:
        date_str: 分納の日付文字列
        is_ship_rule: 直送ルール（出荷日表示）
        days_to_add: 加算営業日数
        holidays: 祝日辞書
        cache: キャッシュストア
        order_number: 受発注伝票
        detail_number: 明細
        is_rosenbin: 路線便フラグ
        today: 基準日（テスト用。Noneなら今日）

    Returns:
        表示用の納期文字列
    """
    if today is None:
        today = datetime.date.today()

    # --- 未定の場合 ---
    if date_str == "未定":
        # 確認中一覧の受注納期をチェック
        if cache and order_number and detail_number:
            confirmed_date = get_confirmed_delivery_date(
                order_number, detail_number, cache
            )
            if confirmed_date is not None:
                return _format_bunno_with_confirmed(
                    confirmed_date, is_ship_rule, days_to_add,
                    holidays, is_rosenbin, today,
                )

        return "確認中"

    # --- 「予定」を含む（○旬予定等） ---
    if "予定" in date_str:
        # まず確認中一覧をチェック
        if cache and order_number and detail_number:
            confirmed_date = get_confirmed_delivery_date(
                order_number, detail_number, cache
            )
            if confirmed_date is not None:
                return _format_bunno_with_confirmed(
                    confirmed_date, is_ship_rule, days_to_add,
                    holidays, is_rosenbin, today,
                )

        return date_str

    # --- M/D形式の日付 ---
    date_str = date_str.replace("／", "/")
    slash_pos = date_str.find("/")

    if slash_pos < 0:
        return "確認中"

    try:
        month_num = int(date_str[:slash_pos])
        day_num = int(date_str[slash_pos + 1:])
    except ValueError:
        return "確認中"

    if not (1 <= month_num <= 12 and 1 <= day_num <= 31):
        return "確認中"

    try:
        bunno_date = datetime.date(today.year, month_num, day_num)
    except ValueError:
        return "確認中"

    # 180日以上過去なら翌年
    if bunno_date < today and (today - bunno_date).days > 180:
        try:
            bunno_date = datetime.date(today.year + 1, month_num, day_num)
        except ValueError:
            return "確認中"

    return _format_bunno_date(
        bunno_date, is_ship_rule, days_to_add,
        holidays, is_rosenbin, today,
    )


def _format_bunno_with_confirmed(
    confirmed_date: datetime.date,
    is_ship_rule: bool,
    days_to_add: int,
    holidays: HolidayMap | None,
    is_rosenbin: bool,
    today: datetime.date,
) -> str:
    """確認中一覧の確定日から分納の表示文字列を生成する。"""
    if is_ship_rule:
        # 直送 → 確定日そのまま出荷日
        suffix = "出荷済み" if confirmed_date <= today else "出荷予定"
        return f"{format_date_japanese(confirmed_date)}{suffix}"
    elif is_rosenbin:
        # 路線便 → 確定日 + (daysToAdd - 1) 営業日
        delivery_date = add_business_days(
            confirmed_date, max(days_to_add - 1, 0), holidays
        )
        suffix = "出荷済み" if delivery_date <= today else "出荷予定"
        return f"{format_date_japanese(delivery_date)}{suffix}"
    else:
        # 自社便 → 確定日 + daysToAdd 営業日
        delivery_date = add_business_days(
            confirmed_date, days_to_add, holidays
        )
        suffix = "配達済み" if delivery_date <= today else "配達予定"
        return f"{format_date_japanese(delivery_date)}{suffix}"


def _format_bunno_date(
    bunno_date: datetime.date,
    is_ship_rule: bool,
    days_to_add: int,
    holidays: HolidayMap | None,
    is_rosenbin: bool,
    today: datetime.date,
) -> str:
    """分納の日付から表示文字列を生成する。"""
    if is_ship_rule:
        # 直送 → 日付そのまま出荷日
        suffix = "出荷済み" if bunno_date <= today else "出荷予定"
        return f"{format_date_japanese(bunno_date)}{suffix}"
    elif is_rosenbin:
        # 路線便 → 日付 + (daysToAdd - 1) 営業日
        delivery_date = add_business_days(
            bunno_date, max(days_to_add - 1, 0), holidays
        )
        suffix = "出荷済み" if delivery_date <= today else "出荷予定"
        return f"{format_date_japanese(delivery_date)}{suffix}"
    else:
        # 自社便 → 日付 + daysToAdd 営業日
        delivery_date = add_business_days(
            bunno_date, days_to_add, holidays
        )
        suffix = "配達済み" if delivery_date <= today else "配達予定"
        return f"{format_date_japanese(delivery_date)}{suffix}"
