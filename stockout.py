"""欠品処理モジュール

VBAの以下の関数を移植:
- ExtractApproxDelivery (L7224): 概算納期抽出（「○月上旬」「○/○頃」等）
- RemoveStockoutText (L7354): 欠品テキスト除去
- GetStoragePlaceFromSameOrder (L7418): 同一注番から保管場所取得
"""

from __future__ import annotations

import re

from nouki_kaitou.models import CacheStore
from nouki_kaitou.utils import convert_to_half_width


# ============================================
# VBA: ExtractApproxDelivery (L7224-7349)
# 「欠品中」の後ろからアバウト納期を抽出
# ============================================
def extract_approx_delivery(comment: str) -> str:
    """コメントから欠品の概算納期を抽出する。

    「欠品中」の後ろにある納期パターンを認識:
    - パターン1: 「○月上旬/中旬/下旬」→ 「○月上旬入荷予定」
    - パターン2: 「○/○頃」「○月○日頃」→ 「○月○日頃入荷予定」
    - パターン3: 「○月末」→ 「○月末入荷予定」

    Args:
        comment: コメント文字列（コメント（明細）等）

    Returns:
        概算納期文字列。抽出できなければ空文字。
    """
    if not comment:
        return ""

    # 「欠品中」の位置を探す
    stockout_pos = comment.find("欠品中")
    if stockout_pos < 0:
        return ""

    after = comment[stockout_pos + 3:].strip()
    if not after:
        return ""

    # 全角→半角変換
    after = convert_to_half_width(after)

    # パターン1: 「○月上旬/中旬/下旬」（初旬→上旬、半ば→中旬として扱う）
    # 範囲指定（～/〜/から）がある場合は後半を採用
    _PERIOD_NORMALIZE = {"初旬": "上旬", "半ば": "中旬"}
    range_match = re.search(
        r"(\d{1,2})月(上旬|中旬|下旬|初旬|半ば)[～〜](上旬|中旬|下旬|初旬|半ば)",
        after,
    )
    if not range_match:
        range_match = re.search(
            r"(\d{1,2})月(上旬|中旬|下旬|初旬|半ば)から(上旬|中旬|下旬|初旬|半ば)",
            after,
        )
    if range_match:
        month = int(range_match.group(1))
        period_from = range_match.group(2)
        period_to = range_match.group(3)
        period_from = _PERIOD_NORMALIZE.get(period_from, period_from)
        period_to = _PERIOD_NORMALIZE.get(period_to, period_to)
        return f"{month}月{period_from}～{period_to}入荷予定"

    match_junme = re.search(
        r"(\d{1,2})月(上旬|中旬|下旬|初旬|半ば)", after
    )
    if match_junme:
        month = int(match_junme.group(1))
        period = match_junme.group(2)
        period = _PERIOD_NORMALIZE.get(period, period)
        return f"{month}月{period}入荷予定"

    # パターン2: 「○/○頃」or「○月○日頃」
    if "頃" in after:
        # 「○/○頃」
        match_slash = re.search(r"(\d{1,2})/(\d{1,2})", after)
        if match_slash:
            month = int(match_slash.group(1))
            day = int(match_slash.group(2))
            if 1 <= month <= 12 and 1 <= day <= 31:
                return f"{month}月{day}日頃入荷予定"

        # 「○月○日頃」
        match_md = re.search(r"(\d{1,2})月(\d{1,2})日", after)
        if match_md:
            month = int(match_md.group(1))
            day = int(match_md.group(2))
            if 1 <= month <= 12 and 1 <= day <= 31:
                return f"{month}月{day}日頃入荷予定"

    # パターン3: 「○月末」
    match_matsu = re.search(r"(\d{1,2})月末", after)
    if match_matsu:
        month = int(match_matsu.group(1))
        return f"{month}月末入荷予定"

    return ""


# ============================================
# VBA: RemoveStockoutText (L7354-7414)
# 欠品テキストを除去（アバウト納期も含めて）
# ============================================
def remove_stockout_text(text: str) -> str:
    """コメントから欠品テキストを除去する。

    「欠品中」とその後ろのアバウト納期パターンを除去。
    例: 「欠品中 1月上旬予定」→ 全部除去

    Args:
        text: 元テキスト

    Returns:
        欠品テキスト除去後の文字列
    """
    if not text:
        return ""

    start = text.find("欠品中")
    if start < 0:
        return text

    # 「欠品中」の後ろを確認
    end = start + 3

    # スペースをスキップ
    while end < len(text) and text[end] in (" ", "\u3000"):
        end += 1

    # 納期パターンがあれば終端まで探す
    if end < len(text):
        check_text = text[end:]
        pattern_end = 0
        has_pattern = False

        if "月" in check_text or "/" in check_text:
            # 終端キーワードを探す
            for keyword in ["予定", "頃", "上旬", "中旬", "下旬", "初旬", "半ば", "月末"]:
                pos = check_text.find(keyword)
                if pos >= 0:
                    candidate = pos + len(keyword)
                    if candidate > pattern_end:
                        pattern_end = candidate
                    has_pattern = True

            if has_pattern:
                end = end + pattern_end

    result = text[:start] + text[end:]
    return result.strip()


# ============================================
# VBA: GetStoragePlaceFromSameOrder (L7418-7456)
# 同一注番から保管場所を取得
# ============================================
def get_storage_place_from_same_order(
    order_number: str,
    cache: CacheStore,
) -> str:
    """同一注番の別明細から保管場所を取得する。

    キャッシュから注番をキーに保管場所を取得。
    送料行など保管場所が空の明細のために使用。

    Args:
        order_number: 受発注伝票
        cache: キャッシュストア

    Returns:
        保管場所（見つからなければ空文字）
    """
    if not order_number:
        return ""

    return cache.storage.get(order_number, "")
