"""メーカー・品目Group参照モジュール

VBAの以下の関数を移植:
- GetManufacturerName (L4965): 品目GroupCodeからメーカー名取得
- GetDeliveryDaysToAdd (L4996): 品目GroupCodeから配送加算日数取得
"""

from __future__ import annotations

from nouki_kaitou.models import CacheStore
from nouki_kaitou.utils import normalize_item_group_code


# ============================================
# VBA: GetManufacturerName (L4965-4994)
# 品目GroupCodeからメーカー名を取得
# ============================================
def get_manufacturer_name(
    item_group_code: str,
    cache: CacheStore,
) -> str:
    """品目GroupCodeからメーカー名を取得する。

    Args:
        item_group_code: 品目GroupCode（例: "D01", "0075"）
        cache: キャッシュストア

    Returns:
        メーカー名（見つからなければ空文字）
    """
    code = normalize_item_group_code(item_group_code)
    if not code:
        return ""

    return cache.mfg_name.get(code, "")


# ============================================
# VBA: GetDeliveryDaysToAdd (L4996-5029)
# 品目GroupCodeから配送加算日数を取得
# ============================================
def get_delivery_days_to_add(
    item_group_code: str,
    cache: CacheStore,
) -> int:
    """品目GroupCodeから配送加算日数を取得する。

    Args:
        item_group_code: 品目GroupCode
        cache: キャッシュストア

    Returns:
        加算日数（デフォルト2）
    """
    code = normalize_item_group_code(item_group_code)
    if not code:
        return 2

    return cache.mfg_days.get(code, 2)
