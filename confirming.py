"""確認中一覧参照モジュール

VBAの以下の関数を移植:
- GetConfirmedDeliveryDate (L5035): 確認中一覧から確定納期取得
- GetConfirmingStatus (L5085): 確認中一覧からステータス取得
"""

from __future__ import annotations

import datetime
from typing import Optional

from nouki_kaitou.models import CacheStore


# ============================================
# VBA: GetConfirmedDeliveryDate (L5035-5080)
# 確認中一覧から確定納期（受注納期列）を取得
# ============================================
def get_confirmed_delivery_date(
    order_number: str,
    detail_number: str,
    cache: CacheStore,
) -> Optional[datetime.date]:
    """確認中一覧から確定納期を取得する。

    確認中一覧のJ列（受注納期）に手入力された日付を取得。
    分納の確定日情報としてCalculateBunnoDateで使用される。

    Args:
        order_number: 受発注伝票
        detail_number: 明細
        cache: キャッシュストア

    Returns:
        確定納期（日付入力あり）、またはNone
    """
    if not order_number or not detail_number:
        return None

    key = f"{order_number}|{detail_number}"
    entry = cache.confirm.get(key)
    if entry is None:
        return None

    # entry = (問合せ状況, ステータス, 受注納期date)
    return entry[2]


# ============================================
# VBA: GetConfirmingStatus (L5085-5127)
# 確認中一覧からステータスを取得
# ============================================
def get_confirming_status(
    order_number: str,
    detail_number: str,
    cache: CacheStore,
) -> str:
    """確認中一覧からステータスを取得する。

    確認中一覧のI列（ステータス）の値を返す。
    値の例: "分納", "欠品中", "未処理" 等

    Args:
        order_number: 受発注伝票
        detail_number: 明細
        cache: キャッシュストア

    Returns:
        ステータス文字列（見つからなければ空文字）
    """
    if not order_number or not detail_number:
        return ""

    key = f"{order_number}|{detail_number}"
    entry = cache.confirm.get(key)
    if entry is None:
        return ""

    # entry = (問合せ状況, ステータス, 受注納期date)
    return entry[1]
