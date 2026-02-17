"""基本フィルタによる除外テスト

【受注】得意先返品、【受注】売上値引、【受注】売上値増などの
想定外の伝票タイプが納期回答書生成時に正しく除外されることを確認する。
"""

import datetime
import pytest

from nouki_kaitou.models import OrderRow, CacheStore, BranchSettings
from nouki_kaitou.report_generator import _pass_basic_filter, build_report_row


class TestBasicFilter:
    """_pass_basic_filter() のテスト"""

    def test_stock_sales_passes(self):
        """【受注】在庫販売は通過する"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="【受注】在庫販売",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is True

    def test_direct_sales_passes(self):
        """【受注】直送販売は通過する"""
        row = OrderRow(
            order_number="GL2F446767",
            detail_number="10",
            document_type="【受注】直送販売",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is True

    def test_customer_return_excluded(self):
        """【受注】得意先返品は除外される"""
        row = OrderRow(
            order_number="366068",
            detail_number="10",
            document_type="【受注】得意先返品",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is False

    def test_sales_discount_excluded(self):
        """【受注】売上値引は除外される"""
        row = OrderRow(
            order_number="365930",
            detail_number="10",
            document_type="【受注】売上値引",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is False

    def test_sales_increase_excluded(self):
        """【受注】売上値増は除外される"""
        row = OrderRow(
            order_number="366070",
            detail_number="10",
            document_type="【受注】売上値増",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is False

    def test_detail_deleted_excluded(self):
        """拒否理由が「明細削除」は除外される"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="【受注】在庫販売",
            rejection_reason="明細削除",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is False

    def test_hash_hash_excluded(self):
        """コメント（社内）に##がある場合は除外される"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="【受注】在庫販売",
            rejection_reason="",
            comment_internal="##除外",
        )
        assert _pass_basic_filter(row) is False

    def test_fullwidth_hash_hash_excluded(self):
        """コメント（社内）に全角＃＃がある場合も除外される"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="【受注】在庫販売",
            rejection_reason="",
            comment_internal="＃＃除外",
        )
        assert _pass_basic_filter(row) is False

    def test_empty_document_type_excluded(self):
        """伝票タイプが空は除外される"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is False


class TestExclusionWithRealData:
    """実データの注番を使った除外テスト"""

    @pytest.fixture
    def sample_excluded_orders(self) -> list[OrderRow]:
        """validation_result.xlsxの「その他」カテゴリの注番"""
        excluded_types = [
            ("366068", "【受注】得意先返品"),
            ("366069", "【受注】得意先返品"),
            ("365930", "【受注】売上値引"),
            ("365931", "【受注】売上値引"),
            ("366070", "【受注】売上値増"),
            ("366568", "【受注】売上値増"),
        ]
        return [
            OrderRow(
                order_number=order_num,
                detail_number="10",
                document_type=doc_type,
                rejection_reason="",
                comment_internal="",
            )
            for order_num, doc_type in excluded_types
        ]

    def test_all_excluded_orders_are_filtered(self, sample_excluded_orders):
        """全ての除外対象注番がフィルタで除外される"""
        for row in sample_excluded_orders:
            assert _pass_basic_filter(row) is False, (
                f"注番 {row.order_number} (伝票タイプ: {row.document_type}) が除外されていない"
            )

    def test_excluded_count(self, sample_excluded_orders):
        """除外件数の確認"""
        excluded_count = sum(
            1 for row in sample_excluded_orders
            if not _pass_basic_filter(row)
        )
        assert excluded_count == len(sample_excluded_orders)


class TestFilterIntegration:
    """フィルタの統合テスト"""

    def test_multiple_conditions(self):
        """複数の除外条件が同時に適用される"""
        # 伝票タイプが返品で、かつ##除外がある
        row = OrderRow(
            order_number="366068",
            detail_number="10",
            document_type="【受注】得意先返品",
            rejection_reason="",
            comment_internal="##除外",
        )
        # どちらか一方でも除外される
        assert _pass_basic_filter(row) is False

    def test_valid_order_with_hash_comment(self):
        """有効な伝票タイプでも##があれば除外"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="【受注】在庫販売",
            rejection_reason="",
            comment_internal="ｆ　##テスト除外",
        )
        assert _pass_basic_filter(row) is False

    def test_whitespace_handling(self):
        """伝票タイプの前後空白は除去される"""
        row = OrderRow(
            order_number="GL2C447163",
            detail_number="10",
            document_type="  【受注】在庫販売  ",
            rejection_reason="",
            comment_internal="",
        )
        assert _pass_basic_filter(row) is True
