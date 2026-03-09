"""run_gui.py のユニットテスト

コンソール待機処理のテスト。
"""

from unittest.mock import patch, MagicMock
import sys


class TestMainConsoleWait:
    """main()がEnterキー待機するかのテスト"""

    @patch("run_gui.select_source_file", return_value=None)
    @patch("builtins.input", return_value="")
    def test_cancel_waits_for_input(self, mock_input, mock_select):
        """キャンセル時にinput()が呼ばれる"""
        from run_gui import main

        try:
            main()
        except SystemExit:
            pass

        # input()が呼ばれていること（Enterキーで終了...）
        mock_input.assert_called()
        calls = [str(c) for c in mock_input.call_args_list]
        assert any("Enterキーで終了" in c for c in calls)

    @patch("run_gui.app_main")
    @patch("run_gui.select_source_file", return_value="/tmp/test.xls")
    @patch("builtins.input", return_value="")
    def test_normal_completion_waits_for_input(self, mock_input, mock_select, mock_app):
        """正常終了時にinput()が呼ばれる"""
        from run_gui import main

        main()

        mock_input.assert_called()
        calls = [str(c) for c in mock_input.call_args_list]
        assert any("Enterキーで終了" in c for c in calls)

    @patch("run_gui.app_main", side_effect=RuntimeError("テストエラー"))
    @patch("run_gui.select_source_file", return_value="/tmp/test.xls")
    @patch("builtins.input", return_value="")
    def test_error_waits_for_input(self, mock_input, mock_select, mock_app):
        """エラー時にinput()が呼ばれる"""
        from run_gui import main

        try:
            main()
        except SystemExit:
            pass

        mock_input.assert_called()
        calls = [str(c) for c in mock_input.call_args_list]
        assert any("Enterキーで終了" in c for c in calls)
