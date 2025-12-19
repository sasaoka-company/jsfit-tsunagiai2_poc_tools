
"""
Word文書マーカー抽出プログラムのテストコード

word_to_text.pyの各関数をテストする
"""

# 改行文字の定数化（テスト用）
LINE_BREAK = "\n"

import unittest
from unittest.mock import Mock, patch, MagicMock
import sys
from io import StringIO

# テスト対象モジュールをインポート
import word_to_text


class TestExtractionState(unittest.TestCase):
    """ExtractionStateクラスのテスト"""
    
    def test_initial_state(self):
        """初期状態の確認"""
        state = word_to_text.ExtractionState()
        self.assertFalse(state.in_parent)
        self.assertFalse(state.in_skip)
        self.assertEqual(state.found_parent_count, 0)
    
    def test_process_marker_parent(self):
        """[PARENT]マーカー処理のテスト"""
        state = word_to_text.ExtractionState()
        state.process_marker("PARENT")
        
        self.assertTrue(state.in_parent)
        self.assertFalse(state.in_skip)
        self.assertEqual(state.found_parent_count, 1)
        
        # 2回目の[PARENT]
        state.process_marker("PARENT")
        self.assertEqual(state.found_parent_count, 2)
    
    def test_process_marker_skip(self):
        """[SKIP]マーカー処理のテスト"""
        state = word_to_text.ExtractionState()
        state.in_parent = True
        state.process_marker("SKIP")
        
        self.assertTrue(state.in_parent)  # in_parentは維持
        self.assertTrue(state.in_skip)
    
    def test_process_marker_child(self):
        """[CHILD]マーカー処理のテスト"""
        state = word_to_text.ExtractionState()
        state.in_parent = True
        state.in_skip = True
        state.process_marker("CHILD")
        
        self.assertTrue(state.in_parent)  # in_parentは維持
        self.assertFalse(state.in_skip)   # in_skipは解除


class TestCheckMarkerType(unittest.TestCase):
    """check_marker_type関数のテスト"""
    
    def test_parent_marker(self):
        """[PARENT]マーカーの検出"""
        self.assertEqual(word_to_text.check_marker_type("[PARENT]"), "PARENT")
        self.assertEqual(word_to_text.check_marker_type("テキスト [PARENT] 続き"), "PARENT")
    
    def test_child_marker(self):
        """[CHILD]マーカーの検出"""
        self.assertEqual(word_to_text.check_marker_type("[CHILD]"), "CHILD")
        self.assertEqual(word_to_text.check_marker_type("テキスト [CHILD]"), "CHILD")
    
    def test_skip_marker(self):
        """[SKIP]マーカーの検出"""
        self.assertEqual(word_to_text.check_marker_type("[SKIP]"), "SKIP")
        self.assertEqual(word_to_text.check_marker_type("[SKIP] テキスト"), "SKIP")
    
    def test_no_marker(self):
        """マーカーなしの場合"""
        self.assertIsNone(word_to_text.check_marker_type("通常のテキスト"))
        self.assertIsNone(word_to_text.check_marker_type(""))
    
    def test_priority_parent_first(self):
        """複数マーカー時は[PARENT]が優先"""
        text = "[PARENT] と [SKIP]"
        self.assertEqual(word_to_text.check_marker_type(text), "PARENT")


class TestGetCombinedMarker(unittest.TestCase):
    """get_combined_marker関数のテスト"""
    
    def test_first_text_has_marker(self):
        """最初のテキストにマーカーがある場合"""
        result = word_to_text.get_combined_marker("[PARENT]", "通常テキスト")
        self.assertEqual(result, "PARENT")
    
    def test_second_text_has_marker(self):
        """2番目のテキストにマーカーがある場合"""
        result = word_to_text.get_combined_marker("通常テキスト", "[SKIP]")
        self.assertEqual(result, "SKIP")
    
    def test_no_marker(self):
        """どちらにもマーカーがない場合"""
        result = word_to_text.get_combined_marker("テキスト1", "テキスト2")
        self.assertIsNone(result)
    
    def test_empty_text(self):
        """空文字列が含まれる場合"""
        result = word_to_text.get_combined_marker("", "[CHILD]")
        self.assertEqual(result, "CHILD")
    
    def test_none_text(self):
        """Noneが含まれる場合"""
        result = word_to_text.get_combined_marker(None, "[PARENT]")
        self.assertEqual(result, "PARENT")


class TestGetTableMarker(unittest.TestCase):
    """get_table_marker関数のテスト"""
    
    def create_mock_table(self, cell_texts):
        """モックテーブルを作成
        
        Args:
            cell_texts: [[row1_cell1, row1_cell2], [row2_cell1, row2_cell2], ...]
        """
        mock_table = Mock()
        mock_rows = []
        
        for row_texts in cell_texts:
            mock_row = Mock()
            mock_cells = []
            for text in row_texts:
                mock_cell = Mock()
                mock_cell.text = text
                mock_cells.append(mock_cell)
            mock_row.cells = mock_cells
            mock_rows.append(mock_row)
        
        mock_table.rows = mock_rows
        return mock_table
    
    def test_single_parent_marker(self):
        """[PARENT]マーカーのみ"""
        table = self.create_mock_table([["[PARENT]", "データ"]])
        self.assertEqual(word_to_text.get_table_marker(table), "PARENT")
    
    def test_single_skip_marker(self):
        """[SKIP]マーカーのみ"""
        table = self.create_mock_table([["[SKIP]", "データ"]])
        self.assertEqual(word_to_text.get_table_marker(table), "SKIP")
    
    def test_single_child_marker(self):
        """[CHILD]マーカーのみ"""
        table = self.create_mock_table([["[CHILD]", "データ"]])
        self.assertEqual(word_to_text.get_table_marker(table), "CHILD")
    
    def test_priority_parent_over_skip(self):
        """優先順位: PARENT > SKIP"""
        table = self.create_mock_table([
            ["[PARENT]", "データ"],
            ["[SKIP]", "データ2"]
        ])
        self.assertEqual(word_to_text.get_table_marker(table), "PARENT")
    
    def test_priority_parent_over_child(self):
        """優先順位: PARENT > CHILD"""
        table = self.create_mock_table([
            ["[CHILD]", "[PARENT]"]
        ])
        self.assertEqual(word_to_text.get_table_marker(table), "PARENT")
    
    def test_priority_skip_over_child(self):
        """優先順位: SKIP > CHILD"""
        table = self.create_mock_table([
            ["[SKIP]", "データ"],
            ["[CHILD]", "データ2"]
        ])
        self.assertEqual(word_to_text.get_table_marker(table), "SKIP")
    
    def test_no_marker(self):
        """マーカーなし"""
        table = self.create_mock_table([["データ1", "データ2"]])
        self.assertIsNone(word_to_text.get_table_marker(table))


class TestPrintTable(unittest.TestCase):
    """print_table関数のテスト"""
    
    def create_mock_table(self, rows_data):
        """モックテーブルを作成
        
        Args:
            rows_data: [[[para1, para2], [para3]], [[para4]]] のような段落データ
        """
        mock_table = Mock()
        mock_rows = []
        
        for row_data in rows_data:
            mock_row = Mock()
            mock_cells = []
            for cell_paras in row_data:
                mock_cell = Mock()
                mock_paragraphs = []
                for para_text in cell_paras:
                    mock_para = Mock()
                    mock_para.text = para_text
                    mock_paragraphs.append(mock_para)
                mock_cell.paragraphs = mock_paragraphs
                mock_cells.append(mock_cell)
            mock_row.cells = mock_cells
            mock_rows.append(mock_row)
        
        mock_table.rows = mock_rows
        return mock_table
    
    def test_simple_table(self):
        """シンプルな表の出力"""
        table = self.create_mock_table([
            [["A"], ["B"], ["C"]],
            [["D"], ["E"], ["F"]]
        ])
        
        captured_output = StringIO()
        sys.stdout = captured_output
        word_to_text.print_table(table)
        sys.stdout = sys.__stdout__
        
        output = captured_output.getvalue()
        self.assertIn("A | B | C", output)
        self.assertIn("D | E | F", output)
    
    def test_cell_with_multiple_paragraphs(self):
        """セル内に複数段落がある場合"""
        table = self.create_mock_table([
            [[f"行1{LINE_BREAK}行2"], ["データ"]]
        ])
        
        captured_output = StringIO()
        sys.stdout = captured_output
        word_to_text.print_table(table)
        sys.stdout = sys.__stdout__
        
        output = captured_output.getvalue()
        self.assertIn(f"行1{LINE_BREAK}行2 | データ", output)


class TestMarkerProcessingFlow(unittest.TestCase):
    """マーカー処理フローの統合テスト"""
    
    def test_parent_section_output(self):
        """[PARENT]セクションが正しく出力されるか"""
        state = word_to_text.ExtractionState()
        
        # [PARENT]前は出力しない
        self.assertFalse(state.in_parent and not state.in_skip)
        
        # [PARENT]で出力開始
        state.process_marker("PARENT")
        self.assertTrue(state.in_parent and not state.in_skip)
    
    def test_skip_section_blocks_output(self):
        """[SKIP]セクションで出力が停止するか"""
        state = word_to_text.ExtractionState()
        state.process_marker("PARENT")
        
        # [PARENT]内で出力中
        self.assertTrue(state.in_parent and not state.in_skip)
        
        # [SKIP]で出力停止
        state.process_marker("SKIP")
        self.assertFalse(state.in_parent and not state.in_skip)
    
    def test_child_resumes_output(self):
        """[CHILD]で出力が再開するか"""
        state = word_to_text.ExtractionState()
        state.process_marker("PARENT")
        state.process_marker("SKIP")
        
        # [SKIP]中は出力しない
        self.assertFalse(state.in_parent and not state.in_skip)
        
        # [CHILD]で出力再開
        state.process_marker("CHILD")
        self.assertTrue(state.in_parent and not state.in_skip)
    
    def test_multiple_parent_sections(self):
        """複数の[PARENT]セクション"""
        state = word_to_text.ExtractionState()
        
        state.process_marker("PARENT")
        self.assertEqual(state.found_parent_count, 1)
        
        state.process_marker("PARENT")
        self.assertEqual(state.found_parent_count, 2)
        # 新しい[PARENT]でも出力継続
        self.assertTrue(state.in_parent and not state.in_skip)


class TestEdgeCases(unittest.TestCase):
    """エッジケースのテスト"""
    
    def test_skip_without_parent(self):
        """[PARENT]なしで[SKIP]が出現"""
        state = word_to_text.ExtractionState()
        state.process_marker("SKIP")
        
        # in_parentがFalseなので出力されない
        self.assertFalse(state.in_parent and not state.in_skip)
    
    def test_child_without_parent(self):
        """[PARENT]なしで[CHILD]が出現"""
        state = word_to_text.ExtractionState()
        state.process_marker("CHILD")
        
        # in_parentがFalseなので出力されない
        self.assertFalse(state.in_parent and not state.in_skip)
    
    def test_consecutive_skip_markers(self):
        """連続する[SKIP]マーカー"""
        state = word_to_text.ExtractionState()
        state.process_marker("PARENT")
        state.process_marker("SKIP")
        state.process_marker("SKIP")
        
        # 2回目の[SKIP]でも状態は変わらない
        self.assertTrue(state.in_skip)


def run_tests():
    """テストを実行"""
    # テストスイートを作成
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # すべてのテストケースを追加
    suite.addTests(loader.loadTestsFromTestCase(TestExtractionState))
    suite.addTests(loader.loadTestsFromTestCase(TestCheckMarkerType))
    suite.addTests(loader.loadTestsFromTestCase(TestGetCombinedMarker))
    suite.addTests(loader.loadTestsFromTestCase(TestGetTableMarker))
    suite.addTests(loader.loadTestsFromTestCase(TestPrintTable))
    suite.addTests(loader.loadTestsFromTestCase(TestMarkerProcessingFlow))
    suite.addTests(loader.loadTestsFromTestCase(TestEdgeCases))
    
    # テスト実行
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # 結果サマリー
    print(LINE_BREAK + "="*70)
    print(f"テスト実行結果: {result.testsRun}件")
    print(f"成功: {result.testsRun - len(result.failures) - len(result.errors)}件")
    print(f"失敗: {len(result.failures)}件")
    print(f"エラー: {len(result.errors)}件")
    print("="*70)
    
    return result.wasSuccessful()


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)

