# TODO: implement this file
"""
社團檔案標準化命名系統測試
"""
import unittest
import os
import sys
import pandas as pd
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.club_file_manager import ClubFileManager

class TestClubFileManager(unittest.TestCase):
    """測試檔案管理系統"""
    
    def setUp(self):
        """設置測試環境"""
        self.manager = ClubFileManager()
        # 創建測試數據
        test_data = {
            '主分類代碼': ['A', 'B', 'C', 'D'],
            '主分類名稱': ['行政管理', '活動相關', '財務', '會議'],
            '細分類代碼': ['GN', 'PL', 'PH', 'RC'],
            '細分類名稱': ['一般', '企劃', '照片', '報帳'],
            '活動代碼': ['HP', 'FL', 'GN'],
            '活動名稱': ['哈盆', '仙湖', '一般'],
        }
        self.test_df = pd.DataFrame(test_data)
        self.manager.naming_data = self.test_df
        self.manager._generate_mapping_tables()
    
    def test_generate_filename(self):
        """測試檔案名稱生成"""
        filename = self.manager.generate_filename("活動相關", "企劃", "哈盆", "企劃書", "2024-03-01")
        self.assertEqual(filename, "B_PL_HP_企劃書_2024-03-01")

if __name__ == '__main__':
    unittest.main()