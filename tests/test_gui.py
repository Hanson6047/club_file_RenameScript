"""
社團檔案標準化命名系統 - GUI測試
=========================
測試 ClubFileManagerGUI 類的基本功能
注意：GUI測試通常需要人工互動，這裡只測試一些基本功能
"""

import unittest
import os
import sys
import tempfile
import shutil
import tkinter as tk
from tkinter import ttk
import logging
from unittest.mock import patch, MagicMock

# 添加上層目錄到 Python 路徑
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 導入要測試的模組
try:
    from src.gui import ClubFileManagerGUI
    from src.club_file_manager import ClubFileManager
except ImportError as e:
    print(f"導入錯誤: {e}")
    print("如果gui.py不在src目錄下，請調整導入路徑")
    sys.exit(1)


class TestGUI(unittest.TestCase):
    """測試GUI功能的基本單元測試"""
    
    @classmethod
    def setUpClass(cls):
        """設置測試環境"""
        # 設置日誌
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # 創建臨時測試目錄
        cls.test_dir = tempfile.mkdtemp()
        cls.logger = logging.getLogger('test_gui')
        cls.logger.info(f"創建測試臨時目錄: {cls.test_dir}")
        
        # 創建測試配置和文件路徑
        cls.config_dir = os.path.join(cls.test_dir, "config")
        os.makedirs(cls.config_dir, exist_ok=True)
        cls.config_file = os.path.join(cls.config_dir, "example_config.xlsx")
        
        # 創建測試文件目錄
        cls.files_dir = os.path.join(cls.test_dir, "files")
        os.makedirs(cls.files_dir, exist_ok=True)
    
    @classmethod
    def tearDownClass(cls):
        """清理測試環境"""
        cls.logger.info(f"刪除測試臨時目錄: {cls.test_dir}")
        shutil.rmtree(cls.test_dir)
    
    def setUp(self):
        """每個測試前的準備工作"""
        # 創建Tkinter根窗口
        self.root = tk.Tk()
        # 創建一個臨時配置文件
        self._create_test_config_file()
        # 創建測試文件
        self._create_test_files()
        
        # 使用模擬對象替代真實的文件對話框
        self.patcher = patch('tkinter.filedialog')
        self.mock_filedialog = self.patcher.start()
        
        # 設置模擬值
        self.mock_filedialog.askopenfilename.return_value = os.path.join(self.files_dir, "test1.docx")
        self.mock_filedialog.askdirectory.return_value = self.files_dir
        
        # 初始化GUI
        self.gui = ClubFileManagerGUI(self.root)
        
        # 手動設置配置文件路徑
        self.gui.excel_path.set(self.config_file)
    
    def tearDown(self):
        """每個測試後的清理工作"""
        # 停止模擬
        self.patcher.stop()
        # 銷毀Tkinter窗口
        self.root.destroy()
    
    def _create_test_config_file(self):
        """創建測試配置文件"""
        # 創建一個臨時的ClubFileManager實例來導出模板
        manager = ClubFileManager()
        manager.export_naming_rules_template(self.config_file)
        self.logger.info(f"已創建測試配置文件: {self.config_file}")
    
    def _create_test_files(self):
        """創建測試文件"""
        # 創建幾個測試文件
        test_files = {
            'test1.docx': "測試文件1",
            'test2.xlsx': "測試表格1",
            'photo1.jpg': "測試照片1",
            'B_PL_HP_測試企劃書_2024-03-01.docx': "已命名的測試文件"
        }
        
        for filename, content in test_files.items():
            file_path = os.path.join(self.files_dir, filename)
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
                
        self.logger.info(f"已創建 {len(test_files)} 個測試文件")
    
    # 以下是測試方法
    
    def test_gui_initialization(self):
        """測試GUI初始化"""
        # 檢查GUI實例是否創建成功
        self.assertIsNotNone(self.gui)
        # 檢查各主要組件是否存在
        self.assertIsNotNone(self.gui.rename_frame)
        self.assertIsNotNone(self.gui.batch_frame)
        self.assertIsNotNone(self.gui.photo_frame)
        self.assertIsNotNone(self.gui.search_frame)
        self.assertIsNotNone(self.gui.settings_frame)
    
    def test_load_config(self):
        """測試載入配置功能"""
        with patch.object(self.gui, 'update_combo_boxes') as mock_update:
            # 調用載入配置方法
            self.gui.load_config()
            
            # 檢查file_manager是否創建
            self.assertIsNotNone(self.gui.file_manager)
            
            # 檢查是否調用了更新下拉框方法
            mock_update.assert_called_once()
            
            # 檢查配置載入狀態
            self.assertTrue(self.gui.config_loaded)
    
    def test_preview_filename(self):
        """測試文件名預覽功能"""
        # 先載入配置
        with patch.object(self.gui, 'messagebox'):  # 抑制消息框
            self.gui.load_config()
        
        # 設置參數
        self.gui.category_var.set("活動相關")
        self.gui.subcategory_var.set("企劃")
        self.gui.activity_var.set("哈盆")
        self.gui.description_var.set("測試預覽")
        self.gui.date_var.set("2024-03-01")
        self.gui.file_path_var.set(os.path.join(self.files_dir, "test1.docx"))
        
        # 調用預覽方法
        self.gui.preview_filename()
        
        # 檢查預覽結果
        self.assertEqual(self.gui.preview_var.get(), "B_PL_HP_測試預覽_2024-03-01.docx")
    
    def test_browse_file(self):
        """測試瀏覽文件功能"""
        # 調用瀏覽文件方法
        self.gui.browse_file()
        
        # 檢查是否調用了文件對話框
        self.mock_filedialog.askopenfilename.assert_called_once()
        
        # 檢查文件路徑是否設置正確
        self.assertEqual(self.gui.file_path_var.get(), os.path.join(self.files_dir, "test1.docx"))
    
    def test_browse_folder(self):
        """測試瀏覽文件夾功能"""
        # 調用瀏覽文件夾方法
        self.gui.browse_folder()
        
        # 檢查是否調用了目錄對話框
        self.mock_filedialog.askdirectory.assert_called_once()
        
        # 檢查文件夾路徑是否設置正確
        self.assertEqual(self.gui.folder_path_var.get(), self.files_dir)
    
    @patch('tkinter.messagebox.showerror')
    def test_rename_file_without_config(self, mock_showerror):
        """測試未載入配置時重命名文件"""
        # 確保file_manager為None
        self.gui.file_manager = None
        
        # 調用重命名方法
        self.gui.rename_file()
        
        # 檢查是否顯示錯誤消息
        mock_showerror.assert_called_once()
        args, _ = mock_showerror.call_args
        self.assertIn("請先載入命名對照表", args[1])
    
    def test_setup_log_handler(self):
        """測試日誌處理器設置"""
        # 創建自定義日誌記錄
        logger = logging.getLogger('test_log_handler')
        
        # 發送測試日誌消息
        logger.info("這是一條測試日誌消息")
        
        # 檢查日誌文本框是否存在
        self.assertIsNotNone(self.gui.log_text)
    
    @patch('os.path.exists')
    @patch('os.rename')
    def test_rename_file_success(self, mock_rename, mock_exists):
        """測試文件重命名成功的情況"""
        # 配置模擬行為
        mock_exists.return_value = True
        
        # 先載入配置
        with patch.object(self.gui, 'messagebox'):  # 抑制消息框
            self.gui.load_config()
        
        # 設置參數
        self.gui.file_path_var.set(os.path.join(self.files_dir, "test1.docx"))
        self.gui.category_var.set("活動相關")
        self.gui.subcategory_var.set("企劃")
        self.gui.activity_var.set("哈盆")
        self.gui.description_var.set("重命名測試")
        self.gui.date_var.set("2024-03-01")
        
        # 模擬file_manager.rename_file方法
        with patch.object(self.gui.file_manager, 'rename_file') as mock_rename_file:
            new_path = os.path.join(self.files_dir, "B_PL_HP_重命名測試_2024-03-01.docx")
            mock_rename_file.return_value = new_path
            
            # 調用重命名方法
            with patch('tkinter.messagebox.showinfo'):  # 抑制成功消息框
                self.gui.rename_file()
            
            # 檢查是否調用了重命名方法
            mock_rename_file.assert_called_once()
            
            # 檢查參數是否正確
            args, _ = mock_rename_file.call_args
            self.assertEqual(args[0], os.path.join(self.files_dir, "test1.docx"))
            self.assertEqual(args[1], "活動相關")
            self.assertEqual(args[2], "企劃")
            self.assertEqual(args[3], "哈盆")
            self.assertEqual(args[4], "重命名測試")
            self.assertEqual(args[5], "2024-03-01")
            
            # 檢查文件路徑是否更新
            self.assertEqual(self.gui.file_path_var.get(), new_path)


if __name__ == '__main__':
    unittest.main()