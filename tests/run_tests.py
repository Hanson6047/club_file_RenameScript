#!/usr/bin/env python
"""
社團檔案標準化命名系統 - 測試運行器
=========================
運行所有單元測試
"""

import unittest
import sys
import os
import logging
import argparse

# 設置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tests.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 添加項目根目錄到Python路徑
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# 導入測試模組
from test_manager import TestClubFileManager
from test_gui import TestGUI


def run_tests(test_type='all'):
    """
    運行測試
    
    Args:
        test_type: 要運行的測試類型，可選值為 'all', 'manager' 或 'gui'
    """
    logger = logging.getLogger('run_tests')
    logger.info(f"開始運行測試: {test_type}...")
    
    # 創建測試套件
    suite = unittest.TestSuite()
    
    if test_type in ['all', 'manager']:
        # 添加ClubFileManager測試
        logger.info("添加ClubFileManager測試...")
        suite.addTest(unittest.makeSuite(TestClubFileManager))
    
    if test_type in ['all', 'gui']:
        # 添加GUI測試
        logger.info("添加GUI測試...")
        suite.addTest(unittest.makeSuite(TestGUI))
    
    # 運行測試
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # 輸出測試結果摘要
    logger.info(f"測試完成: 運行 {result.testsRun} 個測試")
    logger.info(f"成功: {result.testsRun - len(result.errors) - len(result.failures)}")
    logger.info(f"失敗: {len(result.failures)}")
    logger.info(f"錯誤: {len(result.errors)}")
    
    if result.failures:
        logger.warning("失敗的測試:")
        for test, traceback in result.failures:
            logger.warning(f"- {test}")
    
    if result.errors:
        logger.warning("錯誤的測試:")
        for test, traceback in result.errors:
            logger.warning(f"- {test}")
    
    # 返回測試結果，以便在腳本中使用
    return result.wasSuccessful()


if __name__ == '__main__':
    # 解析命令行參數
    parser = argparse.ArgumentParser(description='運行社團檔案標準化命名系統測試')
    parser.add_argument('--type', choices=['all', 'manager', 'gui'], default='all',
                        help='要運行的測試類型: all, manager 或 gui (預設: all)')
    parser.add_argument('--failfast', action='store_true',
                        help='設置此選項會在第一個測試失敗時停止測試')
    args = parser.parse_args()
    
    # 運行測試
    success = run_tests(args.type)
    
    # 設置退出代碼
    sys.exit(0 if success else 1)