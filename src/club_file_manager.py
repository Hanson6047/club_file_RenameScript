"""
社團檔案標準化命名與分類系統
=========================
此腳本用於標準化社團檔案命名與分類，支援從Google Sheet讀取命名規則，
並提供檔案重命名、資料夾創建和檔案查詢功能。

命名格式: [主分類代碼]_[細分類代碼]_[活動代碼]_[描述]_[日期].[副檔名]
例如: B_PL_HP_哈盆企劃書_2024-03-01.docx
"""

import os
import re
import shutil
import pandas as pd
from datetime import datetime
import logging
from typing import Dict, List, Optional, Tuple, Union
from pathlib import Path

# 如果使用Google Sheet，才導入這些模塊
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False


class ClubFileManager:
    """社團檔案管理系統主類"""
    
    def __init__(self, config_file=None, use_google_sheet=False, sheet_name="社團檔案命名對照表", credentials_file=None):
        """
        初始化檔案管理系統
        
        Args:
            config_file: Excel配置文件路徑或Google Sheet ID
            use_google_sheet: 是否使用Google Sheet
            sheet_name: Google Sheet名稱
            credentials_file: Google API 認證檔案路徑
        """
        # 設置日誌
        self.logger = logging.getLogger('club_file_system.manager')
        
        self.config_file = config_file
        self.use_google_sheet = use_google_sheet
        self.sheet_name = sheet_name
        self.credentials_file = credentials_file
        
        # 命名數據相關
        self.naming_data = None
        
        # 映射表
        self.category_map = {}  # 主分類映射: 名稱 -> 代碼
        self.subcategory_map = {}  # 細分類映射: 名稱 -> 代碼
        self.activity_map = {}  # 活動映射: 名稱 -> 代碼
        
        # 反向映射表 (代碼 -> 名稱)
        self.category_code_map = {}  # 主分類映射: 代碼 -> 名稱
        self.subcategory_code_map = {}  # 細分類映射: 代碼 -> 名稱
        self.activity_code_map = {}  # 活動映射: 代碼 -> 名稱
        
        # 如果提供配置檔，則嘗試載入
        if config_file:
            if use_google_sheet and GSPREAD_AVAILABLE and credentials_file:
                self.connect_to_google_sheet(credentials_file, sheet_name)
            else:
                self.load_from_excel(config_file)
    
    def connect_to_google_sheet(self, credentials_file=None, sheet_name=None):
        """
        連接到Google Sheet獲取命名數據
        
        Args:
            credentials_file: Google API認證檔案路徑
            sheet_name: Google Sheet名稱
            
        Returns:
            bool: 是否成功連接
        """
        if not GSPREAD_AVAILABLE:
            self.logger.error("未安裝gspread或oauth2client模塊，無法連接Google Sheet")
            return False
            
        try:
            # 使用參數值或實例屬性
            cred_file = credentials_file or self.credentials_file
            sheet = sheet_name or self.sheet_name
            
            if not cred_file or not os.path.exists(cred_file):
                self.logger.error(f"Google認證檔案不存在: {cred_file}")
                return False
            
            # 定義API範圍
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            
            # 認證
            creds = ServiceAccountCredentials.from_json_keyfile_name(cred_file, scope)
            client = gspread.authorize(creds)
            
            # 打開工作表
            spreadsheet = client.open(sheet)
            
            # 讀取各分頁
            try:
                main_categories = spreadsheet.worksheet("主分類").get_all_records()
                sub_categories = spreadsheet.worksheet("細分類").get_all_records()
                activities = spreadsheet.worksheet("活動").get_all_records()
                
                # 轉換為DataFrame
                main_df = pd.DataFrame(main_categories)
                sub_df = pd.DataFrame(sub_categories)
                activity_df = pd.DataFrame(activities)
                
                # 處理數據
                self._process_config_data(main_df, sub_df, activity_df)
                
                self.logger.info(f"成功從Google Sheet '{sheet}'載入命名數據")
                return True
                
            except Exception as e:
                self.logger.error(f"讀取Google Sheet分頁時出錯: {str(e)}")
                return False
                
        except Exception as e:
            self.logger.error(f"連接Google Sheet時出錯: {str(e)}")
            return False
    
    def load_from_excel(self, excel_file):
        """
        從Excel檔案載入命名數據
        
        Args:
            excel_file: Excel檔案路徑
            
        Returns:
            bool: 是否成功載入
        """
        try:
            if not os.path.exists(excel_file):
                self.logger.error(f"Excel檔案不存在: {excel_file}")
                return False
                
            # 讀取Excel各表格
            main_df = pd.read_excel(excel_file, sheet_name="主分類")
            sub_df = pd.read_excel(excel_file, sheet_name="細分類")
            activity_df = pd.read_excel(excel_file, sheet_name="活動")
            
            # 處理數據
            self._process_config_data(main_df, sub_df, activity_df)
            
            self.logger.info(f"成功從Excel檔案 '{excel_file}' 載入命名數據")
            return True
            
        except Exception as e:
            self.logger.error(f"讀取Excel檔案時出錯: {str(e)}")
            return False
    
    def _process_config_data(self, main_df, sub_df, activity_df):
        """
        處理從配置文件載入的數據
        
        Args:
            main_df: 主分類數據
            sub_df: 細分類數據
            activity_df: 活動數據
        """
        # 清理列名
        main_df.columns = [col.strip() if isinstance(col, str) else col for col in main_df.columns]
        sub_df.columns = [col.strip() if isinstance(col, str) else col for col in sub_df.columns]
        activity_df.columns = [col.strip() if isinstance(col, str) else col for col in activity_df.columns]
        
        # 存儲數據
        self.naming_data = {
            'main_category': main_df,
            'sub_category': sub_df,
            'activity': activity_df
        }
        
        # 生成映射表
        self._generate_mapping_tables()
    
    def _generate_mapping_tables(self):
        """生成各種代碼和名稱的映射表"""
        # 檢查數據是否已載入
        if not self.naming_data:
            self.logger.warning("無法生成映射表: 數據未載入")
            return
        
        # 生成主分類映射
        main_df = self.naming_data['main_category']
        if '主分類代碼' in main_df.columns and '主分類名稱' in main_df.columns:
            for _, row in main_df.iterrows():
                code = row['主分類代碼'] if isinstance(row['主分類代碼'], str) else str(row['主分類代碼'])
                name = row['主分類名稱'] if isinstance(row['主分類名稱'], str) else str(row['主分類名稱'])
                self.category_map[name.strip()] = code.strip()
                self.category_code_map[code.strip()] = name.strip()
        
        # 生成細分類映射
        sub_df = self.naming_data['sub_category']
        if '細分類代碼' in sub_df.columns and '細分類名稱' in sub_df.columns:
            for _, row in sub_df.iterrows():
                code = row['細分類代碼'] if isinstance(row['細分類代碼'], str) else str(row['細分類代碼'])
                name = row['細分類名稱'] if isinstance(row['細分類名稱'], str) else str(row['細分類名稱'])
                self.subcategory_map[name.strip()] = code.strip()
                self.subcategory_code_map[code.strip()] = name.strip()
        
        # 生成活動映射
        activity_df = self.naming_data['activity']
        if '活動代碼' in activity_df.columns and '活動名稱' in activity_df.columns:
            for _, row in activity_df.iterrows():
                code = row['活動代碼'] if isinstance(row['活動代碼'], str) else str(row['活動代碼'])
                name = row['活動名稱'] if isinstance(row['活動名稱'], str) else str(row['活動名稱'])
                self.activity_map[name.strip()] = code.strip()
                self.activity_code_map[code.strip()] = name.strip()
        
        self.logger.debug(f"生成映射表完成: {len(self.category_map)}個主分類, {len(self.subcategory_map)}個細分類, {len(self.activity_map)}個活動")
    
    def get_main_categories(self):
        """獲取所有主分類名稱"""
        return list(self.category_map.keys())
    
    def get_sub_categories(self):
        """獲取所有細分類名稱"""
        return list(self.subcategory_map.keys())
    
    def get_activities(self):
        """獲取所有活動名稱"""
        return list(self.activity_map.keys())
    
    def generate_filename(self, category, subcategory, activity, description, date=None):
        """
        生成標準化檔案名稱
        
        Args:
            category: 主分類名稱或代碼
            subcategory: 細分類名稱或代碼
            activity: 活動名稱或代碼
            description: 檔案描述
            date: 日期 (如果未提供則使用今天日期)
            
        Returns:
            標準化的檔案名稱(不含副檔名)
        """
        # 將輸入轉換為字符串
        category = str(category).strip()
        subcategory = str(subcategory).strip()
        activity = str(activity).strip()
        description = str(description).strip()
        
        # 轉換代碼 (支持輸入名稱或代碼)
        category_code = self.category_map.get(category, category)
        subcategory_code = self.subcategory_map.get(subcategory, subcategory)
        activity_code = self.activity_map.get(activity, activity)
        
        # 檢查代碼是否有效 (可選)
        if category_code not in self.category_code_map:
            self.logger.warning(f"未找到主分類 '{category}' 對應的有效代碼")
        if subcategory_code not in self.subcategory_code_map:
            self.logger.warning(f"未找到細分類 '{subcategory}' 對應的有效代碼")
        if activity_code not in self.activity_code_map:
            self.logger.warning(f"未找到活動 '{activity}' 對應的有效代碼")
        
        # 處理日期
        if not date:
            date = datetime.now().strftime('%Y-%m-%d')
        elif isinstance(date, datetime):
            date = date.strftime('%Y-%m-%d')
        
        # 生成檔案名稱
        filename = f"{category_code}_{subcategory_code}_{activity_code}_{description}_{date}"
        return filename
    
    def rename_file(self, file_path, category, subcategory, activity, description, date=None):
        """
        重命名檔案
        
        Args:
            file_path: 原始檔案路徑
            category, subcategory, activity, description, date: 與 generate_filename 相同
            
        Returns:
            新檔案路徑 或 None (如果失敗)
        """
        # 檢查檔案是否存在
        if not os.path.exists(file_path):
            self.logger.error(f"檔案不存在: {file_path}")
            return None
        
        # 檢查是否為照片檔案
        ext = os.path.splitext(file_path)[1].lower()
        photo_exts = ['.jpg', '.jpeg', '.png', '.heic', '.gif']
        if ext in photo_exts:
            self.logger.info(f"跳過照片檔案: {os.path.basename(file_path)} (照片檔案不重命名)")
            return file_path
        
        # 獲取檔案資訊
        file_dir = os.path.dirname(file_path)
        file_ext = os.path.splitext(file_path)[1]
        
        # 生成新檔名
        new_filename = self.generate_filename(category, subcategory, activity, description, date)
        new_file_path = os.path.join(file_dir, new_filename + file_ext)
        
        # 檢查目標檔案是否已存在
        if os.path.exists(new_file_path) and os.path.abspath(file_path) != os.path.abspath(new_file_path):
            self.logger.warning(f"目標檔案已存在: {os.path.basename(new_file_path)}")
            # 生成唯一檔名
            base_name = new_filename
            counter = 1
            while os.path.exists(new_file_path):
                new_filename = f"{base_name}_{counter}"
                new_file_path = os.path.join(file_dir, new_filename + file_ext)
                counter += 1
            self.logger.info(f"使用唯一檔名: {os.path.basename(new_file_path)}")
        
        # 執行重命名
        try:
            os.rename(file_path, new_file_path)
            self.logger.info(f"檔案已重命名: {os.path.basename(file_path)} -> {os.path.basename(new_file_path)}")
            return new_file_path
        except Exception as e:
            self.logger.error(f"重命名檔案時出錯: {str(e)}")
            return None
    
    def batch_rename(self, directory, file_info_list):
        """
        批次重命名檔案
        
        Args:
            directory: 檔案目錄
            file_info_list: 檔案信息列表，每項包含:
                            [原始檔名, 主分類, 細分類, 活動, 描述, 日期(可選)]
        
        Returns:
            重命名結果字典 {原始檔名: 新檔名}
        """
        results = {}
        
        for file_info in file_info_list:
            # 解析檔案信息
            if len(file_info) < 5:
                self.logger.warning(f"檔案信息不完整: {file_info}")
                continue
                
            orig_filename = file_info[0]
            orig_path = os.path.join(directory, orig_filename)
            
            # 檢查檔案是否存在
            if not os.path.exists(orig_path):
                self.logger.warning(f"檔案不存在: {orig_path}")
                results[orig_filename] = None
                continue
            
            # 檢查檔案是否為照片
            ext = os.path.splitext(orig_filename)[1].lower()
            photo_exts = ['.jpg', '.jpeg', '.png', '.heic', '.gif']
            
            if ext in photo_exts:
                self.logger.info(f"跳過照片檔案: {orig_filename} (照片檔案不重命名)")
                results[orig_filename] = orig_filename
                continue
            
            # 重命名非照片檔案
            date = file_info[5] if len(file_info) > 5 else None
            new_path = self.rename_file(orig_path, file_info[1], file_info[2], 
                                       file_info[3], file_info[4], date)
            
            if new_path:
                results[orig_filename] = os.path.basename(new_path)
            else:
                results[orig_filename] = None
        
        return results
    
    def create_photo_folder(self, base_dir, activity, date=None):
        """
        創建標準命名的照片資料夾
        
        Args:
            base_dir: 基礎目錄
            activity: 活動名稱或代碼
            date: 日期 (如果未提供則使用今天日期)
            
        Returns:
            照片資料夾路徑 或 None (如果失敗)
        """
        try:
            # 活動代碼處理
            activity_str = str(activity).strip()
            activity_code = self.activity_map.get(activity_str, activity_str)
            activity_name = self.activity_code_map.get(activity_code, activity_str)
            
            # 處理日期
            if not date:
                date = datetime.now().strftime('%Y-%m-%d')
            elif isinstance(date, datetime):
                date = date.strftime('%Y-%m-%d')
            
            # 照片資料夾使用固定的主分類和細分類
            folder_name = f"B_PH_{activity_code}_{activity_name}照片_{date}"
            folder_path = os.path.join(base_dir, folder_name)
            
            # 創建資料夾
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                self.logger.info(f"已創建照片資料夾: {folder_name}")
            else:
                self.logger.info(f"照片資料夾已存在: {folder_name}")
            
            return folder_path
            
        except Exception as e:
            self.logger.error(f"創建照片資料夾時出錯: {str(e)}")
            return None
    
    def search_files(self, directory, pattern):
        """
        使用正規表示式搜尋檔案
        
        Args:
            directory: 搜尋目錄
            pattern: 正規表示式模式
            
        Returns:
            匹配的檔案/資料夾路徑列表
        """
        try:
            if not os.path.exists(directory):
                self.logger.warning(f"搜尋目錄不存在: {directory}")
                return []
                
            matches = []
            regex = re.compile(pattern)
            
            for root, dirs, files in os.walk(directory):
                # 檢查目錄名稱
                for d in dirs:
                    if regex.search(d):
                        dir_path = os.path.join(root, d)
                        matches.append(dir_path)
                
                # 檢查檔案名稱
                for f in files:
                    if regex.search(f):
                        file_path = os.path.join(root, f)
                        matches.append(file_path)
            
            self.logger.info(f"使用模式 '{pattern}' 找到 {len(matches)} 個匹配項")
            return matches
            
        except Exception as e:
            self.logger.error(f"搜尋檔案時出錯: {str(e)}")
            return []
    
    def parse_filename(self, filename):
        """
        解析標準化檔案名
        
        Args:
            filename: 檔案名
            
        Returns:
            字典 {'main_code': X, 'main_category': X, ...} 或 None
        """
        try:
            # 移除路徑和副檔名
            basename = os.path.basename(filename)
            name, _ = os.path.splitext(basename)
            
            # 匹配標準格式
            pattern = r'^([A-Z])_([A-Z]{2})_([A-Z]{2})_(.+)_(\d{4}-\d{2}-\d{2})$'
            match = re.match(pattern, name)
            
            if match:
                main_code, subcategory_code, activity_code, description, date = match.groups()
                
                return {
                    'main_code': main_code,
                    'main_category': self.category_code_map.get(main_code, ''),
                    'subcategory_code': subcategory_code,
                    'subcategory': self.subcategory_code_map.get(subcategory_code, ''),
                    'activity_code': activity_code,
                    'activity': self.activity_code_map.get(activity_code, ''),
                    'description': description,
                    'date': date,
                    'is_standard': True
                }
            else:
                self.logger.warning(f"檔案名不符合標準格式: {basename}")
                return {
                    'filename': basename,
                    'is_standard': False
                }
                
        except Exception as e:
            self.logger.error(f"解析檔案名時出錯: {str(e)}")
            return None
    
    def export_naming_rules_template(self, output_path):
        """
        導出命名規則模板
        
        Args:
            output_path: 輸出路徑
            
        Returns:
            bool: 是否成功導出
        """
        try:
            # 主分類表
            main_categories = pd.DataFrame({
                '主分類代碼': ['A', 'B', 'C', 'D'],
                '主分類名稱': ['行政管理', '活動相關', '財務', '會議'],
                '說明': ['社團行政、人員管理', '活動相關檔案', '財務相關檔案', '會議相關檔案']
            })
            
            # 細分類表
            sub_categories = pd.DataFrame({
                '細分類代碼': ['GN', 'PL', 'PH', 'MT', 'RC'],
                '細分類名稱': ['一般', '企劃', '照片', '會議', '報帳'],
                '適用主分類': ['A,B,C,D', 'B', 'B', 'D', 'C']
            })
            
            # 活動表
            activities = pd.DataFrame({
                '活動代碼': ['HP', 'FL', 'GN', 'AG'],
                '活動名稱': ['哈盆', '仙湖', '一般', '社員大會'],
                '活動日期': ['2024-03-01', '2024-08-30', '', '2024-05-20']
            })
            
            # 導出到Excel的不同分頁
            with pd.ExcelWriter(output_path) as writer:
                main_categories.to_excel(writer, sheet_name='主分類', index=False)
                sub_categories.to_excel(writer, sheet_name='細分類', index=False)
                activities.to_excel(writer, sheet_name='活動', index=False)
            
            self.logger.info(f"命名規則模板已導出至 {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"導出命名規則模板時出錯: {str(e)}")
            return False


# 使用範例
if __name__ == "__main__":
    # 設置日誌
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 初始化檔案管理器
    file_manager = ClubFileManager()
    
    # 測試導出模板
    file_manager.export_naming_rules_template("命名規則模板.xlsx")
    
    # 從Excel加載
    file_manager.load_from_excel("命名規則模板.xlsx")
    
    # 範例: 生成檔案名稱
    filename = file_manager.generate_filename("活動相關", "企劃", "哈盆", "企劃書")
    print(f"生成的檔案名稱: {filename}")