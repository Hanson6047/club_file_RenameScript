# TODO: implement this file
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
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pathlib import Path


class ClubFileManager:
    """社團檔案管理系統主類"""
    
    def __init__(self, credentials_file=None, sheet_name="社團檔案命名對照表"):
        """
        初始化檔案管理系統
        
        Args:
            credentials_file: Google API 認證檔案路徑
            sheet_name: Google Sheet 名稱
        """
        self.sheet_name = sheet_name
        self.credentials_file = credentials_file
        self.naming_data = None
        self.category_map = None
        self.subcategory_map = None
        self.activity_map = None
        
        # 如果提供認證檔，則嘗試連接 Google Sheet
        if credentials_file and os.path.exists(credentials_file):
            self.connect_to_google_sheet()
        else:
            print("未提供Google認證檔或檔案不存在，將以本地模式運行")
    
    def connect_to_google_sheet(self):
        """連接到Google Sheet獲取命名數據"""
        try:
            # 定義API範圍
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            
            # 認證
            creds = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_file, scope)
            client = gspread.authorize(creds)
            
            # 打開工作表
            sheet = client.open(self.sheet_name).sheet1
            
            # 獲取數據
            data = sheet.get_all_records()
            self.naming_data = pd.DataFrame(data)
            
            # 生成映射表
            self._generate_mapping_tables()
            
            print(f"成功從Google Sheet '{self.sheet_name}'中讀取命名數據")
        except Exception as e:
            print(f"連接Google Sheet時出錯: {e}")
    
    def load_from_excel(self, excel_file):
        """從Excel檔案加載命名數據"""
        try:
            self.naming_data = pd.read_excel(excel_file)
            self._generate_mapping_tables()
            print(f"成功從Excel檔案 '{excel_file}' 中讀取命名數據")
        except Exception as e:
            print(f"讀取Excel檔案時出錯: {e}")
    
    def _generate_mapping_tables(self):
        """從命名數據生成映射表"""
        if self.naming_data is not None:
            # 主分類映射
            self.category_map = dict(zip(
                self.naming_data['主分類名稱'], 
                self.naming_data['主分類代碼']
            ))
            
            # 細分類映射
            self.subcategory_map = dict(zip(
                self.naming_data['細分類名稱'], 
                self.naming_data['細分類代碼']
            ))
            
            # 活動映射
            self.activity_map = dict(zip(
                self.naming_data['活動名稱'], 
                self.naming_data['活動代碼']
            ))
    
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
        # 轉換代碼
        category_code = self.category_map.get(category, category)
        subcategory_code = self.subcategory_map.get(subcategory, subcategory)
        activity_code = self.activity_map.get(activity, activity)
        
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
        重新命名檔案
        
        Args:
            file_path: 原始檔案路徑
            category, subcategory, activity, description, date: 與 generate_filename 相同
            
        Returns:
            新檔案路徑
        """
        # 檢查檔案是否存在
        if not os.path.exists(file_path):
            print(f"檔案不存在: {file_path}")
            return None
        
        # 獲取檔案資訊
        file_dir = os.path.dirname(file_path)
        file_ext = os.path.splitext(file_path)[1]
        
        # 生成新檔名
        new_filename = self.generate_filename(category, subcategory, activity, description, date)
        new_file_path = os.path.join(file_dir, new_filename + file_ext)
        
        # 執行重命名
        try:
            os.rename(file_path, new_file_path)
            print(f"檔案已重命名: {os.path.basename(file_path)} -> {os.path.basename(new_file_path)}")
            return new_file_path
        except Exception as e:
            print(f"重命名檔案時出錯: {e}")
            return None
    
    def create_photo_folder(self, base_dir, activity, date=None, create_if_not_exists=True):
        """
        創建標準命名的照片資料夾
        
        Args:
            base_dir: 基礎目錄
            activity: 活動名稱或代碼
            date: 日期
            create_if_not_exists: 如果不存在是否創建
            
        Returns:
            照片資料夾路徑
        """
        # 照片資料夾使用固定的主分類和細分類
        folder_name = self.generate_filename("B", "PH", activity, f"{activity_code}照片", date)
        folder_path = os.path.join(base_dir, folder_name)
        
        # 創建資料夾
        if create_if_not_exists and not os.path.exists(folder_path):
            try:
                os.makedirs(folder_path)
                print(f"已創建照片資料夾: {folder_name}")
            except Exception as e:
                print(f"創建照片資料夾時出錯: {e}")
                return None
        
        return folder_path
    
    def search_files(self, directory, pattern):
        """
        使用正規表示式搜尋檔案
        
        Args:
            directory: 搜尋目錄
            pattern: 正規表示式模式
            
        Returns:
            匹配的檔案列表
        """
        matches = []
        regex = re.compile(pattern)
        
        for root, dirs, files in os.walk(directory):
            # 檢查目錄名稱
            for d in dirs:
                dir_path = os.path.join(root, d)
                rel_path = os.path.relpath(dir_path, directory) + '/'
                if regex.match(rel_path):
                    matches.append(rel_path)
            
            # 檢查檔案名稱
            for f in files:
                if regex.match(f):
                    matches.append(os.path.join(os.path.relpath(root, directory), f))
        
        return matches
    
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
                print(f"檔案信息不完整: {file_info}")
                continue
                
            orig_filename = file_info[0]
            orig_path = os.path.join(directory, orig_filename)
            
            # 檢查檔案是否為照片
            ext = os.path.splitext(orig_filename)[1].lower()
            photo_exts = ['.jpg', '.jpeg', '.png', '.heic', '.gif']
            
            if ext in photo_exts:
                print(f"跳過照片檔案: {orig_filename} (照片檔案不重命名)")
                results[orig_filename] = orig_filename
                continue
            
            # 重命名非照片檔案
            date = file_info[5] if len(file_info) > 5 else None
            new_path = self.rename_file(orig_path, file_info[1], file_info[2], 
                                       file_info[3], file_info[4], date)
            
            if new_path:
                results[orig_filename] = os.path.basename(new_path)
        
        return results
    
    def export_naming_rules_template(self, output_path):
        """
        導出命名規則模板
        
        Args:
            output_path: 輸出路徑
        """
        template = pd.DataFrame({
            '主分類代碼': ['A', 'B', 'C', 'D'],
            '主分類名稱': ['行政管理', '活動相關', '財務', '會議'],
            '細分類代碼': ['GN', 'PL', 'PH', 'RC'],
            '細分類名稱': ['一般', '企劃', '照片', '報帳'],
            '活動代碼': ['HP', 'FL', 'GN'],
            '活動名稱': ['哈盆', '仙湖', '一般'],
        })
        
        # 導出為 Excel
        template.to_excel(output_path, index=False)
        print(f"命名規則模板已導出至 {output_path}")


# 使用範例
if __name__ == "__main__":
    # 初始化檔案管理器
    file_manager = ClubFileManager()
    
    # 如果未連接Google Sheet，可以從Excel加載
    file_manager.load_from_excel("命名對照表.xlsx")
    
    # 範例1: 生成檔案名稱
    filename = file_manager.generate_filename("活動相關", "企劃", "哈盆", "企劃書")
    print(f"生成的檔案名稱: {filename}")
    
    # 範例2: 重命名檔案
    # file_manager.rename_file("原始檔案.docx", "B", "PL", "HP", "哈盆企劃書", "2024-03-01")
    
    # 範例3: 創建照片資料夾
    # photo_folder = file_manager.create_photo_folder("./社團檔案", "哈盆", "2024-03-01")
    
    # 範例4: 搜尋企劃書
    # matches = file_manager.search_files("./社團檔案", "^B_PL_.*\.docx$")
    # print(f"找到 {len(matches)} 個企劃書檔案")
    
    # 導出命名規則模板
    file_manager.export_naming_rules_template("命名規則模板.xlsx")