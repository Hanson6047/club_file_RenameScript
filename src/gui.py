# TODO: implement this file
"""
社團檔案標準化命名系統 GUI 工具
==========================
此工具提供圖形化界面，幫助社團幹部簡單地標準化檔案命名並進行分類管理。
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
from pathlib import Path
import re
import threading

# 導入主要邏輯類
from club_file_manager import ClubFileManager


class ClubFileManagerGUI:
    """社團檔案管理系統圖形界面"""
    
    def __init__(self, root):
        """初始化GUI界面"""
        self.root = root
        self.file_manager = ClubFileManager()
        self.root.title("社團檔案標準化命名工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 設置界面樣式
        style = ttk.Style()
        style.theme_use('clam')  # 使用較現代的主題
        
        # 創建主框架
        self.create_widgets()
        
        # 狀態追踪
        self.excel_data = None
        self.config_loaded = False
    
    def create_widgets(self):
        """創建界面組件"""
        # 創建頁籤控件
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 頁籤1: 單檔重命名
        self.rename_frame = ttk.Frame(notebook)
        notebook.add(self.rename_frame, text="單檔重命名")
        self.setup_rename_tab()
        
        # 頁籤2: 批次重命名
        self.batch_frame = ttk.Frame(notebook)
        notebook.add(self.batch_frame, text="批次重命名")
        self.setup_batch_tab()
        
        # 頁籤3: 照片資料夾
        self.photo_frame = ttk.Frame(notebook)
        notebook.add(self.photo_frame, text="照片資料夾")
        self.setup_photo_tab()
        
        # 頁籤4: 檔案搜尋
        self.search_frame = ttk.Frame(notebook)
        notebook.add(self.search_frame, text="檔案搜尋")
        self.setup_search_tab()
        
        # 頁籤5: 設定
        self.settings_frame = ttk.Frame(notebook)
        notebook.add(self.settings_frame, text="設定")
        self.setup_settings_tab()
        
        # 建立狀態欄
        self.status_var = tk.StringVar()
        self.status_var.set("就緒，請先載入命名對照表")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_rename_tab(self):
        """設置單檔重命名頁籤"""
        frame = ttk.LabelFrame(self.rename_frame, text="檔案重命名")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 檔案選擇區
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(file_frame, text="選擇檔案:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="瀏覽...", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 命名參數區
        params_frame = ttk.LabelFrame(frame, text="命名參數")
        params_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 主分類
        ttk.Label(params_frame, text="主分類:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(params_frame, textvariable=self.category_var, state="readonly", width=15)
        self.category_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 細分類
        ttk.Label(params_frame, text="細分類:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.subcategory_var = tk.StringVar()
        self.subcategory_combo = ttk.Combobox(params_frame, textvariable=self.subcategory_var, state="readonly", width=15)
        self.subcategory_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 活動
        ttk.Label(params_frame, text="活動:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.activity_var = tk.StringVar()
        self.activity_combo = ttk.Combobox(params_frame, textvariable=self.activity_var, state="readonly", width=15)
        self.activity_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 描述
        ttk.Label(params_frame, text="描述:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.description_var = tk.StringVar()
        ttk.Entry(params_frame, textvariable=self.description_var, width=20).grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 日期
        ttk.Label(params_frame, text="日期 (YYYY-MM-DD):").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_var = tk.StringVar()
        self.date_var.set(datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(params_frame, textvariable=self.date_var, width=15).grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 預覽
        preview_frame = ttk.LabelFrame(frame, text="預覽")
        preview_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(preview_frame, text="新檔名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.preview_var = tk.StringVar()
        ttk.Entry(preview_frame, textvariable=self.preview_var, state="readonly", width=60).grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        
        ttk.Button(preview_frame, text="預覽", command=self.preview_filename).grid(row=1, column=1, sticky=tk.E, padx=5, pady=5)
        ttk.Button(preview_frame, text="重命名", command=self.rename_file).grid(row=1, column=2, sticky=tk.E, padx=5, pady=5)
    
    def setup_batch_tab(self):
        """設置批次重命名頁籤"""
        frame = ttk.LabelFrame(self.batch_frame, text="批次檔案重命名")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 資料夾選擇區
        folder_frame = ttk.Frame(frame)
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(folder_frame, text="選擇資料夾:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.folder_path_var = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.folder_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(folder_frame, text="瀏覽...", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=5)
        
        # 檔案列表
        list_frame = ttk.LabelFrame(frame, text="檔案列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 建立樹狀檢視表格
        columns = ("原始檔名", "主分類", "細分類", "活動", "描述", "日期", "新檔名")
        self.file_table = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        # 設定欄位標題
        for col in columns:
            self.file_table.heading(col, text=col)
            self.file_table.column(col, width=100)
        
        # 添加捲動條
        scrollbar = ttk.Scroll