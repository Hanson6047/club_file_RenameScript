"""
社團檔案標準化命名系統 GUI 工具
==========================
此工具提供圖形化界面，幫助社團幹部簡單地標準化檔案命名並進行分類管理。
"""

import os  # 操作系統相關功能，如檔案路徑處理
import sys  # 系統相關功能，如退出程式
import tkinter as tk  # Python標準GUI庫
from tkinter import ttk, filedialog, messagebox, scrolledtext  # tkinter的進階控件和對話框
import pandas as pd  # 數據處理庫
from datetime import datetime  # 日期時間處理
from pathlib import Path  # 現代化的路徑處理（比os.path更直觀）
import re  # 正則表達式處理
import threading  # 多線程處理（用於長時間運行的任務）
import logging  # 日誌記錄
from typing import List, Dict, Optional, Tuple, Any  # 類型提示（幫助IDE提供更好的代碼補全）

# 導入主要邏輯類 (假設club_file_manager.py在同一個包中)
try:
    from .club_file_manager import ClubFileManager  # 相對導入，適用於作為包安裝
except ImportError:
    from club_file_manager import ClubFileManager  # 直接導入，適用於直接運行


class ClubFileManagerGUI:
    """社團檔案管理系統圖形界面"""
    
    def __init__(self, root):
        """
        初始化GUI界面
        
        Args:
            root: tkinter的根窗口
        """
        # 設置日誌 - 使用階層式名稱(logging的最佳實踐)
        self.logger = logging.getLogger('club_file_system.gui')
        
        # 保存根窗口引用
        self.root = root
        self.file_manager = None  # 稍後初始化
        
        # 設置窗口標題和大小
        self.root.title("社團檔案標準化命名工具")
        self.root.geometry("800x600")  # 設置初始窗口大小
        self.root.resizable(True, True)  # 允許調整窗口大小
        
        # 設置界面樣式 (ttk是tkinter的主題化版本)
        style = ttk.Style()
        style.theme_use('clam')  # 使用較現代的主題，其他選項有: 'alt', 'default', 'classic', 等
        
        # 設置全局字體 (使用Microsoft JhengHei UI可以更好地顯示中文)
        default_font = ('Microsoft JhengHei UI', 10)  # (字體名稱, 字體大小)
        style.configure('.', font=default_font)  # '.'表示所有控件
        
        # 初始化實例變量 - 使用StringVar等可以自動更新UI
        # 配置相關變量
        self.excel_path = tk.StringVar()  # Excel檔案路徑
        self.google_sheet_id = tk.StringVar()  # Google Sheet ID
        self.use_google_sheet = tk.BooleanVar(value=False)  # 是否使用Google Sheet
        self.credentials_file = tk.StringVar()  # Google API憑證檔案路徑
        
        # 檔案選擇相關變量
        self.file_path_var = tk.StringVar()  # 單一檔案路徑
        self.folder_path_var = tk.StringVar()  # 資料夾路徑(批次處理)
        self.search_folder_var = tk.StringVar()  # 搜尋資料夾路徑
        self.photo_folder_var = tk.StringVar()  # 照片資料夾路徑
        
        # 命名參數相關變量
        self.category_var = tk.StringVar()  # 主分類
        self.subcategory_var = tk.StringVar()  # 細分類
        self.activity_var = tk.StringVar()  # 活動
        self.description_var = tk.StringVar()  # 描述
        self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))  # 日期，預設今天
        self.preview_var = tk.StringVar()  # 檔名預覽
        self.search_pattern_var = tk.StringVar()  # 搜尋模式
        
        # 批次處理相關
        self.selected_files = []  # 選擇的檔案列表
        self.batch_files = []  # 批次處理的檔案
        
        # 創建主框架和所有UI元素
        self.create_widgets()
        
        # 狀態追踪
        self.config_loaded = False
        
        # 嘗試載入默認配置
        self.try_load_default_config()
    
    def create_widgets(self):
        """創建所有界面組件"""
        # 創建頁籤控件 (Notebook) - 允許在單一窗口中顯示多個頁面
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        # fill=tk.BOTH: 填充所有可用空間，expand=True: 如果窗口變大也跟著變大
        # padx/pady: 控件外部的水平/垂直填充空間
        
        # 頁籤1: 單檔重命名
        self.rename_frame = ttk.Frame(notebook)  # 創建框架
        notebook.add(self.rename_frame, text="單檔重命名")  # 添加到頁籤中
        self.setup_rename_tab()  # 設置頁籤內容
        
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
        
        # 建立狀態欄 - 顯示程式狀態
        self.status_var = tk.StringVar()
        self.status_var.set("就緒，請先載入命名對照表")
        status_bar = ttk.Label(
            self.root, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN,  # 使其看起來像是嵌入
            anchor=tk.W  # 文字靠左對齊 (W=西=左)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)  # 放在窗口底部，水平填充
    
    def setup_rename_tab(self):
        """設置單檔重命名頁籤的UI元素"""
        # 創建帶標籤的框架
        frame = ttk.LabelFrame(self.rename_frame, text="檔案重命名")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ======== 檔案選擇區 ========
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill=tk.X, padx=5, pady=5)  # X方向填充
        
        # 標籤 + 文字輸入框 + 按鈕的組合
        ttk.Label(file_frame, text="選擇檔案:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        # grid: 使用網格佈局管理器，更靈活地排列控件
        # sticky: 控件在網格中的對齊方式，W=西=左對齊
        
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        # Entry: 單行文字輸入框
        # textvariable: 連結到StringVar，方便獲取和設置值
        
        ttk.Button(file_frame, text="瀏覽...", command=self.browse_file).grid(
            row=0, column=2, padx=5, pady=5
        )
        # command: 按鈕點擊時調用的函數
        
        # ======== 命名參數區 ========
        params_frame = ttk.LabelFrame(frame, text="命名參數")
        params_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 主分類選擇
        ttk.Label(params_frame, text="主分類:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.category_combo = ttk.Combobox(
            params_frame, 
            textvariable=self.category_var,
            state="readonly",  # 只能選擇，不能輸入
            width=15
        )
        self.category_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        # 設置選擇事件 (當選擇改變時更新預覽)
        self.category_combo.bind("<<ComboboxSelected>>", lambda e: self.preview_filename())
        
        # 細分類選擇
        ttk.Label(params_frame, text="細分類:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.subcategory_combo = ttk.Combobox(
            params_frame, 
            textvariable=self.subcategory_var,
            state="readonly",
            width=15
        )
        self.subcategory_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.subcategory_combo.bind("<<ComboboxSelected>>", lambda e: self.preview_filename())
        
        # 活動選擇
        ttk.Label(params_frame, text="活動:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.activity_combo = ttk.Combobox(
            params_frame, 
            textvariable=self.activity_var,
            state="readonly",
            width=15
        )
        self.activity_combo.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        self.activity_combo.bind("<<ComboboxSelected>>", lambda e: self.preview_filename())
        
        # 描述輸入
        ttk.Label(params_frame, text="描述:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        description_entry = ttk.Entry(params_frame, textvariable=self.description_var, width=20)
        description_entry.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)
        # 綁定鍵盤事件以即時更新預覽
        description_entry.bind("<KeyRelease>", lambda e: self.preview_filename())
        
        # 日期輸入
        ttk.Label(params_frame, text="日期 (YYYY-MM-DD):").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        date_entry = ttk.Entry(params_frame, textvariable=self.date_var, width=15)
        date_entry.grid(row=4, column=1, sticky=tk.W, padx=5, pady=5)
        date_entry.bind("<KeyRelease>", lambda e: self.preview_filename())
        
        # ======== 預覽區 ========
        preview_frame = ttk.LabelFrame(frame, text="預覽")
        preview_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(preview_frame, text="新檔名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(
            preview_frame, 
            textvariable=self.preview_var,
            state="readonly",  # 只讀，防止手動修改
            width=60
        ).grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        # columnspan=2: 佔用2個列位置
        
        # 按鈕區域
        ttk.Button(preview_frame, text="預覽", command=self.preview_filename).grid(
            row=1, column=1, sticky=tk.E, padx=5, pady=5
        )
        ttk.Button(preview_frame, text="重命名", command=self.rename_file).grid(
            row=1, column=2, sticky=tk.E, padx=5, pady=5
        )
    
    def setup_batch_tab(self):
        """設置批次重命名頁籤的UI元素"""
        frame = ttk.LabelFrame(self.batch_frame, text="批次檔案重命名")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ======== 資料夾選擇區 ========
        folder_frame = ttk.Frame(frame)
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(folder_frame, text="選擇資料夾:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(folder_frame, textvariable=self.folder_path_var, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(folder_frame, text="瀏覽...", command=self.browse_folder).grid(
            row=0, column=2, padx=5, pady=5
        )
        ttk.Button(folder_frame, text="載入檔案", command=self.load_folder_files).grid(
            row=0, column=3, padx=5, pady=5
        )
        
        # ======== 檔案列表區 ========
        list_frame = ttk.LabelFrame(frame, text="檔案列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 建立樹狀檢視表格 (支持多列顯示的列表)
        columns = ("原始檔名", "主分類", "細分類", "活動", "描述", "日期", "新檔名")
        self.file_table = ttk.Treeview(
            list_frame, 
            columns=columns,  # 定義列
            show="headings",  # 只顯示標題和資料，不顯示第一列的樹狀結構
            selectmode="browse"  # 只允許單選
        )
        
        # 設定列標題和寬度
        for col in columns:
            self.file_table.heading(col, text=col)  # 設置列標題
            self.file_table.column(col, width=100)  # 設置列寬度
        
        # 添加垂直捲動條
        v_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_table.yview)
        self.file_table.configure(yscrollcommand=v_scrollbar.set)  # 連結表格和捲動條
        
        # 添加水平捲動條
        h_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=self.file_table.xview)
        self.file_table.configure(xscrollcommand=h_scrollbar.set)  # 連結表格和捲動條
        
        # 排列表格和捲動條
        self.file_table.grid(row=0, column=0, sticky="nsew")  # nsew = 四個方向都填充
        v_scrollbar.grid(row=0, column=1, sticky="ns")  # ns = 垂直填充
        h_scrollbar.grid(row=1, column=0, sticky="ew")  # ew = 水平填充
        
        # 設置網格權重，使表格能隨視窗大小調整
        list_frame.columnconfigure(0, weight=1)  # 列0可擴展
        list_frame.rowconfigure(0, weight=1)  # 行0可擴展
        
        # 綁定選擇事件
        self.file_table.bind("<<TreeviewSelect>>", self.on_file_select)
        
        # ======== 批次命名參數區 ========
        batch_params_frame = ttk.LabelFrame(frame, text="批次命名參數")
        batch_params_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 主分類
        ttk.Label(batch_params_frame, text="主分類:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.batch_category_combo = ttk.Combobox(
            batch_params_frame,
            state="readonly",
            width=15
        )
        self.batch_category_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 細分類
        ttk.Label(batch_params_frame, text="細分類:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        self.batch_subcategory_combo = ttk.Combobox(
            batch_params_frame,
            state="readonly",
            width=15
        )
        self.batch_subcategory_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        # 活動
        ttk.Label(batch_params_frame, text="活動:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.batch_activity_combo = ttk.Combobox(
            batch_params_frame,
            state="readonly",
            width=15
        )
        self.batch_activity_combo.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 描述前綴
        ttk.Label(batch_params_frame, text="描述前綴:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
        self.batch_prefix_var = tk.StringVar()
        ttk.Entry(batch_params_frame, textvariable=self.batch_prefix_var, width=15).grid(
            row=1, column=3, sticky=tk.W, padx=5, pady=5
        )
        
        # 日期
        ttk.Label(batch_params_frame, text="日期:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.batch_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(batch_params_frame, textvariable=self.batch_date_var, width=15).grid(
            row=2, column=1, sticky=tk.W, padx=5, pady=5
        )
        
        # ======== 操作按鈕區 ========
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Button(button_frame, text="應用到所有檔案", command=self.apply_to_all_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="應用到選中檔案", command=self.apply_to_selected_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="預覽變更", command=self.preview_batch_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="執行批次重命名", command=self.execute_batch_rename).pack(side=tk.RIGHT, padx=5)
    
    def setup_photo_tab(self):
        """設置照片資料夾頁籤的UI元素"""
        frame = ttk.LabelFrame(self.photo_frame, text="照片資料夾管理")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ======== 基礎目錄選擇 ========
        folder_frame = ttk.Frame(frame)
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(folder_frame, text="基礎目錄:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(folder_frame, textvariable=self.photo_folder_var, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(folder_frame, text="瀏覽...", command=self.browse_photo_folder).grid(
            row=0, column=2, padx=5, pady=5
        )
        
        # ======== 照片資料夾參數 ========
        params_frame = ttk.LabelFrame(frame, text="資料夾命名參數")
        params_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 活動選擇
        ttk.Label(params_frame, text="活動:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.photo_activity_combo = ttk.Combobox(
            params_frame,
            state="readonly",
            width=15
        )
        self.photo_activity_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        # 日期
        ttk.Label(params_frame, text="日期:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.photo_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(params_frame, textvariable=self.photo_date_var, width=15).grid(
            row=1, column=1, sticky=tk.W, padx=5, pady=5
        )
        
        # 預覽
        ttk.Label(params_frame, text="資料夾名稱預覽:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.photo_preview_var = tk.StringVar()
        ttk.Entry(params_frame, textvariable=self.photo_preview_var, state="readonly", width=50).grid(
            row=2, column=1, columnspan=3, sticky=tk.W+tk.E, padx=5, pady=5
        )
        
        # 預覽按鈕
        ttk.Button(params_frame, text="預覽", command=self.preview_photo_folder).grid(
            row=3, column=1, sticky=tk.W, padx=5, pady=5
        )
        
        # 創建資料夾按鈕
        ttk.Button(params_frame, text="創建照片資料夾", command=self.create_photo_folder).grid(
            row=3, column=2, sticky=tk.E, padx=5, pady=5
        )
        
        # ======== 現有照片資料夾列表 ========
        folders_frame = ttk.LabelFrame(frame, text="現有照片資料夾")
        folders_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 資料夾列表
        self.photo_folders_listbox = tk.Listbox(folders_frame, height=8)
        scrollbar = ttk.Scrollbar(folders_frame, orient="vertical", command=self.photo_folders_listbox.yview)
        self.photo_folders_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.photo_folders_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # 刷新按鈕
        ttk.Button(folders_frame, text="刷新列表", command=self.refresh_photo_folders).pack(
            side=tk.BOTTOM, padx=5, pady=5
        )
    
    def setup_settings_tab(self):
        """設置設定頁籤的UI元素"""
        frame = ttk.LabelFrame(self.settings_frame, text="系統設定")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ======== 命名對照表設定 ========
        config_frame = ttk.LabelFrame(frame, text="命名對照表設定")
        config_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 使用本地Excel
        ttk.Radiobutton(
            config_frame, 
            text="使用本地Excel檔案", 
            variable=self.use_google_sheet, 
            value=False,
            command=self.toggle_config_source
        ).grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(config_frame, text="Excel檔案路徑:").grid(row=1, column=0, sticky=tk.W, padx=20, pady=5)
        ttk.Entry(config_frame, textvariable=self.excel_path, width=50).grid(
            row=1, column=1, padx=5, pady=5, sticky=tk.W+tk.E
        )
        ttk.Button(config_frame, text="瀏覽...", command=self.browse_excel_file).grid(
            row=1, column=2, padx=5, pady=5
        )
        
        # 使用Google Sheet
        ttk.Radiobutton(
            config_frame, 
            text="使用Google Sheet", 
            variable=self.use_google_sheet, 
            value=True,
            command=self.toggle_config_source
        ).grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(config_frame, text="Google Sheet ID:").grid(row=3, column=0, sticky=tk.W, padx=20, pady=5)
        ttk.Entry(config_frame, textvariable=self.google_sheet_id, width=50).grid(
            row=3, column=1, padx=5, pady=5, sticky=tk.W+tk.E
        )
        
        ttk.Label(config_frame, text="認證檔案路徑:").grid(row=4, column=0, sticky=tk.W, padx=20, pady=5)
        ttk.Entry(config_frame, textvariable=self.credentials_file, width=50).grid(
            row=4, column=1, padx=5, pady=5, sticky=tk.W+tk.E
        )
        ttk.Button(config_frame, text="瀏覽...", command=self.browse_credentials_file).grid(
            row=4, column=2, padx=5, pady=5
        )
        
        # 設置列的權重
        config_frame.columnconfigure(1, weight=1)
        
        # ======== 載入和匯出按鈕 ========
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Button(button_frame, text="載入配置", command=self.load_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="導出模板", command=self.export_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="儲存設定", command=self.save_settings).pack(side=tk.RIGHT, padx=5)
        
        # ======== 日誌查看 ========
        log_frame = ttk.LabelFrame(frame, text="系統日誌")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)  # 設為只讀
        
        # 添加日誌處理器，將日誌輸出到文本框
        self.setup_log_handler()
    
    def setup_search_tab(self):
        """設置檔案搜尋頁籤的UI元素"""
        frame = ttk.LabelFrame(self.search_frame, text="檔案搜尋")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ======== 搜尋目錄選擇 ========
        folder_frame = ttk.Frame(frame)
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(folder_frame, text="搜尋目錄:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(folder_frame, textvariable=self.search_folder_var, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(folder_frame, text="瀏覽...", command=self.browse_search_folder).grid(
            row=0, column=2, padx=5, pady=5
        )
        
        # ======== 搜尋模式 ========
        pattern_frame = ttk.Frame(frame)
        pattern_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(pattern_frame, text="搜尋模式(正則表達式):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(pattern_frame, textvariable=self.search_pattern_var, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(pattern_frame, text="搜尋", command=self.search_files).grid(
            row=0, column=2, padx=5, pady=5
        )
        
        # ======== 快速搜尋範本 ========
        template_frame = ttk.LabelFrame(frame, text="快速搜尋範本")
        template_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 第一行範本
        ttk.Button(template_frame, text="所有企劃書", 
                  command=lambda: self.search_pattern_var.set(r'^B_PL.*\.docx$')).grid(
            row=0, column=0, padx=5, pady=5
        )
        ttk.Button(template_frame, text="所有照片資料夾", 
                  command=lambda: self.search_pattern_var.set(r'^B_PH.*')).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(template_frame, text="哈盆活動檔案", 
                  command=lambda: self.search_pattern_var.set(r'.*_HP_.*')).grid(
            row=0, column=2, padx=5, pady=5
        )
        
        # 第二行範本
        ttk.Button(template_frame, text="今年檔案", 
                  command=lambda: self.search_pattern_var.set(
                      rf'.*_{datetime.now().year}-.*')).grid(
            row=1, column=0, padx=5, pady=5
        )
        ttk.Button(template_frame, text="會議記錄", 
                  command=lambda: self.search_pattern_var.set(r'^D_MT.*')).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Button(template_frame, text="報帳單", 
                  command=lambda: self.search_pattern_var.set(r'^C_RC.*')).grid(
            row=1, column=2, padx=5, pady=5
        )
        
        # ======== 搜尋結果 ========
        results_frame = ttk.LabelFrame(frame, text="搜尋結果")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 結果列表
        self.search_results = scrolledtext.ScrolledText(results_frame, height=10, width=80, wrap=tk.WORD)
        self.search_results.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 添加右鍵選單
        self.search_result_menu = tk.Menu(self.search_results, tearoff=0)
        self.search_result_menu.add_command(label="打開檔案", command=self.open_selected_file)
        self.search_result_menu.add_command(label="打開資料夾", command=self.open_containing_folder)
        self.search_result_menu.add_separator()
        self.search_result_menu.add_command(label="複製路徑", command=self.copy_file_path)
        
        # 綁定右鍵選單
        self.search_results.bind("<Button-3>", self.show_search_result_menu)
        # 綁定雙擊事件
        self.search_results.bind("<Double-Button-1>", lambda e: self.open_selected_file())
    
    # ===== 檔案處理方法 =====
    
    def browse_file(self):
        """瀏覽並選擇單個檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇檔案",
            filetypes=[("所有檔案", "*.*")]
        )
        if file_path:  # 如果用戶選擇了檔案（未取消對話框）
            self.file_path_var.set(file_path)
            # 嘗試解析檔案名，如果是標準格式則自動填充表單
            self.try_parse_filename(file_path)
    
    def try_parse_filename(self, file_path):
        """嘗試解析檔案名，如果是標準格式則自動填充表單"""
        if not self.file_manager:
            return
            
        # 獲取檔案名
        filename = os.path.basename(file_path)
        # 解析檔案名
        file_info = self.file_manager.parse_filename(filename)
        
        if file_info and file_info.get('is_standard', False):
            # 如果是標準格式，自動填充表單
            if file_info.get('main_category'):
                self.category_var.set(file_info['main_category'])
            if file_info.get('subcategory'):
                self.subcategory_var.set(file_info['subcategory'])
            if file_info.get('activity'):
                self.activity_var.set(file_info['activity'])
            if file_info.get('description'):
                self.description_var.set(file_info['description'])
            if file_info.get('date'):
                self.date_var.set(file_info['date'])
                
            # 更新預覽
            self.preview_filename()
    
    def browse_folder(self):
        """瀏覽並選擇資料夾（用於批次重命名）"""
        folder_path = filedialog.askdirectory(title="選擇資料夾")
        if folder_path:
            self.folder_path_var.set(folder_path)
    
    def browse_search_folder(self):
        """瀏覽並選擇搜尋資料夾"""
        folder_path = filedialog.askdirectory(title="選擇搜尋目錄")
        if folder_path:
            self.search_folder_var.set(folder_path)
    
    def browse_photo_folder(self):
        """瀏覽並選擇照片基礎資料夾"""
        folder_path = filedialog.askdirectory(title="選擇照片基礎目錄")
        if folder_path:
            self.photo_folder_var.set(folder_path)
            # 刷新照片資料夾列表
            self.refresh_photo_folders()
    
    def browse_excel_file(self):
        """瀏覽並選擇Excel命名對照表"""
        file_path = filedialog.askopenfilename(
            title="選擇Excel命名對照表",
            filetypes=[("Excel檔案", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path.set(file_path)
    
    def browse_credentials_file(self):
        """瀏覽並選擇Google API認證檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇Google API認證檔案",
            filetypes=[("JSON檔案", "*.json")]
        )
        if file_path:
            self.credentials_file.set(file_path)
    
    # ===== 配置管理方法 =====
    
    def try_load_default_config(self):
        """嘗試載入默認配置"""
        # 檢查默認Excel配置文件
        default_excel = "config/example_config.xlsx"
        if os.path.exists(default_excel):
            self.excel_path.set(default_excel)
            self.load_config()
        else:
            self.logger.warning(f"未找到默認配置文件: {default_excel}")
            # 可以考慮創建默認模板
            try:
                os.makedirs("config", exist_ok=True)
                self.file_manager = ClubFileManager()  # 創建臨時實例
                self.file_manager.export_naming_rules_template(default_excel)
                self.excel_path.set(default_excel)
                self.load_config()
                messagebox.showinfo("提示", f"已創建並載入默認配置模板: {default_excel}")
            except Exception as e:
                self.logger.error(f"創建默認配置模板失敗: {str(e)}")
    
    def toggle_config_source(self):
        """切換配置來源（本地Excel或Google Sheet）"""
        is_google = self.use_google_sheet.get()
        self.logger.info(f"切換配置來源: {'Google Sheet' if is_google else '本地Excel'}")
        # 根據選擇啟用/禁用相關控件（可以在這裡實現）
    
    def load_config(self):
        """載入命名對照表配置"""
        try:
            # 檢查是否已存在file_manager實例
            if not self.file_manager:
                self.file_manager = ClubFileManager()
            
            # 根據使用者選擇載入配置
            use_google = self.use_google_sheet.get()
            
            if use_google:
                # 使用Google Sheet
                sheet_id = self.google_sheet_id.get()
                cred_file = self.credentials_file.get()
                
                if not sheet_id or not cred_file:
                    messagebox.showerror("錯誤", "請提供Google Sheet ID和認證檔案路徑")
                    return
                
                if not os.path.exists(cred_file):
                    messagebox.showerror("錯誤", f"找不到認證檔案: {cred_file}")
                    return
                
                success = self.file_manager.connect_to_google_sheet(cred_file, sheet_id)
                if not success:
                    messagebox.showerror("錯誤", "連接Google Sheet失敗，請檢查日誌")
                    return
            else:
                # 使用本地Excel
                excel_path = self.excel_path.get()
                
                if not excel_path:
                    messagebox.showerror("錯誤", "請提供Excel檔案路徑")
                    return
                
                if not os.path.exists(excel_path):
                    messagebox.showerror("錯誤", f"找不到Excel檔案: {excel_path}")
                    return
                
                success = self.file_manager.load_from_excel(excel_path)
                if not success:
                    messagebox.showerror("錯誤", "載入Excel檔案失敗，請檢查日誌")
                    return
            
            # 更新下拉選單
            self.update_combo_boxes()
            
            # 更新狀態
            self.config_loaded = True
            self.status_var.set("命名對照表載入成功")
            messagebox.showinfo("成功", "命名對照表載入成功")
            
        except Exception as e:
            self.logger.error(f"載入配置時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"載入配置時出錯: {str(e)}")
    
    def update_combo_boxes(self):
        """更新所有下拉選單的選項"""
        if not self.file_manager:
            return
            
        # 獲取選項列表
        main_categories = self.file_manager.get_main_categories()
        sub_categories = self.file_manager.get_sub_categories()
        activities = self.file_manager.get_activities()
        
        # 更新單檔重命名頁籤的下拉選單
        self.category_combo['values'] = main_categories
        self.subcategory_combo['values'] = sub_categories
        self.activity_combo['values'] = activities
        
        # 更新批次重命名頁籤的下拉選單
        self.batch_category_combo['values'] = main_categories
        self.batch_subcategory_combo['values'] = sub_categories
        self.batch_activity_combo['values'] = activities
        
        # 更新照片資料夾頁籤的下拉選單
        self.photo_activity_combo['values'] = activities
        
        # 如果有選項，預設選擇第一個
        if main_categories:
            self.category_var.set(main_categories[0])
            self.batch_category_combo.set(main_categories[0])
        if sub_categories:
            self.subcategory_var.set(sub_categories[0])
            self.batch_subcategory_combo.set(sub_categories[0])
        if activities:
            self.activity_var.set(activities[0])
            self.batch_activity_combo.set(activities[0])
            self.photo_activity_combo.set(activities[0])
    
    def export_template(self):
        """導出命名規則模板"""
        try:
            # 選擇保存位置
            file_path = filedialog.asksaveasfilename(
                title="保存命名規則模板",
                defaultextension=".xlsx",
                filetypes=[("Excel檔案", "*.xlsx")]
            )
            
            if not file_path:
                return  # 用戶取消了選擇
            
            # 創建臨時ClubFileManager實例（如果尚未創建）
            if not self.file_manager:
                self.file_manager = ClubFileManager()
            
            # 導出模板
            success = self.file_manager.export_naming_rules_template(file_path)
            
            if success:
                self.logger.info(f"命名規則模板已導出至: {file_path}")
                messagebox.showinfo("成功", f"命名規則模板已導出至:\n{file_path}")
            else:
                messagebox.showerror("錯誤", "導出模板失敗，請檢查日誌")
                
        except Exception as e:
            self.logger.error(f"導出模板時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"導出模板時出錯: {str(e)}")
    
    def save_settings(self):
        """保存設定到配置文件"""
        try:
            # 創建配置目錄（如果不存在）
            os.makedirs("config", exist_ok=True)
            
            # 準備設定數據
            settings = {
                "use_google_sheet": str(self.use_google_sheet.get()),
                "excel_path": self.excel_path.get(),
                "google_sheet_id": self.google_sheet_id.get(),
                "credentials_file": self.credentials_file.get()
            }
            
            # 寫入設定文件
            with open("config/settings.ini", "w", encoding="utf-8") as f:
                for key, value in settings.items():
                    f.write(f"{key}={value}\n")
            
            self.logger.info("設定已保存")
            messagebox.showinfo("成功", "設定已保存")
            
        except Exception as e:
            self.logger.error(f"保存設定時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"保存設定時出錯: {str(e)}")
    
    def setup_log_handler(self):
        """設置日誌處理器，將日誌輸出到GUI文本框"""
        # 創建自定義日誌處理器
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget
                
            def emit(self, record):
                msg = self.format(record)
                
                def append():
                    self.text_widget.configure(state=tk.NORMAL)
                    self.text_widget.insert(tk.END, msg + "\n")
                    self.text_widget.configure(state=tk.DISABLED)
                    self.text_widget.see(tk.END)  # 自動捲動到最新日誌
                
                # 使用主線程更新GUI
                self.text_widget.after(0, append)
        
        # 創建日誌處理器
        text_handler = TextHandler(self.log_text)
        text_handler.setLevel(logging.INFO)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        
        # 添加處理器到根日誌器
        root_logger = logging.getLogger()
        root_logger.addHandler(text_handler)
    
    # ===== 功能實現方法 =====
    
    def preview_filename(self):
        """預覽生成的檔案名稱"""
        if not self.file_manager:
            messagebox.showerror("錯誤", "請先載入命名對照表")
            return
            
        # 獲取參數
        category = self.category_var.get()
        subcategory = self.subcategory_var.get()
        activity = self.activity_var.get()
        description = self.description_var.get()
        date = self.date_var.get()
        
        # 檢查必要參數
        if not category or not subcategory or not activity or not description:
            self.preview_var.set("")
            return
        
        # 生成檔案名
        try:
            filename = self.file_manager.generate_filename(
                category, subcategory, activity, description, date
            )
            
            # 獲取原始檔案的副檔名
            file_path = self.file_path_var.get()
            if file_path:
                ext = os.path.splitext(file_path)[1]
                filename += ext
            
            # 更新預覽
            self.preview_var.set(filename)
            
        except Exception as e:
            self.logger.error(f"預覽檔案名時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"預覽檔案名時出錯: {str(e)}")
    
    def rename_file(self):
        """重命名單一檔案"""
        if not self.file_manager:
            messagebox.showerror("錯誤", "請先載入命名對照表")
            return
            
        # 獲取檔案路徑
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("錯誤", "請選擇有效的檔案")
            return
        
        # 獲取參數
        category = self.category_var.get()
        subcategory = self.subcategory_var.get()
        activity = self.activity_var.get()
        description = self.description_var.get()
        date = self.date_var.get()
        
        # 檢查必要參數
        if not category or not subcategory or not activity or not description:
            messagebox.showerror("錯誤", "請填寫所有命名參數")
            return
        
        # 執行重命名
        try:
            new_path = self.file_manager.rename_file(
                file_path, category, subcategory, activity, description, date
            )
            
            if new_path:
                self.logger.info(f"檔案已重命名: {os.path.basename(file_path)} -> {os.path.basename(new_path)}")
                self.file_path_var.set(new_path)  # 更新檔案路徑
                messagebox.showinfo("成功", f"檔案已重命名為:\n{os.path.basename(new_path)}")
            else:
                messagebox.showerror("錯誤", "重命名失敗，請檢查日誌")
                
        except Exception as e:
            self.logger.error(f"重命名檔案時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"重命名檔案時出錯: {str(e)}")
    
    def load_folder_files(self):
        """載入資料夾中的檔案（用於批次重命名）"""
        folder_path = self.folder_path_var.get()
        
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("錯誤", "請選擇有效的資料夾")
            return
        
        try:
            # 清空現有檔案列表
            for item in self.file_table.get_children():
                self.file_table.delete(item)
            
            # 獲取資料夾中的檔案
            files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
            
            # 更新批次檔案列表
            self.batch_files = files
            
            # 填充表格
            for i, file in enumerate(files):
                # 嘗試解析檔案名（如果是標準格式）
                file_info = {}
                if self.file_manager:
                    file_info = self.file_manager.parse_filename(file) or {}
                
                # 檢查是否為照片格式
                ext = os.path.splitext(file)[1].lower()
                is_photo = ext in ['.jpg', '.jpeg', '.png', '.heic', '.gif']
                
                if is_photo:
                    status = "照片檔案（將不重命名）"
                elif file_info.get('is_standard', False):
                    status = "已符合標準格式"
                else:
                    status = "待重命名"
                
                # 添加到表格
                values = [
                    file,
                    file_info.get('main_category', ''),
                    file_info.get('subcategory', ''),
                    file_info.get('activity', ''),
                    file_info.get('description', ''),
                    file_info.get('date', ''),
                    status
                ]
                
                self.file_table.insert('', 'end', text=str(i), values=values)
            
            self.logger.info(f"已載入資料夾 {folder_path} 中的 {len(files)} 個檔案")
            self.status_var.set(f"已載入 {len(files)} 個檔案")
            
        except Exception as e:
            self.logger.error(f"載入資料夾檔案時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"載入資料夾檔案時出錯: {str(e)}")
    
    def on_file_select(self, event):
        """處理檔案表格中的選擇事件"""
        # 獲取選中項目
        selection = self.file_table.selection()
        if not selection:
            return
            
        # 獲取選中項目的值
        item = selection[0]
        values = self.file_table.item(item, 'values')
        
        # 如果是標準格式檔案，填充資料
        if values[1]:  # 主分類值非空
            if values[1] in self.batch_category_combo['values']:
                self.batch_category_combo.set(values[1])
            if values[2] in self.batch_subcategory_combo['values']:
                self.batch_subcategory_combo.set(values[2])
            if values[3] in self.batch_activity_combo['values']:
                self.batch_activity_combo.set(values[3])
            if values[4]:
                self.batch_prefix_var.set(values[4])
            if values[5]:
                self.batch_date_var.set(values[5])
    
    def apply_to_selected_file(self):
        """將命名參數應用到選中的檔案"""
        # 獲取選中項目
        selection = self.file_table.selection()
        if not selection:
            messagebox.showinfo("提示", "請先選擇檔案")
            return
            
        # 獲取命名參數
        category = self.batch_category_combo.get()
        subcategory = self.batch_subcategory_combo.get()
        activity = self.batch_activity_combo.get()
        prefix = self.batch_prefix_var.get()
        date = self.batch_date_var.get()
        
        # 檢查參數
        if not category or not subcategory or not activity:
            messagebox.showerror("錯誤", "請選擇主分類、細分類和活動")
            return
        
        # 更新選中項目
        for item in selection:
            values = list(self.file_table.item(item, 'values'))
            
            # 檢查是否為照片
            filename = values[0]
            ext = os.path.splitext(filename)[1].lower()
            if ext in ['.jpg', '.jpeg', '.png', '.heic', '.gif']:
                continue  # 跳過照片
            
            # 更新參數
            values[1] = category
            values[2] = subcategory
            values[3] = activity
            values[4] = prefix
            values[5] = date
            
            # 生成新檔名預覽
            if self.file_manager:
                try:
                    new_filename = self.file_manager.generate_filename(
                        category, subcategory, activity, f"{prefix}_{item}", date
                    )
                    ext = os.path.splitext(filename)[1]
                    values[6] = new_filename + ext
                except Exception as e:
                    self.logger.error(f"生成預覽檔名時出錯: {str(e)}")
                    values[6] = "預覽失敗"
            
            # 更新表格
            self.file_table.item(item, values=values)
    
    def apply_to_all_files(self):
        """將命名參數應用到所有檔案"""
        # 獲取所有項目
        all_items = self.file_table.get_children()
        if not all_items:
            messagebox.showinfo("提示", "表格中沒有檔案")
            return
        
        # 選中所有項目
        self.file_table.selection_set(all_items)
        
        # 調用應用到選中檔案的方法
        self.apply_to_selected_file()
    
    def preview_batch_changes(self):
        """預覽批次重命名的變更"""
        # 獲取所有項目
        all_items = self.file_table.get_children()
        if not all_items:
            messagebox.showinfo("提示", "表格中沒有檔案")
            return
        
        # 計算變更數量
        change_count = 0
        for item in all_items:
            values = self.file_table.item(item, 'values')
            if values[6] and values[6] != values[0] and values[6] != "照片檔案（將不重命名）":
                change_count += 1
        
        # 顯示預覽訊息
        if change_count > 0:
            messagebox.showinfo("批次重命名預覽", f"將重命名 {change_count} 個檔案\n\n請查看表格中的「新檔名」列以預覽變更")
        else:
            messagebox.showinfo("批次重命名預覽", "沒有檔案需要重命名")
    
    def execute_batch_rename(self):
        """執行批次重命名"""
        if not self.file_manager:
            messagebox.showerror("錯誤", "請先載入命名對照表")
            return
        
        # 獲取資料夾路徑
        folder_path = self.folder_path_var.get()
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("錯誤", "請選擇有效的資料夾")
            return
        
        # 收集重命名信息
        rename_info = []
        for item in self.file_table.get_children():
            values = self.file_table.item(item, 'values')
            filename = values[0]
            
            # 跳過照片或已符合標準格式的檔案
            ext = os.path.splitext(filename)[1].lower()
            if ext in ['.jpg', '.jpeg', '.png', '.heic', '.gif']:
                continue
            
            # 如果有完整的命名參數，加入重命名列表
            if all([values[1], values[2], values[3], values[4]]):
                rename_info.append([
                    filename, 
                    values[1],  # 主分類
                    values[2],  # 細分類
                    values[3],  # 活動
                    values[4],  # 描述
                    values[5]   # 日期
                ])
        
        # 檢查是否有檔案要重命名
        if not rename_info:
            messagebox.showinfo("提示", "沒有檔案需要重命名")
            return
        
        # 確認操作
        confirm = messagebox.askyesno(
            "確認操作", 
            f"確定要重命名 {len(rename_info)} 個檔案嗎？\n此操作無法撤銷。"
        )
        if not confirm:
            return
        
        try:
            # 執行批次重命名
            results = self.file_manager.batch_rename(folder_path, rename_info)
            
            # 統計結果
            success_count = sum(1 for v in results.values() if v is not None)
            fail_count = len(results) - success_count
            
            # 顯示結果
            self.logger.info(f"批次重命名完成: {success_count} 成功, {fail_count} 失敗")
            messagebox.showinfo("完成", f"批次重命名完成\n成功: {success_count}\n失敗: {fail_count}")
            
            # 重新載入資料夾檔案
            self.load_folder_files()
            
        except Exception as e:
            self.logger.error(f"批次重命名時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"批次重命名時出錯: {str(e)}")
    
    def preview_photo_folder(self):
        """預覽照片資料夾名稱"""
        if not self.file_manager:
            messagebox.showerror("錯誤", "請先載入命名對照表")
            return
        
        # 獲取參數
        activity = self.photo_activity_combo.get()
        date = self.photo_date_var.get()
        
        if not activity:
            messagebox.showerror("錯誤", "請選擇活動")
            return
        
        try:
            # 生成資料夾名
            activity_code = self.file_manager.activity_map.get(activity, activity)
            folder_name = f"B_PH_{activity_code}_{activity}照片_{date}"
            
            # 更新預覽
            self.photo_preview_var.set(folder_name)
            
        except Exception as e:
            self.logger.error(f"預覽照片資料夾名時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"預覽照片資料夾名時出錯: {str(e)}")
    
    def create_photo_folder(self):
        """創建照片資料夾"""
        if not self.file_manager:
            messagebox.showerror("錯誤", "請先載入命名對照表")
            return
        
        # 獲取基礎目錄
        base_dir = self.photo_folder_var.get()
        if not base_dir or not os.path.isdir(base_dir):
            messagebox.showerror("錯誤", "請選擇有效的基礎目錄")
            return
        
        # 獲取參數
        activity = self.photo_activity_combo.get()
        date = self.photo_date_var.get()
        
        if not activity:
            messagebox.showerror("錯誤", "請選擇活動")
            return
        
        try:
            # 創建照片資料夾
            folder_path = self.file_manager.create_photo_folder(base_dir, activity, date)
            
            if folder_path:
                self.logger.info(f"已創建照片資料夾: {os.path.basename(folder_path)}")
                messagebox.showinfo("成功", f"已創建照片資料夾:\n{os.path.basename(folder_path)}")
                
                # 刷新照片資料夾列表
                self.refresh_photo_folders()
            else:
                messagebox.showerror("錯誤", "創建照片資料夾失敗，請檢查日誌")
                
        except Exception as e:
            self.logger.error(f"創建照片資料夾時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"創建照片資料夾時出錯: {str(e)}")
    
    def refresh_photo_folders(self):
        """刷新照片資料夾列表"""
        # 清空列表
        self.photo_folders_listbox.delete(0, tk.END)
        
        # 獲取基礎目錄
        base_dir = self.photo_folder_var.get()
        if not base_dir or not os.path.isdir(base_dir):
            return
        
        try:
            # 搜尋照片資料夾
            photo_folders = []
            for item in os.listdir(base_dir):
                item_path = os.path.join(base_dir, item)
                if os.path.isdir(item_path) and re.match(r'^B_PH_.*照片_', item):
                    photo_folders.append(item)
            
            # 填充列表
            for folder in sorted(photo_folders):
                self.photo_folders_listbox.insert(tk.END, folder)
            
            self.logger.info(f"找到 {len(photo_folders)} 個照片資料夾")
            
        except Exception as e:
            self.logger.error(f"刷新照片資料夾列表時出錯: {str(e)}")
    
    def search_files(self):
        """搜尋檔案"""
        if not self.file_manager:
            messagebox.showerror("錯誤", "請先載入命名對照表")
            return
        
        # 獲取搜尋目錄
        search_dir = self.search_folder_var.get()
        if not search_dir or not os.path.isdir(search_dir):
            messagebox.showerror("錯誤", "請選擇有效的搜尋目錄")
            return
        
        # 獲取搜尋模式
        pattern = self.search_pattern_var.get()
        if not pattern:
            messagebox.showerror("錯誤", "請輸入搜尋模式")
            return
        
        try:
            # 執行搜尋
            matches = self.file_manager.search_files(search_dir, pattern)
            
            # 清空結果框
            self.search_results.config(state=tk.NORMAL)
            self.search_results.delete(1.0, tk.END)
            
            # 顯示結果
            if matches:
                self.search_results.insert(tk.END, f"找到 {len(matches)} 個匹配項：\n\n")
                for i, match in enumerate(matches, 1):
                    full_path = os.path.join(search_dir, match)
                    self.search_results.insert(tk.END, f"{i}. {match}\n")
                    # 儲存完整路徑為標籤，方便後續操作
                    self.search_results.tag_add(f"path_{i}", f"{i+1}.0", f"{i+2}.0")
                    self.search_results.tag_config(f"path_{i}", foreground="blue", underline=1)
                    self.search_results.tag_bind(f"path_{i}", "<Button-1>", 
                                              lambda e, path=full_path: self.open_file(path))
            else:
                self.search_results.insert(tk.END, "未找到符合模式的檔案。")
            
            self.search_results.config(state=tk.DISABLED)
            self.logger.info(f"搜尋完成, 找到 {len(matches)} 個匹配項")
            
        except Exception as e:
            self.logger.error(f"搜尋檔案時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"搜尋檔案時出錯: {str(e)}")
    
    def show_search_result_menu(self, event):
        """顯示搜尋結果右鍵選單"""
        # 獲取選中的文本行
        try:
            index = self.search_results.index(f"@{event.x},{event.y}")
            line = int(float(index))
            
            # 獲取該行的標籤
            for tag in self.search_results.tag_names(index):
                if tag.startswith("path_"):
                    # 記錄當前選中的路徑
                    self.current_selected_path = self.get_path_from_tag(tag)
                    # 顯示選單
                    self.search_result_menu.post(event.x_root, event.y_root)
                    break
        except Exception as e:
            self.logger.error(f"顯示選單時出錯: {str(e)}")
    
    def get_path_from_tag(self, tag):
        """從標籤獲取檔案路徑"""
        try:
            # 從標籤名稱中提取索引
            idx = int(tag.split("_")[1])
            # 獲取搜尋目錄
            search_dir = self.search_folder_var.get()
            # 獲取對應行的文本（去除行號和點）
            line_text = self.search_results.get(f"{idx+1}.3", f"{idx+1}.end")
            # 組合完整路徑
            return os.path.join(search_dir, line_text.strip())
        except Exception as e:
            self.logger.error(f"獲取路徑時出錯: {str(e)}")
            return None
    
    def open_selected_file(self):
        """打開選中的檔案"""
        if hasattr(self, 'current_selected_path') and self.current_selected_path:
            self.open_file(self.current_selected_path)
    
    def open_file(self, path):
        """打開檔案"""
        try:
            if os.path.exists(path):
                # 使用操作系統默認程式打開檔案
                if sys.platform == 'win32':
                    os.startfile(path)
                elif sys.platform == 'darwin':  # macOS
                    os.system(f'open "{path}"')
                else:  # Linux
                    os.system(f'xdg-open "{path}"')
                self.logger.info(f"已打開檔案: {path}")
            else:
                self.logger.warning(f"檔案不存在: {path}")
                messagebox.showerror("錯誤", f"檔案不存在: {path}")
        except Exception as e:
            self.logger.error(f"打開檔案時出錯: {str(e)}")
            messagebox.showerror("錯誤", f"打開檔案時出錯: {str(e)}")
    
    def open_containing_folder(self):
        """打開包含檔案的資料夾"""
        if hasattr(self, 'current_selected_path') and self.current_selected_path:
            try:
                path = self.current_selected_path
                folder = os.path.dirname(path)
                
                if os.path.exists(folder):
                    # 使用操作系統默認檔案瀏覽器打開資料夾
                    if sys.platform == 'win32':
                        os.startfile(folder)
                    elif sys.platform == 'darwin':  # macOS
                        os.system(f'open "{folder}"')
                    else:  # Linux
                        os.system(f'xdg-open "{folder}"')
                    self.logger.info(f"已打開資料夾: {folder}")
                else:
                    self.logger.warning(f"資料夾不存在: {folder}")
                    messagebox.showerror("錯誤", f"資料夾不存在: {folder}")
            except Exception as e:
                self.logger.error(f"打開資料夾時出錯: {str(e)}")
                messagebox.showerror("錯誤", f"打開資料夾時出錯: {str(e)}")
    
    def copy_file_path(self):
        """複製檔案路徑到剪貼簿"""
        if hasattr(self, 'current_selected_path') and self.current_selected_path:
            try:
                # 清除剪貼簿
                self.root.clipboard_clear()
                # 設置剪貼簿內容
                self.root.clipboard_append(self.current_selected_path)
                self.logger.info(f"已複製路徑到剪貼簿: {self.current_selected_path}")
            except Exception as e:
                self.logger.error(f"複製路徑時出錯: {str(e)}")
                messagebox.showerror("錯誤", f"複製路徑時出錯: {str(e)}")


# 主程式入口
def main():
    """主程式入口函數"""
    # 設置日誌
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('club_file_system.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    # 創建主視窗
    root = tk.Tk()
    app = ClubFileManagerGUI(root)
    
    # 進入事件迴圈
    root.mainloop()


# 當直接執行此模組時才運行主程式
if __name__ == "__main__":
    main()