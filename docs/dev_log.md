## 2024-05-02 開發日誌
🛠️ 昨日補記進度（2024-05-01）
- 建立專案結構與 create_project.py 腳本
- 建立 example_config.xlsx，規劃欄位設計（主分類、細分類、活動代碼等）
- 設定虛擬環境並安裝 pandas、openpyxl

✅ 今日完成事項
- 調整虛擬環境啟用相關設定（強化顯示提示等）
- 在 setting.json 設定 PowerShell 啟動時自動執行 activate.ps1
- 加入程式執行日誌功能
- 恢復並整理核心程式碼（club_file_manager.py / gui.py / main.py）
- 新增測試程式檔案：tests/test_manager.py

### 🧠 學到
- 如何判斷虛擬環境有無啟用
- avtivate.ps1 -> powershell 在看
- avtivate.bat -> cmd 在看
- 執行 git commit 時，不管你之前有沒有 push 到 GitHub，Git 都會：
     把目前 stage（已 add）區域 的檔案變動，打包成一個新的 commit
     這個 commit 會在你上次的 commit「後面」，成為下一版歷史
    不會覆蓋上次的 commit，而是會「追加在後面」

### 📎 相關 commit / PR
- commit: `70c1b05 建立專案骨架`
- commit: `c3d51e9 chore: 移除過多追蹤檔案`

### ⏭️ 下一步
- 了解code 在幹嘛
- 了解預期介面會怎樣
當然可以，以下是將你剛剛的表格內容整理成清楚的待辦清單（List 形式）：

---

## ✅ 近期待辦事項清單

### 🔥 高優先
- [ ] **修正 `ModuleNotFoundError` 問題**
  - 檢查 `gui.py` 中的 `import` 是否應改為相對路徑（如 `from .club_file_manager import ClubFileManager`）
  - 確認 `src/` 資料夾中有 `__init__.py`，使其成為 Python 套件
  - 確保 `main.py` 是從專案根目錄執行的（不要進入 `src/` 執行）

---

### 🔧 中優先
- [ ] **測試 GUI 是否可成功讀取 `example_config.xlsx`**
  - 確認下拉選單有讀到主分類、細分類、活動代碼
  - 測試選項變動是否能正確更新預覽命名結果

- [ ] **確認 CLI 模式下 `ClubFileManager` 是否可穩定產生命名**
  - 在 `main.py` 中手動呼叫 `.generate_filename()` 做輸出測試
  - 例：印出 `B_PL_HP_哈盆企劃書_2024-03-01`

---

### 📂 低優先
- [ ] **整理 `test_data/processed` 為命名後輸出區**
  - 將重命名完成的檔案集中放入 `/processed`
  - 方便日後回顧或自動化導出驗收

---

### 📜 低優先
- [ ] **撰寫 `README.md` 初版**
  - 專案目的、安裝方式、執行方式、使用範例

- [ ] **撰寫 `docs/dev_log.md` 日誌**
  - 每日完成紀錄、問題與修正過程、下一步規劃


### 長遠階段進程
- 第一階段：以 Excel / Google Sheet 為命名對照來源
- 第二階段（你期待的樣子）：設計圖形化命名介面




已完成項目
1. 專案基本架構設計

✅ 確定了專案名稱：「社團檔案標準化命名與分類系統」
✅ 規劃了基本的資料夾結構（src, config, docs, tests 等）
✅ 決定使用 VS Code 作為開發環境

2. 核心命名規則設計(不確定有無完成)

✅ 定義了標準化命名格式：[主分類代碼]_[細分類代碼]_[活動代碼]_[描述]_[日期].[副檔名]
✅ 例如：B_PL_HP_哈盆企劃書_2024-03-01.docx
✅ 決定照片檔案不進行重命名，但使用標準化資料夾進行管理

3. 資料管理結構設計(不確定有無完成)

✅ 設計了分類對照表結構（主分類、細分類、活動代碼）
✅ 選擇使用 Google Sheet 作為首選資料存儲和管理方式
✅ 規劃了 Regex 搜尋規則結構

4. 程式核心功能設計

✅ 設計了 ClubFileManager 核心類別，包含檔案命名、重命名等功能
✅ 設計了 GUI 界面基本結構，包含單檔重命名、批次處理等頁籤
✅ 確定與 Google API 的連接方式

5. 技術選型

✅ 確定使用 Python 作為主要開發語言
✅ 選擇了主要依賴庫：pandas, gspread, oauth2client, tkinter 等
✅ 確定使用虛擬環境進行開發環境管理

已生成的基本程式碼

✅ club_file_manager.py：核心檔案管理邏輯類
✅ file_renamer_gui.py：GUI 界面設計（部分完成）
✅ Google Sheet 模板設計：包含主分類表、細分類表、活動代碼表等

尚未完成的部分
1. 程式碼開發

❌ 完成 GUI 界面程式碼的剩餘部分
❌ 整合核心邏輯與界面
❌ 設計並實現配置文件讀取邏輯
❌ 完善錯誤處理機制

2. 測試與除錯

❌ 編寫單元測試
❌ 進行功能整合測試
❌ 邊緣案例測試（特殊字符、長檔名等）

3. 文檔與部署

❌ 完成詳細的使用者手冊
❌ 建立專案 README 文件
❌ 設計安裝腳本或打包為執行檔

4. 實際資料準備

❌ 建立實際的命名對照表內容
❌ 設置 Google API 憑證（如需要）
❌ 測試資料集準備

下一步優先任務
根據現有進度，建議優先完成這些任務：

完成核心功能開發：

完善 GUI 界面程式碼
整合核心邏輯與界面
實現 Google Sheet 連接功能


建立實際資料與配置：

創建實際的命名對照表（主分類、細分類、活動代碼）
設置必要的 API 憑證（如果計劃使用 Google API）


進行小規模測試：

準備測試資料
測試基本功能
調整與優化


完善文檔：

編寫簡明的使用說明
記錄設計決策和技術細節
準備未來的使用者培訓資料

# 2024-05-03 開發日誌
 
## 語法
 - python 'self.' == C++ 裡的 'this->'

## 問題
 - 解決ModuleNotFoundError問題
 ```
  | `from club_file_manager import ...`  | 
    ->從專案根目錄（絕對路徑）找 `club_file_manager.py` 

  | `from .club_file_manager import ...` | 
    -> 從目前這個模組所在資料夾中（也就是 `src/`）找 `club_file_manager.py`
```
- 遇到AttributeError: '' object has no attribute 'browse_file'
```
  加入 def browse_file(self):
```

# 2024-05-08 開發日誌

### 核心功能模組開發
- 重新建立 `ClubFileManager` 類 (src/club_file_manager.py)
  - 檔案命名生成功能
  - 檔案重命名功能
  - 批次檔案處理功能
  - 照片資料夾建立功能
  - 正規表示式檔案搜尋功能
- 處理 Google Sheet 與本地 Excel 的互通

### GUI 介面開發
- 重新建立 `ClubFileManagerGUI` 類 (src/gui.py)
  - 設計多標籤界面架構
  - 實現單檔重命名頁籤
  - 實現批次重命名頁籤
  - 實現照片資料夾頁籤
  - 實現檔案搜尋頁籤
  - 實現設定頁籤

### 基礎測試
- 建立基礎測試環境
- 建立 `test_manager.py` 
  - 專注於測試系統的核心功能邏輯
  - 不涉及使用者界面，僅測試背後的數據處理和檔案操作
  - 包含針對各個具體功能的測試案例，如檔案命名、重命名、搜尋等
- 建立 `test_gui.py`
  - 專注於測試界面元素和使用者互動
  - 使用模擬(mock)替代實際檔案對話框等
  - 測試界面元素的初始化、狀態變化、用戶操作響應等
- 建立 `run_tests.py`
  - 整合所有測試並提供統一運行方式
  - 支援命令行參數，可選擇性運行不同類型的測試
  - 配置測試環境，如日誌設置等

❌ 尚未測試檔案名稱生成功能

## 遇到問題
```
warning: LF will be replaced by CRLF the next time Git touches it
```
> 意思：行結尾符號（Line Ending）警告
> 正常現象，不會影響程式功能，也不是錯誤，只是 Git 在提醒你「有可能會出現行結尾差異」

### 解決方式
#### 在 .gitattributes 中統一行結尾
在專案根目錄下新增一個 .gitattributes 檔案，並加入以下內容：
```
* text=auto
```
這代表：
-  Git 自動根據作業系統決定使用哪一種行結尾
- 開發中使用系統習慣的格式，但提交時 Git 會統一處理

## ⏭️ 下一步
### 1. 測試：
* 單元測試：驗證核心功能是否正常工作
* 集成測試：確保GUI和核心邏輯協同工作

運行這些測試，您可以使用以下命令：
bash
#### 運行所有測試
 ``` 
python tests/run_tests.py
 ``` 
#### 只運行 ClubFileManager 測試
 ``` 
python tests/run_tests.py --type manager
 ``` 
####  只運行 GUI 測試
``` 
python tests/run_tests.py --type gui
``` 
####  在第一個測試失敗時停止
``` 
python tests/run_tests.py --failfast
``` 


== 您也可以使用 Python 的標準測試運行器：bash ==

####  運行特定測試文件
``` 
python -m unittest tests/test_manager.py
``` 
####  運行特定測試類
``` 
python -m unittest tests.test_manager.TestClubFileManager
``` 
####  運行特定測試方法
``` 
python -m unittest tests.test_manager.TestClubFileManager.test_generate_filename
``` 

這些測試可以幫助您確保系統的穩定性、可靠性，並在進行修改時避免意外引入錯誤。測試也提供了一種文檔形式，說明了各個功能的預期行為。

### 2. 安装脚本：
setup.py：使项目可安装
requirements.txt：列出依赖项

### 3. 打包：
使用PyInstaller将应用打包为可执行文件


> 1. 完善單元測試架構
> 2. 增加 GUI 測試
> 3. 進行整合測試
> 4. 添加錯誤處理和用戶反饋機制
> 5. 改進界面美觀性和用戶體驗
> 6. 考慮打包為可執行文件