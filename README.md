# 文字轉 PowerPoint 工具

將特定格式的文字自動轉換成 PowerPoint 簡報的工具集合。

## 📦 包含三個版本

### 1️⃣ Windows VBA 版本（推薦 Windows 用戶）
- **檔案**: `TextToPPT_VBA.bas`
- **優點**: 完整功能，支援剪貼簿和檔案輸入
- **缺點**: 僅限 Windows

### 2️⃣ Mac VBA 版本（Mac 用戶專用）
- **檔案**: `TextToPPT_VBA_Mac.bas`
- **優點**: Mac 相容，無需安裝其他軟體
- **缺點**: 只支援檔案輸入（不支援剪貼簿）

### 3️⃣ Python 跨平台版本（推薦跨平台使用者）
- **檔案**: `text_to_ppt.py`
- **優點**: Mac/Windows/Linux 都可用，功能強大
- **缺點**: 需要安裝 Python 和 python-pptx

---

## 🎯 文字格式說明

```
##主題標題
這是主題頁面的內容

#內文標題
第一個重點
第二個重點
第三個重點

##另一個主題
繼續您的簡報
```

**格式規則**:
- `##` 開頭 → 建立**主題頁面**（藍色背景，大標題，置中）
- `#` 開頭 → 建立**內文頁面**（灰色背景，標題+條列內容）
- 一般文字 → 加入到前一張投影片的內容區

---

## 🚀 快速開始

### Windows 用戶

1. 開啟 PowerPoint
2. 按 `Alt+F11` 開啟 VBA 編輯器
3. 插入 → 模組
4. 複製 `TextToPPT_VBA.bas` 的內容並貼上
5. 複製 `範例輸入文字.txt` 的內容到剪貼簿
6. 按 `F5` 執行 `ConvertTextToPPT`

### Mac 用戶

**⚠️ 重要：Mac 版本只支援從檔案讀取**

1. 開啟 PowerPoint
2. 工具 → 巨集 → Visual Basic 編輯器
3. 插入 → 模組
4. 複製 `TextToPPT_VBA_Mac.bas` 的內容並貼上
5. 按 `▶️` 執行 `ConvertTextFileToPPT`（檔案讀取）
6. 選擇 `範例輸入文字.txt` 檔案

**其他方法：**
- `ConvertTextFromNotes`：從投影片備註欄讀取
- 詳細說明請參考 `Mac使用說明.txt`

### Python 用戶（Mac/Windows/Linux）

```bash
# 1. 安裝相依套件
pip install python-pptx

# 2. 執行轉換
python text_to_ppt.py 範例輸入文字.txt

# 或指定輸出檔名
python text_to_ppt.py 範例輸入文字.txt 我的簡報.pptx
```

---

## 📊 功能比較表

| 功能 | Windows VBA | Mac VBA | Python |
|------|-------------|---------|--------|
| **剪貼簿輸入** | ✅ | ❌ | ❌ |
| **檔案輸入** | ✅ | ✅ | ✅ |
| **對話框輸入** | ✅ | ✅ | ❌ |
| **命令列執行** | ❌ | ❌ | ✅ |
| **批次處理** | 🟡 | 🟡 | ✅ |
| **自訂樣式** | 🟡 程式碼 | 🟡 程式碼 | ✅ 程式碼 |
| **跨平台** | ❌ | ❌ | ✅ |

---

## 🎨 投影片樣式

### 主題頁面（`##`）
- **背景顏色**: 淺藍色 `RGB(230, 240, 255)`
- **標題字型**: 微軟正黑體 44pt 粗體
- **對齊方式**: 置中

### 內文頁面（`#`）
- **背景顏色**: 淺灰色 `RGB(245, 245, 245)`
- **標題字型**: 微軟正黑體 32pt 粗體
- **內文字型**: 微軟正黑體 18pt
- **條列符號**: 自動啟用

---

## ⚙️ 自訂設定

### VBA 版本
修改程式碼中的以下部分：

```vba
' 修改背景顏色
sld.Background.Fill.ForeColor.RGB = RGB(230, 240, 255)

' 修改字型
.Font.Name = "微軟正黑體"

' 修改字型大小
.Font.Size = 44
```

### Python 版本
修改 `text_to_ppt.py` 的 `__init__` 方法：

```python
self.title_bg_color = RGBColor(230, 240, 255)  # 主題頁背景
self.content_bg_color = RGBColor(245, 245, 245)  # 內文頁背景
self.font_name = "微軟正黑體"  # 字型名稱
```

---

## 📁 檔案說明

| 檔案 | 說明 |
|------|------|
| `TextToPPT_VBA.bas` | Windows VBA 版本（完整功能） |
| `TextToPPT_VBA_Mac.bas` | Mac VBA 版本（相容性優化） |
| `text_to_ppt.py` | Python 跨平台版本 |
| `範例輸入文字.txt` | 範例文字檔案 |
| `使用說明.txt` | 詳細使用手冊（Windows 為主） |
| `Mac使用說明.txt` | **Mac 用戶專用說明** |
| `Python_安裝指南.txt` | Python 環境設定指南 |
| `README.md` | 本說明文件 |

---

## ❓ 常見問題

### Q: Mac 版為什麼不能從剪貼簿讀取？
**A**: Mac 版 Office 的 VBA 不支援 Windows 的剪貼簿 API。建議使用檔案輸入或改用 Python 版本。

### Q: Python 版本如何安裝？
**A**: 
```bash
# 檢查是否已安裝 Python
python --version  # 或 python3 --version

# 安裝 python-pptx
pip install python-pptx  # 或 pip3 install python-pptx
```

### Q: 可以使用現有的 PPT 模板嗎？
**A**: 
- **VBA 版本**: 先開啟您的模板檔案，再執行巨集，新投影片會加到最後
- **Python 版本**: 需要修改程式碼，載入現有模板：
  ```python
  self.prs = Presentation('您的模板.pptx')
  ```

### Q: 如何批次處理多個文字檔？
**A**: 使用 Python 版本最方便：
```bash
# Windows PowerShell
Get-ChildItem *.txt | ForEach-Object { python text_to_ppt.py $_.Name }

# Mac/Linux
for file in *.txt; do python text_to_ppt.py "$file"; done
```

### Q: 字型在不同電腦上會不會跑掉？
**A**: 可能會。建議：
- 使用常見字型（如 Arial、微軟正黑體）
- 或將字型嵌入到 PPT 中（檔案 → 選項 → 儲存 → 將字型嵌入檔案）

---

## 🔒 安全性提示

執行 VBA 巨集時，PowerPoint 可能會顯示安全性警告：

1. 檔案 → 選項 → 信任中心
2. 信任中心設定
3. 巨集設定 → 啟用所有巨集（建議僅在信任的文件中使用）

---

## 📞 建議使用情境

| 情境 | 推薦版本 |
|------|----------|
| 僅使用 Windows | Windows VBA |
| 僅使用 Mac | Mac VBA 或 Python |
| 兩個平台都用 | **Python**（最推薦） |
| 需要批次處理 | **Python** |
| 不想安裝額外軟體 | VBA |
| 需要進階自訂 | **Python** |

---

## 🎉 開始使用

1. 查看 `範例輸入文字.txt` 了解格式
2. 根據您的平台選擇對應版本
3. 按照「快速開始」的步驟操作
4. 享受自動化簡報製作的便利！

**祝您使用愉快！** 🚀
