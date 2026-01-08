# 文字轉 PowerPoint 工具

將特定格式的文字自動轉換成 PowerPoint 簡報的跨平台工具。

## 📦 多種工具版本

### 🖥️ GUI 版本（推薦一般使用者）
- **檔案**: `text_to_ppt_gui.py`
- **優點**: 圖形化介面，操作簡單直覺
- **功能**: 文字輸入或檔案選擇，即時預覽

### ⌨️ 命令列版本（推薦開發者）
- **檔案**: `text_to_ppt.py`
- **優點**: 適合批次處理和自動化
- **功能**: 命令列執行，支援腳本整合

### 📘 Word 轉 PPT 版本（新功能！）
- **檔案**: `docx_to_ppt.py`
- **優點**: 從 Word 文件提取藍色文字，一鍵轉換成 PPT
- **功能**: 自動識別藍色標記的重點內容
- **使用**: `python docx_to_ppt.py 文件.docx`

**系統需求**: Python 3.6+、python-pptx、python-docx

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

### 方法一：Word 轉 PPT（最新功能！）

#### 🎯 簡化流程（推薦）

```bash
# 1. 安裝相依套件
pip install python-pptx python-docx

# 2. 將 Word 文件命名為 input.docx（或使用任意檔名）

# 3. 提取藍色文字
python extract_blue_text_from_docx.py
# 或指定檔名：python extract_blue_text_from_docx.py 你的文件.docx

# 4. 編輯 output.txt 中的變數區塊（日期、主題、經文等）

# 5. 生成 PPT
python generate_ppt_from_template.py
# 輸出：output.pptx
```

**預設檔名：**
- 輸入 Word：`input.docx`
- 輸出文字：`output.txt`
- 輸出 PPT：`output.pptx`

#### 📝 一鍵版本（舊版，不含變數）

```bash
# 直接從 Word 轉 PPT（不含變數模板）
python docx_to_ppt.py 你的文件.docx

# 或使用預設檔名
python docx_to_ppt.py
```

### 方法二：使用 GUI 版本

```bash
# 1. 安裝相依套件
pip install python-pptx

# 2. 啟動 GUI
python text_to_ppt_gui.py
```

然後在視窗中：
1. 直接輸入文字或點擊「選擇檔案」載入 `examples/範例輸入文字.txt`
2. 點擊「轉換為 PPT」
3. 選擇儲存位置
4. 完成！

### 方法三：使用命令列版本

```bash
# 1. 安裝相依套件
pip install python-pptx

# 2. 執行轉換
python text_to_ppt.py examples/範例輸入文字.txt

# 或指定輸出檔名
python text_to_ppt.py examples/範例輸入文字.txt 我的簡報.pptx
```

### 方法四：快速測試腳本

```bash
# Mac/Linux 使用者
bash scripts/測試工具.sh

# 或使用快速開始腳本（適合 GitHub 部署）
bash scripts/快速開始.sh
```

---

## 📊 功能比較表

| 功能 | Word 轉 PPT | GUI 版本 | 命令列版本 |
|------|------------|---------|-----------|
| **Word 文件輸入** | ✅ | ❌ | ❌ |
| **藍色文字識別** | ✅ | ❌ | ❌ |
| **圖形化介面** | ❌ | ✅ | ❌ |
| **文字輸入** | ❌ | ✅ | ❌ |
| **檔案輸入** | ✅ | ✅ | ✅ |
| **命令列執行** | ✅ | 🟡 | ✅ |
| **批次處理** | ✅ | ❌ | ✅ |
| **腳本整合** | ✅ | 🟡 | ✅ |
| **跨平台** | ✅ | ✅ | ✅ |
| **使用難度** | 簡單 | 簡單 | 中等 |

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

## 📁 專案結構

```
text-to-ppt/
├── text_to_ppt.py                      # 命令列版本
├── text_to_ppt_gui.py                  # GUI 版本
├── docx_to_ppt.py                      # Word 轉 PPT（一鍵版本）
├── extract_blue_text_from_docx.py      # Word 藍色文字提取器
├── requirements.txt                    # Python 套件需求
├── README.md                           # 本說明文件
│
├── docs/                               # 📚 說明文件
│   ├── 使用說明.txt
│   ├── Python_安裝指南.txt
│   ├── GitHub_Actions_打包教學.md
│   └── ...更多文件
│
├── scripts/                            # 🔧 工具腳本
│   ├── 快速開始.sh
│   ├── 測試工具.sh
│   └── build_windows_exe.py
│
├── examples/                           # 📝 範例檔案
│   ├── 範例輸入文字.txt
│   └── 使用範例_Word轉PPT.md
│
└── .github/workflows/                  # ⚙️ GitHub Actions
    └── build.yml
```

---

## ❓ 常見問題

### Q: Word 轉 PPT 找不到藍色文字？
**A**: 
1. 確認 Word 中的文字是用「字型顏色」設定為藍色，不是螢光筆
2. 確認顏色是標準藍色（RGB 接近 0,0,255）
3. 可以調整容差值：修改 `extract_blue_text_from_docx.py` 中的 `tolerance` 參數

### Q: 如何安裝 Python 環境？
**A**: 
```bash
# 檢查是否已安裝 Python
python --version  # 或 python3 --version

# 安裝 python-pptx
pip install python-pptx  # 或 pip3 install python-pptx
```

### Q: 可以使用現有的 PPT 模板嗎？
**A**: 需要修改程式碼，在 `text_to_ppt.py` 的 `__init__` 方法中載入現有模板：
```python
self.prs = Presentation('您的模板.pptx')
```

### Q: 如何批次處理多個文字檔？
**A**: 使用命令列版本：
```bash
# Windows PowerShell
Get-ChildItem examples/*.txt | ForEach-Object { python text_to_ppt.py $_.FullName }

# Mac/Linux
for file in examples/*.txt; do python text_to_ppt.py "$file"; done
```

### Q: 字型在不同電腦上會不會跑掉？
**A**: 可能會。建議：
- 使用常見字型（如 Arial、微軟正黑體）
- 或將字型嵌入到 PPT 中（檔案 → 選項 → 儲存 → 將字型嵌入檔案）

---

## 📞 建議使用情境

| 情境 | 推薦版本 |
|------|----------|
| Word 文件有藍色標記 | **Word 轉 PPT** 🆕 |
| 從 Word 提取重點 | **Word 轉 PPT** 🆕 |
| 第一次使用 | **GUI 版本** |
| 快速建立簡報 | **GUI 版本** |
| 批次處理多個檔案 | **命令列版本** 或 **Word 轉 PPT** |
| 整合到自動化腳本 | **命令列版本** |
| 需要進階自訂 | **命令列版本**（修改程式碼） |

---

## 🎉 開始使用

1. **選擇適合的版本**：
   - 有 Word 文件？→ 使用 **Word 轉 PPT** 🆕
   - 初學者？→ 使用 **GUI 版本**
   - 開發者？→ 使用 **命令列版本**

2. **安裝必要套件**：
   ```bash
   # Word 轉 PPT 需要
   pip install python-pptx python-docx
   
   # 其他版本只需要
   pip install python-pptx
   ```

3. **開始使用**：
   ```bash
   # Word 轉 PPT
   python docx_to_ppt.py 你的文件.docx
   
   # GUI 版本
   python text_to_ppt_gui.py
   
   # 命令列版本
   python text_to_ppt.py examples/範例輸入文字.txt
   ```

4. **查看範例和文件**：
   - 文字格式範例：`examples/範例輸入文字.txt`
   - Word 轉 PPT 教學：`examples/使用範例_Word轉PPT.md`
   - 更多文件：`docs/` 資料夾

**祝您使用愉快！** 🚀

---

## 📚 更多資源

- [Word 轉 PPT 使用範例](examples/使用範例_Word轉PPT.md) 🆕
- [完整方案總覽](docs/README_完整方案總覽.md)
- [Python 安裝指南](docs/Python_安裝指南.txt)
- [GitHub Actions 打包教學](docs/GitHub_Actions_打包教學.md)
- [使用說明](docs/使用說明.txt)
