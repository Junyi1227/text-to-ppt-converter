# Word 藍色文字轉 PPT 使用範例

## 📖 情境說明

當你有一個 Word 文件，其中重要內容用**藍色**標記時，可以使用此工具快速提取並轉換成 PPT。

## 🚀 快速開始

### 步驟 1：準備 Word 文件

在 Word 中將重要內容標記為藍色：

```
這是普通文字
這是藍色的標題          <-- 藍色
這是藍色的重點內容      <-- 藍色
這又是普通文字
```

### 步驟 2：安裝套件

```bash
pip install python-docx
```

### 步驟 3：提取藍色文字

```bash
# 基本用法（自動產生輸出檔名）
python extract_blue_text_from_docx.py 20251231.docx

# 指定輸出檔名
python extract_blue_text_from_docx.py 20251231.docx my_content.txt

# 指定輸出檔名和主標題
python extract_blue_text_from_docx.py 20251231.docx my_content.txt "2025年度報告"
```

### 步驟 4：轉換成 PPT

```bash
python text_to_ppt.py 20251231_blue_text.txt
```

## 🎨 自訂藍色容差

如果你的藍色不是標準藍色，可以修改腳本：

```python
# 在程式碼中修改容差值（預設是 50）
extractor = BlueTextExtractor(tolerance=80)  # 容許更多變化
```

## 📋 輸出格式

提取的文字會自動格式化為：

```
## 主標題

# 短文字（當作小標題）
長文字內容會當作投影片內容
另一段內容
```

## 💡 進階用法

### 一鍵轉換

```bash
# 提取並直接轉換成 PPT
python extract_blue_text_from_docx.py 20251231.docx output.txt && \
python text_to_ppt.py output.txt
```

### 批次處理多個 Word 檔案

```bash
# Mac/Linux
for file in *.docx; do
    python extract_blue_text_from_docx.py "$file"
    python text_to_ppt.py "${file%.docx}_blue_text.txt"
done

# Windows PowerShell
Get-ChildItem *.docx | ForEach-Object {
    python extract_blue_text_from_docx.py $_.Name
    $txtFile = $_.BaseName + "_blue_text.txt"
    python text_to_ppt.py $txtFile
}
```

## ⚠️ 注意事項

1. **顏色識別**：腳本會識別「典型藍色」（RGB 接近 0,0,255）
2. **文字長度**：短於 30 字元的會被當作小標題
3. **檔案格式**：僅支援 .docx 格式（不支援 .doc）
4. **字型顏色**：必須是直接設定的 RGB 顏色（不是主題顏色）

## 🔧 故障排除

### 找不到藍色文字？

1. 確認 Word 中的文字是用「字型顏色」設定，不是螢光筆
2. 確認顏色是標準藍色或接近藍色
3. 嘗試增加容差值：`tolerance=80` 或 `tolerance=100`

### 提取的格式不對？

可以手動編輯生成的 .txt 檔案：
- `##` 開頭：主題頁（藍色背景）
- `#` 開頭：內文頁標題（灰色背景）
- 一般文字：投影片內容

## 📝 範例

假設你的 Word 文件內容：

```
會議記錄
日期：2025年12月31日        <-- 藍色
主題：年度總結              <-- 藍色

今年完成的項目：
1. 專案 A                   <-- 藍色
   - 完成度 100%
2. 專案 B                   <-- 藍色
   - 完成度 95%
```

提取後會生成：

```
## 年度報告

# 日期：2025年12月31日
# 主題：年度總結
# 1. 專案 A
# 2. 專案 B
```

轉換成 PPT 後，會有 5 張投影片：
1. 主題頁：年度報告
2. 內文頁：日期：2025年12月31日
3. 內文頁：主題：年度總結
4. 內文頁：1. 專案 A
5. 內文頁：2. 專案 B
