# PPT 生成系統 V2 範例檔案

本資料夾包含 V2 系統的範例檔案。

## 📁 檔案說明

### 範例檔案
- **config_範例.txt** - Config 檔案範例（定義頁面結構）
- **input_範例.txt** - 輸入檔案範例（含變數模板）
- **範例輸出.pptx** - 生成的 PPT 範例

### 測試檔案
- **config_output_test.txt** - 測試用 config 檔案
- **output_test.txt** - 測試用輸入檔案
- **output.txt** - 從 Word 提取的藍色文字範例

## 🚀 使用方式

### 基本使用
```bash
python generate_ppt_from_template_v2.py template.pptx input_範例.txt config_範例.txt output.pptx
```

### 完整流程
```bash
# 步驟 1: 從 Word 提取藍色文字
python extract_blue_text_from_docx.py input.docx output.txt

# 步驟 2: 編輯 output.txt 填入變數

# 步驟 3: 創建 config.txt 定義頁面結構

# 步驟 4: 生成 PPT
python generate_ppt_from_template_v2.py template.pptx output.txt config.txt final.pptx
```

## 📖 詳細說明

請參閱 [使用說明_V2.md](../../docs/使用說明_V2.md) 獲取完整文檔。
