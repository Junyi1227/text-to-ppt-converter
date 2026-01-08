#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 藍色文字直接轉 PPT（一鍵完成）
Extract blue text from Word and convert to PPT in one step
"""

import sys
import os
from extract_blue_text_from_docx import BlueTextExtractor
from text_to_ppt import TextToPPTConverter


def main():
    """主程式：一鍵從 Word 轉換成 PPT"""
    
    # 參數 1：輸入 Word 檔案（可選，預設 input.docx）
    input_file = sys.argv[1] if len(sys.argv) >= 2 else "input.docx"
    
    # 參數 2：輸出 PPT 檔案（可選，預設 output.pptx）
    output_ppt = sys.argv[2] if len(sys.argv) >= 3 else "output.pptx"
    
    # 參數 3：主標題（選用）
    title = sys.argv[3] if len(sys.argv) >= 4 else "簡報標題"
    
    # 顯示使用說明（如果沒有任何參數）
    if len(sys.argv) == 1:
        print("🔄 Word 轉 PPT 工具（一鍵完成）")
        print("=" * 70)
        print()
        print("使用方式：")
        print("  python docx_to_ppt.py [Word檔案.docx] [輸出PPT.pptx]")
        print()
        print("預設值：")
        print("  Word檔案.docx = input.docx")
        print("  輸出PPT.pptx  = output.pptx")
        print()
        print("範例：")
        print("  python docx_to_ppt.py")
        print("    → 從 input.docx 提取，生成 output.txt，提示使用模板生成 PPT")
        print()
        print("  python docx_to_ppt.py 20251231.docx")
        print("    → 從 20251231.docx 提取，生成 output.txt")
        print()
        print("  python docx_to_ppt.py 20251231.docx sermon.pptx")
        print("    → 從 20251231.docx 提取，準備生成 sermon.pptx")
        print()
        print("=" * 70)
        print()
        print("功能：")
        print("  1. 自動提取 Word 中的藍色文字")
        print("  2. 儲存為含變數模板的 TXT 格式")
        print("  3. 提示使用 generate_ppt_from_template.py 生成 PPT")
        print()
        sys.exit(0)
    
    # 檢查輸入檔案
    if not os.path.exists(input_file):
        print(f"❌ 錯誤：找不到檔案 '{input_file}'")
        sys.exit(1)
    
    print("=" * 60)
    print("🔄 Word 藍色文字 → PowerPoint 轉換器")
    print("=" * 60)
    print()
    
    # 步驟 1：提取藍色文字
    print("📖 步驟 1/3：讀取 Word 檔案...")
    print(f"   來源：{input_file}")
    extractor = BlueTextExtractor(tolerance=50)
    extractor.extract_from_docx(input_file)
    
    if not extractor.extracted_text:
        print("❌ 沒有找到藍色文字！")
        print("   提示：請確認 Word 中有用藍色標記的文字")
        sys.exit(1)
    
    print(f"   ✅ 找到 {len(extractor.extracted_text)} 段藍色文字")
    print()
    
    # 步驟 2：儲存為 TXT（含變數模板）
    print("✏️  步驟 2/2：儲存為 TXT 格式（含變數模板）...")
    output_txt = "output.txt"
    extractor.save_to_file(output_txt, title)
    print(f"   ✅ 已儲存到：{output_txt}")
    print()
    
    print("=" * 60)
    print("✅ 提取完成！")
    print("=" * 60)
    print()
    print("📝 下一步：請編輯 output.txt 中的變數區塊，然後執行：")
    print()
    print("   python generate_ppt_from_template.py")
    print()
    print("   這將使用 template.pptx + output.txt 生成 output.pptx")
    print()
    print("💡 或指定輸出檔名：")
    print(f"   python generate_ppt_from_template.py template.pptx output.txt {output_ppt}")
    return
    print()
    
    print("=" * 60)
    print("🎉 轉換完成！")
    print("=" * 60)
    print(f"📁 輸出檔案：{output_ppt}")
    print(f"📊 投影片數：{len(converter.prs.slides)} 張")
    print(f"📝 藍色文字：{len(extractor.extracted_text)} 段")
    print()
    print("💡 提示：")
    print("   - 可以直接用 PowerPoint 開啟檔案")
    print("   - 如需調整格式，請編輯 text_to_ppt.py")
    print("   - 如需調整藍色識別，請編輯 extract_blue_text_from_docx.py")


if __name__ == "__main__":
    main()
