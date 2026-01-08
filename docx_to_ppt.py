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
    
    if len(sys.argv) < 2:
        print("使用方式：")
        print("  python docx_to_ppt.py <Word檔案.docx> [輸出PPT.pptx] [主標題]")
        print()
        print("範例：")
        print("  python docx_to_ppt.py 20251231.docx")
        print("  python docx_to_ppt.py 20251231.docx 我的簡報.pptx")
        print("  python docx_to_ppt.py 20251231.docx 我的簡報.pptx '2025年度報告'")
        print()
        print("功能：")
        print("  1. 自動提取 Word 中的藍色文字")
        print("  2. 轉換格式為 PPT 可用格式")
        print("  3. 生成 PowerPoint 簡報")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # 判斷輸出檔名
    if len(sys.argv) >= 3:
        output_ppt = sys.argv[2]
    else:
        base_name = os.path.splitext(input_file)[0]
        output_ppt = f"{base_name}.pptx"
    
    # 判斷主標題
    title = sys.argv[3] if len(sys.argv) >= 4 else "簡報標題"
    
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
    print("✏️  步驟 2/3：儲存為 TXT 格式（含變數模板）...")
    temp_txt = f"{os.path.splitext(output_ppt)[0]}_temp.txt"
    extractor.save_to_file(temp_txt, title)
    print(f"   ✅ 已儲存到：{temp_txt}")
    print()
    
    # 步驟 3：使用模板生成 PPT
    print("📊 步驟 3/3：使用模板生成 PowerPoint 簡報...")
    print(f"   目標：{output_ppt}")
    print()
    print("⚠️  注意：請手動編輯變數區塊後，使用以下指令生成 PPT：")
    print(f"   python generate_ppt_from_template.py template.pptx {temp_txt} {output_ppt}")
    print()
    print("💡 或使用舊版直接轉換（不含變數）：")
    print(f"   python text_to_ppt.py {temp_txt} {output_ppt}")
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
