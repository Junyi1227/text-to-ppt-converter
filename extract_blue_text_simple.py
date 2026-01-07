#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
從 Word 文件中提取藍色文字（簡單版本）
只輸出純文字清單，每段一行
"""

from docx import Document
from docx.shared import RGBColor
import sys
import os


def is_blue(rgb, tolerance=50):
    """判斷顏色是否為藍色"""
    if rgb is None:
        return False
    
    if isinstance(rgb, RGBColor):
        r, g, b = rgb
    elif isinstance(rgb, tuple) and len(rgb) == 3:
        r, g, b = rgb
    else:
        return False
    
    # 藍色判斷：B 值高，R 和 G 值低
    return (b > 150 and r < tolerance and g < tolerance)


def is_verse_reference(text):
    """判斷是否為經文章節標記"""
    import re
    # 匹配格式：〈章節〉 或 <章節>
    pattern = r'^[〈<].+[〉>]\s*$'
    return re.match(pattern, text.strip())


def extract_blue_text(docx_path, tolerance=50):
    """提取所有藍色文字，並合併經文章節和內容"""
    try:
        doc = Document(docx_path)
        blue_texts = []
        pending_verse_ref = None  # 暫存的經文章節
        pending_verse_content = []  # 暫存的經文內容（多段）
        
        for paragraph in doc.paragraphs:
            para_blue_text = []
            
            for run in paragraph.runs:
                # 檢查文字顏色
                if run.font.color and run.font.color.type == 1:  # RGB 顏色
                    rgb = run.font.color.rgb
                    if is_blue(rgb, tolerance):
                        text = run.text.strip()
                        if text:
                            para_blue_text.append(text)
            
            if para_blue_text:
                current_text = ' '.join(para_blue_text)
                
                # 檢查是否為經文章節
                if is_verse_reference(current_text):
                    # 如果之前有未處理的經文，先完成它
                    if pending_verse_ref:
                        if pending_verse_content:
                            blue_texts.append(pending_verse_ref + '\n' + ' '.join(pending_verse_content))
                        else:
                            blue_texts.append(pending_verse_ref)
                    
                    # 開始新的經文章節
                    pending_verse_ref = current_text
                    pending_verse_content = []
                else:
                    # 一般文字或經文內容
                    if pending_verse_ref:
                        # 如果之前有經文章節，這是經文內容，繼續收集
                        pending_verse_content.append(current_text)
                    else:
                        # 一般文字，直接加入
                        blue_texts.append(current_text)
        
        # 處理最後可能剩餘的經文
        if pending_verse_ref:
            if pending_verse_content:
                blue_texts.append(pending_verse_ref + '\n' + ' '.join(pending_verse_content))
            else:
                blue_texts.append(pending_verse_ref)
        
        return blue_texts
    
    except Exception as e:
        print(f"❌ 讀取文件時發生錯誤: {e}")
        sys.exit(1)


def save_to_file(blue_texts, output_path):
    """儲存為純文字檔案，用空行分隔每段"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            for i, text in enumerate(blue_texts):
                f.write(text)
                # 段落之間用空行分隔（最後一段不加）
                if i < len(blue_texts) - 1:
                    f.write('\n\n')
        
        print(f"✅ 成功提取 {len(blue_texts)} 段藍色文字")
        print(f"📝 已儲存到：{output_path}")
        return True
    
    except Exception as e:
        print(f"❌ 儲存檔案時發生錯誤: {e}")
        return False


def main():
    """主程式"""
    if len(sys.argv) < 2:
        print("使用方式：")
        print("  python extract_blue_text_simple.py <Word檔案.docx> [輸出檔案.txt]")
        print()
        print("範例：")
        print("  python extract_blue_text_simple.py 20251231.docx")
        print("  python extract_blue_text_simple.py 20251231.docx blue_text.txt")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # 判斷輸出檔名
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_blue_text.txt"
    
    # 檢查輸入檔案
    if not os.path.exists(input_file):
        print(f"❌ 錯誤：找不到檔案 '{input_file}'")
        sys.exit(1)
    
    # 執行提取
    print(f"📖 讀取 Word 檔案：{input_file}")
    blue_texts = extract_blue_text(input_file)
    
    if not blue_texts:
        print("⚠️  沒有找到藍色文字")
        sys.exit(0)
    
    # 顯示預覽
    print(f"\n找到 {len(blue_texts)} 段藍色文字：")
    print("-" * 70)
    for i, text in enumerate(blue_texts[:10], 1):
        preview = text[:60] + "..." if len(text) > 60 else text
        print(f"{i}. {preview}")
    if len(blue_texts) > 10:
        print(f"... 還有 {len(blue_texts) - 10} 段")
    print("-" * 70)
    
    # 儲存結果
    save_to_file(blue_texts, output_file)


if __name__ == "__main__":
    main()
