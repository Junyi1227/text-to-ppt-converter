#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
從 Word 文件中提取藍色文字並轉換為 PPT 格式
Extract blue text from Word document and convert to PPT format
"""

from docx import Document
from docx.shared import RGBColor
import sys
import os


class BlueTextExtractor:
    """藍色文字提取器"""
    
    def __init__(self, tolerance=50):
        """
        初始化提取器
        
        Args:
            tolerance: 顏色容差，用於判斷是否為藍色（0-255）
        """
        self.tolerance = tolerance
        self.extracted_text = []
    
    def is_blue(self, rgb):
        """
        判斷顏色是否為藍色
        
        Args:
            rgb: RGBColor 物件或 tuple (r, g, b)
        
        Returns:
            bool: 是否為藍色
        """
        if rgb is None:
            return False
        
        # 獲取 RGB 值
        if isinstance(rgb, RGBColor):
            r, g, b = rgb
        elif isinstance(rgb, tuple) and len(rgb) == 3:
            r, g, b = rgb
        else:
            return False
        
        # 藍色判斷邏輯：B 值高，R 和 G 值低
        # 典型藍色：(0, 0, 255)，容許一些變化
        return (b > 150 and 
                r < self.tolerance and 
                g < self.tolerance)
    
    def extract_from_paragraph(self, paragraph):
        """
        從段落中提取藍色文字
        
        Args:
            paragraph: docx 段落物件
        
        Returns:
            str: 提取的藍色文字（如果有）
        """
        blue_text = []
        
        for run in paragraph.runs:
            # 檢查文字顏色
            if run.font.color and run.font.color.type == 1:  # RGB 顏色
                rgb = run.font.color.rgb
                if self.is_blue(rgb):
                    text = run.text.strip()
                    if text:
                        blue_text.append(text)
        
        return ' '.join(blue_text) if blue_text else None
    
    def extract_from_docx(self, docx_path):
        """
        從 Word 文件中提取所有藍色文字
        
        Args:
            docx_path: Word 文件路徑
        
        Returns:
            list: 提取的藍色文字列表
        """
        try:
            doc = Document(docx_path)
            self.extracted_text = []
            
            for paragraph in doc.paragraphs:
                blue_text = self.extract_from_paragraph(paragraph)
                if blue_text:
                    self.extracted_text.append(blue_text)
            
            return self.extracted_text
        
        except Exception as e:
            print(f"❌ 讀取文件時發生錯誤: {e}")
            sys.exit(1)
    
    def format_for_ppt(self, title="簡報標題"):
        """
        將提取的文字格式化為 text_to_ppt.py 可用的格式
        
        Args:
            title: 主標題
        
        Returns:
            str: 格式化後的文字
        """
        if not self.extracted_text:
            return ""
        
        # 基本格式：
        # ## 主標題（藍色背景）
        # # 小標題（灰色背景）
        # 內容行
        
        formatted = f"## {title}\n\n"
        
        for i, text in enumerate(self.extracted_text, 1):
            # 判斷是否為標題（可以根據實際情況調整）
            if len(text) < 30:  # 短文字當作小標題
                formatted += f"# {text}\n"
            else:  # 長文字當作內容
                formatted += f"{text}\n"
        
        return formatted
    
    def save_to_file(self, output_path, title="簡報標題"):
        """
        儲存提取的文字到檔案
        
        Args:
            output_path: 輸出檔案路徑
            title: 主標題
        """
        formatted_text = self.format_for_ppt(title)
        
        if not formatted_text:
            print("⚠️  沒有找到藍色文字")
            return False
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(formatted_text)
            
            print(f"✅ 成功提取 {len(self.extracted_text)} 段藍色文字")
            print(f"📝 已儲存到：{output_path}")
            return True
        
        except Exception as e:
            print(f"❌ 儲存檔案時發生錯誤: {e}")
            return False


def main():
    """主程式"""
    if len(sys.argv) < 2:
        print("使用方式：")
        print("  python extract_blue_text_from_docx.py <Word檔案.docx> [輸出檔案.txt] [主標題]")
        print()
        print("範例：")
        print("  python extract_blue_text_from_docx.py 20251231.docx")
        print("  python extract_blue_text_from_docx.py 20251231.docx output.txt")
        print("  python extract_blue_text_from_docx.py 20251231.docx output.txt '我的簡報'")
        print()
        print("提取完成後，可直接使用：")
        print("  python text_to_ppt.py output.txt")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # 判斷輸出檔名
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        # 自動產生輸出檔名
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_blue_text.txt"
    
    # 判斷主標題
    title = sys.argv[3] if len(sys.argv) >= 4 else "簡報標題"
    
    # 檢查輸入檔案是否存在
    if not os.path.exists(input_file):
        print(f"❌ 錯誤：找不到檔案 '{input_file}'")
        sys.exit(1)
    
    # 執行提取
    print(f"📖 讀取 Word 檔案：{input_file}")
    extractor = BlueTextExtractor(tolerance=50)
    extractor.extract_from_docx(input_file)
    
    # 顯示提取結果
    if extractor.extracted_text:
        print(f"\n找到 {len(extractor.extracted_text)} 段藍色文字：")
        print("-" * 50)
        for i, text in enumerate(extractor.extracted_text, 1):
            preview = text[:60] + "..." if len(text) > 60 else text
            print(f"{i}. {preview}")
        print("-" * 50)
    
    # 儲存結果
    if extractor.save_to_file(output_file, title):
        print(f"\n🎉 完成！現在可以執行：")
        print(f"   python text_to_ppt.py {output_file}")


if __name__ == "__main__":
    main()
