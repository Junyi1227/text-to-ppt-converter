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
        從 Word 文件中提取所有藍色文字（連續的藍色段落會合併）
        
        Args:
            docx_path: Word 文件路徑
        
        Returns:
            list: 提取的藍色文字列表
        """
        try:
            doc = Document(docx_path)
            self.extracted_text = []
            current_group = []  # 用來收集連續的藍色段落
            
            for paragraph in doc.paragraphs:
                blue_text = self.extract_from_paragraph(paragraph)
                
                if blue_text:
                    # 如果是藍色段落，加入當前組
                    current_group.append(blue_text)
                else:
                    # 如果不是藍色段落，將之前收集的組合併並加入結果
                    if current_group:
                        merged_text = '\n'.join(current_group)
                        self.extracted_text.append(merged_text)
                        current_group = []
            
            # 處理最後一組（如果文件結尾是藍色段落）
            if current_group:
                merged_text = '\n'.join(current_group)
                self.extracted_text.append(merged_text)
            
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
        儲存提取的文字到檔案（包含變數模板）
        
        Args:
            output_path: 輸出檔案路徑
            title: 主標題
        """
        if not self.extracted_text:
            print("⚠️  沒有找到藍色文字")
            return False
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                # 寫入變數模板
                f.write("[變數]\n")
                f.write("日期=2026年1月1日\n")
                f.write("禮拜類型=週三禮拜\n")
                f.write("主題=我是主題\n")
                f.write("經文章節=【箴言27章12節、詩篇46篇1節】\n")
                f.write("經文1=〈箴言27章12節〉XXXXXXXX。\n")
                f.write("經文2=〈詩篇46篇1節〉OOOOOOOO。\n")
                f.write("[變數結束]\n\n")
                
                # 寫入提取的藍色文字內容
                for text in self.extracted_text:
                    f.write(f"{text}\n\n")
            
            print(f"✅ 成功提取 {len(self.extracted_text)} 段藍色文字")
            print(f"📝 已儲存到：{output_path}")
            return True
        
        except Exception as e:
            print(f"❌ 儲存檔案時發生錯誤: {e}")
            return False


def main():
    """主程式"""
    # 參數 1：輸入 Word 檔案（可選，預設 input.docx）
    input_file = sys.argv[1] if len(sys.argv) >= 2 else "input.docx"
    
    # 參數 2：輸出 TXT 檔案（可選，預設 output.txt）
    output_file = sys.argv[2] if len(sys.argv) >= 3 else "output.txt"
    
    # 參數 3：主標題（選用，目前未使用）
    title = sys.argv[3] if len(sys.argv) >= 4 else "簡報標題"
    
    # 顯示使用說明（如果沒有任何參數）
    if len(sys.argv) == 1:
        print("📖 藍色文字提取工具")
        print("=" * 70)
        print()
        print("使用方式：")
        print("  python extract_blue_text_from_docx.py [Word檔案.docx] [輸出檔案.txt]")
        print()
        print("預設值：")
        print("  Word檔案.docx = input.docx")
        print("  輸出檔案.txt  = output.txt")
        print()
        print("範例：")
        print("  python extract_blue_text_from_docx.py")
        print("    → 從 input.docx 提取，輸出到 output.txt")
        print()
        print("  python extract_blue_text_from_docx.py 20251231.docx")
        print("    → 從 20251231.docx 提取，輸出到 output.txt")
        print()
        print("  python extract_blue_text_from_docx.py 20251231.docx sermon.txt")
        print("    → 從 20251231.docx 提取，輸出到 sermon.txt")
        print()
        print("=" * 70)
        print()
        print("💡 提取完成後，可直接執行：")
        print("   python generate_ppt_from_template.py")
        print()
        sys.exit(0)
    
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
