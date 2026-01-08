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
    """特定顏色文字提取器"""
    
    def __init__(self, target_color=None, tolerance=50):
        """
        初始化提取器
        
        Args:
            target_color: 目標顏色 (r, g, b) 或 "#RRGGBB"，預設為藍色
            tolerance: 顏色容差（0-255）
        """
        self.tolerance = tolerance
        self.extracted_text = []
        
        # 設定目標顏色（預設藍色）
        if target_color is None:
            self.target_color = (0, 0, 255)  # 預設藍色
        elif isinstance(target_color, str) and target_color.startswith('#'):
            # 16進位格式轉換
            hex_color = target_color.lstrip('#')
            self.target_color = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        elif isinstance(target_color, tuple) and len(target_color) == 3:
            self.target_color = target_color
        else:
            raise ValueError("target_color 必須是 (r, g, b) tuple 或 '#RRGGBB' 格式")
    
    def is_target_color(self, rgb):
        """
        判斷顏色是否為目標顏色（在容差範圍內）
        
        Args:
            rgb: RGBColor 物件或 tuple (r, g, b)
        
        Returns:
            bool: 是否為目標顏色
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
        
        # 判斷是否在目標顏色的容差範圍內
        target_r, target_g, target_b = self.target_color
        return (abs(r - target_r) <= self.tolerance and
                abs(g - target_g) <= self.tolerance and
                abs(b - target_b) <= self.tolerance)
    
    # 保留舊方法名稱以維持向下相容
    def is_blue(self, rgb):
        """向下相容的方法，實際調用 is_target_color"""
        return self.is_target_color(rgb)
    
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
    
    # 固定輸出檔案為 output.txt
    output_file = "output.txt"
    
    # 從 config.txt 讀取顏色設定（可選，預設藍色）
    target_color = None
    config_file = "config.txt"
    
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                in_color_section = False
                for line in f:
                    line = line.strip()
                    
                    if line == '[顏色設定]':
                        in_color_section = True
                        continue
                    
                    if line.startswith('[') and line.endswith(']'):
                        in_color_section = False
                        continue
                    
                    if in_color_section and line.startswith('提取文字顏色'):
                        if '=' in line:
                            _, value = line.split('=', 1)
                            value = value.strip()
                            
                            if value.startswith('#'):
                                target_color = value
                            else:
                                rgb = tuple(int(c.strip()) for c in value.split(','))
                                if len(rgb) == 3:
                                    target_color = rgb
                            break
        except Exception as e:
            print(f"⚠️  警告：讀取 config.txt 時發生錯誤: {e}")
            print(f"    使用預設藍色")
    
    # 顯示使用說明（如果沒有任何參數）
    if len(sys.argv) == 1:
        print("📖 特定顏色文字提取工具")
        print("=" * 70)
        print()
        print("使用方式：")
        print("  python 1_extract.py [Word檔案]")
        print()
        print("參數說明：")
        print("  Word檔案  - Word 文件路徑（預設：input.docx）")
        print()
        print("固定設定：")
        print("  輸出檔案：output.txt（固定）")
        print("  顏色設定：從 config.txt 讀取「提取文字顏色」（預設：藍色）")
        print()
        print("Config 顏色設定範例（在 config.txt 中）：")
        print("  [顏色設定]")
        print("  提取文字顏色 = 0,0,255        # 藍色（預設）")
        print("  提取文字顏色 = 255,0,0        # 紅色")
        print("  提取文字顏色 = #FF0000        # 紅色（16進位）")
        print()
        print("範例：")
        print("  python 1_extract.py")
        print("    → 從 input.docx 提取文字，輸出到 output.txt")
        print()
        print("  python 1_extract.py 20251231.docx")
        print("    → 從 20251231.docx 提取文字，輸出到 output.txt")
        print()
        print("=" * 70)
        print()
        print("💡 提取完成後，可直接執行：")
        print("   python 2_generate.py")
        print()
        sys.exit(0)
    
    # 檢查輸入檔案是否存在
    if not os.path.exists(input_file):
        print(f"❌ 錯誤：找不到檔案 '{input_file}'")
        sys.exit(1)
    
    # 執行提取
    print(f"📖 讀取 Word 檔案：{input_file}")
    if target_color:
        if isinstance(target_color, str):
            print(f"🎨 目標顏色：{target_color}")
        else:
            print(f"🎨 目標顏色：RGB{target_color}")
    else:
        print(f"🎨 目標顏色：藍色（預設）")
    
    extractor = BlueTextExtractor(target_color=target_color, tolerance=50)
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
    if extractor.save_to_file(output_file):
        print(f"\n🎉 完成！現在可以執行：")
        print(f"   python 2_generate.py")


if __name__ == "__main__":
    main()
