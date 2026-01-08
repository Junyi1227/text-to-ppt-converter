#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 生成程式 V2 - 基於 template.pptx 的彈性化版本

使用方式：
    python generate_ppt_from_template_v2.py template.pptx input.txt config.txt output.pptx

功能：
    - 支援彈性化的頁面結構定義（透過 config）
    - 支援變數模板（從 TXT 讀取）
    - 自動識別經文格式
    - 支援多種頁面類型：COVER, TITLE, CONTENT, BIBLE, AUTOCONTENT
"""

import sys
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE


class PPTGeneratorV2:
    """PPT 生成器 V2"""
    
    def __init__(self, template_path, output_path):
        """
        初始化 PPT 生成器
        
        Args:
            template_path: 模板 PPT 路徑（必須包含 4 頁）
            output_path: 輸出 PPT 路徑
        """
        # 先複製 template 到 output
        import shutil
        shutil.copy2(template_path, output_path)
        
        # 開啟輸出檔案（包含模板的 4 頁）
        self.output_prs = Presentation(output_path)
        self.output_path = output_path
        
        # 確認模板有 4 頁
        if len(self.output_prs.slides) < 4:
            raise ValueError(f"模板必須包含至少 4 頁，目前只有 {len(self.output_prs.slides)} 頁")
        
        # 注意：不刪除模板頁，稍後生成時會用到
        
        # 變數字典
        self.variables = {}
        # 內容列表
        self.content_lines = []
        # 頁面結構
        self.page_structure = []
        # 記錄需要刪除的模板頁索引
        self.template_page_count = len(self.output_prs.slides)
    
    def load_variables_and_content(self, txt_path):
        """
        從 TXT 檔案讀取變數和內容（使用空行分隔頁面）
        
        Args:
            txt_path: TXT 檔案路徑
        """
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        in_variables = False
        in_content = False
        current_block = []
        
        for line in lines:
            line = line.rstrip('\n')
            
            # 檢查變數區開始
            if line.strip() == '[變數]':
                in_variables = True
                continue
            
            # 檢查變數區結束
            if line.strip() == '[變數結束]':
                in_variables = False
                in_content = True
                continue
            
            # 讀取變數
            if in_variables and '=' in line:
                key, value = line.split('=', 1)
                self.variables[key.strip()] = value.strip()
            
            # 讀取內容（使用空行分隔不同頁面）
            elif in_content:
                if line.strip():
                    # 有內容的行，加入當前區塊
                    current_block.append(line.strip())
                else:
                    # 空行，表示一個區塊結束
                    if current_block:
                        # 將區塊合併成一個項目（用換行符連接）
                        self.content_lines.append('\n'.join(current_block))
                        current_block = []
        
        # 處理最後一個區塊（如果檔案結尾沒有空行）
        if current_block:
            self.content_lines.append('\n'.join(current_block))
        
        print(f"✅ 讀取變數: {len(self.variables)} 個")
        print(f"✅ 讀取內容區塊: {len(self.content_lines)} 個（用空行分隔）")
    
    def load_config(self, config_path):
        """
        從 config 檔案讀取頁面結構
        
        Args:
            config_path: config 檔案路徑
        """
        with open(config_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        in_structure = False
        
        for line in lines:
            line = line.strip()
            
            # 跳過空行和註解
            if not line or line.startswith('#'):
                continue
            
            # 檢查頁面結構區開始
            if line == '[頁面結構]':
                in_structure = True
                continue
            
            # 讀取頁面結構
            if in_structure:
                # 解析頁面類型和參數
                if '=' in line:
                    parts = line.split('=', 1)
                    page_type = parts[0].strip()
                    param = parts[1].strip()
                    self.page_structure.append((page_type, param))
                else:
                    page_type = line.strip()
                    self.page_structure.append((page_type, None))
        
        print(f"✅ 讀取頁面結構: {len(self.page_structure)} 頁")
    
    def is_verse_format(self, text):
        """
        判斷是否為經文格式
        支援兩種格式：
        1. 單行：〈創19:17〉領他們出來...
        2. 多行：第一行是章節，第二行是內容
        
        Args:
            text: 要判斷的文字
            
        Returns:
            Match object 如果匹配，否則 None
        """
        # 單行格式：〈章節〉內容
        pattern = r'^[〈<]([^〉>]+)[〉>](.+)$'
        match = re.match(pattern, text)
        return match
    
    def convert_verse_reference(self, verse_ref):
        """
        轉換經文章節格式
        創19:17 → 創世記19章17節
        箴言27章12節 → 箴言27章12節（不變）
        
        Args:
            verse_ref: 原始章節格式
            
        Returns:
            轉換後的章節格式
        """
        # 如果已經包含「章」「節」，直接返回
        if '章' in verse_ref and '節' in verse_ref:
            return verse_ref
        
        # 轉換簡化格式（例如：創19:17）
        pattern = r'^([^0-9]+)(\d+):(\d+)$'
        match = re.match(pattern, verse_ref)
        
        if match:
            book = match.group(1)
            chapter = match.group(2)
            verse = match.group(3)
            
            # 書卷名稱轉換（如果需要）
            book_map = {
                '創': '創世記',
                '出': '出埃及記',
                '利': '利未記',
                '民': '民數記',
                '申': '申命記',
                '書': '約書亞記',
                '士': '士師記',
                '得': '路得記',
                '撒上': '撒母耳記上',
                '撒下': '撒母耳記下',
                '王上': '列王紀上',
                '王下': '列王紀下',
                '代上': '歷代志上',
                '代下': '歷代志下',
                '拉': '以斯拉記',
                '尼': '尼希米記',
                '斯': '以斯帖記',
                '伯': '約伯記',
                '詩': '詩篇',
                '箴': '箴言',
                '傳': '傳道書',
                '歌': '雅歌',
                '賽': '以賽亞書',
                '耶': '耶利米書',
                '哀': '耶利米哀歌',
                '結': '以西結書',
                '但': '但以理書',
                '何': '何西阿書',
                '珥': '約珥書',
                '摩': '阿摩司書',
                '俄': '俄巴底亞書',
                '拿': '約拿書',
                '彌': '彌迦書',
                '鴻': '那鴻書',
                '哈': '哈巴谷書',
                '番': '西番雅書',
                '該': '哈該書',
                '亞': '撒迦利亞書',
                '瑪': '瑪拉基書',
                '太': '馬太福音',
                '可': '馬可福音',
                '路': '路加福音',
                '約': '約翰福音',
                '徒': '使徒行傳',
                '羅': '羅馬書',
                '林前': '哥林多前書',
                '林後': '哥林多後書',
                '加': '加拉太書',
                '弗': '以弗所書',
                '腓': '腓立比書',
                '西': '歌羅西書',
                '帖前': '帖撒羅尼迦前書',
                '帖後': '帖撒羅尼迦後書',
                '提前': '提摩太前書',
                '提後': '提摩太後書',
                '多': '提多書',
                '門': '腓利門書',
                '來': '希伯來書',
                '雅': '雅各書',
                '彼前': '彼得前書',
                '彼後': '彼得後書',
                '約壹': '約翰一書',
                '約貳': '約翰二書',
                '約參': '約翰三書',
                '猶': '猶大書',
                '啟': '啟示錄'
            }
            
            full_book = book_map.get(book, book)
            return f"{full_book}{chapter}章{verse}節"
        
        return verse_ref
    
    def create_cover_page(self, subtitle=None):
        """
        建立封面頁（複製 template 第 1 頁並修改內容）
        
        Args:
            subtitle: 小標題（可選）
        """
        # 使用模板第 1 頁的版面配置
        template_slide = self.output_prs.slides[0]
        slide_layout = template_slide.slide_layout
        new_slide = self.output_prs.slides.add_slide(slide_layout)
        
        # 刪除從版面配置繼承的空文字框
        shapes_to_remove = []
        for shape in new_slide.shapes:
            if hasattr(shape, "text_frame") and not shape.text.strip():
                shapes_to_remove.append(shape)
        
        for shape in shapes_to_remove:
            sp = shape.element
            sp.getparent().remove(sp)
        
        # 複製模板頁的所有形狀並修改文字
        for shape in template_slide.shapes:
            if hasattr(shape, "text_frame"):
                # 根據位置判斷是哪個文字框
                # 文字框1: 日期+禮拜類型 (top ≈ 1.23")
                # 文字框2: 小標題 (top ≈ 4.30")
                # 文字框3: 經文章節 (top ≈ 3.40")
                
                if abs(shape.top.inches - 1.23) < 0.1:
                    # 文字框1: 日期+禮拜類型
                    date = self.variables.get('日期', '')
                    service_type = self.variables.get('禮拜類型', '')
                    text = f"{date}\n\n{service_type}"
                    self._create_textbox_with_format(new_slide, shape, text)
                
                elif abs(shape.top.inches - 4.30) < 0.1:
                    # 文字框2: 小標題（只有在有參數時才顯示）
                    if subtitle:
                        self._create_textbox_with_format(new_slide, shape, subtitle)
                
                elif abs(shape.top.inches - 3.40) < 0.1:
                    # 文字框3: 經文章節
                    verse_refs = self.variables.get('經文章節', '')
                    self._create_textbox_with_format(new_slide, shape, verse_refs)
        
        return new_slide
    
    def create_title_page(self, subtitle=None):
        """
        建立主題頁（複製 template 第 2 頁並修改內容）
        
        Args:
            subtitle: 小標題（可選）
        """
        # 使用模板第 2 頁的版面配置
        template_slide = self.output_prs.slides[1]
        slide_layout = template_slide.slide_layout
        new_slide = self.output_prs.slides.add_slide(slide_layout)
        
        # 刪除從版面配置繼承的空文字框
        shapes_to_remove = []
        for shape in new_slide.shapes:
            if hasattr(shape, "text_frame") and not shape.text.strip():
                shapes_to_remove.append(shape)
        
        for shape in shapes_to_remove:
            sp = shape.element
            sp.getparent().remove(sp)
        
        # 複製模板頁的所有形狀並修改文字
        for shape in template_slide.shapes:
            if hasattr(shape, "text_frame"):
                # 文字框1: 日期+禮拜類型 (top ≈ 0.51")
                # 文字框2: 主題 (top ≈ 1.72")
                # 文字框3: 經文章節 (top ≈ 3.76")
                # 文字框4: 小標題 (top ≈ 4.46")
                
                if abs(shape.top.inches - 0.51) < 0.1:
                    # 文字框1: 日期+禮拜類型
                    date = self.variables.get('日期', '')
                    service_type = self.variables.get('禮拜類型', '')
                    text = f"{date} {service_type}"
                    self._create_textbox_with_format(new_slide, shape, text)
                
                elif abs(shape.top.inches - 1.72) < 0.1:
                    # 文字框2: 主題
                    title = self.variables.get('主題', '')
                    self._create_textbox_with_format(new_slide, shape, title)
                
                elif abs(shape.top.inches - 3.76) < 0.1:
                    # 文字框3: 經文章節
                    verse_refs = self.variables.get('經文章節', '')
                    self._create_textbox_with_format(new_slide, shape, verse_refs)
                
                elif abs(shape.top.inches - 4.46) < 0.1:
                    # 文字框4: 小標題（只有在有參數時才顯示）
                    if subtitle:
                        self._create_textbox_with_format(new_slide, shape, subtitle)
        
        return new_slide
    
    def create_content_page(self, text):
        """
        建立內文頁（複製 template 第 3 頁並修改內容）
        
        Args:
            text: 內容文字
        """
        # 使用模板第 3 頁的版面配置
        template_slide = self.output_prs.slides[2]
        slide_layout = template_slide.slide_layout
        new_slide = self.output_prs.slides.add_slide(slide_layout)
        
        # 刪除從版面配置繼承的空文字框
        shapes_to_remove = []
        for shape in new_slide.shapes:
            if hasattr(shape, "text_frame") and not shape.text.strip():
                shapes_to_remove.append(shape)
        
        for shape in shapes_to_remove:
            sp = shape.element
            sp.getparent().remove(sp)
        
        # 找到模板頁的第一個文字框並複製
        for shape in template_slide.shapes:
            if hasattr(shape, "text_frame"):
                # 使用模板的位置和大小（不要寫死）
                self._create_textbox_with_format(new_slide, shape, text)
                break
        
        return new_slide
    
    def create_verse_page(self, verse_ref, verse_text):
        """
        建立經文頁（複製 template 第 4 頁並修改內容）
        
        Args:
            verse_ref: 經文章節
            verse_text: 經文內容
        """
        # 使用模板第 4 頁的版面配置
        template_slide = self.output_prs.slides[3]
        slide_layout = template_slide.slide_layout
        new_slide = self.output_prs.slides.add_slide(slide_layout)
        
        # 刪除從版面配置繼承的空文字框
        shapes_to_remove = []
        for shape in new_slide.shapes:
            if hasattr(shape, "text_frame") and not shape.text.strip():
                shapes_to_remove.append(shape)
        
        for shape in shapes_to_remove:
            sp = shape.element
            sp.getparent().remove(sp)
        
        # 找到模板頁的第一個文字框
        source_shape = None
        for shape in template_slide.shapes:
            if hasattr(shape, "text_frame"):
                source_shape = shape
                break
        
        if source_shape:
            # 使用模板的位置和大小（不要寫死）
            new_shape = new_slide.shapes.add_textbox(
                source_shape.left,
                source_shape.top,
                source_shape.width,
                source_shape.height
            )
            
            # 清空預設文字
            new_shape.text_frame.clear()
            
            # 複製文字框屬性
            new_shape.text_frame.word_wrap = source_shape.text_frame.word_wrap
            new_shape.text_frame.vertical_anchor = source_shape.text_frame.vertical_anchor
            new_shape.text_frame.auto_size = source_shape.text_frame.auto_size
            
            # 轉換章節格式
            verse_ref_formatted = self.convert_verse_reference(verse_ref)
            
            # 第一段：經文章節（從模板複製格式）
            p1 = new_shape.text_frame.paragraphs[0]
            p1.text = verse_ref_formatted
            
            # 複製第一段格式（如果模板有的話）
            if source_shape.text_frame.paragraphs:
                source_p = source_shape.text_frame.paragraphs[0]
                p1.alignment = source_p.alignment
                
                if source_p.runs:
                    source_run = source_p.runs[0]
                    for run in p1.runs:
                        if source_run.font.size:
                            run.font.size = source_run.font.size
                        if source_run.font.bold is not None:
                            run.font.bold = source_run.font.bold
                        if source_run.font.name:
                            run.font.name = source_run.font.name
                        # 經文章節使用淺藍色
                        run.font.color.rgb = RGBColor(121, 155, 193)
            
            # 第二段：經文內容
            p2 = new_shape.text_frame.add_paragraph()
            p2.text = verse_text
            
            # 複製第二段格式（如果模板有多個段落的話）
            if len(source_shape.text_frame.paragraphs) > 1:
                source_p2 = source_shape.text_frame.paragraphs[1]
                p2.alignment = source_p2.alignment
                
                if source_p2.runs:
                    source_run2 = source_p2.runs[0]
                    for run in p2.runs:
                        if source_run2.font.size:
                            run.font.size = source_run2.font.size
                        if source_run2.font.bold is not None:
                            run.font.bold = source_run2.font.bold
                        if source_run2.font.name:
                            run.font.name = source_run2.font.name
                        # 經文內容使用深藍色
                        run.font.color.rgb = RGBColor(27, 54, 106)
            else:
                # 如果模板只有一段，使用第一段的格式
                if source_shape.text_frame.paragraphs:
                    source_p = source_shape.text_frame.paragraphs[0]
                    p2.alignment = source_p.alignment
                    
                    if source_p.runs:
                        source_run = source_p.runs[0]
                        for run in p2.runs:
                            if source_run.font.size:
                                run.font.size = source_run.font.size
                            if source_run.font.bold is not None:
                                run.font.bold = source_run.font.bold
                            if source_run.font.name:
                                run.font.name = source_run.font.name
                            # 經文內容使用深藍色
                            run.font.color.rgb = RGBColor(27, 54, 106)
        
        return new_slide
    
    def _create_textbox_with_format(self, slide, source_shape, text):
        """
        創建文字框並複製格式（支援多段落）
        
        Args:
            slide: 目標投影片
            source_shape: 來源形狀（用於複製位置和格式）
            text: 要填入的文字
        """
        # 創建新文字框
        new_shape = slide.shapes.add_textbox(
            source_shape.left,
            source_shape.top,
            source_shape.width,
            source_shape.height
        )
        
        # 設定文字
        new_shape.text = text
        
        # 複製文字框屬性
        new_shape.text_frame.word_wrap = source_shape.text_frame.word_wrap
        new_shape.text_frame.vertical_anchor = source_shape.text_frame.vertical_anchor
        new_shape.text_frame.auto_size = source_shape.text_frame.auto_size
        
        # 複製所有段落的格式
        text_paragraphs = text.split('\n')
        target_paragraphs = new_shape.text_frame.paragraphs
        
        # 確保目標有足夠的段落
        while len(target_paragraphs) < len(text_paragraphs):
            new_shape.text_frame.add_paragraph()
            target_paragraphs = new_shape.text_frame.paragraphs
        
        # 為每個段落複製對應的格式
        for i, target_p in enumerate(target_paragraphs):
            # 找到對應的源段落（如果沒有就用最後一個）
            source_para_index = min(i, len(source_shape.text_frame.paragraphs) - 1)
            if source_para_index >= 0 and source_para_index < len(source_shape.text_frame.paragraphs):
                source_p = source_shape.text_frame.paragraphs[source_para_index]
                
                # 複製段落對齊
                target_p.alignment = source_p.alignment
                
                # 複製字體格式
                if source_p.runs and target_p.runs:
                    source_run = source_p.runs[0]
                    for target_run in target_p.runs:
                        if source_run.font.size:
                            target_run.font.size = source_run.font.size
                        if source_run.font.bold is not None:
                            target_run.font.bold = source_run.font.bold
                        if source_run.font.name:
                            target_run.font.name = source_run.font.name
                        if source_run.font.color and source_run.font.color.rgb:
                            target_run.font.color.rgb = source_run.font.color.rgb
        
        return new_shape
    
    def _copy_text_format(self, source_shape, target_shape):
        """
        複製文字格式
        
        Args:
            source_shape: 來源形狀
            target_shape: 目標形狀
        """
        if not source_shape.text_frame.paragraphs or not target_shape.text_frame.paragraphs:
            return
        
        source_p = source_shape.text_frame.paragraphs[0]
        target_p = target_shape.text_frame.paragraphs[0]
        
        # 複製段落對齊
        target_p.alignment = source_p.alignment
        
        # 複製字體格式
        if source_p.runs and target_p.runs:
            source_run = source_p.runs[0]
            for target_run in target_p.runs:
                if source_run.font.size:
                    target_run.font.size = source_run.font.size
                if source_run.font.bold is not None:
                    target_run.font.bold = source_run.font.bold
                if source_run.font.name:
                    target_run.font.name = source_run.font.name
                if source_run.font.color and source_run.font.color.rgb:
                    target_run.font.color.rgb = source_run.font.color.rgb
    
    def generate(self):
        """
        根據頁面結構生成 PPT
        """
        content_index = 0  # 追蹤 AUTOCONTENT 的當前索引
        
        for page_type, param in self.page_structure:
            print(f"生成頁面: {page_type}" + (f" = {param}" if param else ""))
            
            if page_type == "COVER":
                # 封面頁
                self.create_cover_page(subtitle=param)
            
            elif page_type == "TITLE":
                # 主題頁
                self.create_title_page(subtitle=param)
            
            elif page_type == "CONTENT":
                # 內文頁（固定內容）
                if param:
                    self.create_content_page(param)
            
            elif page_type == "BIBLE":
                # 經文頁（讀取變數區的經文1, 經文2, ...）
                verse_num = 1
                while True:
                    verse_key = f"經文{verse_num}"
                    if verse_key not in self.variables:
                        break
                    
                    verse_data = self.variables[verse_key]
                    # 用 〉 分隔章節和內容
                    if '〉' in verse_data:
                        verse_ref, verse_text = verse_data.split('〉', 1)
                        verse_ref = verse_ref.lstrip('〈<')
                        verse_text = verse_text.strip()
                        
                        print(f"  生成經文頁 {verse_num}: {verse_ref}")
                        self.create_verse_page(verse_ref, verse_text)
                    
                    verse_num += 1
            
            elif page_type == "AUTOCONTENT":
                # 自動內容頁（從內容區讀取，每個區塊是一頁）
                while content_index < len(self.content_lines):
                    block = self.content_lines[content_index]
                    content_index += 1
                    
                    # 檢查區塊的第一行是否為經文格式
                    lines_in_block = block.split('\n')
                    first_line = lines_in_block[0] if lines_in_block else ""
                    
                    # 檢查是否為經文格式（單行）
                    verse_match = self.is_verse_format(first_line)
                    if verse_match:
                        # 單行經文格式：〈章節〉內容
                        verse_ref = verse_match.group(1)
                        verse_text = verse_match.group(2).strip()
                        print(f"  生成經文頁: {verse_ref}")
                        self.create_verse_page(verse_ref, verse_text)
                    elif first_line.startswith('〈') or first_line.startswith('<'):
                        # 多行經文格式：第一行是章節，後面是內容
                        verse_ref = first_line.lstrip('〈<').rstrip('〉>')
                        verse_text = '\n'.join(lines_in_block[1:]) if len(lines_in_block) > 1 else ""
                        print(f"  生成經文頁: {verse_ref}")
                        self.create_verse_page(verse_ref, verse_text)
                    else:
                        # 一般內容（整個區塊）
                        print(f"  生成內文頁")
                        self.create_content_page(block)
        
        # 刪除前面的模板頁（4 頁）
        print(f"\n刪除模板頁...")
        for i in range(self.template_page_count - 1, -1, -1):
            rId = self.output_prs.slides._sldIdLst[i].rId
            self.output_prs.part.drop_rel(rId)
            del self.output_prs.slides._sldIdLst[i]
        
        # 儲存 PPT
        self.output_prs.save(self.output_path)
        print(f"\n✅ PPT 生成完成！")
        print(f"📊 總共生成 {len(self.output_prs.slides)} 張投影片")
        print(f"💾 已儲存到：{self.output_path}")


def main():
    """主程式"""
    if len(sys.argv) != 5:
        print("使用方式：")
        print("  python generate_ppt_from_template_v2.py template.pptx input.txt config.txt output.pptx")
        print()
        print("參數說明：")
        print("  template.pptx  - 模板 PPT（必須包含 4 頁）")
        print("  input.txt      - 輸入文字檔（包含變數和內容）")
        print("  config.txt     - 設定檔（定義頁面結構）")
        print("  output.pptx    - 輸出 PPT 檔名")
        sys.exit(1)
    
    template_path = sys.argv[1]
    input_path = sys.argv[2]
    config_path = sys.argv[3]
    output_path = sys.argv[4]
    
    print("=" * 60)
    print("PPT 生成程式 V2")
    print("=" * 60)
    print(f"模板檔案：{template_path}")
    print(f"輸入文字：{input_path}")
    print(f"設定檔案：{config_path}")
    print(f"輸出檔案：{output_path}")
    print("=" * 60)
    print()
    
    try:
        # 建立生成器（會先複製 template 到 output）
        generator = PPTGeneratorV2(template_path, output_path)
        
        # 載入變數和內容
        generator.load_variables_and_content(input_path)
        
        # 載入設定
        generator.load_config(config_path)
        
        # 生成 PPT
        generator.generate()
        
    except Exception as e:
        print(f"❌ 錯誤：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
