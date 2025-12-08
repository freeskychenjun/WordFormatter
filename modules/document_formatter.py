import logging
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

from .logger import global_logger


class DocumentFormatter:
    def __init__(self, config, log_callback=None):
        self.config = config
        self.log_callback = log_callback
        self.logger = global_logger

    def _log(self, message):
        if self.log_callback:
            self.log_callback(message)
        else:
            self.logger.info(message)

    def _set_run_font(self, run, font_name, size_pt, set_color=False, is_bold=False, use_times_roman_for_ascii=False):
        """设置单个run的字体属性
        
        参数:
            use_times_roman_for_ascii: 如果为True，将ASCII字符（英文、数字、符号）设置为Times New Roman字体
        """
        # 尝试多种字体名称设置方式
        # 1. 设置高级API的font.name属性
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.font.bold = is_bold

        if set_color:
            run.font.color.rgb = RGBColor(0, 0, 0)
        
        # 2. 直接设置XML的rFonts元素，确保所有字符类型都使用相同字体
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        
        # 确定要使用的字体
        ascii_font = "Times New Roman" if use_times_roman_for_ascii else font_name
        
        # 设置所有字符类型的字体
        rFonts.set(qn('w:ascii'), ascii_font)      # ASCII字符（英文、数字、符号）
        rFonts.set(qn('w:hAnsi'), ascii_font)      # 高ASCII字符
        rFonts.set(qn('w:eastAsia'), font_name)    # 中文字体
        
        # 3. 额外设置cs（复杂脚本）字体，确保所有语言都能正确显示
        rFonts.set(qn('w:cs'), font_name)
        
        # 4. 清除可能存在的字体主题设置，确保使用指定的字体
        if hasattr(rFonts, 'themeFont'):
            rFonts.themeFont = None
        if hasattr(rFonts, 'themeFontAscii'):
            rFonts.themeFontAscii = None
        if hasattr(rFonts, 'themeFontHAnsi'):
            rFonts.themeFontHAnsi = None
        if hasattr(rFonts, 'themeFontEastAsia'):
            rFonts.themeFontEastAsia = None
        if hasattr(rFonts, 'themeFontCs'):
            rFonts.themeFontCs = None

    def _set_run_font_without_size(self, run, font_name, set_color=False, use_times_roman_for_ascii=False):
        """设置单个run的字体属性，但不修改字体大小
        
        参数:
            use_times_roman_for_ascii: 如果为True，将ASCII字符（英文、数字、符号）设置为Times New Roman字体
        """
        # 尝试多种字体名称设置方式
        # 1. 设置高级API的font.name属性
        run.font.name = font_name
        # 不设置字体大小，保持原始大小不变
        # run.font.size = Pt(size_pt)  # 这行被注释掉，不设置字体大小
        # 不设置字体粗细，保持原始粗细不变
        # run.font.bold = is_bold  # 这行被注释掉，不设置字体粗细

        if set_color:
            run.font.color.rgb = RGBColor(0, 0, 0)
        
        # 2. 直接设置XML的rFonts元素，确保所有字符类型都使用相同字体
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        
        # 确定要使用的字体
        ascii_font = "Times New Roman" if use_times_roman_for_ascii else font_name
        
        # 设置所有字符类型的字体
        rFonts.set(qn('w:ascii'), ascii_font)      # ASCII字符（英文、数字、符号）
        rFonts.set(qn('w:hAnsi'), ascii_font)      # 高ASCII字符
        rFonts.set(qn('w:eastAsia'), font_name)    # 中文字体
        
        # 3. 额外设置cs（复杂脚本）字体，确保所有语言都能正确显示
        rFonts.set(qn('w:cs'), font_name)
        
        # 4. 清除可能存在的字体主题设置，确保使用指定的字体
        if hasattr(rFonts, 'themeFont'):
            rFonts.themeFont = None
        if hasattr(rFonts, 'themeFontAscii'):
            rFonts.themeFontAscii = None
        if hasattr(rFonts, 'themeFontHAnsi'):
            rFonts.themeFontHAnsi = None
        if hasattr(rFonts, 'themeFontEastAsia'):
            rFonts.themeFontEastAsia = None
        if hasattr(rFonts, 'themeFontCs'):
            rFonts.themeFontCs = None

    def _apply_font_to_runs(self, para, font_name, size_pt, set_color=False, is_bold=False, use_times_roman_for_ascii=False):
        """应用字体设置到段落的所有runs"""
        for run in para.runs:
            self._set_run_font(run, font_name, size_pt, set_color=set_color, is_bold=is_bold, use_times_roman_for_ascii=use_times_roman_for_ascii)

    def _get_paragraph_font_info(self, para):
        """获取段落主要字体和字号信息"""
        if not para.runs:
            return None, None

        # 获取第一个非空run的字体信息
        for run in para.runs:
            if run.text.strip():
                font_name = run.font.name
                font_size = run.font.size.pt if run.font.size else None
                return font_name, font_size
        return None, None

    def _strip_leading_whitespace(self, para):
        if not para.runs:
            return
        while para.runs and not para.runs[0].text.strip():
            p = para._p
            p.remove(para.runs[0]._r)
        if not para.runs:
            return
        first_run = para.runs[0]
        original_text = first_run.text
        stripped_text = original_text.lstrip()
        if original_text != stripped_text:
            first_run.text = stripped_text
            self._log("  > 已移除段落前的多余空格。")

    def _reset_pagination_properties(self, para):
        para.paragraph_format.widow_control = False
        para.paragraph_format.keep_with_next = False
        para.paragraph_format.keep_lines_together = False
        para.paragraph_format.page_break_before = False
        para.paragraph_format.keep_together = False

    def _get_outline_level(self, para):
        """
        读取段落的当前大纲级别
        返回: 0-8 表示级别1-9，None 表示未设置
        """
        pPr = para._p.get_or_add_pPr()
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is not None:
            val = outlineLvl.get(qn('w:val'))
            if val is not None:
                try:
                    level = int(val)
                    # 确保返回值在有效范围内
                    if 0 <= level <= 8:
                        return level
                except ValueError:
                    pass
        return None

    def _set_outline_level(self, para, level):
        """
        直接设置段落的大纲级别，不通过样式，不影响字体字号等格式
        level: 1-9 的整数，表示大纲级别
        返回: 原有的大纲级别 (0-8) 或 None
        """
        if level < 1 or level > 9:
            self._log(f"  > 警告：大纲级别 {level} 超出范围 (1-9)，已跳过设置")
            return None

        # 读取原有大纲级别
        original_level = self._get_outline_level(para)

        # 确保级别值正确（1-9）
        level_value = max(1, min(9, level))
        
        # 设置新的大纲级别 (Word内部用0-8表示1-9级)
        pPr = para._p.get_or_add_pPr()
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is None:
            outlineLvl = OxmlElement('w:outlineLvl')
            pPr.append(outlineLvl)
        # 确保设置的值在0-8范围内
        outlineLvl.set(qn('w:val'), str(level_value - 1))

        return original_level

    def _apply_text_indent_and_align(self, para):
        # 对于所有标题，无论之前是否标记过，都强制执行缩进清除
        # 首先通过python-docx的API清除所有缩进
        para.paragraph_format.first_line_indent = None
        para.paragraph_format.left_indent = Pt(0)
        para.paragraph_format.right_indent = Pt(0)
        para.paragraph_format.hanging_indent = Pt(0)
        
        # 确保彻底清除所有缩进相关设置 - 使用更健壮的方式
        try:
            pPr = para._p.get_or_add_pPr()
            
            # 完全移除现有的缩进元素，然后创建一个全新的
            existing_ind = pPr.find(qn('w:ind'))
            if existing_ind is not None:
                pPr.remove(existing_ind)
            
            # 创建一个全新的、干净的缩进元素
            new_ind = OxmlElement('w:ind')
            
            # 显式设置所有可能的缩进属性为0，使用多种单位确保彻底移除缩进
            new_ind.set(qn('w:firstLineChars'), '0')
            new_ind.set(qn('w:leftChars'), '0')
            new_ind.set(qn('w:rightChars'), '0')
            new_ind.set(qn('w:firstLine'), '0')
            new_ind.set(qn('w:left'), '0')
            new_ind.set(qn('w:right'), '0')
            new_ind.set(qn('w:hanging'), '0')
            new_ind.set(qn('w:hangingChars'), '0')
            
            # 添加到pPr元素
            pPr.append(new_ind)
            
            # 检查并移除可能影响缩进的其他元素
            for element_name in ['w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars',
                                'w:right', 'w:rightChars', 'w:hanging', 'w:hangingChars']:
                element = pPr.find(qn(element_name))
                if element is not None:
                    pPr.remove(element)
            
            # 设置标记，表明这个段落已经明确设置为无缩进
            para._has_no_indent = True
            
        except Exception as e:
            self._log(f"设置标题缩进时出错: {e}")
            # 即使发生异常，仍然尝试通过简单的API调用确保没有缩进
            try:
                para.paragraph_format.first_line_indent = None
                para.paragraph_format.left_indent = Pt(0)
                # 额外添加这一行，直接设置为负数以强制覆盖可能存在的缩进
                para.paragraph_format.left_indent = Pt(-0.01)
                # 然后再设置为0
                para.paragraph_format.left_indent = Pt(0)
            except:
                pass

    def _apply_body_text_indent_and_align(self, para):
        # 正文首行缩进2字符
        # para.paragraph_format.left_indent = Cm(self.config['left_indent_cm'])
        # para.paragraph_format.right_indent = Cm(self.config['right_indent_cm'])
        # 设置首行缩进2字符（200表示2个字符）
        ind = para._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn("w:firstLineChars"), "200")
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # 确保没有其他样式影响缩进
        try:
            # 清除可能存在的大纲级别设置，确保正文段落不受标题样式影响
            pPr = para._p.get_or_add_pPr()
            outlineLvl = pPr.find(qn('w:outlineLvl'))
            if outlineLvl is not None:
                pPr.remove(outlineLvl)
        except Exception as e:
            self._log(f"清除正文段落大纲级别时出错: {e}")

    def _iter_block_items(self, parent):
        parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)