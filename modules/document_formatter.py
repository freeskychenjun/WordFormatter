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

    def _set_run_font(self, run, font_name, size_pt, set_color=False, is_bold=False):
        """设置单个run的字体属性"""
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.font.bold = is_bold

        if set_color:
            run.font.color.rgb = RGBColor(0, 0, 0)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:eastAsia'), font_name)

    def _apply_font_to_runs(self, para, font_name, size_pt, set_color=False, is_bold=False):
        """应用字体设置到段落的所有runs"""
        for run in para.runs:
            self._set_run_font(run, font_name, size_pt, set_color=set_color, is_bold=is_bold)

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
                return int(val)
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

        # 设置新的大纲级别 (Word内部用0-8表示1-9级)
        pPr = para._p.get_or_add_pPr()
        outlineLvl = pPr.find(qn('w:outlineLvl'))
        if outlineLvl is None:
            outlineLvl = OxmlElement('w:outlineLvl')
            pPr.append(outlineLvl)
        outlineLvl.set(qn('w:val'), str(level - 1))

        return original_level

    def _apply_text_indent_and_align(self, para):
        # 检查是否是图表标题，如果是则跳过，防止覆盖之前设置的缩进
        if hasattr(para, '_has_no_indent') and para._has_no_indent:
            return

        # 标题不缩进，确保完全移除所有缩进设置
        para.paragraph_format.first_line_indent = None
        para.paragraph_format.left_indent = Pt(0)
        para.paragraph_format.right_indent = Pt(0)
        para.paragraph_format.hanging_indent = Pt(0)
        # 不设置首行缩进
        # para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # 确保彻底清除所有缩进相关设置
        try:
            pPr = para._p.get_or_add_pPr()

            # 获取或创建缩进元素
            if pPr.find(qn('w:ind')) is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            else:
                ind = pPr.find(qn('w:ind'))

            # 清除所有可能的缩进属性
            for attr in ['w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars',
                         'w:right', 'w:rightChars', 'w:hanging', 'w:hangingChars']:
                if attr in ind.attrib:
                    del ind.attrib[attr]

            # 显式设置为0 - 使用不同单位确保彻底移除缩进
            ind.set(qn('w:firstLineChars'), '0')
            ind.set(qn('w:leftChars'), '0')
            ind.set(qn('w:rightChars'), '0')
            ind.set(qn('w:firstLine'), '0')
            ind.set(qn('w:left'), '0')
            ind.set(qn('w:right'), '0')
        except Exception as e:
            self._log(f"设置标题缩进时出错: {e}")

    def _apply_body_text_indent_and_align(self, para):
        # 正文首行缩进2字符
        # para.paragraph_format.left_indent = Cm(self.config['left_indent_cm'])
        # para.paragraph_format.right_indent = Cm(self.config['right_indent_cm'])
        # 设置首行缩进2字符（200表示2个字符）
        ind = para._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn("w:firstLineChars"), "200")
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    def _iter_block_items(self, parent):
        parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)