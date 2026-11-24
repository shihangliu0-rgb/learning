# -*- coding: utf-8 -*-
"""《机器人学导论》期末复习精要 — Word 文档生成器
知识点梳理 + 习题详解（含解题步骤与答案）+ 速记表
内容来源：课程 PPT（第1/2/3/4/5章）与《机器人学导论》（贠超、王伟，机械工业出版社，2023）
"""
import copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = '1F4E79'      # 深蓝（章节色）
DARK = '262626'      # 正文深灰
RED = 'C00000'       # 强调红
KP_HL = 'FFF2A8'     # 考点高亮（淡黄）——比纯黄更耐看
LIGHT = 'DEEAF6'     # 浅蓝底（表头）
GRAY = 'F2F2F2'      # 浅灰底

_CN_BODY = '宋体'
_CN_HEAD = '黑体'
_LATIN = 'Times New Roman'
_MONO = 'Consolas'


def set_run_font(run, cn=_CN_BODY, latin=_LATIN, size=None, bold=None, color=None, italic=None):
    run.font.name = latin
    run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    if size is not None: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if color is not None: run.font.color.rgb = RGBColor.from_string(color)
    if italic is not None: run.font.italic = italic


def set_highlight(run, hex_color):
    """给 run 设置任意十六进制底色（w:highlight 仅支持预设色，改用 w:shd 字符底纹）。"""
    rPr = run._element.get_or_add_rPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_color)
    rPr.append(shd)


def set_para_shading(paragraph, hex_fill):
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_fill)
    pPr.append(shd)


def set_para_bottom_border(paragraph, color=NAVY, sz=18, space=6, style='single'):
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = pPr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = OxmlElement('w:pBdr'); pPr.append(pbdr)
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), style); bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), str(space)); bottom.set(qn('w:color'), color)
    pbdr.append(bottom)


def set_table_borders(table, color='BFBFBF', sz=6, inside=True, left_right=True, top_bottom=True):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    def edge(tag, on):
        el = OxmlElement('w:%s' % tag)
        if on:
            el.set(qn('w:val'), 'single'); el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0'); el.set(qn('w:color'), color)
        else:
            el.set(qn('w:val'), 'none'); el.set(qn('w:sz'), '0'); el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto')
        borders.append(el)
    edge('top', top_bottom); edge('bottom', top_bottom)
    edge('left', left_right); edge('right', left_right)
    edge('insideH', inside); edge('insideV', inside)
    tblPr.append(borders)


def set_cell_shading(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_fill)
    tcPr.append(shd)


def set_col_widths(table, widths):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is not None:
        for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
            gc.set(qn('w:w'), str(int(w * 567)))
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths):
                cell.width = Cm(widths[i])


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    trPr.append(tblHeader)


def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for tag, val in (('top', top), ('bottom', bottom), ('start', left), ('end', right)):
        el = OxmlElement('w:%s' % tag)
        el.set(qn('w:w'), str(val)); el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tcPr.append(mar)


def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:cantSplit')
    trPr.append(el)


def add_field_run(paragraph, fieldcode):
    fld_begin = OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = fieldcode
    fld_end = OxmlElement('w:fldChar'); fld_end.set(qn('w:fldCharType'), 'end')
    r1 = paragraph.add_run(); r1._r.append(fld_begin)
    r2 = paragraph.add_run(); r2._r.append(instr)
    r3 = paragraph.add_run(); r3._r.append(fld_end)
    return (r1, r2, r3)


class Doc:
    def __init__(self):
        self.doc = Document()
        self._setup_page()
        self._setup_styles()

    # ---------- 页面与样式 ----------
    def _setup_page(self):
        for section in self.doc.sections:
            section.page_height, section.page_width = Cm(29.7), Cm(21.0)
            section.top_margin, section.bottom_margin = Cm(2.0), Cm(1.9)
            section.left_margin, section.right_margin = Cm(1.9), Cm(1.9)
            section.header_distance, section.footer_distance = Cm(1.0), Cm(1.0)

    def _setup_styles(self):
        styles = self.doc.styles
        normal = styles['Normal']
        normal.font.name = _LATIN
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = RGBColor.from_string(DARK)
        normal._element.rPr.rFonts.set(qn('w:eastAsia'), _CN_BODY)
        pf = normal.paragraph_format
        pf.space_before = Pt(2); pf.space_after = Pt(4)
        pf.line_spacing = 1.25
        for name, size, color, before, after in [
            ('Heading 1', 17, NAVY, 18, 10),
            ('Heading 2', 13.5, NAVY, 12, 6),
            ('Heading 3', 11.5, DARK, 8, 4),
            ('Heading 4', 10.5, DARK, 6, 3),
        ]:
            st = styles[name]
            st.font.name = _LATIN; st.font.bold = True
            st.font.size = Pt(size); st.font.color.rgb = RGBColor.from_string(color)
            st._element.rPr.rFonts.set(qn('w:eastAsia'), _CN_HEAD)
            st.paragraph_format.space_before = Pt(before); st.paragraph_format.space_after = Pt(after)
            st.paragraph_format.keep_with_next = True
            st.paragraph_format.line_spacing = 1.2

    # ---------- 基础构件 ----------
    def add_field_simple(self, paragraph, fieldcode):
        fld_begin = OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = fieldcode
        fld_end = OxmlElement('w:fldChar'); fld_end.set(qn('w:fldCharType'), 'end')
        r1 = paragraph.add_run(); r1._r.append(fld_begin)
        r2 = paragraph.add_run(); r2._r.append(instr)
        r3 = paragraph.add_run(); r3._r.append(fld_end)

    def body_par(self, text=None, size=10.5, align=None, space_after=4, space_before=2):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        if align is not None: p.alignment = align
        if text is not None:
            r = p.add_run(text)
            set_run_font(r, size=size)
        return p

    def h1(self, text, pagebreak=False):
        if pagebreak:
            self.doc.add_page_break()
        p = self.doc.add_heading(text, level=1)
        set_para_bottom_border(p, NAVY, sz=16, space=8)
        return p

    def h2(self, text):
        return self.doc.add_heading(text, level=2)

    def h3(self, text):
        return self.doc.add_heading(text, level=3)

    def h4(self, text):
        return self.doc.add_heading(text, level=4)

    def bullets(self, items, indent=0.75, hang=True):
        """items: list of either str, or (segments) where segments = list of (text, kind).
        kind: '' normal | 'b' bold | 'kp' 考点(bold+highlight) | 'i' italic | 'red' red bold"""
        for item in items:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(indent)
            p.paragraph_format.first_line_indent = Cm(-0.35 if hang else 0)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.25
            r = p.add_run('▪  ')
            set_run_font(r, size=9, color=NAVY, bold=True)
            self._add_segments(p, item)

    def _add_segments(self, p, item):
        if isinstance(item, str):
            item = [(item, '')]
        for text, kind in item:
            r = p.add_run(text)
            set_run_font(r, size=10.5)
            if kind == 'b':
                r.font.bold = True
            elif kind == 'kp':
                r.font.bold = True
                set_highlight(r, KP_HL)
            elif kind == 'kpblue':
                r.font.bold = True
                r.font.color.rgb = RGBColor.from_string(NAVY)
                set_highlight(r, KP_HL)
            elif kind == 'i':
                r.font.italic = True
            elif kind == 'red':
                r.font.bold = True; r.font.color.rgb = RGBColor.from_string(RED)
            elif kind == 'sub':
                r.font.subscript = True; r.font.italic = True
            elif kind == 'sup':
                r.font.superscript = True; r.font.italic = True
            elif kind == 'subn':
                r.font.subscript = True
            elif kind == 'supn':
                r.font.superscript = True

    def math_par(self, segments, align_center=True, size=10.5, space_after=4):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(space_after)
        if align_center: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._add_segments(p, segments)
        for r in p.runs:
            r.font.name = _LATIN
        return p

    def note(self, text_segments, fill=GRAY):
        """浅灰底提示框"""
        p = self.body_par(space_after=6, space_before=4)
        set_para_shading(p, fill)
        self._add_segments(p, text_segments if isinstance(text_segments, list) else [(text_segments, '')])
        return p

    def mono_block(self, lines, caption=None):
        """等宽字体公式块（矩阵等）"""
        for i, line in enumerate(lines):
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1.2)
            p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            if i == 0: set_para_shading(p, 'F7F8FA')
            if caption and i == 0:
                pass
            r = p.add_run(line)
            set_run_font(r, cn=_MONO, latin=_MONO, size=9)
        self.body_par(space_after=4)

    def table(self, headers, rows, widths, header_fill=LIGHT, size=9.5, align_center=False):
        t = self.doc.add_table(rows=len(rows) + 1, cols=len(headers))
        for i, htext in enumerate(headers):
            cell = t.cell(0, i)
            set_cell_shading(cell, header_fill)
            set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(1); p.paragraph_format.line_spacing = 1.15
            r = p.add_run(htext)
            set_run_font(r, cn=_CN_HEAD, size=size, bold=True, color=NAVY)
        for ri, row in enumerate(rows):
            for i, cellval in enumerate(row):
                cell = t.cell(ri + 1, i)
                set_cell_margins(cell)
                if ri % 2 == 1:
                    set_cell_shading(cell, 'FAFBFD')
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(1); p.paragraph_format.line_spacing = 1.18
                if isinstance(cellval, list):
                    self._add_segments(p, cellval)
                    for r in p.runs: set_run_font(r, size=size)
                else:
                    r = p.add_run(str(cellval))
                    set_run_font(r, size=size)
        set_table_borders(t, color='C9D3E0', sz=6, inside=True)
        set_col_widths(t, widths)
        set_repeat_table_header(t.rows[0])
        for row in t.rows: set_row_cant_split(row)
        self.body_par(space_after=6)
        return t


def add_header_footer(doc):
    for section in doc.sections:
        # header
        hp = section.header.paragraphs[0]
        hp.alignment = 2  # right
        r = hp.add_run('机器人学导论 · 期末复习精要')
        set_run_font(r, cn=_CN_HEAD, size=8, color='808080')
        set_para_bottom_border(hp, 'D9D9D9', sz=6, space=2)
        # footer: left title, right page numbers
        fp = section.footer.paragraphs[0]
        fp.alignment = 1  # center
        r1 = fp.add_run('《机器人学导论》期末复习精要   ·   第 ')
        set_run_font(r1, size=8, color='808080')
        add_field_run(fp, 'PAGE')
        r2 = fp.add_run(' 页 / 共 ')
        set_run_font(r2, size=8, color='808080')
        add_field_run(fp, 'NUMPAGES')
        r3 = fp.add_run(' 页')
        set_run_font(r3, size=8, color='808080')


def build_cover(d):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm as _Cm
    # 顶部小字
    p = d.body_par(space_before=6, space_after=0)
    r = p.add_run('机 器 人 学 · 课 程 复 习 资 料')
    set_run_font(r, cn=_CN_HEAD, size=10, color='A6A6A6', bold=True)
    p2 = d.body_par(space_after=2)
    set_para_bottom_border(p2, NAVY, sz=22, space=10)

    # 主标题
    p = d.doc.add_paragraph()
    p.paragraph_format.space_before = Pt(70); p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('机器人学导论')
    set_run_font(r, cn=_CN_HEAD, size=40, bold=True, color=NAVY)

    p = d.doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run('期 末 复 习 精 要')
    set_run_font(r, cn=_CN_HEAD, size=23, bold=True, color=DARK)

    p = d.doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('— 知识点梳理　·　习题详解　·　速记表 —')
    set_run_font(r, size=12, color='595959')

    # 中部说明框
    d.body_par(space_before=26, space_after=2)
    tbl = d.doc.add_table(rows=4, cols=2)
    info = [
        ('参考教材', '《机器人学导论》，贠超、王伟，机械工业出版社，2023'),
        ('内容依据', '课程 PPT（第1章 机器人基础知识 · 第2章 机器人的机械结构 · 第3章 机器人的运动学和动力学 · 第4章 机器人雅可比矩阵 · 第5章 机器人的轨迹规划）'),
        ('覆盖范围', '第1～5章知识点梳理；各章 PPT 例题与“思考与习题 / 思考与练习”的解题步骤与答案；期末速记表'),
        ('编制日期', '2026年11月'),
    ]
    for i, (k, v) in enumerate(info):
        c0, c1 = tbl.cell(i, 0), tbl.cell(i, 1)
        set_cell_shading(c0, GRAY)
        set_cell_margins(c0, top=90, bottom=90); set_cell_margins(c1, top=90, bottom=90)
        p0 = c0.paragraphs[0]; p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k); set_run_font(r0, cn=_CN_HEAD, size=10, bold=True, color=NAVY)
        p1 = c1.paragraphs[0]; p1.paragraph_format.space_after = Pt(0); p1.paragraph_format.line_spacing = 1.3
        r1 = p1.add_run(v); set_run_font(r1, size=9.5)
    set_table_borders(tbl, color='D9D9D9', sz=6, inside=True)
    set_col_widths(tbl, [3.2, 13.2])

    # 图例说明
    d.body_par(space_before=18)
    p = d.body_par(space_before=2)
    set_para_shading(p, 'FFFDF0')
    d._add_segments(p, [('使用说明　', 'b'),
                        ('正文列出各章知识点（直接列表）；', ''),
                        ('加粗 + 淡黄底纹', 'kp'),
                        ('标记为高频考点；', ''),
                        ('红色加粗', 'red'),
                        ('标记为易错 / 易混点；★ 表示课堂强调。习题部分每题均给出解题步骤与答案，可直接用于自测对照。', '')])
    d.body_par(space_after=0)
    p = d.body_par()
    d._add_segments(p, [('声明　', 'b'),
                        ('本资料由课程 PPT 内容整理而成，仅供课程复习参考，请以教材与课堂讲授为准。', '')])


def build_toc(d):
    d.h1('目　录')
    toc_items = [
        ('第一部分', '知识点梳理', [
            '第1章 机器人基础知识',
            '第2章 机器人的机械结构',
            '第3章 机器人的运动学和动力学',
            '第4章 机器人雅可比矩阵',
            '第5章 机器人的轨迹规划',
        ]),
        ('第二部分', '习题详解', [
            '第1章 思考与习题（填空 · 判断 · 简答）',
            '第2章 思考与习题（填空 · 判断 · 简答）',
            '第3章 思考与练习（填空 · 判断 · 计算 · 简答）',
            '第4章 思考与练习（简答）',
            '第5章 思考与练习（判断 · 简答 · 计算）',
        ]),
        ('第三部分', '速记表', [
            '速记 1–4  机器人四大家族 · 坐标形式速查 · 机身/臂驱动 · 谐波 vs RV',
            '速记 5–7  齐次变换口诀 · 雅可比矩阵 · 三次/五次多项式系数',
            '速记 8–9  易混点辨析 · 例题数字一览（直接背答案）',
            '速记 10–12 判断题高频 · 答题模板 · 考前 5 分钟 checklist',
        ]),
    ]
    for part, title, chapters in toc_items:
        p = d.body_par(space_before=10, space_after=2)
        r = p.add_run(part + '　')
        set_run_font(r, cn=_CN_HEAD, size=13, bold=True, color=RED)
        r = p.add_run(title)
        set_run_font(r, cn=_CN_HEAD, size=13, bold=True, color=NAVY)
        for ch in chapters:
            p = d.body_par(space_after=6)
            p.paragraph_format.left_indent = Cm(0.8)
            r = p.add_run('▸  ')
            set_run_font(r, size=10, color=NAVY, bold=True)
            r = p.add_run(ch)
            set_run_font(r, size=11, color=DARK)
    d.note([('阅读建议　', 'b'), ('先通读“第一部分 知识点梳理”建立框架，再用“第二部分 习题详解”自测；考前对照“第三部分 速记表”集中记忆公式与易混点。', '')])


def part_title(d, num, title, subtitle):
    d.h1('%s　%s' % (num, title), pagebreak=False)
    p = d.body_par(space_after=10)
    r = p.add_run(subtitle)
    set_run_font(r, size=9.5, color='7F7F7F')


# =====================================================================
# 第一部分：知识点梳理
# =====================================================================

def build_part1(d):
    part_title(d, '第一部分', '知识点梳理', '按 PPT 章节顺序整理；加粗 + 淡黄底纹 = 考点；红色加粗 = 易错/易混')
    build_ch1(d)
    build_ch2(d)
    build_ch3(d)
    build_ch4(d)
    build_ch5(d)


def build_ch1(d):
    d.h1('第1章 机器人基础知识', pagebreak=True)
    d.h2('1.1 机器人的发展简史')
    d.bullets([
        [('1920年，捷克作家卡雷尔·卡佩克在科幻剧本《洛桑万能机器人》中首次使用“机器人”一词，源自捷克语 Robota（奴隶、苦力）。', '')],
        [('1950年，科幻作家阿西莫夫在《我是机器人》中提出', ''), ('“机器人三原则”', 'kp'), ('：①机器人不能伤害人类；②机器人应遵守人类的命令（与第一条相违背的命令除外）；③机器人应能保护自己（与第一条相抵触的命令除外）。', '')],
        [('1954年，美国', ''), ('戴沃尔', 'kp'), ('最早提出工业机器人概念并申请专利。', '')],
        [('1959年，恩格尔伯格（“机器人之父”）研制出', ''), ('世界上第一台工业机器人 Unimation', 'kp'), ('（液压驱动），建立 Unimation 公司（1958）；1983年转战服务机器人，创立 TRC 公司，推出“护士助手”机器人。', '')],
        [('20世纪60年代：液压驱动；70年代：计算机控制 + 电机驱动；80年代：技术趋于成熟；90年代：Delta 机器人，以', ''), ('在线示教', 'b'), ('为主。', '')],
        [('我国古代：鲁班“飞鸟”、偃师“伶人”（最早记载）、张衡“记里鼓车”、诸葛亮“木牛流马”。', '')],
    ])
    d.bullets([
        [('机器人发展的三个阶段：', 'b'), ('①示教再现型机器人（第1代，20世纪60年代起，广泛应用）→ ②有感觉的机器人（第2代，具有不同程度感知能力）→ ③智能机器人（第3代，具备识别、推理、规划、学习能力，把感知与行动智能化结合）。', '')],
        [('全球工业机器人“四大家族”：', ''), ('瑞士 ABB、德国 KUKA、日本 FANUC、日本 YASKAWA（安川）', 'kp'), ('。中国：中国制造2025；美国：机器人发展线路图（2013）；欧盟：Horizon 计划；日本：《机器人新战略》（2015）。', '')],
    ])

    d.h2('1.2 机器人的定义与特点')
    d.h3('1.2.1 四种权威定义')
    d.table(['机构', '定义要点'], [
        ['美国机器人协会（RIA）', '用于移动各种材料、零件、工具或专用装置，通过可编程序动作来执行各种任务的、具有编程能力的多功能机械手（manipulator）'],
        ['日本工业机器人协会（JIRA）', '装备有记忆装置和末端执行器，能够转动并通过自动完成各种移动来代替人类劳动的通用机器'],
        ['美国国家标准局（NBS）', '能够进行编程并在自动控制下执行某些操作和移动作业任务的机械装置（较广义的工业机器人定义）'],
        ['国际标准化组织（ISO）', '自动的、位置可控的、具有编程能力的多功能机械手；具有几个轴，可处理材料、零件、工具和专用装置，执行种种任务'],
    ], [4.6, 12.2], size=9)
    d.bullets([
        [('我国科学家的定义：机器人是一种自动化的机器，具备与人或生物相似的', ''), ('感知能力、规划能力、动作能力和协同能力', 'kp'), ('，是一种具有高度灵活性的自动化机器。', '')],
        [('机器人共同点：', 'b'), ('①外形像人或像人的上肢，能模仿人的动作；②是人造的机器或机电装置；③具有一定的智力、感觉与识别性。', '')],
    ])
    d.h3('1.2.2 机器人的特点（四点）')
    d.bullets([
        [('可编程', 'kp'), ('：可随工作环境变化再编程，是柔性制造系统（FMS）的重要组成部分。', '')],
        [('拟人化', 'kp'), ('：机械结构类似人的行走、腰转、大臂、小臂、手腕、手爪；控制上有“电脑”；智能化机器人还有类似人类的“生物传感器”，提高自适应能力。', '')],
        [('通用性', 'kp'), ('：①机构可变（改变末端操作器、自由度数量）；②程序可变；能完成多样的简单任务。', '')],
        [('适应性', 'kp'), ('：①自动化（不依赖人的干预）；②智能化（记忆、语言理解、图像识别、判断推理能力）；能感测环境、自主执行未经完全指定的任务。', '')],
    ])

    d.h2('1.3 机器人的分类')
    d.h3('1.3.1 按坐标形式分类（★★ 高频填空考点）')
    d.table(['类型', '代号', '特点'], [
        ['直角坐标型', 'PPP', '位置精度高、空间轨迹易于求解；作用范围较小、结构庞大、无法调节工具姿态、灵活性差'],
        ['圆柱坐标型', 'RPP', '控制简单；结构庞大；沿轴线方向需移动空间，空间利用率低；难与其他机器人协调工作，两移动轴设计复杂'],
        ['球坐标型', 'RRP', '占地面积小、结构紧凑；能与其他机器人协调工作；避障性差、有平衡问题，位置误差与臂长有关'],
        ['关节坐标型', 'RRR', '由前、后臂俯仰及立柱回转构成；结构最紧凑、灵活性大、占地面积最小、工作空间最大，能与其他机器人协调工作，应用最多'],
    ], [2.8, 2.0, 12.0], size=9)
    d.bullets([
        [('结论：', 'b'), ('直角坐标型精度高但庞大；', ''), ('关节坐标型应用最广', 'kp'), ('；填空“四类坐标形式” = 直角、圆柱、球、关节坐标型。', '')],
    ])
    d.h3('1.3.2 按控制方式分类（JIRA）')
    d.bullets([
        [('①人工操作装置 → ②固定顺序机器人 → ③可变顺序机器人 → ④示教再现（playback）机器人 → ⑤感知机器人 → ⑥智能机器人。（美国标准强调“易于再编程的装置”。）', '')],
    ])
    d.h3('1.3.3 按应用环境分类')
    d.table(['类型', '要点'], [
        ['工业机器人', '模拟人的手、臂动作，按预定程序、轨迹及其他要求，实现抓取、搬运工件或操作的自动化装置；可通过更改程序快速改变工作内容（柔性）'],
        ['服务机器人', '多为可移动性，移动平台 + 手臂操作 + 力觉/视觉/超声测距传感器；如迎宾、扫地、导盲、手术机器人（达芬奇）'],
        ['特种机器人', '用于人难以进入的特殊环境（核电站、海底、宇宙空间）；军用、消防、保安、无人机、建筑、农业机器人'],
        ['智能机器人', '具备内外部传感器感觉系统：感知内部（关节速度、力）参数和外部环境信息，自主决策完成任务；类人型 vs 非类人型'],
    ], [2.6, 14.2], size=9)

    d.h2('1.4 机器人的基本组成')
    d.bullets([
        [('工业机器人系统一般由', ''), ('三部分、六个子系统', 'kp'), ('组成。', '')],
    ])
    d.table(['子系统', '核心内容'], [
        ['驱动系统', '驱动机械系统的装置；按驱动源分为电动、液压、气动；可直接相连，也可经同步带、链条、齿轮、谐波减速器等间接相连'],
        ['机械系统', '腰部（机身）+ 臂部 + 腕部 + 手部（末端执行器）'],
        ['感知系统', [('内部传感器', 'b'), ('：关节位置/速度/力/扭矩传感器；', ''), ('外部传感器', 'b'), ('：听觉、视觉、触觉传感器', '')]],
        ['机器人-环境交互系统', '实现机器人与外部设备相互联系和协调的系统'],
        ['控制系统', '控制计算机 + 示教编程器 + 操作面板；根据指令程序和传感器反馈信号支配执行机构完成运动；程序控制（开环）/ 自适应控制（闭环）/ 智能控制（神经网络、遗传/蚁群算法、模糊控制）；按运动形式分为点位控制和轨迹控制'],
        ['人机交互系统', '使操作人员参与机器人控制并与之联系的装置；标准终端、指令控制台、信息显示板、危险信号报警器；交互方式：指令、语音、动作/手势、眼动、触觉/力觉、脑机接口'],
    ], [2.8, 14.0], size=9)
    d.bullets([
        [('电动驱动（对比记忆）：', 'b'), ('①三相异步电机（结构简单、价格低、但不能精准定位）；②步进电机（脉冲驱动、断电自锁、但高速丢步/力矩衰减）；③直流伺服电机（有刷：启动力矩大、调速平滑；无刷：寿命长、成本高）；④交流伺服电机（编码器闭环反馈，定位最精准，用于机器人关节、数控机床）；⑤关节模组（电机+减速器+驱动器+编码器高度集成，成本可占整机 50% 以上；分谐波/行星/摆线关节模组）。', '')],
    ])

    d.h2('1.5 机器人的技术参数')
    d.table(['参数', '定义 / 要点'], [
        ['自由度', '独立坐标轴运动的数目（不含手爪/末端执行器的自由度），衡量操作灵活性；越高越灵活'],
        ['工作空间', '手臂末端或手腕中心所能到达的所有点的集合（工作区域）；含灵巧点（任意姿态可达）与非灵巧点（边界极限位置）'],
        ['承载能力', '工作范围内任意位姿上能承受的最大质量；含末端执行器质量；与运行速度、加速度的大小和方向有关；指标为高速运行时的承载能力'],
        ['精度', '位置量的绝对度量：手部/参考点实际到达位置与理论位置的差距；影响：机械误差（传动、关节间隙、配合、轴承）、控制算法误差、分辨率系统误差'],
        ['重复定位精度', '相同指令下连续若干次相同运动轨迹间的度量（统计量）；工业机器人重复定位精度应高于绝对定位精度'],
        ['最大工作速度', '速度越高效率越高，但升降速耗时、对最大加/减速度要求越高'],
        ['其他', '控制方式、驱动方式、安装方式、动力源容量、本体质量、环境参数、安全注意事项'],
    ], [3.0, 13.8], size=9)
    d.note([('考点　', 'red'), ('“自由度不包括手爪自由度”“工作空间含灵巧点/非灵巧点”“承载能力为高速运行指标且含末端执行器质量”“重复定位精度高于绝对定位精度”——均为填空/简答高频点。', '')])


def build_ch2(d):
    d.h1('第2章 机器人的机械结构')
    d.h2('2.1 机身结构')
    d.bullets([
        [('机身是直接连接、支承和传动手臂及行走机构的部件，是机器人的基础部分（固定式或移动式），由臂部运动（升降、平移、回转和俯仰）机构、导向装置和支撑件组成。', '')],
        [('四类机身结构（对应四种坐标类型）：', 'b')],
        [('①回转与升降型（圆柱坐标型）', 'b'), ('：实现臂部回转 + 升降；升降在下回转在上时回转缸驱动力矩可小，反之升降缸需加大；常用丝杆螺母、蜗轮蜗杆、液压缸、单/双杠活塞缸 + 同步带轮/链轮。', '')],
        [('②回转与俯仰型（球坐标型）', 'b'), ('：手臂左右回转 + 上下俯仰，用俯仰部件代替升降部件（俯仰缸 + 连杆机构）。', '')],
        [('③直移型（直角坐标型）', 'b'), ('：多为悬挂式，机身即悬挂手臂的横梁；导轨是关键构件。', '')],
        [('④类人型', 'b'), ('：机身上还装有驱动腿部运动的装置和腰部关节。', '')],
    ])

    d.h2('2.2 臂部结构')
    d.bullets([
        [('臂部（手臂部件）是机器人的主要执行部件，作用是', ''), ('支承腕部和手部，并带动它们在空间运动', 'kp'), ('；臂部的运动和参数决定腕部的空间位置及工作空间。', '')],
        [('组成：臂杆 + 与其伸缩/屈伸/自转有关的构件（传动机构、驱动装置、导向定位装置、支撑联接、位置检测元件）。', '')],
        [('分类：伸缩型、转动伸缩型、屈伸型、其他专用传动臂部。', '')],
    ])
    d.h3('2.2.1 机身和臂部的配置形式（填空考点）')
    d.bullets([
        [('横梁式配置', 'b'), ('（单臂悬挂式 / 双臂悬挂式）：悬挂手臂部件，占地面积小、有效利用空间、动作简单直观。', '')],
        [('立柱式配置', 'b'), ('（单臂式 / 双臂式）：回转型、俯仰型或屈伸型运动；占地面积小而工作范围大。', '')],
        [('机座（基座）式配置', 'b'), ('（固定式 / 移动式）：单臂回转式、双臂回转式、多臂回转式。', '')],
        [('屈伸式配置', 'b'), ('（平面屈伸式 / 立体屈伸式）：大小臂组成，大小臂间相对运动，回转式臂部的基础形式。', '')],
    ])
    d.h3('2.2.2 臂部运动方式')
    d.bullets([
        [('手臂直线运动机构', 'b'), ('：活塞缸、齿轮齿条、丝杆螺母。', '')],
        [('臂部俯仰机构', 'b'), ('：一般采用活塞（气）缸 + 连杆机构。', '')],
        [('臂部回转与升降机构', 'b'), ('：回转缸与升降缸单独驱动（升降行程短、回转角度<360°）；或升降缸 + 气动马达-锥齿轮传动。', '')],
    ])
    d.h3('2.2.3 机身和臂部设计注意事项')
    d.bullets([
        [('刚度', 'b'), ('：抵抗弹性变形能力；合理选择截面形状与轮廓尺寸（封闭空心截面抗弯抗扭 > 实心/开口截面；工字钢抗弯 > 圆截面）；提高支撑刚度和接触刚度；合理布置作用力位置减小弯曲变形。', '')],
        [('精度', 'b'), ('：装配精度、导向精度、刚度、耐磨性。', '')],
        [('平稳性', 'b'), ('：结构紧凑、质量轻、减小惯性力；注意重心位置。', '')],
        [('其他', 'b'), ('：传动链力求简短，元件布置紧凑合理。', '')],
    ])

    d.h2('2.3 腕部结构')
    d.bullets([
        [('位置：机器人手部与臂部之间；作用：支撑手部，', ''), ('调节或改变手部的作业方向', 'kp'), ('（姿态调整），具有独立自由度，满足复杂姿态调整。', '')],
    ])
    d.h3('2.3.1 腕部的运动方式（填空考点）')
    d.table(['方式', '定义'], [
        ['臂转', '腕部绕小臂轴线方向旋转，也称腕部旋转（对应翻转关节 R）'],
        ['腕摆', '手部绕垂直小臂轴线方向旋转；分为俯仰（pitch）和偏转（yaw）；同时具有俯仰和偏转的称作双腕摆（对应俯仰/偏转关节 B）'],
        ['手转', '手部绕自身轴线方向旋转（对应移动关节 T 之外的自转）'],
    ], [2.2, 14.6], size=9.5)
    d.bullets([
        [('常用组合方式：', ''), ('臂转—腕摆—手转', 'kp'), ('结构；', ''), ('臂转—双腕摆—手转', 'kp'), ('结构。', '')],
        [('轴线与手臂纵轴线共线的腕关节称为', ''), ('滚转关节（R）', 'kp'), ('；垂直的称为', ''), ('俯仰关节（B）/偏转关节（B）', 'kp'), ('（垂直于手臂轴线者）。', '')],
    ])
    d.h3('2.3.2 腕部按自由度数分类')
    d.bullets([
        [('单自由度手腕：R（翻转）、B（俯仰）、B（偏转）、T（移动）。', '')],
        [('二自由度手腕：BR、BB、RR 关节。', '')],
        [('三自由度手腕（最常见：实现空间任意姿态）。', '')],
    ])
    d.bullets([
        [('柔顺手腕', 'b'), ('：用于精密装配（轴孔装配），补偿零件不一致、定位夹具与定位精度误差。主动柔顺 = 边校正边装配；', ''), ('被动柔顺 = 在腕部配置可调角度柔顺环节', 'kp'), ('（如移动摆动柔顺手腕）。', '')],
    ])

    d.h2('2.4 手部结构（末端执行器）')
    d.bullets([
        [('手部特点（四点）：', 'b'), ('①与手腕相连处可拆卸（手部结构随夹持对象不同而有差异，通常一机多手，要求接头通用性、互换性）；②是机器人末端操作器（类人手爪，或喷枪、焊枪等专用工具）；③通用性比较差（一种手爪往往只能抓握一种或几种形状/尺寸/重量相近似的工件，一种工具只能执行一种作业）；④是一个独立的部件（机身、手臂、手部——机器人机械系统三大件）。', '')],
    ])
    d.h3('2.4.1 手部的分类')
    d.table(['分类方式', '类型'], [
        ['按用途', '手爪（有一定通用性：抓住—握持—释放工件）；专用操作器（工具，如喷枪、焊枪）'],
        ['按夹持方式', '外夹式、内撑式、内外夹持式'],
        ['按工作原理', '夹持类（摩擦力夹持、吊钩承重）；吸附类（磁力吸盘—铁磁材料；真空吸盘—表面平整规则板材；静电吸盘—薄型纤细件；范德华力吸盘—仿壁虎，3C 微器件）'],
        ['按手指/吸盘数目', '二指手爪；多指手爪'],
        ['按智能化', '普通式手爪（无传感器）；智能化手爪（集成力/触觉/滑觉传感器）'],
    ], [3.2, 13.6], size=9)
    d.h3('2.4.2 夹持式手部')
    d.bullets([
        [('夹钳式手部（组成：手指 + 传动机构 + 驱动机构 + 连接与支承元件）', 'b'), ('：', '')],
        [('手指：直接与工件接触；指端有 V 形指、平面指、尖指、特形指。', '')],
        [('传动机构（向手指传递运动动力，实现夹紧/松开）：回转型（一支点 / 多支点回转；斜楔杠杆式、滑槽式杠杆回转型、双支点连接杠杆式、齿条齿轮杠杆式；又分摆动回转型和平动回转型）、平移型（平面平行移动机构、直线往复移动机构，常用于夹持箱体等具有平行平面的工件）。', '')],
        [('钩托式手部：不靠夹紧力，利用手指钩、托、捧动作托持工件（无驱动 / 有驱动）。', '')],
        [('弹簧式手部：靠弹簧力夹紧工件，无需专用驱动装置，适于夹持轻小工件。', '')],
    ])
    d.h3('2.4.3 吸附式手部')
    d.table(['类型', '原理 / 适用'], [
        ['真空吸附', '碟形橡胶吸盘贴合 + 真空抽气形成负压；放料通大气失去真空。结构简单可靠、对工件表面无损伤，适用非金属板材（表面平整光滑、无孔凹槽）'],
        ['气流负压', '压缩空气高速流经喷嘴出口处气压低于吸盘腔内气压，腔内气体被带走形成负压；切断气源即释放'],
        ['挤压排气', '吸盘压紧工件挤出腔内空气，上升靠橡胶恢复力形成负压；压下拉杆通大气释放'],
        ['磁吸附', '永久磁铁/电磁铁通电产生电磁吸力；只能吸附铁磁物料；某些不允许有剩磁的零件禁用'],
    ], [2.6, 14.2], size=9)
    d.bullets([
        [('仿生多指灵巧手：柔性手（牵引钢丝绳 + 摩擦滚轮，两根钢丝绳一拉一松）；多指灵巧手（每手指 3 个回转关节、独立控制，能模仿拧螺钉、弹钢琴等复杂动作）。', '')],
    ])

    d.h2('2.5 驱动与传动')
    d.h3('2.5.1 三种驱动装置对比（★★）')
    d.table(['驱动', '优点', '缺点 / 适用'], [
        ['液压', '输出力矩大、结构紧凑、刚度好', '响应较慢、精度不高；需液压源；易泄漏；适用于重载（如 Atlas）'],
        ['气动', '结构简单、价格低', '工作压强低，难以精确位置控制；仅用于中小负荷末端执行器（气爪、旋转气缸、气动吸盘）'],
        ['电动', '能源简单、速度变化范围大、效率高、速度与位置精度高', '多与减速装置相连；分直流（DC）/交流（AC）伺服电机和步进电机驱动'],
    ], [2.0, 6.2, 8.6], size=9)
    d.h3('2.5.2 传动机构')
    d.table(['类型', '要点'], [
        ['直线：齿轮齿条', '旋转运动→直线运动；回差大、齿轮反转有间隙'],
        ['直线：滚珠丝杠', '摩擦力很小、响应快、传动效率高，消除低速爬行现象'],
        ['直线：液压 / 同步带滑台', '液压传动、同步带滑台'],
        ['旋转：轮系', '两个及以上齿轮组成；传递角位移、角速度、力和力矩'],
        ['旋转：同步齿形带', '柔性好、价格便宜、摩擦力好不打滑'],
        ['旋转：谐波减速器', '刚性齿轮 + 谐波发生器 + 柔性齿轮；如刚轮100齿、柔轮少2齿，谐波发生器转50圈柔轮转1圈，实现 1:50 减速；精密、结构紧凑；关节市占率约80%；置于小臂、腕部、手部'],
        ['旋转：RV 减速器', '行星齿轮减速 + 摆线针轮减速；结构紧凑、传动比大、振动小、噪音低、刚度与回转精度高；置于机座、大臂、肩部等重负载位置'],
    ], [3.4, 13.4], size=9)
    d.note([('易混　', 'red'), ('机座/大臂/肩部等重负载 → RV 减速器；小臂/腕部/手部 → 谐波减速器。（判断题常考）', '')])


def build_ch3(d):
    d.h1('第3章 机器人的运动学和动力学')
    d.h2('3.1 空间描述与变换')

    d.h3('3.1.1 位姿描述')
    d.bullets([
        [('运动学正问题', 'b'), ('：已知杆件几何参数 + 关节角矢量，求末端执行器相对参考坐标系的位姿。', ''), ('运动学逆问题', 'b'), ('：已知末端期望位姿，求关节矢量（判断能否到达预期位姿）。', '')],
        [('点的位置描述：在选定直角坐标系{A}中，空间任一点 P 用 (3×1) 列阵表示，左上标表示参考坐标系。', '')],
        [('齐次坐标：用 (4×1) 列阵表示三维空间点', ''), ('P = [x y z 1]ᵀ', 'kp'), ('；方向（单位矢量）用', ''), ('[a b c 0]ᵀ', 'kp'), ('表示，且 a²+b²+c²=1。', '')],
        [('普通坐标 ↔ 齐次坐标：点 (x,y,z) → (x,y,z,1)；矢量 (x,y,z) → (x,y,z,0)。反向同理。', '')],
        [('坐标轴单位矢量的方向列阵：', ''), ('i=[1 0 0 0]ᵀ，j=[0 1 0 0]ᵀ，k=[0 0 1 0]ᵀ', 'b'), ('；一般矢量单位矢量 h 的方向列阵 h = [cosα cosβ cosγ 0]ᵀ。', '')],
        [('原点的齐次列阵定义为 o = [0 0 0 α]ᵀ（α≠0）。', '')],
    ])
    d.bullets([
        [('动坐标系（动系）的位姿描述 = 原点位置 P + 坐标轴方向 n、o、a：', 'b')],
        [('n、o、a 分别为动系 X′、Y′、Z′ 轴单位矢量在静系上的分量（方向余弦），以齐次列阵表示。', '')],
        [('动系相对固定系的位姿矩阵（', ''), ('4×4 齐次变换矩阵', 'kp'), ('）：', '')],
    ])
    d.mono_block([
        '        ┌                    ┐',
        '        │  nx  ox  ax   Px   │',
        '        │  ny  oy  ay   Py   │   ← 前三列表示手部姿态',
        '        │  nz  oz  az   Pz   │   第四列表示手部中心位置',
        '        │   0   0   0    1   │',
        '        └                    ┘',
    ])
    d.note([('书写规范　', 'b'), ('表示连杆/手部姿态：用 n、o、a（X′、Y′、Z′ 轴的方向余弦）+ 位置列阵 P；目标物位姿可用其顶点齐次坐标（如楔块 6 个顶点）描述。', '')])

    d.h3('3.1.2 齐次变换和运算')
    d.bullets([
        [('多次简单运动 = 多个变换矩阵的积（齐次变换矩阵）；连杆初始位姿矩阵 × 齐次变换矩阵 = 最终位姿矩阵。', '')],
        [('平移算子 Trans(a,b,c)：', 'b'), ('对点、矢量、坐标系、物体均适用。', '')],
    ])
    d.mono_block([
        '        ┌                 ┐',
        '        │  1   0   0   a  │',
        'Trans = │  0   1   0   b  │',
        '        │  0   0   1   c  │',
        '        │  0   0   0   1  │',
        '        └                 ┘',
    ])
    d.bullets([
        [('旋转算子 Rot(轴,θ)（右手定则，逆时针为正）：', 'b')],
    ])
    d.mono_block([
        '        ┌                ┐        ┌                 ┐        ┌                ┐',
        '        │  1   0     0  0 │        │  cθ   0   sθ  0 │        │  cθ  −sθ  0  0 │',
        'Rot(x,θ)│  0  cθ  −sθ  0 │ Rot(y,θ)│  0    1    0  0 │ Rot(z,θ)│  sθ   cθ  0  0 │',
        '        │  0  sθ   cθ  0 │        │ −sθ   0   cθ  0 │        │   0    0  1  0 │',
        '        │  0   0     0  1 │        │   0   0    0  1 │        │   0    0  0  1 │',
        '        └                ┘        └                 ┘        └                ┘',
        '   cθ = cosθ，sθ = sinθ',
    ])
    d.bullets([
        [('算子左乘', 'kp'), ('：相对于', ''), ('固定（参考）坐标系', 'b'), ('进行的变换（如 T = T₃·T₂·T₁，T₃ 为最后一次相对固定系的变换写在最左）。', '')],
        [('算子右乘', 'kp'), ('：相对于', ''), ('动（自身）坐标系', 'b'), ('进行的变换（如 T = T₁·T₂·T₃，每次变换相对当前动系）。', '')],
        [('该规则同样适用于坐标系与物体的平移/旋转变换；复合变换中不同顺序结果不同。', '')],
    ])

    d.h2('3.2 机器人正运动学')
    d.bullets([
        [('正运动学：已知关节变量 [θ₁ … θₙ]ᵀ，求末端位姿（4×4）。', ''), ('正运动学模型 = 关节变量 → 末端位姿的映射。', '')],
        [('意义：运动学数学描述体系，是逆运动学、速度分析、动力学的理论基础；设计阶段评估工作空间、末端速度/加速度。', '')],
        [('串联机器人构型：基座到末端开链；相邻两连杆之间的相对位姿仅取决于连接关节 → 逐次累乘。', '')],
    ])
    d.h3('3.2.1 D-H 参数法（★★ 核心考点）')
    d.bullets([
        [('1955年 Denavit 和 Hartenberg 提出 D-H 参数法（后置坐标系法）；建立在前向关节上称前置坐标系法（Modified DH）。', '')],
        [('连杆几何参数：', 'b'), ('连杆长度 aₙ（两端关节轴线公垂线段的长）、连杆扭角 αₙ（两关节轴线夹角）。', '')],
        [('连杆关系参数：', 'b'), ('连杆距离 dₙ（沿关节 n 轴线两个公垂线间的距离）、连杆转角 θₙ（垂直于关节 n 轴线平面内两个公垂线的夹角）。', '')],
        [('旋转关节：θₙ 为关节变量，其余三个固定；移动关节：dₙ 为关节变量，其余三个固定。', '')],
        [('末端连杆 aₙ = a₀ = 0，αₙ = α₀ = 0；关节1若为转动关节取 d₁ = 0（θ₁ 零位任意），若为移动关节取 θ₁ = 0（d₁ 零位任意）。', '')],
        [('连杆 i 坐标系建立：原点位于 i+1 关节轴线上（公垂线交点）；Zᵢ 轴与 i+1 关节轴线重合；Xᵢ 轴与公垂线重合（从 i 指向 i+1）；Yᵢ 轴按右手法则。', '')],
    ])
    d.bullets([
        [('从 i−1 系到 i 系的四次变换（均相对动坐标系，故连乘顺序为右乘）：', 'b')],
    ])
    d.math_par([('A', 'i'), ('(i−1, i)', 'subn'), (' = Rot(z, θᵢ) · Trans(0,0,dᵢ) · Trans(aᵢ,0,0) · Rot(x, αᵢ)', '')])
    d.mono_block([
        '        ┌                                  ┐',
        '        │  cosθᵢ   −sinθᵢ·cosαᵢ   sinθᵢ·sinαᵢ   aᵢ·cosθᵢ │',
        'Aᵢ =    │  sinθᵢ    cosθᵢ·cosαᵢ  −cosθᵢ·sinαᵢ   aᵢ·sinθᵢ │',
        '        │    0         sinαᵢ         cosαᵢ        dᵢ    │',
        '        │    0           0             0          1    │',
        '        └                                  ┘',
    ])
    d.bullets([
        [('机器人运动学方程：⁰Tₙ = A₁A₂…Aₙ', 'kp'), ('（ⁱ⁻¹Tₙ = AᵢAᵢ₊₁…Aₙ）；前三列 = 手部姿态，第四列 = 手部位置。', '')],
        [('SCARA 实例：A₁ = Rot(z,θ₁)·Trans(l₁,0,0)，A₂ = Rot(z,θ₂)·Trans(l₂,0,0)，A₃ = Rot(z,θ₃)·Trans(l₃,0,0)，T₃ = A₁A₂A₃。', '')],
    ])

    d.h2('3.3 机器人逆运动学')
    d.bullets([
        [('逆运动学：已知末端期望位姿（笛卡尔空间，又称操作空间），求关节矢量（→ 关节空间变换）。', '')],
        [('特点：含三角函数，是超越方程；存在解的不存在性、多解性；求解方法多样性（封闭解法 / 数值解法；代数法=分离变量法、几何法）。', '')],
        [('代数解法（分离变量法）：用变换矩阵的逆左乘，找出右端为常数的元素，令其与左端对应元素相等，得到可求解的三角函数方程，逐次解出全部关节角（PUMA560 示例）。', '')],
    ])
    d.h3('3.3.1 逆解的多解问题与取舍')
    d.bullets([
        [('逆解多重性来自反三角函数方程；一个真实机器人只有一组解对应实际情况，必须取舍。', '')],
        [('剔除多余解的方法：', 'b')],
        [('① 根据关节运动空间选择合适的解（如关节限位 −150°～150°，两解 ±30° 只能取其一）；', '')],
        [('② 选择与前一采样时间最接近的解（连续性约束）；', '')],
        [('③ 根据避障要求选择合适的解；', '')],
        [('④ 逐级剔除多余解（n 个关节的全部解构成树形结构，逐级剪枝）。', '')],
    ])
    d.h3('3.3.2 二自由度平面关节型机器人（2R）示例')
    d.bullets([
        [('末端位置：x = l₁cosθ₁ + l₂cos(θ₁+θ₂)，y = l₁sinθ₁ + l₂sin(θ₁+θ₂)', ''), ('（作业题，见习题详解）', 'i')],
        [('逆解（几何法）：cosθ₂ = (x²+y²−l₁²−l₂²) / (2l₁l₂)，θ₂ 取正负两组解（肘上/肘下），θ₁ = atan2(y,x) − atan2(k₂,k₁)。', '')],
    ])
    d.note([('考点　', 'red'), ('①齐次变换矩阵 [n o a P] 的结构；②左乘（固定系）/ 右乘（动系）；③Rot/Trans 算子矩阵；④DH 四个参数与 Aᵢ 公式；⑤逆运动学多解及剔除方法——期末计算/填空必考。', '')])


def build_ch4(d):
    d.h1('第4章 机器人雅可比矩阵')
    d.h2('4.1 雅可比矩阵的意义')
    d.bullets([
        [('应用背景：焊缝跟踪中要求手部以恒定速度 v 沿特定方向运动 → 需计算任意时刻每个关节的速度，使总运动等于手部期望速度。', '')],
        [('定义：', 'b'), ('机器人手部在笛卡尔空间的', ''), ('操作速度与其关节空间角速度的线性变换特性', 'kp'), ('称为雅可比矩阵。', '')],
        [('机器人操作臂在某一姿态下，关节微小变化与末端操作器姿态微小变化之间的关系可用雅可比矩阵线性化表示；关节角速度与末端操作器速度及姿态变化之间的关系也可用雅可比矩阵表示。', '')],
        [('意义：雅可比矩阵可将单个关节的微分运动或速度转换为末端执行器的微分运动或速度。', '')],
        [('雅可比矩阵并非机器人运动学独有：只要输出可表示为输入的函数，都存在雅可比矩阵。', '')],
    ])
    d.h3('4.1.1 雅可比矩阵的定义式')
    d.bullets([
        [('设 y = [y₁ … yₙ]ᵀ，x = [x₁ … xₘ]ᵀ，则雅可比矩阵 J 为 m×n 矩阵，元素 Jᵢⱼ = ∂yᵢ/∂xⱼ。', '')],
        [('6 自由度操作臂：关节角矢量 q = [q₁ … q₆]ᵀ，末端位置矢量 p = [pₓ p_y p_z]ᵀ，正运动学 p = f(q)；', '')],
    ])
    d.math_par([('J(q) = ∂p / ∂q = [∂p/∂q₁  ∂p/∂q₂  …  ∂p/∂q₆]', '')])
    d.bullets([
        [('得到 3×6 的雅可比矩阵（各元素为操作空间参数对各关节变量的偏导数）。关节角矢量写为 q：既区别于只有回转关节的机器人，也因为 q 中可能出现移动副关节。', '')],
        [('位置矢量相对于关节矢量的雅可比矩阵称为', ''), ('位置雅可比矩阵（Jacobian Matrix）', 'kp'), ('。', '')],
    ])
    d.h3('4.1.2 微分运动关系（★★ 核心公式）')
    d.math_par([('dp = J(q) · dq', ''), ('　　（末端微分位移 = 雅可比 × 关节微分位移）', 'i')])
    d.math_par([('v = J(q) · q̇', ''), ('　　（末端速度 = 雅可比 × 关节速度）', 'i')])

    d.h2('4.2 通用雅可比矩阵')
    d.bullets([
        [('对一般 6 自由度操作臂，将线速度与角速度合并为 6 维空间速度：', '')],
    ])
    d.math_par([('[v; ω] = J(q) · q̇　，　J 为 6×6 通用雅可比矩阵', '')])
    d.bullets([
        [('列的物理意义：J 的第 j 列 = 关节 j 单位角速度（或单位线速度，移动副）在末端产生的线速度与角速度分量。', '')],
        [('行的物理意义：J 的第 i 行 = 末端第 i 个空间速度分量对全部关节变量的敏感度（偏导数）。', '')],
        [('元素的物理意义：Jᵢⱼ = 关节 j 单位运动引起末端第 i 个速度分量的变化率。', '')],
    ])
    d.h3('4.2.1 力雅可比矩阵（★★ 简答高频）')
    d.math_par([('τ = Jᵀ(q) · F', ''), ('　　（关节力矩 = 速度雅可比的转置 × 末端力/力矩）', 'i')])
    d.bullets([
        [('由虚功原理：末端力做的虚功等于关节力矩做的虚功，Fᵀ·δp = τᵀ·δq，代入 δp = J·δq 即得 τ = JᵀF。', '')],
        [('即：', ''), ('力雅可比矩阵 = 速度雅可比矩阵的转置', 'kp'), ('；力控制以此式为基础。', '')],
        [('意义：已知末端受力/力矩 F，可求各关节电机需输出的驱动力矩 τ。', '')],
    ])

    d.h2('4.3 雅可比矩阵的计算方法')
    d.bullets([
        [('①微分法（解析法）：对正运动学方程 p = f(q) 直接求偏导。', '')],
        [('②几何法（矢量积法）：旋转关节列 Jᵢ = [zᵢ₋₁ × (p − pᵢ₋₁); zᵢ₋₁]；移动关节列 Jᵢ = [zᵢ₋₁; 0]（z 为该关节轴线方向单位矢量）。', '')],
    ])
    d.h3('4.3.1 例：2-DOF 操作臂的雅可比矩阵（PPT 例1）')
    d.bullets([
        [('设两连杆长度 l₁、l₂，关节角 θ₁、θ₂（均为转动关节），末端位置 p = [l₁c₁+l₂c₁₂, l₁s₁+l₂s₁₂]ᵀ（c₁=cosθ₁，s₁₂=sin(θ₁+θ₂)）。', '')],
    ])
    d.mono_block([
        '        ┌                              ┐',
        '  J =   │  −l₁s₁−l₂s₁₂   −l₂s₁₂  │',
        '        │   l₁c₁+l₂c₁₂    l₂c₁₂  │',
        '        └                              ┘',
    ])
    d.bullets([
        [('J₁、J₂ 为雅可比矩阵的第一、二列矢量：速度传递关系 dp = J₁dθ₁ + J₂dθ₂；J₁、J₂ 分别表示单位关节速度在末端位置产生的速度分量的传动比。', '')],
        [('令 det(J) = 0 可得奇异形位：l₁l₂s₂ = 0，即 θ₂ = 0° 或 180°（手臂伸直/折叠）时，末端失去沿某些方向运动的能力。', '')],
    ])

    d.h2('4.4 奇异形位与使用注意点')
    d.bullets([
        [('产生条件：det(J) = 0（J 的秩下降），此时手臂处于', ''), ('奇异形位', 'kp'), ('。', '')],
        [('影响：①某些方向的末端速度无法实现；②为获得期望速度，关节速度将趋于无穷大（电机无法承受）；③逆解不唯一/不存在。', '')],
        [('常见奇异：手臂完全伸直或完全折叠（θ₂ = 0 或 180°）；腕关节轴线共线等。', '')],
        [('避免/处理：规划路径时避开奇异形位；奇异附近限制末端速度、采用阻尼最小二乘等数值方法。', '')],
    ])
    d.note([('考点　', 'red'), ('①雅可比矩阵的定义（操作速度与关节角速度的线性变换）；②dp=J·dq、v=J·q̇；③力雅可比 τ=JᵀF（虚功原理）；④奇异形位 det(J)=0 及影响。', '')])


def build_ch5(d):
    d.h1('第5章 机器人的轨迹规划')
    d.h2('5.1 轨迹描述和生成')
    d.bullets([
        [('运动方式：', 'b'), ('点到点运动', 'kp'), ('（只关注起止点位姿，不关注路径，如搬运、装配）；', ''), ('路径跟踪运动', 'kp'), ('（关注起止点位姿及路径、路径点位姿，如弧焊、曲面加工）。', '')],
        [('路径', 'b'), ('：从起始点到终止点的空间曲线；', ''), ('轨迹', 'b'), ('：机器人每一个自由度的位置、速度、加速度与时间的关系（强调时间性）。', '')],
        [('轨迹规划：根据作业任务要求，对机器人末端执行器在工作中位姿变化的路径、取向及其变化速度和加速度的人为设定。', '')],
        [('理想轨迹 = 平滑路径（连续一阶、二阶导数的连续路径）；笨拙、急速的运动会加剧机构磨损、激起操作臂共振。', '')],
        [('规划函数必须', ''), ('连续且平滑', 'kp'), ('；与手臂种类无关，工具坐标系 {T} 相对工件坐标系 {S} 描述（{G} 也可随时间变动，如输送带）。', '')],
        [('轨迹规划可在', ''), ('关节空间', 'b'), ('或', ''), ('直角（笛卡尔）空间', 'b'), ('进行：前者以关节转角函数描述轨迹；后者以手部位姿关于时间的函数描述路径形状，相应关节位移/速度/加速度由手部信息导出。', '')],
    ])

    d.h2('5.2 关节空间的规划方法（★★ 计算大题高频）')
    d.bullets([
        [('将每个目标作业路径点经逆运动学拟合成光滑函数：起始点 i=1，终止点 i=N+1；得到多个关节的平滑运动轨迹。', '')],
    ])
    d.h3('5.2.1 三次多项式插值')
    d.math_par([('θ(t) = a₀ + a₁t + a₂t² + a₃t³', '')])
    d.bullets([
        [('4 个约束条件（起始/终止点的', ''), ('位置约束', 'b'), (' + ', ''), ('速度约束', 'b'), ('）：', '')],
    ])
    d.mono_block([
        '  θ(0) = θ₀，θ̇(0) = θ̇₀，θ(t_f) = θ_f，θ̇(t_f) = θ̇_f',
    ])
    d.bullets([
        [('静止起步、静止停止（θ̇₀=θ̇_f=0）时：', 'b')],
    ])
    d.math_par([('a₀ = θ₀，', ''), ('a₁ = 0，', ''), ('a₂ = 3(θ_f−θ₀)/t_f²，', ''), ('a₃ = −2(θ_f−θ₀)/t_f³', '')])
    d.note([('记忆法　', 'b'), ('位置差 Δθ = θ_f−θ₀：a₂ = 3Δθ/t_f²，a₃ = −2Δθ/t_f³。速度曲线为抛物线，加速度线性变化（端点加速度最大）。', '')])

    d.h3('5.2.2 过路径点的三次多项式插值')
    d.bullets([
        [('把每个关节上相邻的两个路径点分别看做起始点和终止点，确定相应三次多项式；路径点平滑连接；一般起止点速度不再为零（由过路径点条件连续求得）。', '')],
        [('约束：每段起止位置 + 速度连续；多段联立方程求解各段系数。', '')],
    ])

    d.h3('5.2.3 五次多项式插值（高次多项式）')
    d.math_par([('θ(t) = a₀ + a₁t + a₂t² + a₃t³ + a₄t⁴ + a₅t⁵', '')])
    d.bullets([
        [('在起止点同时指定', ''), ('位置、速度、加速度', 'b'), ('共 6 个约束条件 → 6 个系数。', '')],
        [('起止速度为零、起止加速度为零时：', 'b')],
    ])
    d.math_par([('a₀ = θ₀，a₁ = 0，a₂ = 0，', ''), ('a₃ = 10Δθ/t_f³，', ''), ('a₄ = −15Δθ/t_f⁴，', ''), ('a₅ = 6Δθ/t_f⁵', '')])

    d.h3('5.2.4 用抛物线过渡的线性插值（★★）')
    d.bullets([
        [('单纯线性插值：两端点速度不连续、加速度无穷大，造成刚性冲击。', '')],
        [('对策：线性段两端用抛物线段过渡（抛物线二阶导数=常数 → 过渡段加速度恒定），保证位置、速度连续平滑。', '')],
        [('构成：抛物线段 + 线性段 + 抛物线段，速度曲线呈梯形（T 形速度曲线）。', '')],
    ])
    d.bullets([
        [('设过渡时间为 t_a、匀加速段加速度为 a、巡航速度为 V：', 'b')],
    ])
    d.math_par([('V = a·t_a，', ''), ('Δθ = V·(T − t_a) = V·T − V²/a', '')])
    d.bullets([
        [('由 Δθ、T、a 解出 V（取小于 Δθ/T…的合理根），再得 t_a = V/a；速度曲线先线性增加（抛物线过渡）→ 匀速段 → 线性下降（抛物线过渡）。', '')],
        [('经验取值：PPT 在此给出 T 形速度曲线示意，过渡时间由期望最大速度与加速度确定（见例7-3解法）。', '')],
    ])

    d.h2('5.3 笛卡尔空间规划方法')
    d.bullets([
        [('关节空间法轨迹虽保证经过途径点和终止点，但手爪在空间的路径非常复杂（非直线），取决于机器人运动学方程；且关节变量呈线性变化时直角空间轨迹并不形成直线 → 只有无路径要求的作业才能在关节空间直接规划。', '')],
        [('笛卡尔空间规划：用笛卡尔位置和姿态关于时间的函数描述路径形状（直线、圆、正弦曲线等）。', '')],
        [('笛卡尔路径的三个几何问题：①无法到达中间点；②在奇异点附近关节速度增大；③起始点和终止点有不同的解。一般选关节空间法，必要时才用笛卡尔空间路径规划。', '')],
    ])
    d.note([('考点　', 'red'), ('①点到点 vs 路径跟踪；②轨迹三要素：位置、速度、加速度（时间函数）；③三次多项式系数公式（静止起停）；④抛物线过渡线性插值：V、t_a 的计算；⑤笛卡尔空间三个几何问题。', '')])


# =====================================================================
# 第二部分：习题详解
# =====================================================================

def ex_q(d, tag, question):
    p = d.doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.line_spacing = 1.25
    r = p.add_run('【' + tag + '】')
    set_run_font(r, cn=_CN_HEAD, size=10.5, bold=True, color=NAVY)
    if isinstance(question, str):
        question = [(question, '')]
    d._add_segments(p, question)
    return p


def ex_a(d, answer, label='答'):
    p = d.doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.space_after = Pt(2); p.paragraph_format.line_spacing = 1.25
    r = p.add_run(label + '：')
    set_run_font(r, cn=_CN_HEAD, size=10, bold=True, color='1F7A33')
    if isinstance(answer, str):
        answer = [(answer, '')]
    d._add_segments(p, answer)
    return p


def ex_steps(d, steps):
    for i, step in enumerate(steps, 1):
        p = d.doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.5)
        p.paragraph_format.first_line_indent = Cm(-0.45)
        p.paragraph_format.space_after = Pt(2); p.paragraph_format.line_spacing = 1.25
        r = p.add_run('(%d) ' % i)
        set_run_font(r, size=10, bold=True, color=RED)
        if isinstance(step, str):
            step = [(step, '')]
        d._add_segments(p, step)


def ex_ans(d, segments, label='答案'):
    p = d.doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9)
    set_para_shading(p, 'EFF6EF')
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.25
    r = p.add_run(label + '：')
    set_run_font(r, cn=_CN_HEAD, size=10, bold=True, color='1F7A33')
    if isinstance(segments, str):
        segments = [(segments, '')]
    d._add_segments(p, segments)


def build_part2(d):
    part_title(d, '第二部分', '习题详解', '各章 PPT 中的例题、练习与“思考与习题 / 思考与练习”均配解题步骤与答案')
    build_ex_ch1(d)
    build_ex_ch2(d)
    build_ex_ch3(d)
    build_ex_ch4(d)
    build_ex_ch5(d)


def build_ex_ch1(d):
    d.h1('第1章 思考与习题', pagebreak=True)

    d.h2('一、填空题')
    ex_q(d, '填空1', '一台真正意义上的机器人，应该具备三种基本功能：______、______和______。')
    ex_a(d, '感知（感觉）、决策（思考/规划）、执行（行动）。')
    ex_q(d, '填空2', '按照坐标形式分类，工业机器人可分为______、______、______和关节坐标型。')
    ex_a(d, '直角坐标型、圆柱坐标型、球坐标型。')
    ex_q(d, '填空3', '工业机器人四大家族企业分别为______、______、______和______。')
    ex_a(d, '瑞士 ABB、德国 KUKA、日本 FANUC、日本 YASKAWA（安川）。')

    d.h2('二、判断题（先判断对错，再看解析）')
    judge = [
        ('判断1', '数控机床可编程并且能加工多种机械零件，因此数控机床也是一种工业机器人。', '×',
         '数控机床虽可编程、能加工多种零件，但它不具备机器人的通用性与拟人化/智能化特征（不能更改程序便捷地改变作业内容、不具备手爪等执行机构），不属于工业机器人。'),
        ('判断2', '直角坐标机器人具有结构紧凑、灵活、占地空间小等优点，是目前工业机器人大多采用的结构形式。', '×',
         '恰恰相反：直角坐标机器人位置精度高、轨迹易求，但作用范围小、结构庞大、无法调节工具姿态、灵活性差；关节坐标型机器人结构最紧凑、工作空间最大、应用最多。'),
        ('判断3', '工业机器人的自由度越高其运动灵活性越好，承载能力越强。', '×',
         '自由度越高灵活性越好，但承载能力与自由度并无正比关系；自由度越高，结构越复杂、刚性越差，承载能力往往越低。'),
        ('判断4', '为了提高机器人的工作效率，要求机器人的最大工作速度越大越好。', '×',
         '速度越高效率越高，但升降速耗时更长、对最大加/减速度及加速度变化率要求更高，需综合设计，并非越大越好。'),
    ]
    for tag, q, ans, why in judge:
        ex_q(d, tag, q)
        ex_a(d, why)
        ex_ans(d, ans, '答案')

    d.h2('三、简答题')
    ex_q(d, '简答1', '简述工业机器人的定义，说明机器人的主要特点。')
    ex_a(d, '定义（ISO）：自动的、位置可控的、具有编程能力的多功能机械手，具有几个轴，能借助可编程序操作来处理各种材料、零件、工具和专用装置，执行种种任务。')
    ex_steps(d, [
        '可编程：可随工作环境变化再编程，是柔性制造系统（FMS）的重要组成部分。',
        '拟人化：机械结构类似人的行走、腰转、大臂、小臂、手腕、手爪；控制上有“电脑”；智能化机器人还有类似人类的“生物传感器”。',
        '通用性：①机构可变（改变末端操作器、自由度数量）；②程序可变。',
        '适应性：①自动化（不依赖人的干预）；②智能化（记忆、语言理解、图像识别、判断推理能力）。',
    ])
    ex_q(d, '简答2', '简述下面几个术语的含义：自由度、重复定位精度、工作空间和承载能力。')
    ex_steps(d, [
        [('自由度：', 'b'), ('机器人所具有的独立坐标轴运动的数目，不包括手爪（末端执行器）的自由度，用以衡量运动灵活性，越高灵活性越好。', '')],
        [('重复定位精度：', 'b'), ('相同运动位置指令下，机器人连续若干次相同运动轨迹间的度量（统计量）；工业机器人的重复定位精度要求高于绝对定位精度。', '')],
        [('工作空间：', 'b'), ('机器人手臂末端（或手腕中心）所能到达的所有点的集合，也称工作区域；含灵巧点（任意姿态可达）与非灵巧点（接近手臂运动边界的极限点）。', '')],
        [('承载能力：', 'b'), ('机器人在工作范围内任何位姿上所能承受的最大质量；既含负载又含末端执行器质量；与运行速度和加速度的大小、方向有关，为安全起见指高速运行时的承载能力。', '')],
    ])
    ex_q(d, '简答3', '简述机器人的组成。')
    ex_a(d, '工业机器人系统一般由三部分、六个子系统组成：')
    ex_steps(d, [
        [('驱动系统：', 'b'), ('驱动机械系统的装置，按驱动源分为电动、液压、气动；可直接相连，也可经同步带、链条、齿轮、谐波减速器等间接相连。', '')],
        [('机械系统：', 'b'), ('腰部（机身）、臂部、腕部、手部（末端执行器）。', '')],
        [('感知系统：', 'b'), ('内部传感器（关节位置/速度/力/扭矩）+ 外部传感器（听觉、视觉、触觉）。', '')],
        [('机器人-环境交互系统：', 'b'), ('实现机器人与外部设备相互联系和协调的系统。', '')],
        [('控制系统：', 'b'), ('控制计算机 + 示教编程器 + 操作面板；按运动形式分点位控制和轨迹控制；程序控制（开环）/ 自适应控制（闭环）/ 智能控制。', '')],
        [('人机交互系统：', 'b'), ('使操作人员参与控制并与之联系的装置（标准终端、指令控制台、显示板、报警器等）。', '')],
    ])
    ex_q(d, '简答4', '简述机器人的发展历史。')
    ex_steps(d, [
        '1920年，捷克作家卡雷尔·卡佩克在《洛桑万能机器人》中首次使用“机器人”一词（Robota，奴隶）。',
        '1950年，阿西莫夫《我是机器人》提出机器人三原则。',
        '1954年，戴沃尔提出工业机器人概念并申请专利；1956年与恩格尔伯格相识；1959年研制全球首台工业机器人（Unimation，液压驱动）。',
        '20世纪60年代：液压驱动；70年代：小型计算机控制 + 电机驱动；80年代：趋于成熟，外形功能与现代机器人基本相同；90年代：Delta 机器人，在线示教为主。',
        '三个发展阶段：示教再现型（第1代）→ 有感觉的机器人（第2代）→ 智能机器人（第3代）。',
        '我国古代：鲁班“飞鸟”、偃师“伶人”、张衡“记里鼓车”、诸葛亮“木牛流马”。',
    ])
    ex_q(d, '简答5', '查阅资料，谈谈国外和国内工业机器人应用技术现状。（开放题，参考要点）')
    ex_steps(d, [
        [('国外：', 'b'), ('美国提出机器人发展线路图（2013，重点关注创新应用）；欧盟第7框架 / Horizon 研究计划 / 火花计划（布局智能制造应用）；日本《机器人新战略》（2015，世界机器人创新基地、第一应用国家）；工业机器人四大家族：ABB（瑞士）、KUKA（德国）、FANUC（日本）、YASKAWA（日本）。', '')],
        [('国内：', 'b'), ('2014年两院院士大会提出“不仅要把我国机器人水平提高上去，而且要尽可能多占领市场”；中国制造2025 将机器人列为重点发展领域；应用集中于汽车/3C/新能源的搬运、焊接、喷涂、装配等工序，服务机器人与特种机器人加速落地。', '')],
    ])


def build_ex_ch2(d):
    d.h1('第2章 思考与习题')

    d.h2('一、填空题')
    ex_q(d, '填空1', '机器人机身和臂部的配置形式有______、______、______和______。')
    ex_a(d, '横梁式、立柱式、机座（基座）式、屈伸式。')
    ex_q(d, '填空2', '腕部是联接______和______的结构部件，它的主要作用是确定手部的______。')
    ex_a(d, '手臂、手部；姿态（作业方向）。')
    ex_q(d, '填空3', '若手腕的关节轴线与手臂的纵轴线共线，该手腕关节称为______，若手腕的关节轴线与手臂的纵轴线垂直，该手腕关节称为______。')
    ex_a(d, '共线 → 滚转关节（R，臂转）；垂直 → 俯仰关节（B，腕摆）。（同时具有俯仰与偏转者称双腕摆。）')
    ex_q(d, '填空4', '手部按其用途可以分为______和______。')
    ex_a(d, '手爪（具一定通用性）和专用操作器（工具）。')
    ex_q(d, '填空5', '工业机器人主要的三种驱动方式：______、______和______。')
    ex_a(d, '液压、气动、电动。')

    d.h2('二、判断题')
    judge = [
        ('判断1', '臂部是机器人的主要执行部件，它的作用是支撑腕部和手部，确定手部的位姿。', '√',
         '臂部的作用正是支撑腕部和手部并带动它们在空间运动，臂部的运动和参数决定腕部的空间位置及工作空间。'),
        ('判断2', '工业机器人的手部为专用部件，通用性比较差。', '√',
         '一种手爪往往只能抓握一种或几种形状/尺寸/重量相近似的工件；一种工具只能执行一种作业，故通用性差。'),
        ('判断3', '一般将 RV 减速器放置在机座、大臂、肩部等重负载的位置，而将谐波减速器放置在小臂、腕部或手部。', '√',
         'RV 减速器刚度与回转精度更高，适用于重负载位置；谐波减速器精密紧凑，适用于小臂、腕部或手部。'),
    ]
    for tag, q, ans, why in judge:
        ex_q(d, tag, q)
        ex_a(d, why)
        ex_ans(d, ans, '答案')

    d.h2('三、简答题')
    ex_q(d, '简答1', [('如图所示的机器人为何种坐标类型的机器人？（PPT“图2-42 习题1图”，垂直多关节机械臂）', '')])
    d.doc.add_picture('render/ch2_image92.jpeg', width=Cm(5.2))
    d.doc.paragraphs[-1].alignment = 1
    ex_a(d, [('该机器人为', ''), ('关节坐标型（垂直多关节型）机器人', 'kp'), ('：由立柱回转、大臂俯仰、小臂俯仰（屈伸）等转动关节串联构成，结构紧凑、灵活性大、占地面积小、工作空间大，是目前应用最多的一类。', '')])
    ex_q(d, '简答2', '机器人机械结构由哪几部分组成，每一部分的作用是什么？')
    ex_steps(d, [
        [('机身（腰部）：', 'b'), ('机器人的基础部分，直接连接、支承和传动手臂及行走机构；带有臂部升降、平移、回转、俯仰机构及导向装置和支撑件。', '')],
        [('手臂（臂部）：', 'b'), ('主要执行部件，支撑腕部和手部并带动它们在空间运动；其运动和参数决定腕部的空间位置和工作空间；并将载荷传递到基座。', '')],
        [('腕部：', 'b'), ('联接手臂和手部，具有独立自由度，调节/改变手部的作业方向（姿态），满足复杂姿态调整。', '')],
        [('手部（末端执行器）：', 'b'), ('装在手腕末端法兰上，直接抓握工件或执行作业（手爪或专用工具）。', '')],
    ])
    ex_q(d, '简答3', '机器人的俯仰型机身机构是如何工作的？')
    ex_a(d, '俯仰型机身（球坐标型）主要由实现手臂左右回转和上下俯仰的部件组成：它用手臂的俯仰运动部件代替手臂的升降运动部件；手臂俯仰一般采用活塞（气）缸与连杆机构联用实现（俯仰驱动缸安置在臂部/立柱上，铰接活塞缸通过连杆带动手臂上下俯仰），回转由回转缸或摆动缸驱动，两者配合使手臂在球面范围内运动。')
    ex_q(d, '简答4', '机器人腕部机构的转动方式有哪几种？')
    ex_steps(d, [
        [('臂转：', 'b'), ('腕部绕小臂轴线方向旋转，也称腕部旋转（对应翻转关节 R）。', '')],
        [('腕摆：', 'b'), ('手部绕垂直小臂轴线方向旋转；分俯仰（pitch）和偏转（yaw）；同时具有俯仰和偏转的称双腕摆（对应俯仰/偏转关节 B）。', '')],
        [('手转：', 'b'), ('手部绕自身轴线方向旋转。', '')],
        '常用组合：臂转—腕摆—手转；臂转—双腕摆—手转。',
    ])
    ex_q(d, '简答5', '试述机器人的手部机构的特点，机器人手部有哪些种类？')
    ex_steps(d, [
        [('特点（4点）：', 'b'), ('①与手腕相连处可拆卸（手部随夹持对象而异，通常一机多手，要求接头通用性、互换性）；②是机器人末端操作器（类人手爪或喷枪、焊枪等专用工具）；③通用性比较差（一种手爪只能抓握一种或几种形状/尺寸/重量相近的工件，一种工具只能执行一种作业）；④是一个独立的部件（机身、手臂、手部——机械系统三大件）。', '')],
        [('种类（5种分类）：', 'b'), ('按用途：手爪 / 专用操作器（工具）；按夹持方式：外夹式、内撑式、内外夹持式；按工作原理：夹持类（摩擦力夹持、吊钩承重）、吸附类（磁力/真空/静电/范德华力吸盘）；按手指数目：二指、多指；按智能化：普通式手爪、智能化手爪（集成力/触觉/滑觉传感器）。', '')],
    ])
    ex_q(d, '简答6', '夹持类手部、吸附类手部和仿人类手部分别适用于哪些作业场合？')
    ex_steps(d, [
        [('夹持类（机械手爪）：', 'b'), ('抓住—握持—释放形状、尺寸、重量相近似的工件；箱体类、轴类、板材类工件的上下料、抓放（如外夹式夹箱体、内撑式撑内孔）。', '')],
        [('吸附类：', 'b'), ('磁力吸盘 → 铁磁材料工件；真空吸盘 → 表面较平整、形状规则的板材（铁板、铝板、石板、玻璃等）；静电吸盘 → 半导体晶圆、玻璃、金属箔、薄膜等薄型纤细工件；范德华力吸盘 → 3C 电子行业平板类微器件。', '')],
        [('仿人类（多指灵巧手 / 柔性手）：', 'b'), ('类人复杂动作——拧螺钉、弹钢琴、礼仪手势等；人形/拟人机器人及需要高灵巧操作的精密作业。', '')],
    ])
    ex_q(d, '简答7', '夹持类手部有哪几类？夹钳式手部由哪几个部分组成？')
    ex_a(d, '夹持类手部有两类：摩擦力夹持和吊钩承重。夹钳式手部由手指、传动机构（回转型 / 平移型）、驱动机构、连接与支承元件（壳体、支架）四部分组成；手指直接与工件接触，传动机构向手指传递运动和动力实现夹紧与松开。')


def build_ex_ch3(d):
    d.h1('第3章 思考与练习')

    d.h2('一、填空题')
    ex_q(d, '填空1', '机器人运动学问题可分为______和______。')
    ex_a(d, '运动学正问题、运动学逆问题。')
    ex_q(d, '填空2', '在机器人坐标系中，运动时相对于连杆不动的坐标系称为______，跟随连杆运动的坐标系称______。')
    ex_a(d, '静坐标系（静系）、动坐标系（动系）。')
    ex_q(d, '填空3', '目标物位姿的描述，可以用______来描述。')
    ex_a(d, '齐次坐标（齐次变换矩阵），如楔块用 6 个顶点的齐次坐标描述。')

    d.h2('二、判断题')
    judge = [
        ('判断1', '给定机器人末端执行器相对于参考坐标系的期望位置和姿态，求解关节矢量的问题称为运动学正问题。', '×',
         '给定末端期望位姿求关节矢量是逆问题（逆运动学）；正问题是已知关节矢量求末端执行器相对参考坐标系的位姿。'),
        ('判断2', '关节驱动力和力矩与末端执行器施加的力和力矩之间的关系是机器人操作臂力控制的基础。', '√',
         '该关系由力雅可比矩阵（速度雅可比的转置）描述，τ = JᵀF，是力控制的理论基础。'),
        ('判断3', '工业机器人的运动学影响其定位精度，动力学影响其运动稳定性。', '√',
         '运动学决定位置/姿态的几何映射（定位精度）；动力学决定力/力矩与运动的关系（运动稳定性、动态性能）。'),
    ]
    for tag, q, ans, why in judge:
        ex_q(d, tag, q)
        ex_a(d, why)
        ex_ans(d, ans, '答案')

    d.h2('三、计算题（旋转变换）')
    ex_q(d, '计算', [('已知坐标系中点 U 的位置矢量 u = [8 1 4 1]ᵀ，将此点绕 X 轴旋转 90°，再绕 Z 轴旋转 90°，求旋转变换后所得的点 W。', '')])
    ex_steps(d, [
        [('第一步（绕 X 轴 90°）：Rot(x,90°) = [[1,0,0,0],[0,0,−1,0],[0,1,0,0],[0,0,0,1]]，', ''), ('u′ = Rot(x,90°)·u = [8, 0×1+0×1−1×4, 0×1+1×1+0×4, 1]ᵀ = [8, −4, 1, 1]ᵀ。', 'b')],
        [('第二步（绕 Z 轴 90°）：Rot(z,90°) = [[0,−1,0,0],[1,0,0,0],[0,0,1,0],[0,0,0,1]]，', ''), ('W = Rot(z,90°)·u′ = [0×8−1×(−4), 1×8+0×(−4), 1×1, 1]ᵀ = [4, 8, 1, 1]ᵀ。', 'b')],
    ])
    ex_ans(d, [('W = [4  8  1  1]ᵀ。', 'b'), ('提示：复合旋转按“先×后×”顺序左乘，即 W = Rot(z,90°)·Rot(x,90°)·u；若两次都绕固定系轴旋转，左乘固定系算子；若绕动系轴旋转，则右乘动系算子。', 'i')])

    d.h2('四、简答题')
    ex_q(d, '简答1', '工业机器人力雅可比矩阵和速度雅可比矩阵有何关系？')
    ex_a(d, '力雅可比矩阵是速度雅可比矩阵的转置：π = Jᵀ(q)·F。由虚功原理——末端力做的虚功等于关节力矩做的虚功 Fᵀ·δp = τᵀ·δq，代入 δp = J·δq 整理即得 τ = JᵀF。它是机器人操作臂力控制的基础。')
    ex_q(d, '简答2', '简述二自由度平面关节型机械手动力学方程主要包含哪些项？有何物理意义？')
    ex_steps(d, [
        [('二自由度平面关节型机械手的动力学方程：τ = M(q)·q̈ + C(q, q̇)·q̇ + G(q)。', 'b')],
        [('惯性力项 M(q)·q̈：', 'b'), ('与质量矩阵（惯量矩阵）有关的项，反映连杆质量/惯量对关节驱动力矩的要求。', '')],
        [('科氏力和离心力项 C(q, q̇)·q̇：', 'b'), ('由连杆速度产生的耦合项（q̇ 的二次型），反映各连杆运动相互激发的惯性效应。', '')],
        [('重力项 G(q)：', 'b'), ('由连杆重力产生的项，仅与位姿 q 有关，与运动无关，需关节电机持续输出力矩平衡。', '')],
        '（若有摩擦还需加摩擦力项；有外部负载时加上负载力/力矩项。）',
    ])
    ex_q(d, '简答3', '什么是机械臂连杆之间的耦合作用？')
    ex_a(d, '机械臂的动力学方程中惯量矩阵 M(q) 一般为非对角矩阵，且含科氏/离心项：一个关节的加速度/速度会在其他关节上产生惯性力与力矩分量——即各连杆之间惯性力的相互影响，称为连杆之间的耦合作用（动力学耦合）。因此关节控制器若按 SISO 独立设计，各关节运动会互相干扰。')

    d.h2('五、PPT 例题精讲（含解题步骤）')

    ex_q(d, '例3-1', '用齐次坐标分别写出图中矢量 u 的方向列阵（u 与 Y、Z 轴夹 45°）。')
    ex_steps(d, [
        '矢量单位矢量 h 的方向列阵为 h = [cosα cosβ cosγ 0]ᵀ。',
        '由图读出：u 与 X 轴夹 90°，与 Y 轴夹 45°，与 Z 轴夹 45°。',
        'u = [cos90° cos45° cos45° 0]ᵀ = [0  0.7071  0.7071  0]ᵀ。',
    ])
    ex_ans(d, 'u = [0  0.7071  0.7071  0]ᵀ')

    ex_q(d, '练习3-1', '用齐次坐标分别写出图中矢量 v、w 的方向列阵（v：与 X、Z 轴夹 45°；w：与 X、Y 轴夹 60°、Z 轴夹 45°）。')
    ex_steps(d, [
        'v = [cos45° cos90° cos45° 0]ᵀ = [0.7071  0  0.7071  0]ᵀ。',
        'w = [cos60° cos60° cos45° 0]ᵀ = [0.5  0.5  0.7071  0]ᵀ。',
    ])
    ex_ans(d, 'v = [0.7071  0  0.7071  0]ᵀ，　w = [0.5  0.5  0.7071  0]ᵀ')

    ex_q(d, '例3-2', [('如图所示为固连于连杆的坐标系{ B }位于 Oᴮ 点，Xᴮ=2，Yᴮ=1，Zᴮ=0。在 XOY 平面内，坐标系{ B }相对固定坐标系{ A }有一个 30° 的偏转，试写出表示连杆位姿的坐标系{ B }的 4×4 矩阵表达式。', '')])
    ex_steps(d, [
        [('位置列阵：P = [2  1  0  1]ᵀ。', '')],
        [('XB 方向列阵：n = [cos30° cos60° cos90° 0]ᵀ = [0.866  0.500  0  0]ᵀ。', '')],
        [('YB 方向列阵：o = [cos120° cos30° cos90° 0]ᵀ = [−0.500  0.866  0  0]ᵀ。', '')],
        [('ZB 方向列阵：a = [0  0  1  0]ᵀ（XOY 平面内转 30°，Z 轴不变）。', '')],
        [('按 [n o a P] 列排布得位姿矩阵：', 'b')],
    ])
    d.mono_block([
        '        ┌                          ┐',
        '        │  0.866  −0.500   0   2  │',
        '  ᴬTʙ = │  0.500   0.866   0   1  │',
        '        │  0.000   0.000   1   0  │',
        '        │  0       0       0   1  │',
        '        └                          ┘',
    ])

    ex_q(d, '练习3-2', '动坐标系在固定参考坐标系中的表示：P 动坐标系位于参考坐标系中 (3,5,7) 的位置，它的 n 轴与 x 轴平行，o 轴与 y 轴夹 45°，a 轴与 z 轴夹 45°，写出该动坐标系的矩阵表达式。')
    ex_steps(d, [
        [('位置列阵：P = [3  5  7  1]ᵀ。', '')],
        [('n 轴与 x 平行：n = [cos0° cos90° cos90° 0]ᵀ = [1  0  0  0]ᵀ。', '')],
        [('o 轴与 y 轴夹 45° 且 ⊥ n：o = [0  cos45°  sin45°  0]ᵀ = [0  0.7071  0.7071  0]ᵀ。', '')],
        [('a 轴与 z 轴夹 45° 且 a = n × o：a = [0  −sin45°  cos45°  0]ᵀ = [0  −0.7071  0.7071  0]ᵀ。', '')],
    ])
    d.mono_block([
        '        ┌                          ┐',
        '        │  1   0        0       3  │',
        '  ᴬTᴮ = │  0   0.7071  −0.7071  5  │',
        '        │  0   0.7071   0.7071  7  │',
        '        │  0   0        0       1  │',
        '        └                          ┘',
    ])

    ex_q(d, '练习3-3', '右图表示手部抓握物体 Q（边长 2 个单位的正立方体），写出表达该手部位姿的矩阵表达式。')
    ex_a(d, '方法与例3-2 相同：由图读出手部坐标系原点 O′ 相对参考系的位置 P（正方体中心/顶点，依图取值，如 P = [2 2 2 1]ᵀ 之类），再由图读出 X′、Y′、Z′ 轴的三个方向余弦方向列阵 n、o、a，按 [n o a P] 列排布写出 4×4 位姿矩阵。注意 n、o、a 必须两两垂直且满足右手法则（n × o = a），可用此性质验算。')

    ex_q(d, '例3-4', '图3-10 有下面三种情况：1）动坐标系{A}相对固定坐标系作 (−1,2,2) 平移后到 {A′}；2）动坐标系{A}相对自身（动）坐标系作 (−1,2,2) 平移后到 {A″}；3）物体 Q 相对固定坐标系作 (2,6,0) 平移后到 Q′。已知 A、Q，计算 {A′}、{A″} 以及 Q′ 的矩阵表达式。')
    ex_steps(d, [
        [('平移算子 Trans(−1,2,2) 对两种情形相同；关键是左乘还是右乘。', '')],
        [('1）相对固定系 → 左乘：{A′} = Trans(−1,2,2) · A。', '')],
        [('2）相对动系（自身系）→ 右乘：{A″} = A · Trans(−1,2,2)。由所得矩阵可读出 A″ 在 O₀X₀Y₀Z₀ 系中的坐标为 (−1, 2, −1)。', '')],
        [('3）物体 Q 相对固定系平移 → 左乘：Q′ = Trans(2,6,0) · Q。', '')],
    ])
    ex_ans(d, '规则：算子左乘 = 相对固定坐标系变换；算子右乘 = 相对动坐标系变换（该规则适用于平移、旋转以及坐标系和物体的变换）。')

    ex_q(d, '练习一', '设空间点 A(2,3,−4) 相对于固定坐标系做 (−1,2,9) 平移后到 A′ 点，求 A′ 的位置。')
    ex_steps(d, ['A′ = Trans(−1,2,9) · A，即逐坐标相加：x′=2+(−1)=1，y′=3+2=5，z′=−4+9=5。'])
    ex_ans(d, 'A′ = (1, 5, 5)')

    ex_q(d, '练习二', '设空间有一动坐标系 O′，(1) 相对于固定坐标系 O 做 (1,6,−5) 平移后到 O′；(2) 相对于自身坐标系做 (1,6,−5) 平移后到 A″，求 O′ 和 A″ 的矩阵表达式。')
    ex_steps(d, [
        '(1) 固定系 → 左乘：O′ = Trans(1,6,−5) · O。',
        '(2) 动系 → 右乘：A″ = A′ · Trans(1,6,−5)。',
        '两式展开即得矩阵；注意 (2) 中平移量是相对 O′ 自身坐标轴的分量。',
    ])
    ex_ans(d, 'O′ = Trans(1,6,−5)·O；　A″ = A·Trans(1,6,−5)（右乘，动系算子）')

    ex_q(d, '练习题', '1. 设空间点 M(1,2,5) 相对于固定坐标系做 (−1,2,−1) 平移后到 M′ 点，求 M′ 的位置。2. 设空间有一动坐标系 A′，(1) 相对于固定坐标系 O 做 (1,6,−5) 平移后到 O′，(2) 相对于自身坐标系做 (1,6,−5) 平移后到 A″，求 O′ 和 A″ 的矩阵表达式。')
    ex_steps(d, [
        '1. M′ = (1+(−1), 2+2, 5+(−1)) = (0, 4, 4)。',
        '2.(1) O′ = Trans(1,6,−5)·O；(2) A″ = A·Trans(1,6,−5)。',
    ])
    ex_ans(d, 'M′ = (0, 4, 4)；　O′ = Trans(1,6,−5)·O；　A″ = A·Trans(1,6,−5)')

    ex_q(d, '例3-5', '已知坐标系中点 U 的位置矢量 u = [7 3 2 1]ᵀ，将此点绕 Z 轴旋转 90°，再绕 Y 轴旋转 90°，求旋转变换后所得的点 W。')
    ex_steps(d, [
        [('第一步（绕 Z 轴）：u′ = Rot(z,90°)·u = [0·7−1·3, 1·7+0·3, 2, 1]ᵀ = [−3, 7, 2, 1]ᵀ。', '')],
        [('第二步（绕 Y 轴）：W = Rot(y,90°)·u′ = [0×(−3)+0×7+1×2, 1×7, −1×(−3)+0×2, 1]ᵀ = [2, 7, 3, 1]ᵀ。', '')],
    ])
    ex_ans(d, 'W = [2  7  3  1]ᵀ')

    ex_q(d, '例3-6', '如图所示的单操作手臂，手腕也具有一个旋转自由度。已知手部的起始位姿矩阵为 G₁。若手臂绕 Z₀ 轴旋转 90°，则手臂到达 G₂；若手臂不动，则手部绕手腕 Z₁ 轴旋转 90°，则手部到达 G₃。写出手部坐标系 G₂、G₃ 的矩阵表达式。')
    ex_steps(d, [
        '手臂绕固定系（基座系）Z₀ 轴旋转 → 算子左乘：G₂ = Rot(z₀,90°)·G₁。',
        '手臂不动，手部绕动系 Z₁ 轴（腕轴）旋转 → 算子右乘：G₃ = G₁·Rot(z₁,90°)。',
    ])
    ex_ans(d, 'G₂ = Rot(z₀,90°)·G₁；　G₃ = G₁·Rot(z₁,90°)')

    ex_q(d, '综合变换（1）', '动坐标系{O′:u,v,w}与固定坐标系{O:x,y,z}初始重合，经下列坐标变换：①绕 z 轴旋转 90°；②绕 y 轴旋转 90°；③相对于固定坐标系平移位置矢量 [4,−3,7]。试求合成齐次变换矩阵 T。')
    ex_steps(d, [
        '①绕 z 轴 90°（固定系）：T₁ = Rot(z,90°)。',
        '②绕 y 轴 90°（固定系）：T₂ = Rot(y,90°)。',
        '③固定系平移 [4,−3,7]：T₃ = Trans(4,−3,7)。',
        '全部相对固定坐标系 → 算子左乘，且后做的变换写在最左：T = T₃·T₂·T₁。',
    ])
    d.mono_block([
        '        ┌                ┐',
        '        │  0  0  1  4  │',
        '   T =  │  1  0  0 −3  │',
        '        │  0  1  0  7  │',
        '        │  0  0  0  1  │',
        '        └                ┘',
    ])

    ex_q(d, '综合变换（2）', '动坐标系与固定坐标系初始重合，首先相对于固定坐标系平移位置矢量 [4,−3,7]，然后绕活动坐标系 v 轴旋转 90°，最后绕 w 轴旋转 90°，试求其齐次坐标变换矩阵。')
    ex_steps(d, [
        '①固定系平移：T₁ = Trans(4,−3,7)。',
        '②绕动系 v 轴（即动系 y 轴）旋转 90°：T₂ = Rot(v,90°) = Rot(y,90°)，算子右乘。',
        '③绕动系 w 轴（即动系 z 轴）旋转 90°：T₃ = Rot(w,90°) = Rot(z,90°)，算子右乘。',
        '动系算子 → 右乘，先做的写在最左：T = T₁·T₂·T₃。',
    ])
    d.mono_block([
        '        ┌                ┐',
        '        │  0  0  1  4  │',
        '   T =  │  1  0  0 −3  │',
        '        │  0  1  0  7  │',
        '        │  0  0  0  1  │',
        '        └                ┘',
    ])

    ex_q(d, '练习题（1）', '写出齐次坐标变换矩阵 AHB，它表示动坐标系{B}连续相对于固定坐标系{A}做以下变换：①绕固定坐标系 zA 轴旋转 90°；②绕 xA 轴旋转 90°；③移动 [4,−3,7]。')
    ex_steps(d, [
        '固定系算子左乘，后做的写在最左：T = Trans(4,−3,7) · Rot(x,90°) · Rot(z,90°)。',
    ])
    d.mono_block([
        '        ┌                ┐',
        '        │  0 −1  0  4  │',
        '   T =  │  0  0 −1 −3  │',
        '        │  1  0  0  7  │',
        '        │  0  0  0  1  │',
        '        └                ┘',
    ])

    ex_q(d, '练习题（2）', '写出齐次坐标变换矩阵 BHB，它表示动坐标系{B}连续相对于自身坐标系{B}做以下变换：①移动 [3,7,9]；②绕动坐标系 xB 轴旋转 90°；③绕 zB 轴旋转 90°。')
    ex_steps(d, [
        '动系算子右乘，先做的写在最左：T = Trans(3,7,9) · Rot(x,90°) · Rot(z,90°)。',
    ])
    d.mono_block([
        '        ┌                ┐',
        '        │  0 −1  0  3  │',
        '   T =  │  0  0 −1  7  │',
        '        │  1  0  0  9  │',
        '        │  0  0  0  1  │',
        '        └                ┘',
    ])

    ex_q(d, 'SCARA 正解', [('SCARA 型工业机器人的运动学方程（例）：A₁ = Rot(z,θ₁)·Trans(l₁,0,0)；A₂ = Rot(z,θ₂)·Trans(l₂,0,0)；A₃ = Rot(z,θ₃)·Trans(l₃,0,0)，求 ⁰T₃ = A₁A₂A₃。', '')])
    ex_steps(d, [
        '先写出各 Aᵢ；再两两相乘（先 A₁A₂，再乘 A₃）。',
        '记 c123 = cos(θ₁+θ₂+θ₃)，s123 = sin(θ₁+θ₂+θ₃)，c12 = cos(θ₁+θ₂)，s12 = sin(θ₁+θ₂)，c₁ = cosθ₁，s₁ = sinθ₁。',
        '展开得 ⁰T₃ 的各元素：位置列 Px = l₁c₁ + l₂c₁₂ + l₃c₁₂₃，Py = l₁s₁ + l₂s₁₂ + l₃s₁₂₃，Pz = 0（SCARA 臂垂直方向无偏移，d 参数并入 Trans(0,0,dᵢ)）；姿态前三列由 c123/s123 给出。',
    ])
    ex_ans(d, 'T₃ = A₁A₂A₃，展开后即得手部在固定坐标系中的位置和姿态；练习：自行写出 ⁰T₂ 并验证手部姿态矩阵 R = [n o a]。')

    ex_q(d, '作业', '二自由度平面关节型机器人（2R 机器人）：求手部端点的位移方程式和关节角度的方程式，即运动学的正解与逆解。')
    ex_steps(d, [
        [('正解（已知 θ₁、θ₂，求端点坐标）：', 'b'), ('x = l₁cosθ₁ + l₂cos(θ₁+θ₂)，y = l₁sinθ₁ + l₂sin(θ₁+θ₂)。', '')],
        [('逆解（已知端点 (x,y)，求 θ₁、θ₂）：', 'b')],
        [('由余弦定理：cosθ₂ = (x² + y² − l₁² − l₂²) / (2l₁l₂)，', ''), ('θ₂ = ± arccos((x²+y²−l₁²−l₂²)/(2l₁l₂))', 'kp'), ('（+ 为肘上解，− 为肘下解）。', '')],
        'θ₁ = atan2(y,x) − atan2(l₂sinθ₂, l₁ + l₂cosθ₂)。',
        '选解：根据关节运动空间限制、与前一时刻解最接近、避障要求，并逐级剔除多余解（树形结构剪枝）。',
    ])
    ex_ans(d, '正解 x/y 如上；逆解 θ₂ = ± arccos((x²+y²−l₁²−l₂²)/(2l₁l₂))，θ₁ = atan2(y,x) − atan2(l₂sinθ₂, l₁+l₂cosθ₂)。')


def build_ex_ch4(d):
    d.h1('第4章 思考与练习')

    d.h2('一、简答题')
    ex_q(d, '简答1', '工业机器人力雅可比矩阵和速度雅可比矩阵有何关系？')
    ex_a(d, '它们互为转置关系：力雅可比矩阵 J_F = Jᵀ（速度雅可比的转置），即 τ = Jᵀ(q)·F。推导（虚功原理）：末端力做的虚功 Fᵀ·δp 等于关节力矩做的虚功 τᵀ·δq；将 δp = J·δq 代入得 Fᵀ·J·δq = τᵀ·δq，对任意 δq 成立 → τ = JᵀF。这是机器人操作臂力控制的基础。')
    ex_q(d, '简答2', '利用雅可比矩阵控制机器人时，存在的问题有哪些？')
    ex_steps(d, [
        [('奇异形位问题：', 'b'), ('det(J)=0（J 的秩下降）时，某些方向的末端速度无法实现；为获得期望速度关节速度→∞，将损坏电机与机构；逆解不唯一或不存在。', '')],
        [('计算量大：', 'b'), ('每次控制周期都要计算机器人的雅可比矩阵及其逆（或伪逆），对实时性要求高。', '')],
        [('数值病态：', 'b'), ('接近奇异时 J 接近奇异，矩阵求逆放大误差，导致关节速度抖动/发散。', '')],
        [('解的选择：', 'b'), ('雅可比方程解的非唯一性（多解），需结合关节限位、连续性、避障等约束取舍。', '')],
        [('仅运动学层面：', 'b'), ('纯雅可比控制不考虑动力学耦合，高速/高负载时跟踪误差明显，需力控制或动力学补偿。', '')],
    ])

    d.h2('二、PPT 例题精讲')
    ex_q(d, '例1', '求 2-DOF 操作臂的雅可比矩阵（两连杆长度 l₁、l₂，关节角 θ₁、θ₂ 均为转动关节）。')
    ex_steps(d, [
        [('正运动学：末端位置 p = [l₁c₁ + l₂c₁₂, l₁s₁ + l₂s₁₂]ᵀ，c₁=cosθ₁，s₁₂=sin(θ₁+θ₂)。', '')],
        [('对关节变量求偏导：', 'b')],
    ])
    d.mono_block([
        '        ┌                              ┐',
        '  J =   │ ∂x/∂θ₁   ∂x/∂θ₂ │  =  │  −l₁s₁−l₂s₁₂   −l₂s₁₂  │',
        '        │ ∂y/∂θ₁   ∂y/∂θ₂ │     │   l₁c₁+l₂c₁₂    l₂c₁₂  │',
        '        └                              ┘',
    ])
    ex_steps(d, [
        'J₁、J₂ 分别表示单位关节速度在末端位置产生的速度分量的传动比；速度传递关系 dp = J₁dθ₁ + J₂dθ₂。',
        [('求奇异形位：', 'b'), ('det(J) = l₁l₂sinθ₂ = 0 → θ₂ = 0° 或 180°（手臂完全伸直/折叠）时为奇异形位，末端在径向失去运动能力。', '')],
    ])
    ex_ans(d, 'J = [[−l₁s₁−l₂s₁₂, −l₂s₁₂],[l₁c₁+l₂c₁₂, l₂c₁₂]]；奇异形位 θ₂ = 0°/180°')

    ex_q(d, '思考', '雅可比矩阵中各行、各列、各元素的物理意义是什么？')
    ex_steps(d, [
        [('列：', 'b'), ('J 的第 j 列 = 关节 j 单位速度在末端产生的线速度（与角速度）分量——即该关节运动到末端执行器的传动比。', '')],
        [('行：', 'b'), ('J 的第 i 行 = 末端第 i 个空间速度分量对各关节变量的偏导（敏感度）。', '')],
        [('元素：', 'b'), ('Jᵢⱼ = 关节 j 单位速度变化引起末端第 i 个速度分量变化的变化率。', '')],
    ])


def build_ex_ch5(d):
    d.h1('第5章 思考与练习')

    d.h2('一、判断题')
    judge = [
        ('判断1', '机器人控制系统是一个与运动学和动力学原理密切相关的、有耦合的、非线性的多变量控制系统。', '√',
         '正是因为机器人 dynamics 的强耦合/非线性多变量特性，才需要建立运动学/动力学模型进行轨迹规划与控制，机器人控制系统确为运动学与动力学密切相关的耦合、非线性多变量控制系统。'),
        ('判断2', '实际中的工业机器人，大多为串接的连杆结构，其动态特性具有高度的非线性。', '√',
         '串接连杆结构的惯量矩阵随关节位姿变化（非对角耦合），重力、科氏、离心力项均为非线性项，故动态特性具有高度非线性。'),
    ]
    for tag, q, ans, why in judge:
        ex_q(d, tag, q)
        ex_a(d, why)
        ex_ans(d, ans, '答案')

    d.h2('二、简答题')
    ex_q(d, '简答1', '点到点的运动和路径跟踪运动的区别是什么？路径跟踪运动时，能否做到实际运动路径与期望路径严格一致？')
    ex_steps(d, [
        [('点到点（PTP）：', 'b'), ('只对末端运动的起点和终点位姿加以约束，只要求平稳、准确地到达终点，对运动中经过的轨迹不加控制（如点焊、上下料）。', '')],
        [('路径跟踪（CP）：', 'b'), ('对末端运动轨迹上的空间点位姿加以约束，要求末端按给定路径、给定速度/加速度连续平滑地运动（如弧焊、喷涂）。', '')],
        [('能否严格一致：', 'b'), ('不能严格一致。实际路径与期望路径总存在偏差，只能尽可能逼近；原因包括关节限位与死区、模型误差（几何参数、标定）、驱动/控制误差、采样与插补精度有限、动力学耦合与外部扰动等。因此路径跟踪运动中实际运动路径与期望路径只能尽可能一致、不可能绝对一致。', '')],
    ])
    ex_q(d, '简答2', '机器人的运动轨迹生成方式有哪几种？')
    ex_steps(d, [
        [('按空间划分：', 'b'), ('关节空间轨迹规划（直接规划各关节角 θ(t)，再正解得末端轨迹）与笛卡尔空间轨迹规划（规划末端位姿 X(t)，再逆解得各关节角）。', '')],
        [('关节空间规划法：', 'b'), ('三次（3次）多项式插值、过路径点的三次多项式插值、高次（5次）多项式插值、抛物线过渡的线性插值（T 形速度）。', '')],
        [('笛卡尔空间规划法：', 'b'), ('笛卡尔直线插值、笛卡尔圆弧（圆）插值、笛卡尔参数样条曲线（如四元数插值姿态）等。', '')],
    ])
    ex_q(d, '简答3', '简述机器人的运动轨迹生成的过程。')
    ex_steps(d, [
        '明确任务要求：给出末端在起始点/终止点（或路径点）的位姿及运动时间、速度、加速度边界条件。',
        '选择规划空间：关节空间或笛卡尔空间（若在笛卡尔空间则先规划位姿，再逆解出各关节变量）。',
        '选择插补函数：多项式（3次/5次）、抛物线过渡的线性插值等。',
        '由边界条件确定函数系数（列方程求解），并检查位移、速度、加速度是否满足关节极限。',
        '按控制周期采样、插补（在线计算/预生成），将插补点送给位置控制器驱动电机。',
        '加入反馈校正：用编码器/位置传感器实测关节位置，与给定轨迹比较修正，保证精度与稳定性。',
    ])
    ex_q(d, '简答4', '工业机器人的关节空间运动规划法有哪些？试比较多项式插值与抛物线过渡的线性插值的优缺点。')
    ex_steps(d, [
        [('规划法：', 'b'), ('三次多项式插值；过路径点的三次多项式插值；五次（高次）多项式插值；抛物线过渡的线性插值。', '')],
        [('多项式插值优点：', 'b'), ('轨迹平滑（位置/速度/加速度连续）、边界条件可满足、过路径的精度好；五次多项式还可约束端加速度。', '')],
        [('多项式插值缺点：', 'b'), ('计算量较大、需要逆解或运动学模型；高速时加速度峰值大、关节可能饱和；过多路径点时分段繁琐。', '')],
        [('抛物线过渡线性插值优点：', 'b'), ('算法简单、计算量小、速度为 T 形便于控制、转矩脉动小。', '')],
        [('抛物线过渡线性插值缺点：', 'b'), ('加速度分段恒定（非连续光滑过渡）、轨迹不够平滑，柔性较大时易引起振动；不能精确约束端加速度。', '')],
    ])

    d.h2('三、PPT 例题精讲（轨迹规划计算题）')

    ex_q(d, '例5-1（三次多项式，15°→75°，tf=3s）', '具有一个旋转关节的单杆机器人处于静止状态时，初始角位 15°。期望在 3 秒内平滑地运动到终止位置，这时的关节角为 75°。求解满足该运动的一个三次多项式的系数，并且使操作臂在终止位置为静止状态。')
    ex_steps(d, [
        '设 θ(t) = a₀ + a₁t + a₂t² + a₃t³；则 θ′(t) = a₁ + 2a₂t + 3a₃t²。',
        '边界条件：θ(0)=15°，θ′(0)=0，θ(3)=75°，θ′(3)=0。',
        '由 θ(0)=15 → a₀=15；θ′(0)=0 → a₁=0。',
        'θ(3)=15+9a₂+27a₃=75 → 9a₂+27a₃=60，即 a₂+3a₃=6.6667。',
        'θ′(3)=6a₂+27a₃=0 → 2a₂+9a₃=0。',
        '解得 a₃=−(6.6667×2)/(9−6)=−4.4444°/s³，a₂=20°/s²。',
    ])
    ex_ans(d, 'θ(t) = 15 + 20t² − 4.4444 t³　（验证：θ(3)=15+180−120=75° ✓，θ′(3)=120−120=0 ✓）')

    ex_q(d, '练习（六轴关节 1，30°→75°，5s）', '要求一个六轴机器人的第一关节在 5s 内从初始角 30° 运动到终端角 75°，且起始点和终止点速度均为零。用三次多项式规划该关节的运动，并计算在第 1、2、3秒和第 4 秒时关节的角度。')
    ex_steps(d, [
        'θ(t) = a₀ + a₁t + a₂t² + a₃t³，θ′(t) = a₁ + 2a₂t + 3a₃t²。',
        '边界：θ(0)=30，θ′(0)=0，θ(5)=75，θ′(5)=0。',
        'θ(5)=30+25a₂+125a₃=75 → 25a₂+125a₃=45；θ′(5)=10a₂+75a₃=0 → a₂=−7.5a₃；',
        '代入：25×(−7.5a₃)+125a₃=45 → −187.5a₃+125a₃=45 → −62.5a₃=45 → a₃=−0.72，a₂=5.4。',
    ])
    ex_ans(d, 'θ(t) = 30 + 5.4t² − 0.72t³。各时刻：θ(1)=34.68°，θ(2)=45.84°，θ(3)=59.16°，θ(4)=70.32°。（验证：θ(5)=30+135−90=75° ✓）')

    ex_q(d, '例7-2（五次多项式）', '关节 1 在 5s 内从初始角 30° 运动到终止角 75°，起止速度均为零，且已知起始加速度和终止减速度均为 50°/s²。')
    ex_steps(d, [
        '设 θ(t) = a₀ + a₁t + a₂t² + a₃t³ + a₄t⁴ + a₅t⁵（（6个边界条件求6个系数）',
        '边界：θ(0)=30，θ′(0)=0，θ″(0)=50；θ(5)=75，θ′(5)=0，θ″(5)=−50。',
        'a₀=30，a₁=0，2a₂=50 → a₂=25。',
        '其余方程联立解得 a₃=−16.4，a₄=3.92，a₅=−0.3136。',
        '（验证：θ(5)=30+625−2050+2450−980=75 ✓；θ′(5)=0 ✓；θ″(5)=−50 ✓。）',
    ])
    ex_ans(d, 'θ(t) = 30 + 25t² − 16.4 t³ + 3.92 t⁴ − 0.3136 t⁵')

    ex_q(d, '例7-3（抛物线过渡的线性插值）', '在例 7-1（关节 1：30°→70°，5s）中，假设关节 1 以角速度 100°/s² 加速，5 秒内从初始角 30° 运动到目的角 70°。求解所需的过渡时间并绘制关节位置、速度和加速度曲线。')
    ex_steps(d, [
        [('设抛物线过渡段时间为 t_a，匀速段速度为 V：', 'b'), ('两过渡段各占 t_a（加速/减速对称），匀速段时间为 T−2t_a（T=5s）。', '')],
        '速度关系：V = a·t_a = 100·t_a。',
        [('位移关系：30° + 2×(½·a·t_a²) + V·(T−2t_a) = 70°', 'b'), ('，即 2×50·t_a² + 100t_a·(5−2t_a) = 40。', '')],
        '展开：100t_a² + 500t_a − 200t_a² = 40 → −100t_a² + 500t_a − 40 = 0 → 100t_a² − 500t_a + 40 = 0。',
        [('解二次方程：t_a = [500 ± √(250000−16000)] / 200 = [500 ± √234000] / 200 = [500 ± 483.73] / 200。', 'b')],
        't_a = (500−483.73)/200 = 0.0813s（另一根 t_a = 4.92s 不合题意，舍去——加速段不能长于总时间一半）。',
        'V = 100×0.0813 = 8.13°/s；T 形速度曲线：0→8.13°/s（0.0813s）、匀速 8.13°/s、再匀减至 0（0.0813s）。',
    ])
    ex_ans(d, '过渡时间 t_a = 0.0813s，匀速段角速度 V = 8.13°/s（对应 例7-3 的解法：由位移方程 V·(T−t_a)=40 也可直接解得 V = 8.13°/s）。绘制曲线：位置曲线为三段（前抛物线、中间直线、后抛物线），速度曲线呈 T 形，加速度为两端 ±100°/s² 中间 0。')

    ex_q(d, '计算题（−5°→80°，4s）', '具有一个旋转关节的单杆机器人，处于静止状态时，初始角位 −5°。期望在 4s 内平滑地运动到终止位置，这时的关节角为 80°。求出完成此运动并且使目标停在目标点的三次曲线的系数，并用 Matlab 画出位置、速度和加速度的时间函数曲线。')
    ex_steps(d, [
        'θ(t) = a₀ + a₁t + a₂t² + a₃t³；边界：θ(0)=−5，θ′(0)=0，θ(4)=80，θ′(4)=0。',
        'a₀=−5，a₁=0；16a₂+64a₃=85 → a₂+4a₃=5.3125；8a₂+48a₃=0 → a₂=−6a₃。',
        '代入 a₂=−6a₃ 到 a₂+4a₃=5.3125：−2a₃=5.3125 → a₃=−2.65625，a₂=15.9375。',
        '（复核：θ(4)=−5+15.9375×16−2.65625×64 = −5+255−170 = 80 ✓；θ′(4)=31.875×4−7.96875×16 = 127.5−127.5 = 0 ✓。）',
        [('峰值角速度：', 'b'), ('θ′(t) = 31.875t − 7.96875t²，在 t=2s 处取极大值：θ′(2) = 63.75−31.875 = 31.875°/s；端加速度 θ″(0)=+31.875°/s²，θ″(4)=−31.875°/s²。', '')],
    ])
    ex_ans(d, 'θ(t) = −5 + 15.9375 t² − 2.65625 t³；Matlab：t=0:0.01:4; theta = ...; plot(t,theta,t,gradient(theta)/10,t,gradient(gradient(theta))/1)。')


# =====================================================================
# 第三部分：速记表
# =====================================================================

def build_part3(d):
    part_title(d, '第三部分', '速记表', '考前集中记忆：公式 / 对比 / 易错点 一表打尽')

    d.h2('速记 1　机器人四大家族与关键数字')
    d.table(
        ['企业', '国家', '备注'],
        [
            ['ABB', '瑞士', '四大家族之首；6 kg 机器人，1.2 万次重复定位精度 ±0.01 mm 级'],
            ['KUKA', '德国', '汽车焊接产线大户；LBR iiwa 轻型人机协作机器人'],
            ['FANUC', '日本', '数控机床 + 机器人双巨头；全电机驱动、CRX 协作机器人'],
            ['YASKAWA', '日本', '电气驱动机器人鼻祖；MOTOMAN 系列广泛用于弧焊'],
        ],
        [2.2, 1.6, 12.2])

    d.h2('速记 2　四种坐标形式速查（直角/圆柱/球/关节）')
    d.table(
        ['坐标形式', '代号', '自由度构成', '优点', '缺点', '应用'],
        [
            ['直角坐标型', 'PPP', '3 直线', '结构简单、精度高、刚度好', '占地大、灵活性差', '装卸、搬运、数控机床上下料'],
            ['圆柱坐标型', 'RPP', '1 转 + 2 直线', '简单、取料出料直观、位置精度高、垂直升降速度快', '水平臂不能接近立柱、体积大', '大型码垛、搬运、铸造'],
            ['球坐标型', 'RRP', '2 转 + 1 直线', '作用范围大、结构紧凑', '复杂空间轨迹难实现、位置误差随距离增大', '喷涂、点焊、搬运'],
            ['关节坐标型', 'RRR', '多个转动（≥3）', '结构最紧凑、灵活性大、工作空间最大', '数学建模复杂、控制难度大', '弧焊、装配、打磨（最常用）'],
        ],
        [2.2, 1.2, 2.6, 3.6, 3.3, 3.1], size=8)

    d.h2('速记 3　第2章：机身 vs 手臂，驱动方式对比')
    d.table(
        ['机身配置', '特点'],
        [
            ['横梁式', '横梁（龙门）安装，适合大型工件/长行程，常用于焊接、喷涂流水线'],
            ['立柱式', '立柱（回转）+ 悬臂，结构简单，回转半径大，覆盖范围广'],
            ['机座式', '基座固定在机座上，结构最稳固，负载能力强，通用'],
            ['屈伸式', '手臂可屈伸折叠，避撞能力强，灵巧度高'],
        ],
        [3.0, 13.0])
    d.table(
        ['驱动', '优点', '缺点', '典型应用'],
        [
            ['液压', '功率大、动作平稳、易实现直线运动', '油源系统复杂、维护难、有泄漏', '重载搬运、锻造、压力机上下料'],
            ['气动', '气源方便、响应快、成本低、清洁安全', '空气可压缩性→精度差、功率小', '轻小型拾放、冲压上下料、食品/药'],
            ['电动', '控制精度高、响应快、效率高、易编程', '大功率时电机大、需减速装置', '绝大多数工业机器人（最常用）'],
        ],
        [1.8, 5.4, 5.2, 4.0], size=8)

    d.h2('速记 4　谐波减速器 vs RV 减速器')
    d.table(
        ['', '谐波减速器', 'RV 减速器'],
        [
            ['结构', '波发生器 + 柔轮 + 刚轮（三件）', '渐开线行星轮 + 摆线针轮（两级）'],
            ['体积/重量', '小、轻、薄', '大、重'],
            ['精度/刚度', '回差小、柔顺性好', '刚度大、回转精度高'],
            ['负载', '中小负载（适合小臂、腕、手部）', '重负载（适合基座、大臂、肩部）'],
            ['弱点', '柔轮疲劳寿命', '结构复杂、成本高'],
        ],
        [2.0, 6.6, 7.6], size=8)

    d.h2('速记 5　第3章齐次变换 · 运算规则口诀')
    d.note([('口诀　', 'b'), ('“静左动右”：相对固定坐标系变换 → 左乘；相对动坐标系变换 → 右乘。复合变换中，后做的变换写在最左（固定系）或最右（动系）。', '')])
    d.table(
        ['算子', '矩阵', '意义'],
        [
            ['Trans(a,b,c)', '[[1,0,0,a],[0,1,0,b],[0,0,1,c],[0,0,0,1]]', '平移 a（x）、b（y）、c（z）'],
            ['Rot(x,θ)', '[[1,0,0,0],[0,cosθ,−sinθ,0],[0,sinθ,cosθ,0],[0,0,0,1]]', '绕 x 轴旋转 θ'],
            ['Rot(y,θ)', '[[cosθ,0,sinθ,0],[0,1,0,0],[−sinθ,0,cosθ,0],[0,0,0,1]]', '绕 y 轴旋转 θ'],
            ['Rot(z,θ)', '[[cosθ,−sinθ,0,0],[sinθ,cosθ,0,0],[0,0,1,0],[0,0,0,1]]', '绕 z 轴旋转 θ'],
            ['Scale(Sx,Sy,Sz)', '[[Sx,0,0,0],[0,Sy,0,0],[0,0,Sz,0],[0,0,0,1]]', '比例缩放'],
        ],
        [2.6, 8.6, 4.6], size=8)
    d.bullets([
        [('位姿矩阵 [n o a P]：', 'b'), ('n、o、a 分别为 X′、Y′、Z′ 轴的方向（单位）矢量，P 为原点位置；三列 n、o、a 两两正交且 n×o=a（右手法则）。', '')],
        [('DH 四参数：', 'b'), ('θᵢ（关节旋转角）、dᵢ（连杆偏距）、aᵢ（连杆长度）、αᵢ（连杆扭角）；旋转关节：变量 θᵢ、常量 dᵢ aᵢ αᵢ；移动关节：变量 dᵢ、常量 θᵢ aᵢ αᵢ。', '')],
        [('第一变换法 DH：Aᵢ = Rot(z,θᵢ)·Trans(0,0,dᵢ)·Trans(aᵢ,0,0)·Rot(x,αᵢ)。', '')],
        [('左乘（前乘）与右乘（后乘）：', 'b'), ('复合变换中，物体/坐标系相对固定系依次做变换 T₁、T₂、…、Tₙ 时，总变换 = T₁·T₂…Tₙ（后做的左乘进来）；相对动系依次做变换时，总变换 = T₁·T₂…Tₙ（后做的右乘上去）。', '')],
    ])

    d.h2('速记 6　第4章雅可比矩阵速记')
    d.bullets([
        [('定义：', 'b'), ('末端速度 dp/dt 与关节速度 dq/dt 的线性映射：dp = J(q)·dq，即 dp/dt = J(q)·dq/dt。', '')],
        [('构成：', 'b'), ('J = [J₁ J₂ … Jₙ]，第 j 列 = 第 j 个关节单位速度对末端速度的传动分量；对转动关节：J（线速列）= zⱼ × (p_e − pⱼ)；对移动关节：J = zⱼ（zⱼ 为关节轴方向单位矢量，pⱼ 为关节原点）。', '')],
        [('力雅可比：', 'b'), ('τ = Jᵀ·F（虚功原理：Fᵀδp = τᵀδq）；是力控制的基础。', '')],
        [('奇异：', 'b'), ('det(J) = 0 时 J 降秩 → 某些末端速度方向不可达、逆解不唯一；2R 机器人 l₁l₂sinθ₂ = 0（手臂伸直/折叠）为奇异形位。', '')],
        [('用途：', 'b'), ('①速度分解到关节；②奇异分析与避碰；③静力传递与力控制；④逆运动学数值解（牛顿迭代）。', '')],
    ])

    d.h2('速记 7　第5章轨迹规划 · 三次/五次多项式系数速查')
    d.table(
        ['方法', '函数', '边界条件', '关键结果'],
        [
            ['三次多项式', 'θ = a₀+a₁t+a₂t²+a₃t³', '起终位移 + 起终速度（=0）共 4 个', 'a₀、a₁、a₂、a₃ 由边界方程解出；速度连续、加速度端点突变'],
            ['五次多项式', 'θ = a₀+…+a₅t⁵', '起终位移 + 速度 + 加速度共 6 个', '端加速度可约束，轨迹更平滑（冲击小）'],
            ['抛物线过渡+线性插值', '抛物线段（a 恒定）+ 匀速线性段', '位置 + 起终速度=0 + 过渡加速度 a', 'T 形速度；先由位移方程求过渡时间 t_a 与匀速 V：见例 7-3'],
        ],
        [2.6, 4.6, 4.6, 4.6], size=8)
    d.note([('三次多项式通式系数（起终速度为零）：', 'b'), ('a₀=θ₀，a₁=0，a₂=3(θf−θ₀)/tf²，a₃=−2(θf−θ₀)/tf³；则 θ(t)=θ₀+3Δθ·(t/tf)²−2Δθ·(t/tf)³。', '')])

    d.h2('速记 8　易混点一表辨析')
    d.table(
        ['易混点', '正确说法', '常见错误'],
        [
            ['运动学正/逆问题', '已知关节求末端 = 正解；已知末端求关节 = 逆解', '说反'],
            ['绝对定位精度 vs 重复定位精度', '工业机器人要求重复定位精度高', '认为要求绝对定位精度高'],
            ['直角坐标型 vs 关节坐标型', '关节坐标型（RRR）灵活性大、应用最多', '认为直角坐标型最常用'],
            ['左乘 vs 右乘', '固定系 → 左乘；动系 → 右乘', '不分情形一律左乘'],
            ['速度雅可比 vs 力雅可比', '力雅可比 = 速度雅可比的转置，τ=JᵀF', '以为力雅可比=速度雅可比'],
            ['点到点 vs 路径跟踪', 'PTP 不控中间轨迹；CP 控制整条轨迹', '认为 PTP 也控制路径'],
            ['自由度=灵活性≠承载', '自由度高灵活性高，但刚性/承载不一定提升', '认为自由度越高承载越强'],
            ['谐波减速器 vs RV', '重负载用 RV（基/大臂），轻负载用谐波（小臂/腕/手部）', '位置用反'],
        ],
        [3.4, 7.2, 5.2], size=8)

    d.h2('速记 9　例题数字一览（直接背答案）')
    d.table(
        ['题', '答案/结论'],
        [
            ['例5-1　15°→75°，tf=3s，三次多项式', 'θ(t)=15+20t²−(40/9)t³（a₃=−4.4444）'],
            ['练习　关节1　30°→75°，5s，三次多项式', 'θ(t)=30+5.4t²−0.72t³；θ(1)=34.68°，θ(2)=45.84°，θ(3)=59.16°，θ(4)=70.32°'],
            ['例7-2　同上条件 + α₀=αf=50°/s²，五次', 'θ(t)=30+25t²−16.4t³+3.92t⁴−0.3136t⁵'],
            ['例7-3　30°→70°，5s，a=100°/s²', 't_a=0.0813s，V=8.13°/s（T 形速度）'],
            ['计算　−5°→80°，4s，三次多项式', 'θ(t)=−5+15.9375t²−2.65625t³；峰值角速度 θ′(2s)=31.875°/s'],
            ['综合变换（1）T=Trans(4,−3,7)·Ry(90°)·Rz(90°)', 'R 部分 [[0,0,1],[1,0,0],[0,1,0]]，P=[4,−3,7]ᵀ'],
            ['综合变换（2）先平移再绕动系 v、w 旋转', '与（1）结果相同：R=[[0,0,1],[1,0,0],[0,1,0]]，P=[4,−3,7]ᵀ'],
            ['练习题（1）固定系：绕 z 90°→绕 x 90°→平移[4,−3,7]', 'R=[[0,−1,0],[0,0,−1],[1,0,0]]，P=[4,−3,7]ᵀ'],
            ['练习题（2）动系：平移[3,7,9]→绕 xB 90°→绕 zB 90°', 'R=[[0,−1,0],[0,0,−1],[1,0,0]]，P=[3,7,9]ᵀ'],
            ['例3-5　u=[7 3 2 1]ᵀ 绕 z 90°→绕 y 90°', 'W=[2, 7, 3, 1]ᵀ'],
            ['2R 机器人（l₁=l₂=1）雅可比', 'J=[[−sinθ₁−sin(θ₁+θ₂), −sin(θ₁+θ₂)],[cosθ₁+cos(θ₁+θ₂), cos(θ₁+θ₂)]]；θ₂=0/180° 为奇异'],
        ],
        [7.2, 9.2], size=8)

    d.h2('速记 10　判断题高频考点（√/× 一题一句话）')
    d.bullets([
        '数控机床不是工业机器人（没有通用性/可变程序特征）',
        '直角坐标机器人并非应用最多的结构（关节坐标型 RRR 才是）',
        '自由度越高 ≠ 承载能力越强（灵活 vs 刚性/承载要兼顾）',
        '最大工作速度并非越大越好（要兼顾加/减速度与变加速度能力）',
        '工业机器人重复定位精度高于绝对定位精度',
        '运动学正/逆问题不可混淆；力雅可比 = 速度雅可比转置',
        'PTP 运动不控制中间轨迹；CP（路径跟踪）才跟踪整条路径',
        '实际路径不可能与期望路径严格一致，只能逼近',
        '串接连杆结构动态特性高度非线性、强耦合',
        'RV 放重负载位置（机座/大臂），谐波放轻负载位置（小臂/腕/手部）',
    ])

    d.h2('速记 11　答题模板（论述/简答通用）')
    d.bullets([
        [('问“定义”：', 'b'), ('先给出 ISO 定义（自动、位置可控、可编程、多功能机械手）→ 再答三点要素：可编程、拟人化、通用性。', '')],
        [('问“特点/三原则”：', 'b'), ('拟人化 + 通用性 + 适应性 + 独立自主运动；阿西莫夫三原则可顺带补。', '')],
        [('问“组成”：', 'b'), ('三部分（驱动、执行、控制）或六子系统（机械系统、驱动系统、感知系统、控制系统、人机交互系统、机器人-环境交互系统）。', '')],
        [('问“对比（驱动/减速器/坐标型）”：', 'b'), ('先结论后理由：“重负载用 RV，因为…；轻负载用谐波，因为…”。', '')],
        [('问“为什么关节坐标型应用最广”：', 'b'), ('结构最紧凑、体积小、工作空间大、灵活性高、靠近关节无干涉；代价是建模/控制复杂。', '')],
        [('问“奇异”：', 'b'), ('det(J)=0 → 降秩 → 末端某些方向失去运动能力 → 关节速度发散 → 需避开奇异形位或阻尼最小二乘法。', '')],
        [('问“轨迹规划步骤”：', 'b'), ('选空间 → 选插补函数 → 定系数 → 采样插补 → 反馈修正。', '')],
    ])

    d.h2('速记 12　考前 5 分钟 checklist')
    d.bullets([
        '能默写 ISO 对工业机器人的定义（自动、位置可控、可编程、多功能机械手）',
        '能默写四大家族：ABB、KUKA、FANUC、YASKAWA',
        '能区分 PPP/RPP/RRP/RRR 及各自优缺点',
        '能写出 Trans / Rotx / Roty / Rotz 的 4×4 矩阵',
        '能口算：绕固定系左乘、绕动系右乘',
        '能默写 DH 四参数 θ/d/a/α 及 Aᵢ 的通式',
        '能写 2R 机器人雅可比并说明奇异（θ₂=0°/180°）',
        '能默写三次多项式系数公式 a₂=3Δθ/tf²、a₃=−2Δθ/tf³',
        '能背例5-1/例7-2/例7-3 的关键数字',
        '能区分 PTP 与 CP；能区分速度雅可比与力雅可比',
    ])


# =====================================================================
# 生成
# =====================================================================

def main():
    d = Doc()
    add_header_footer(d.doc)
    build_cover(d)
    build_toc(d)
    build_part1(d)
    build_part2(d)
    build_part3(d)
    out = '《机器人学导论》期末复习精要.docx'
    d.doc.save(out)
    print('saved', out)


if __name__ == '__main__':
    main()
