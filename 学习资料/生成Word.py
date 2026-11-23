"""将同目录复习讲义转换为排版后的 Word。依赖：python-docx。"""
from pathlib import Path
import re
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / '机器人学_知识点总结与练习题.md'
OUTPUT = ROOT / '机器人学_知识点总结与练习题（含答案）.docx'

def font(style, name, size, color=None):
    style.font.name = name
    style.font.size = Pt(size)
    style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), name)
    if color:
        style.font.color.rgb = RGBColor.from_string(color)

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), fill); tcPr.append(shd)

def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(2), Cm(1.8)
    sec.left_margin, sec.right_margin = Cm(2), Cm(2)
    sec.header_distance, sec.footer_distance = Cm(.8), Cm(.8)
    font(doc.styles['Normal'], '微软雅黑', 10.5, '253247')
    normal = doc.styles['Normal'].paragraph_format
    normal.space_after = Pt(6)
    normal.line_spacing = 1.22
    for name, size in [('Title', 27), ('Heading 1', 19), ('Heading 2', 13)]:
        font(doc.styles[name], '微软雅黑', size, '163B65')
    doc.styles['Heading 1'].paragraph_format.space_before = Pt(15)
    doc.styles['Heading 2'].paragraph_format.space_before = Pt(11)
    font(doc.styles['Subtitle'], '微软雅黑', 14, '51728B')
    code_style = doc.styles.add_style('Code Block', 1)
    font(code_style, 'Consolas', 9)
    code_style.paragraph_format.line_spacing = 1.05
    code_style.paragraph_format.space_after = Pt(1)
    code_style.paragraph_format.left_indent = Cm(.25)
    header = sec.header.paragraphs[0]
    header.text = 'ROBOTICS  /  机器人学复习讲义'
    header.runs[0].font.size = Pt(8)
    header.runs[0].font.color.rgb = RGBColor.from_string('51728B')
    footer = sec.footer.paragraphs[0]
    footer.alignment = 2
    footer.add_run('知识总结 · 60 题 · 参考答案    |    ')
    field = OxmlElement('w:fldSimple'); field.set(qn('w:instr'), 'PAGE'); footer._p.append(field)
    for r in footer.runs: r.font.size = Pt(8)
    doc.core_properties.title = '机器人学：知识点总结与练习题（含答案）'
    doc.core_properties.subject = '基于五章 PPT 的课程复习及专业拓展'
    doc.core_properties.author = '学习资料整理'
    lines = SOURCE.read_text().splitlines()
    i = 0; in_code = False; code_lines=[]
    while i < len(lines):
        line = lines[i]; i += 1
        if line.startswith('```'):
            if in_code:
                p = doc.add_paragraph(style='Code Block')
                p.add_run('\n'.join(code_lines))
                p.paragraph_format.keep_together = True
                pp = p._p.get_or_add_pPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'F0F4F8'); pp.append(shd)
                code_lines=[]
            in_code = not in_code
            continue
        if in_code:
            code_lines.append(line); continue
        if not line.strip(): continue
        if line.startswith('|'):
            rows = [line]
            while i < len(lines) and lines[i].startswith('|'):
                rows.append(lines[i]); i += 1
            rows = [r for r in rows if not re.fullmatch(r'[|\- :]+',r)]
            parsed = [[c.strip() for c in r.strip('|').split('|')] for r in rows]
            table=doc.add_table(rows=0, cols=len(parsed[0])); table.style='Table Grid'
            for ri, row in enumerate(parsed):
                cells=table.add_row().cells
                for ci, text in enumerate(row):
                    cells[ci].text=text
                    for p in cells[ci].paragraphs:
                        p.paragraph_format.space_after=Pt(4)
                        p.paragraph_format.space_before=Pt(4)
                        p.paragraph_format.line_spacing=1.1
                        for r in p.runs:
                            r.font.size=Pt(9)
                            if ri==0: r.bold=True; r.font.color.rgb=RGBColor(255,255,255)
                    shade(cells[ci], '163B65' if ri==0 else ('F0F4F8' if ri%2 else 'FFFFFF'))
                trPr=table.rows[-1]._tr.get_or_add_trPr()
                cant=OxmlElement('w:cantSplit'); trPr.append(cant)
                if ri==0:
                    repeat=OxmlElement('w:tblHeader'); trPr.append(repeat)
            doc.add_paragraph().paragraph_format.space_after=Pt(1)
        elif line.startswith('# '):
            doc.add_paragraph(line[2:], 'Title')
        elif line.startswith('## '):
            title=line[3:]
            if title.startswith('基于'):
                doc.add_paragraph(title, 'Subtitle')
            else:
                p=doc.add_heading(title, level=1)
                p.paragraph_format.page_break_before=True
        elif line.startswith('### '):
            p=doc.add_heading(line[4:],level=2)
            if line[4:].startswith(('A 组答案','B 组答案','C 组答案','D 组答案','B 组：','C 组：','D 组：')):
                p.paragraph_format.page_break_before=True
        else:
            p=doc.add_paragraph(line)
            if re.match(r'^\d+\.【',line):
                p.paragraph_format.space_before=Pt(6)
                p.paragraph_format.keep_together=True
            if line.startswith(('【课件主线】','本节超出')):
                for r in p.runs:
                    r.italic=True; r.font.color.rgb=RGBColor.from_string('51728B')
    doc.save(OUTPUT)
    print(OUTPUT)
    print('Paragraphs:',len(doc.paragraphs),'Tables:',len(doc.tables))

if __name__=='__main__': main()
