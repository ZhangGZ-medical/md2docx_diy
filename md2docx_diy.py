#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown → 竖版A4中文DOCX转换器
支持：标题H1-H6、粗体、斜体、粗斜体、代码、引用块、代码块、表格、列表、分隔线、锚点书签
"""
import sys, re, os
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# XML辅助函数
# ============================================================

def shd(cell_or_para, hex_color):
    """为单元格或段落设置背景色。"""
    el = cell_or_para._element
    if el.tag.endswith('}tc'):
        pPr = el.get_or_add_tcPr()
    else:
        pPr = el.get_or_add_pPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), hex_color)
    pPr.append(s)


def cell_margins(cell, top=0.15, bottom=0.15, left=0.2, right=0.2):
    """设置单元格内边距。"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    m = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        e = OxmlElement(f'w:{side}')
        e.set(qn('w:w'), str(int(val * 567)))
        e.set(qn('w:type'), 'dxa')
        m.append(e)
    tcPr.append(m)


def para_spacing(para, before=2, after=3):
    """设置段落间距和行距。"""
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15


def set_cell_border(cell):
    """设置单元格边框。"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), 'CCCCCC')
        tcBorders.append(b)
    tcPr.append(tcBorders)


def run_font(run, fname, fsize, bold=False, italic=False, color=None, eastAsia=None):
    """设置run的字体属性。"""
    run.font.name = fname
    run._element.rPr.rFonts.set(qn('w:eastAsia'), eastAsia or fname)
    run.font.size = Pt(fsize)
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_left_border(para, color='2C5AA0', sz='12'):
    """为段落添加左边框（用于引用块）。"""
    pPr = para._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), sz)
    left.set(qn('w:space'), '8')
    left.set(qn('w:color'), color)
    pBdr.append(left)
    pPr.append(pBdr)


def add_bottom_border(para, color='4A90C2', sz='6'):
    """为段落添加底边框（用于标题）。"""
    pPr = para._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), sz)
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_page_number(footer_para):
    """在页脚添加动态页码字段。"""
    run = footer_para.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = ' PAGE '
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run.font.size = Pt(8)


def add_bookmark(para, bookmark_name, text):
    """在段落后添加Word书签，保留目录跳转功能。"""
    # 在段落的XML元素中添加bookmarkStart和bookmarkEnd
    p = para._element
    # 生成唯一ID
    bm_id = abs(hash(bookmark_name)) % 65534

    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), str(bm_id))
    bm_start.set(qn('w:name'), bookmark_name)
    p.append(bm_start)

    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), str(bm_id))
    p.append(bm_end)


# ============================================================
# 内联格式解析（字符扫描方式，避免regex冲突）
# ============================================================

def _add_run(para, text, fname, fsize, bold=False, italic=False,
             strike=False, color=None, eastAsia=None):
    if not text:
        return
    run = para.add_run(text)
    run_font(run, fname, fsize, bold=bold, italic=italic, color=color,
             eastAsia=eastAsia)
    if strike:
        run.font.strike = True


def format_text(para, text, fname='宋体', fsize=10.5, eastAsia='宋体'):
    """解析内联Markdown标记并分run添加到段落。

    处理顺序（优先级从高到低）：
    1. 图片和链接（先去除，保留文本）
    2. 粗斜体 ***text***
    3. 粗体 **text**
    4. 代码 `text`
    5. 斜体 *text*
    6. 删除线 ~~text~~
    """
    # 去除图片和链接标记
    text = re.sub(r'!\[([^\]]*)\]\([^\)]+\)', r'\1', text)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)

    i, n, current = 0, len(text), ''

    while i < n:
        ch = text[i]

        # 粗斜体 ***text*** 或 ___text___
        if i + 5 <= n and text[i:i+3] in ('***', '___'):
            end = text.find(text[i:i+3], i + 3)
            if end != -1:
                if current:
                    _add_run(para, current, fname, fsize, eastAsia=eastAsia)
                    current = ''
                _add_run(para, text[i+3:end], fname, fsize,
                         bold=True, italic=True, eastAsia=eastAsia)
                i = end + 3
                continue

        # 粗体 **text** 或 __text__
        if i + 3 <= n and text[i:i+2] in ('**', '__'):
            end = text.find(text[i:i+2], i + 2)
            if end != -1:
                if current:
                    _add_run(para, current, fname, fsize, eastAsia=eastAsia)
                    current = ''
                _add_run(para, text[i+2:end], fname, fsize,
                         bold=True, eastAsia=eastAsia)
                i = end + 2
                continue

        # 代码 `text`
        if ch == '`':
            end = text.find('`', i + 1)
            if end != -1:
                if current:
                    _add_run(para, current, fname, fsize, eastAsia=eastAsia)
                    current = ''
                _add_run(para, text[i+1:end], 'Courier New', fsize - 1,
                         color='0066AA', eastAsia=eastAsia)
                i = end + 1
                continue

        # 斜体 *text*（不在粗体内）
        if ch in ('*', '_') and i + 1 < n \
           and text[i+1] != ch and text[i+1] not in (' ', '\n'):
            end = text.find(ch, i + 1)
            if end != -1 and end < n - 1 and text[end+1] != ch:
                if current:
                    _add_run(para, current, fname, fsize, eastAsia=eastAsia)
                    current = ''
                _add_run(para, text[i+1:end], fname, fsize,
                         italic=True, eastAsia='楷体')
                i = end + 1
                continue

        # 删除线 ~~text~~
        if i + 3 <= n and text[i:i+2] == '~~':
            end = text.find('~~', i + 2)
            if end != -1:
                if current:
                    _add_run(para, current, fname, fsize, eastAsia=eastAsia)
                    current = ''
                _add_run(para, text[i+2:end], fname, fsize,
                         strike=True, eastAsia=eastAsia)
                i = end + 2
                continue

        current += ch
        i += 1

    if current:
        _add_run(para, current, fname, fsize, eastAsia=eastAsia)


# ============================================================
# 表格检测与渲染
# ============================================================

def is_md_table_start(lines, i):
    """判断lines[i]是否为Markdown表格首行。"""
    if i >= len(lines):
        return False
    line = lines[i].strip()
    if not line.startswith('|') or line.count('|') < 2:
        return False
    if i + 1 >= len(lines):
        return False
    sep = lines[i + 1].strip()
    return bool(re.match(r'^[\|\-\s:]+$', sep))


def parse_md_table(lines, start):
    """解析Markdown表格，返回(rows, next_index)。"""
    rows, i = [], start
    while i < len(lines):
        line = lines[i].strip()
        # 跳过分隔行
        if re.match(r'^[\|\-\s:]+$', line):
            i += 1
            continue
        # 非表格行则停止
        if not line.startswith('|'):
            break
        line = line.strip('|')
        cells = [c.strip() for c in line.split('|')]
        if any(cells):
            rows.append(cells)
        i += 1
    return rows, i


def render_table(doc, rows):
    """将二维列表渲染为Word表格。"""
    if not rows:
        return
    n_cols = len(rows[0])
    avail = 25.0  # 可用宽度(cm)

    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            cell.width = Cm(avail / n_cols)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            cell_margins(cell)
            set_cell_border(cell)

            if r_idx == 0:
                shd(cell, '1A3A5C')
                para = cell.paragraphs[0]
                para.clear()
                run = para.add_run(cell_text)
                run_font(run, '黑体', 9.5, bold=True, color='FFFFFF',
                         eastAsia='黑体')
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_spacing(para, 2, 2)
            else:
                bg = 'F0F7FC' if r_idx % 2 == 0 else 'FFFFFF'
                shd(cell, bg)
                para = cell.paragraphs[0]
                para.clear()
                format_text(para, cell_text, fname='宋体', fsize=9.5,
                            eastAsia='宋体')
                para_spacing(para, 2, 2)


# ============================================================
# 块级元素
# ============================================================

def add_code_block(doc, code_lines):
    """代码块：灰色背景框。"""
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.width = Cm(25)
    shd(cell, 'F5F5F5')
    cell_margins(cell, 0.2, 0.2, 0.3, 0.3)
    para = cell.paragraphs[0]
    para.clear()
    run = para.add_run('\n'.join(code_lines))
    run_font(run, 'Courier New', 8.5, color='333333', eastAsia='宋体')
    para_spacing(para, 2, 2)


def add_blockquote(doc, text):
    """引用块：浅蓝背景+蓝色左边框。"""
    para = doc.add_paragraph()
    shd(para, 'F0F7FC')
    add_left_border(para)
    para.paragraph_format.left_indent = Cm(0.3)
    format_text(para, text.strip(), fname='宋体', fsize=10,
                eastAsia='宋体')
    para_spacing(para, before=3, after=6)


def add_hr(doc):
    """分隔线。"""
    para = doc.add_paragraph()
    run = para.add_run('─' * 70)
    run.font.color.rgb = RGBColor(74, 144, 194)
    run.font.size = Pt(7)
    para_spacing(para, before=3, after=3)


def add_bullet(doc, text, indent=0):
    """项目列表。"""
    para = doc.add_paragraph()
    base_indent = Cm(0.5 + indent * 0.5)
    para.paragraph_format.left_indent = base_indent
    para.paragraph_format.first_line_indent = Cm(-0.35)
    run = para.add_run('• ')
    run_font(run, '宋体', 10, color='2C5AA0', eastAsia='宋体')
    format_text(para, text.strip(), fname='宋体', fsize=10,
                eastAsia='宋体')
    para_spacing(para, before=1, after=2)


def add_heading(doc, text, level, anchor=None):
    """标题（支持生成书签）。"""
    para = doc.add_paragraph()
    add_bottom_border(para)

    # 去除行内链接标记
    text_clean = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)

    run = para.add_run(text_clean.strip())
    if level == 1:
        run_font(run, '黑体', 16, bold=True, color='1A3A5C', eastAsia='黑体')
        para_spacing(para, before=14, after=6)
    elif level == 2:
        run_font(run, '黑体', 13, bold=True, color='2C5AA0', eastAsia='黑体')
        para_spacing(para, before=10, after=4)
    elif level == 3:
        run_font(run, '黑体', 11, bold=True, color='2C5AA0', eastAsia='黑体')
        para_spacing(para, before=8, after=3)
    elif level == 4:
        run_font(run, '黑体', 10, bold=True, color='2C5AA0', eastAsia='黑体')
        para_spacing(para, before=6, after=2)
    else:
        run_font(run, '黑体', 9, bold=True, color='555555', eastAsia='黑体')
        para_spacing(para, before=4, after=2)

    # 生成书签（保留目录跳转）
    if anchor:
        add_bookmark(para, anchor, text_clean.strip())


# ============================================================
# 主转换函数
# ============================================================

def md_to_docx(md_path, docx_path, header_text=None):
    """将Markdown文件转换为横版A4中文DOCX。

    Args:
        md_path: 输入的Markdown文件路径
        docx_path: 输出的DOCX文件路径
        header_text: 页眉文本（默认使用文件名）
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(29.7)
    sec.page_height = Cm(21.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.5)

    # 页眉
    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = hp.add_run(header_text or f'文件：{os.path.basename(md_path)}')
    run_font(run, '宋体', 8, color='808080', eastAsia='宋体')

    # 页脚
    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run(f'{os.path.basename(md_path)} | 第 ')
    run_font(run, '宋体', 8, color='808080', eastAsia='宋体')
    add_page_number(fp)
    run = fp.add_run(' 页')
    run_font(run, '宋体', 8, color='808080', eastAsia='宋体')

    lines = content.split('\n')
    n = len(lines)
    i = 0

    while i < n:
        line = lines[i]

        # 跳过目录链接行
        if re.match(r'^\d+\.\s+\[', line):
            i += 1
            continue

        # ---- 标题（含锚点处理）----
        hm = re.match(r'^(#{1,6})\s+(.+)', line)
        if hm:
            level = len(hm.group(1))
            raw_text = hm.group(2).strip()
            # 提取锚点名
            anchor = None
            # 锚点格式1: <a name="xxx">
            am = re.search(r'<a\s+name="([^"]+)"[^>]*>', raw_text)
            if am:
                anchor = am.group(1)
            # 锚点格式2: {#anchor-name}
            im = re.search(r'\{\#([^}]+)\}', raw_text)
            if im:
                anchor = im.group(1)
            add_heading(doc, raw_text, level, anchor=anchor)
            i += 1
            continue

        # ---- 表格 ----
        if is_md_table_start(lines, i):
            rows, next_i = parse_md_table(lines, i)
            if rows:
                render_table(doc, rows)
                sp = doc.add_paragraph()
                para_spacing(sp, before=0, after=4)
            i = next_i
            continue

        # ---- 代码块 ----
        if line.strip().startswith('```'):
            i += 1
            code_lines = []
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1
            add_code_block(doc, code_lines)
            continue

        # ---- 分隔线 ----
        if re.match(r'^---+$', line.strip()):
            add_hr(doc)
            i += 1
            continue

        # ---- 引用块 ----
        if line.strip().startswith('>'):
            bq_lines = []
            while i < n and lines[i].strip().startswith('>'):
                bq_lines.append(lines[i].strip()[1:].strip())
                i += 1
            add_blockquote(doc, ' '.join(bq_lines))
            continue

        # ---- 列表 ----
        bm = re.match(r'^(\s*)[-*]\s+(.+)', line)
        if bm:
            indent = len(bm.group(1)) // 2
            add_bullet(doc, bm.group(2), indent)
            i += 1
            continue

        # ---- 空行 ----
        if not line.strip():
            i += 1
            continue

        # ---- 普通段落 ----
        para = doc.add_paragraph()
        format_text(para, line, fname='宋体', fsize=10.5, eastAsia='宋体')
        para_spacing(para, before=2, after=3)
        i += 1

    doc.save(docx_path)
    kb = os.path.getsize(docx_path) / 1024
    print(f'  Saved: {docx_path} ({kb:.0f} KB)')


# ============================================================
# 批量转换入口
# ============================================================
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Markdown → 横版A4中文DOCX')
    parser.add_argument('md_file', nargs='?', help='输入MD文件路径')
    parser.add_argument('-o', '--output', help='输出DOCX文件路径（默认同目录同名）')
    parser.add_argument('--all', action='store_true',
                        help='转换als_review目录下所有MD文件')
    args = parser.parse_args()

    if args.all:
        base = r'd:\.openclaw\workspace\als_review'
        for f in os.listdir(base):
            if f.endswith('.md'):
                md = os.path.join(base, f)
                docx = md.replace('.md', '.docx')
                print(f'Processing: {f}')
                md_to_docx(md, docx)
    elif args.md_file:
        md = os.path.abspath(args.md_file)
        docx = args.output or md.replace('.md', '.docx')
        md_to_docx(md, docx)
    else:
        parser.print_help()
