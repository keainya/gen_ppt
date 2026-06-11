#!/usr/bin/env python3
"""
build_ppt.py — 将结构化 Markdown 编译为 PPT 演示文稿。

读取 content.md，按 --- 分隔页面，识别封面页、内容页、图片展示页，
生成 output.pptx。
"""

import os
import re
import sys

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ── 常量 ────────────────────────────────────────────────────

SLIDE_WIDTH  = Inches(13.333)   # 16:9 宽屏
SLIDE_HEIGHT = Inches(7.5)
BG_COVER     = "src/cover.png"       # 封面背景
BG_CONTENT   = "src/background.png"       # 内容页 / 图片页背景
INPUT_FILE   = "content.md"
OUTPUT_FILE  = "output.pptx"

WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x33, 0x33, 0x33)

# 字体
FONT_LATIN = "Times New Roman"   # 英文/拉丁字体
FONT_CN_H1 = "SimSun"            # 宋体（一级标题）
FONT_CN_H2 = "SimSun"            # 宋体（二级标题）
FONT_CN_BODY = "FangSong"        # 仿宋（正文）


# ── 工具函数 ────────────────────────────────────────────────

def _set_font(p, cn_font: str, size, bold: bool = False):
    """设置段落的中文字体、英文字体、字号和粗细。"""
    p.font.name = FONT_LATIN
    p.font.size = size
    p.font.bold = bold
    p.font.color.rgb = DARK

    # 通过 XML 设置东亚字体（中文字体）
    pPr = p._p.get_or_add_pPr()
    defRPr = pPr.get_or_add_defRPr()
    ea = defRPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(defRPr, qn("a:ea"))
    ea.set("typeface", cn_font)


def _set_cell_font(cell, cn_font: str, size, bold: bool = False):
    """设置表格单元格的中英文字体、字号和粗细。"""
    for paragraph in cell.text_frame.paragraphs:
        paragraph.font.name = FONT_LATIN
        paragraph.font.size = size
        paragraph.font.bold = bold
        paragraph.font.color.rgb = DARK
        pPr = paragraph._p.get_or_add_pPr()
        defRPr = pPr.get_or_add_defRPr()
        ea = defRPr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(defRPr, qn("a:ea"))
        ea.set("typeface", cn_font)


def _style_table_three_line(tbl, num_rows: int, num_cols: int):
    """应用三段式表格：无填充、无竖线，仅顶/中/底三条横线。"""
    BORDER_W = "19000"   # EMU ≈ 1.5pt
    BORDER_C = "333333"

    # 移除表格默认样式
    tblPr = tbl._tbl.tblPr
    for sid in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(sid)

    for ri in range(num_rows):
        for ci in range(num_cols):
            cell = tbl.cell(ri, ci)
            tcPr = cell._tc.get_or_add_tcPr()

            # 移除所有填充
            for fe in tcPr.findall(qn("a:solidFill")):
                tcPr.remove(fe)

            # 移除所有旧边框
            for tag in ("a:lnT", "a:lnB", "a:lnL", "a:lnR"):
                for el in tcPr.findall(qn(tag)):
                    tcPr.remove(el)

            # 表头行：顶线 + 底线（中间分隔线）
            if ri == 0:
                _add_tc_border(tcPr, "a:lnT", BORDER_W, BORDER_C)
                _add_tc_border(tcPr, "a:lnB", BORDER_W, BORDER_C)

            # 最后一行：底线
            if ri == num_rows - 1:
                _add_tc_border(tcPr, "a:lnB", BORDER_W, BORDER_C)


def _add_tc_border(tcPr, tag: str, width: str, color: str):
    """向单元格属性添加一条边框线。"""
    ln = etree.SubElement(tcPr, qn(tag))
    ln.set("w", width)
    sf = etree.SubElement(ln, qn("a:solidFill"))
    sc = etree.SubElement(sf, qn("a:srgbClr"))
    sc.set("val", color)


def _set_background(slide, image_path: str):
    """设置幻灯片背景：图片存在则铺满全幅，否则纯白。"""
    if os.path.exists(image_path):
        pic = slide.shapes.add_picture(image_path, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
        # 将背景图片移至 z-order 最底层
        sp = pic._element
        spTree = sp.getparent()
        spTree.remove(sp)
        spTree.insert(2, sp)
    else:
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = WHITE


# ── Markdown 解析 ───────────────────────────────────────────

def split_sections(md_text: str) -> list[str]:
    """按行首的 --- 分隔为多个页面段落。"""
    lines = md_text.split("\n")
    sections: list[str] = []
    current: list[str] = []

    for line in lines:
        if line.strip() == "---":
            if current:
                sections.append("\n".join(current).strip())
                current = []
        else:
            current.append(line)

    if current:
        sections.append("\n".join(current).strip())

    return [s for s in sections if s]


def detect_page_type(section: str) -> str:
    """检测页面类型。"""
    has_image = False
    has_structured = False
    for line in section.split("\n"):
        s = line.strip()
        if re.match(r"^#\s+", s) and not s.startswith("## "):  # 一级标题 → 封面
            return "cover"
        if re.match(r"^!\[.*\]\(.+\)", s):  # 图片引用
            has_image = True
        # 结构化内容：二级标题、列表、表格
        if (s.startswith("## ") or re.match(r"^(\d+)\.\s+", s) or
                re.match(r"^[-*]\s+", s) or (s.startswith("|") and s.endswith("|"))):
            has_structured = True

    if has_image and has_structured:
        return "text_image"
    if has_image:
        return "image"
    return "content"


def parse_cover(section: str) -> tuple[str, str]:
    """解析封面页，返回 (title, speaker)。"""
    title = ""
    speaker = ""
    for line in section.split("\n"):
        s = line.strip()
        if s.startswith("# ") and not s.startswith("## "):
            title = s[2:].strip()
        elif s.startswith("## "):
            speaker = s[3:].strip()
    return title, speaker


def parse_content(section: str) -> list[tuple]:
    """解析内容页，返回 [(type, text/headers, ...), ...]。

    type: h2 | ordered | unordered | plain | table
    表格项格式: ("table", headers_list, rows_list)
    """
    items: list[tuple] = []
    lines = section.split("\n")
    i = 0

    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue

        # ── 表格检测：以 | 开头和结尾的连续行 ──
        if s.startswith("|") and s.endswith("|"):
            table_lines = [s]
            j = i + 1
            while j < len(lines):
                nxt = lines[j].strip()
                if nxt.startswith("|") and nxt.endswith("|"):
                    table_lines.append(nxt)
                    j += 1
                else:
                    break

            # 至少需要 header + separator（分隔线）两行
            if len(table_lines) >= 2:
                # 解析表头（跳过首尾空元素）
                headers = [c.strip() for c in table_lines[0].split("|")[1:-1]]
                # 跳过分隔线（index 1），解析数据行
                rows = []
                for tl in table_lines[2:]:
                    row = [c.strip() for c in tl.split("|")[1:-1]]
                    if any(row):  # 跳过全空行
                        rows.append(row)
                items.append(("table", headers, rows))
                i = j
                continue

        # ── 非表格内容 ──
        # 跳过封面专属的一级标题
        if re.match(r"^#\s+", s) and not re.match(r"^##\s+", s):
            i += 1
            continue

        if s.startswith("## "):
            items.append(("h2", s[3:].strip()))
            i += 1
            continue

        m = re.match(r"^(\d+)\.\s+(.+)$", s)
        if m:
            items.append(("ordered", m.group(2).strip(), int(m.group(1))))
            i += 1
            continue

        m = re.match(r"^[-*]\s+(.+)$", s)
        if m:
            items.append(("unordered", m.group(1).strip()))
            i += 1
            continue

        items.append(("plain", s))
        i += 1

    return items


def parse_image_page(section: str) -> tuple[str, str]:
    """解析图片展示页，返回 (caption, image_path)。"""
    caption = ""
    image_path = ""
    for line in section.split("\n"):
        s = line.strip()
        if not s:
            continue
        m = re.match(r"^!\[.*\]\((.+)\)$", s)
        if m:
            image_path = m.group(1).strip()
        elif not s.startswith("#"):
            caption = s
    return caption, image_path


def parse_text_image(section: str) -> tuple[list[tuple], str]:
    """解析图配文页，返回 (items, image_path)。
    文字部分解析逻辑同 content 页，图片引用行单独提取。
    """
    items: list[tuple] = []
    image_path = ""
    lines = section.split("\n")
    i = 0

    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue

        # 图片引用行 → 提取路径，不加入 items
        m = re.match(r"^!\[.*\]\((.+)\)$", s)
        if m:
            image_path = m.group(1).strip()
            i += 1
            continue

        # ── 表格检测：以 | 开头和结尾的连续行 ──
        if s.startswith("|") and s.endswith("|"):
            table_lines = [s]
            j = i + 1
            while j < len(lines):
                nxt = lines[j].strip()
                if nxt.startswith("|") and nxt.endswith("|"):
                    table_lines.append(nxt)
                    j += 1
                else:
                    break

            if len(table_lines) >= 2:
                headers = [c.strip() for c in table_lines[0].split("|")[1:-1]]
                rows = []
                for tl in table_lines[2:]:
                    row = [c.strip() for c in tl.split("|")[1:-1]]
                    if any(row):
                        rows.append(row)
                items.append(("table", headers, rows))
                i = j
                continue

        # 跳过封面专属的一级标题
        if re.match(r"^#\s+", s) and not re.match(r"^##\s+", s):
            i += 1
            continue

        if s.startswith("## "):
            items.append(("h2", s[3:].strip()))
            i += 1
            continue

        m2 = re.match(r"^(\d+)\.\s+(.+)$", s)
        if m2:
            items.append(("ordered", m2.group(2).strip(), int(m2.group(1))))
            i += 1
            continue

        m2 = re.match(r"^[-*]\s+(.+)$", s)
        if m2:
            items.append(("unordered", m2.group(1).strip()))
            i += 1
            continue

        items.append(("plain", s))
        i += 1

    return items, image_path


# ── 幻灯片创建 ──────────────────────────────────────────────

def _new_blank_slide(prs: Presentation):
    """创建一个空白版式的幻灯片。"""
    return prs.slides.add_slide(prs.slide_layouts[6])  # 6 = blank


def create_cover_slide(prs: Presentation, title: str, speaker: str):
    """创建封面幻灯片。"""
    slide = _new_blank_slide(prs)
    _set_background(slide, BG_COVER)

    if title:
        tb = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.333), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.CENTER
        _set_font(p, FONT_CN_H1, Pt(66), bold=True)

    if speaker:
        tb = slide.shapes.add_textbox(Inches(1.5), Inches(4.8), Inches(10.333), Inches(1))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = speaker
        p.alignment = PP_ALIGN.CENTER
        _set_font(p, FONT_CN_H2, Pt(44))


def create_content_slide(prs: Presentation, items: list[tuple]):
    """创建内容幻灯片，支持文本和表格。"""
    slide = _new_blank_slide(prs)
    _set_background(slide, BG_CONTENT)

    if not items:
        return

    # 分离文本项和表格项
    text_items = [it for it in items if it[0] != "table"]
    table_items = [it for it in items if it[0] == "table"]

    # ── 渲染文本 ──
    if text_items:
        # 有表格时文本框高度缩小，为表格留空间
        text_height = Inches(2.5) if table_items else Inches(5.8)
        tb = slide.shapes.add_textbox(Inches(1.5), Inches(0.8), Inches(10.333), text_height)
        tf = tb.text_frame
        tf.word_wrap = True

        first = True
        for item in text_items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            t = item[0]

            if t == "h2":
                p.text = item[1]
                _set_font(p, FONT_CN_H2, Pt(44), bold=True)
                p.space_after = Pt(24)
            elif t == "ordered":
                p.text = f"{item[2]}. {item[1]}"
                _set_font(p, FONT_CN_BODY, Pt(32))
                p.space_after = Pt(12)
            elif t == "unordered":
                p.text = f"• {item[1]}"
                _set_font(p, FONT_CN_BODY, Pt(32))
                p.space_after = Pt(12)
            else:  # plain
                p.text = item[1]
                _set_font(p, FONT_CN_BODY, Pt(32))
                p.space_after = Pt(12)

    # ── 渲染表格 ──
    if table_items:
        table_top = Inches(3.5) if text_items else Inches(0.8)
        for _, headers, rows in table_items:
            num_cols = len(headers)
            num_rows = len(rows) + 1  # +1 for header

            col_width = Inches(10.333 / max(num_cols, 1))
            row_height = Inches(0.5)

            table_shape = slide.shapes.add_table(
                num_rows, num_cols,
                Inches(1.5), table_top,
                col_width * num_cols, row_height * num_rows
            )
            tbl = table_shape.table

            # 设置列宽
            for ci in range(num_cols):
                tbl.columns[ci].width = col_width

            # 表头行
            for ci, header_text in enumerate(headers):
                cell = tbl.cell(0, ci)
                cell.text = header_text
                _set_cell_font(cell, FONT_CN_H2, Pt(32), bold=True)

            # 数据行
            for ri, row in enumerate(rows):
                for ci, cell_text in enumerate(row):
                    if ci < num_cols:
                        cell = tbl.cell(ri + 1, ci)
                        cell.text = cell_text
                        _set_cell_font(cell, FONT_CN_BODY, Pt(32))

            # 应用三段式样式（无填充、仅三条横线）
            _style_table_three_line(tbl, num_rows, num_cols)

            # 后续表格依次向下排列
            table_top += row_height * num_rows + Inches(0.3)


def create_image_slide(prs: Presentation, caption: str, img_path: str):
    """创建图片展示幻灯片。"""
    slide = _new_blank_slide(prs)
    _set_background(slide, BG_CONTENT)

    if caption:
        tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(11.333), Inches(1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = caption
        p.alignment = PP_ALIGN.CENTER
        _set_font(p, FONT_CN_H2, Pt(44), bold=True)

    if img_path and os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(2.5), Inches(1.8), Inches(8), Inches(4.8))
    elif img_path:
        print(f"  ⚠ 图片不存在: {img_path}")


def create_text_image_slide(prs: Presentation, items: list[tuple], img_path: str):
    """创建图配文幻灯片：左侧 2/3 文字，右侧 1/3 图片。
    文字部分支持的数据类型和样式与原"内容页"一致。
    """
    slide = _new_blank_slide(prs)
    _set_background(slide, BG_CONTENT)

    # ── 左侧 2/3：文字区域 ──
    text_left = Inches(0.8)
    text_width = Inches(8.0)

    text_items = [it for it in items if it[0] != "table"]
    table_items = [it for it in items if it[0] == "table"]

    if text_items:
        text_height = Inches(2.5) if table_items else Inches(5.8)
        tb = slide.shapes.add_textbox(text_left, Inches(0.8), text_width, text_height)
        tf = tb.text_frame
        tf.word_wrap = True

        first = True
        for item in text_items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            t = item[0]

            if t == "h2":
                p.text = item[1]
                _set_font(p, FONT_CN_H2, Pt(44), bold=True)
                p.space_after = Pt(24)
            elif t == "ordered":
                p.text = f"{item[2]}. {item[1]}"
                _set_font(p, FONT_CN_BODY, Pt(32))
                p.space_after = Pt(12)
            elif t == "unordered":
                p.text = f"• {item[1]}"
                _set_font(p, FONT_CN_BODY, Pt(32))
                p.space_after = Pt(12)
            else:  # plain
                p.text = item[1]
                _set_font(p, FONT_CN_BODY, Pt(32))
                p.space_after = Pt(12)

    if table_items:
        table_top = Inches(3.5) if text_items else Inches(0.8)
        for _, headers, rows in table_items:
            num_cols = len(headers)
            num_rows = len(rows) + 1

            col_width = Inches(text_width.inches / max(num_cols, 1))
            row_height = Inches(0.5)

            table_shape = slide.shapes.add_table(
                num_rows, num_cols,
                text_left, table_top,
                col_width * num_cols, row_height * num_rows
            )
            tbl = table_shape.table

            for ci in range(num_cols):
                tbl.columns[ci].width = col_width

            for ci, header_text in enumerate(headers):
                cell = tbl.cell(0, ci)
                cell.text = header_text
                _set_cell_font(cell, FONT_CN_H2, Pt(32), bold=True)

            for ri, row in enumerate(rows):
                for ci, cell_text in enumerate(row):
                    if ci < num_cols:
                        cell = tbl.cell(ri + 1, ci)
                        cell.text = cell_text
                        _set_cell_font(cell, FONT_CN_BODY, Pt(32))

            _style_table_three_line(tbl, num_rows, num_cols)
            table_top += row_height * num_rows + Inches(0.3)

    # ── 右侧 1/3：图片区域 ──
    if img_path and os.path.exists(img_path):
        img_left = Inches(9.2)
        slide.shapes.add_picture(img_path, img_left, Inches(1.5), Inches(3.6), Inches(4.5))
    elif img_path:
        print(f"  ⚠ 图片不存在: {img_path}")


# ── 主流程 ──────────────────────────────────────────────────

def create_end_slide(prs: Presentation):
    """创建结束页：「谢谢」。"""
    slide = _new_blank_slide(prs)
    _set_background(slide, BG_COVER)

    tb = slide.shapes.add_textbox(Inches(2), Inches(2.5), Inches(9.333), Inches(2.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "谢谢"
    p.alignment = PP_ALIGN.CENTER
    _set_font(p, FONT_CN_H1, Pt(66), bold=True)


def build(md_path: str = INPUT_FILE, output_path: str = OUTPUT_FILE):
    """读取 Markdown 文件，生成 .pptx。"""
    if not os.path.exists(md_path):
        print(f"错误：找不到输入文件 {md_path}")
        sys.exit(1)

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    sections = split_sections(md_text)
    if not sections:
        print("错误：Markdown 中没有有效内容")
        sys.exit(1)

    prs = Presentation()
    prs.slide_width  = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    for i, sec in enumerate(sections):
        pt = detect_page_type(sec)
        print(f"  第 {i + 1} 页 → {pt}")

        if pt == "cover":
            title, speaker = parse_cover(sec)
            create_cover_slide(prs, title, speaker)
        elif pt == "image":
            caption, img_path = parse_image_page(sec)
            create_image_slide(prs, caption, img_path)
        elif pt == "text_image":
            items, img_path = parse_text_image(sec)
            create_text_image_slide(prs, items, img_path)
        else:
            items = parse_content(sec)
            create_content_slide(prs, items)

    # 自动添加结束页
    create_end_slide(prs)
    print(f"  第 {len(sections) + 1} 页 → end")

    prs.save(output_path)
    print(f"\n已生成 {output_path}（共 {len(sections)} 页）")


if __name__ == "__main__":
    build()
