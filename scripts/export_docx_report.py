#!/usr/bin/env python3
"""
Commercial Film Review Word (.docx) Report Exporter

Turns a structured review JSON (see examples/sample_report.json) plus optional
quantitative visual assets (color_barcode.png, contact_sheet.jpg) into a
publication-ready A4 Word (.docx) review report with:
  - Navy commercial visual hierarchy (Microsoft YaHei + Segoe UI)
  - Callout diagnosis card
  - Quantitative metrics table with embedded color barcode & contact sheet
  - Six-dimension breakdown bullets
  - Color-coded P0 / P1 / P2 priority timecode checklist (repeating headers & cantSplit)
  - Series notes & honest limitations section
"""

import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor


def set_cell_shading(cell, fill_hex: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color="D9E1EC"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="EDF1F7"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def add_bottom_border(paragraph, color="1B365D", sz="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def set_run_font(run, size=10.2, bold=False, color=RGBColor(0x22, 0x2B, 0x38), font_name="微软雅黑"):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Segoe UI")
    rFonts.set(qn("w:hAnsi"), "Segoe UI")
    rFonts.set(qn("w:eastAsia"), font_name)


def add_callout_box(doc, title_text: str, body_lines: list[dict], fill="F4F7FB", border_color="1B365D"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Cm(16.6)
    set_cell_shading(cell, fill)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=180)
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(
        parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="none"/>'
            f'  <w:bottom w:val="none"/>'
            f'  <w:right w:val="none"/>'
            f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
            f'</w:tcBorders>'
        )
    )
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(4)
    p0.paragraph_format.line_spacing = 1.35
    r0 = p0.add_run(title_text)
    set_run_font(r0, size=10.5, bold=True, color=RGBColor(0x1B, 0x36, 0x5D))

    for item in body_lines:
        lead = item.get("lead", "")
        text = item.get("text", "")
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.4
        if lead:
            rl = p.add_run(lead)
            set_run_font(rl, size=9.8, bold=True, color=RGBColor(0x22, 0x2B, 0x38))
        rt = p.add_run(text)
        set_run_font(rt, size=9.8, bold=False, color=RGBColor(0x33, 0x3D, 0x4B))


def add_h1(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_font(r, size=13.5, bold=True, color=RGBColor(0x1B, 0x36, 0x5D))
    add_bottom_border(p, color="1B365D", sz="8")
    return p


def add_h2(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_run_font(r, size=11.2, bold=True, color=RGBColor(0x0F, 0x4C, 0x81))
    return p


def add_bullet(doc, lead: str, text: str, level: int = 0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.55 + level * 0.55)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.paragraph_format.space_before = Pt(1.5)
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.line_spacing = 1.42
    bullet_char = "▪  " if level == 0 else "•  "
    rb = p.add_run(bullet_char)
    set_run_font(
        rb,
        size=9.5,
        bold=True,
        color=RGBColor(0xC8, 0x23, 0x33) if level == 0 else RGBColor(0x5A, 0x6B, 0x82),
    )
    if lead:
        rl = p.add_run(lead)
        set_run_font(rl, size=10.0, bold=True, color=RGBColor(0x1E, 0x26, 0x31))
    rt = p.add_run(text)
    set_run_font(rt, size=10.0, bold=False, color=RGBColor(0x33, 0x3D, 0x4B))
    return p


def add_styled_table(
    doc,
    headers: list[str],
    rows: list[list[str]],
    col_widths_cm: list[float],
    priority_col: int | None = None,
):
    tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl)

    hdr_row = tbl.rows[0]
    hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for j, h in enumerate(headers):
        cell = hdr_row.cells[j]
        cell.width = Cm(col_widths_cm[j])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(cell, "1B365D")
        set_cell_margins(cell, top=95, bottom=95, left=110, right=110)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER
            if (priority_col == j or (j == 0 and len(headers) > 3))
            else WD_ALIGN_PARAGRAPH.LEFT
        )
        r = p.add_run(h)
        set_run_font(r, size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    for i, row_data in enumerate(rows):
        row = tbl.rows[i + 1]
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        bg = "FFFFFF" if i % 2 == 0 else "F7F9FC"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = Cm(col_widths_cm[j])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell_bg = bg
            txt_color = RGBColor(0x2B, 0x34, 0x42)
            is_bold = False
            align = WD_ALIGN_PARAGRAPH.LEFT

            if priority_col is not None and j == priority_col:
                align = WD_ALIGN_PARAGRAPH.CENTER
                is_bold = True
                if val == "P0":
                    cell_bg = "FDECEC"
                    txt_color = RGBColor(0xB9, 0x1C, 0x1C)
                elif val == "P1":
                    cell_bg = "FEF5E7"
                    txt_color = RGBColor(0xB4, 0x53, 0x09)
                elif val == "P2":
                    cell_bg = "ECFDF3"
                    txt_color = RGBColor(0x15, 0x80, 0x3D)
            elif priority_col is not None and j == 2:
                align = WD_ALIGN_PARAGRAPH.CENTER

            set_cell_shading(cell, cell_bg)
            set_cell_margins(cell, top=80, bottom=80, left=105, right=105)
            p = cell.paragraphs[0]
            p.alignment = align
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.32
            r = p.add_run(str(val))
            set_run_font(r, size=9.0 if len(headers) >= 5 else 9.5, bold=is_bold, color=txt_color)

    return tbl


def build_report(spec: dict, out_path: Path, assets_dir: Path | None = None):
    doc = Document()

    header_text = spec.get("header_text", spec.get("title", "专业审片与视听优化报告"))
    footer_text = spec.get("footer_text", "Commercial Film Review · 品牌形象片 / 产品发布片专项审片")

    for sec in doc.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.0)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin = Cm(2.2)
        sec.right_margin = Cm(2.2)

        hdr = sec.header.paragraphs[0]
        hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rh = hdr.add_run(header_text)
        set_run_font(rh, size=8.2, color=RGBColor(0x88, 0x95, 0xA7))

        ftr = sec.footer.paragraphs[0]
        ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rf = ftr.add_run(footer_text)
        set_run_font(rf, size=8.2, color=RGBColor(0x88, 0x95, 0xA7))

    # Cover header block
    p_kicker = doc.add_paragraph()
    p_kicker.paragraph_format.space_before = Pt(0)
    p_kicker.paragraph_format.space_after = Pt(2)
    rk = p_kicker.add_run(spec.get("kicker", "COMMERCIAL FILM REVIEW  |  品牌形象片与产品发布片专业审片"))
    set_run_font(rk, size=8.8, bold=True, color=RGBColor(0xC8, 0x23, 0x33))

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    rt = p_title.add_run(spec.get("title", "专业审片报告"))
    set_run_font(rt, size=17.0, bold=True, color=RGBColor(0x11, 0x24, 0x42))

    if spec.get("meta_line"):
        p_meta = doc.add_paragraph()
        p_meta.paragraph_format.space_before = Pt(0)
        p_meta.paragraph_format.space_after = Pt(10)
        rm = p_meta.add_run(spec["meta_line"])
        set_run_font(rm, size=9.0, color=RGBColor(0x5C, 0x6B, 0x7E))

    # Section 1: 总体诊断
    diag = spec.get("diagnosis", {})
    add_h1(doc, "一、总体诊断")
    add_callout_box(
        doc,
        diag.get("headline", "综合评级"),
        diag.get("bullets", []),
    )

    # Section 2: 量化体检表 + 图表
    add_h1(doc, "二、量化体检表")
    metrics_rows = spec.get("metrics_table", [])
    if metrics_rows:
        add_styled_table(
            doc,
            ["指标", "实测值", "专业参考 / 诊断结论"],
            metrics_rows,
            [4.4, 5.2, 7.0],
        )

    figures = spec.get("figures", {})
    barcode_path = Path(figures["color_barcode"]) if figures.get("color_barcode") else (assets_dir / "color_barcode.png" if assets_dir else None)
    sheet_path = Path(figures["contact_sheet"]) if figures.get("contact_sheet") else (assets_dir / "contact_sheet.jpg" if assets_dir else None)

    if barcode_path and barcode_path.exists():
        p_cap1 = doc.add_paragraph()
        p_cap1.paragraph_format.space_before = Pt(8)
        p_cap1.paragraph_format.space_after = Pt(3)
        p_cap1.paragraph_format.keep_with_next = True
        rc1 = p_cap1.add_run(figures.get("color_barcode_caption", "图 1  全片色彩与影调色带（Color Barcode）"))
        set_run_font(rc1, size=9.0, bold=True, color=RGBColor(0x4A, 0x55, 0x68))
        p_img1 = doc.add_paragraph()
        p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img1.paragraph_format.space_after = Pt(4)
        p_img1.add_run().add_picture(str(barcode_path), width=Cm(16.4))

    if sheet_path and sheet_path.exists():
        p_cap2 = doc.add_paragraph()
        p_cap2.paragraph_format.space_before = Pt(6)
        p_cap2.paragraph_format.space_after = Pt(3)
        p_cap2.paragraph_format.keep_with_next = True
        rc2 = p_cap2.add_run(figures.get("contact_sheet_caption", "图 2  全片时间码与主色值故事板（Contact Sheet）"))
        set_run_font(rc2, size=9.0, bold=True, color=RGBColor(0x4A, 0x55, 0x68))
        p_img2 = doc.add_paragraph()
        p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img2.paragraph_format.space_after = Pt(6)
        p_img2.add_run().add_picture(str(sheet_path), width=Cm(16.4))

    # Section 3: 六维深度拆解
    add_h1(doc, "三、六维深度拆解")
    for dim in spec.get("dimensions", []):
        add_h2(doc, dim.get("heading", ""))
        for b in dim.get("bullets", []):
            add_bullet(doc, b.get("lead", ""), b.get("text", ""), level=b.get("level", 0))

    # Section 4: 按时间码排列的改片清单
    add_h1(doc, "四、按时间码排列的改片清单（可直接进达芬奇 / PR / AE 执行）")
    checklist = spec.get("checklist", [])
    if checklist:
        rows = [
            [
                item.get("priority", "P1"),
                item.get("timecode", ""),
                item.get("dimension", ""),
                item.get("issue", ""),
                item.get("action", ""),
            ]
            if isinstance(item, dict)
            else item
            for item in checklist
        ]
        add_styled_table(
            doc,
            ["优先级", "时间码", "维度", "具体问题", "可执行改法"],
            rows,
            [1.4, 2.7, 1.8, 4.3, 6.4],
            priority_col=0,
        )

    # Section 5: 补充建议与局限说明
    notes = spec.get("notes", [])
    if notes:
        add_h1(doc, spec.get("notes_heading", "五、系列共性提示与局限说明"))
        for n in notes:
            add_bullet(doc, n.get("lead", ""), n.get("text", ""), level=n.get("level", 0))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"SAVED: {out_path} ({out_path.stat().st_size} bytes)")


def main():
    parser = argparse.ArgumentParser(description="Export a structured commercial film review JSON to a formatted Word (.docx) report.")
    parser.add_argument("report_json", help="Path to structured review JSON file")
    parser.add_argument("--out", required=True, help="Output .docx file path")
    parser.add_argument("--assets", help="Optional directory containing color_barcode.png and contact_sheet.jpg")
    args = parser.parse_args()

    json_path = Path(args.report_json).resolve()
    if not json_path.exists():
        print(f"Error: report JSON not found: {json_path}", file=sys.stderr)
        sys.exit(1)

    spec = json.loads(json_path.read_text(encoding="utf-8"))
    assets_dir = Path(args.assets).resolve() if args.assets else None
    build_report(spec, Path(args.out).resolve(), assets_dir)


if __name__ == "__main__":
    main()
