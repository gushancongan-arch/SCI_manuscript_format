#!/usr/bin/env python3
"""Audit structural invariants in the reviewed SCI manuscript template."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = ROOT / "assets" / "SCI_manuscript_template.docx"
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


class Audit:
    def __init__(self) -> None:
        self.results: list[dict[str, object]] = []

    def check(self, name: str, condition: bool, detail: str = "") -> None:
        self.results.append({"name": name, "status": "PASS" if condition else "FAIL", "detail": detail})

    @property
    def failures(self) -> int:
        return sum(item["status"] == "FAIL" for item in self.results)


def close(value, expected, tolerance=0.05) -> bool:
    return value is not None and abs(value - expected) <= tolerance


def point_value(length) -> float | None:
    return None if length is None else length.pt


def parse_part(archive: zipfile.ZipFile, name: str):
    return etree.fromstring(archive.read(name))


def style_is_exposed(style) -> bool:
    element = style._element
    return element.find(qn("w:qFormat")) is not None and element.find(qn("w:semiHidden")) is None


def audit_styles(document: Document, audit: Audit) -> None:
    required = [
        "SCI Body",
        "SCI Abstract Body",
        "SCI Statement Body",
        "图片",
        "图片标题",
        "表格标题",
        "Equation",
        "Reference",
        "Line Number",
        "SCI Three-Line Table",
    ]
    names = {style.name for style in document.styles}
    audit.check("required styles exist", all(name in names for name in required), ", ".join(required))

    for name in ("SCI Body", "SCI Abstract Body", "SCI Statement Body"):
        style = document.styles[name]
        fmt = style.paragraph_format
        valid = (
            close(point_value(style.font.size), 14)
            and fmt.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
            and close(point_value(fmt.first_line_indent), 28)
            and close(point_value(fmt.space_before), 6)
            and close(point_value(fmt.space_after), 0)
            and fmt.line_spacing == 1.5
            and style_is_exposed(style)
        )
        audit.check(f"{name} geometry", valid, "14 pt, justified, 1.5 spacing, 28 pt first-line indent")

    figure_caption = document.styles["图片标题"]
    audit.check(
        "figure caption style",
        close(point_value(figure_caption.font.size), 12)
        and figure_caption.font.italic is True
        and figure_caption.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
        and close(point_value(figure_caption.paragraph_format.first_line_indent), 0),
        "12 pt italic justified",
    )

    table_caption = document.styles["表格标题"]
    audit.check(
        "table caption style",
        close(point_value(table_caption.font.size), 12)
        and table_caption.font.italic is True
        and table_caption.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.LEFT,
        "12 pt italic left aligned",
    )

    line_number = document.styles["Line Number"]
    audit.check(
        "line number style",
        line_number.type == WD_STYLE_TYPE.CHARACTER and close(point_value(line_number.font.size), 11),
        "character style, Times New Roman 11 pt",
    )

    reference = document.styles["Reference"]
    audit.check(
        "reference style",
        close(point_value(reference.font.size), 11)
        and close(point_value(reference.paragraph_format.left_indent), 36)
        and close(point_value(reference.paragraph_format.first_line_indent), -36),
        "11 pt with 1.27 cm hanging indent",
    )


def audit_page(document: Document, document_xml, footer_xml, settings_xml, audit: Audit) -> None:
    section = document.sections[0]
    audit.check(
        "A4 page geometry",
        close(section.page_width.cm, 21.0, 0.02)
        and close(section.page_height.cm, 29.7, 0.02)
        and all(
            close(value.cm, 2.54, 0.02)
            for value in (section.top_margin, section.bottom_margin, section.left_margin, section.right_margin)
        ),
        "21.0 x 29.7 cm, 2.54 cm margins",
    )

    sections = document_xml.xpath("//w:sectPr", namespaces=NS)
    line_number_nodes = document_xml.xpath("//w:sectPr/w:lnNumType", namespaces=NS)
    valid_line_numbers = len(line_number_nodes) == len(sections) and all(
        node.get(qn("w:countBy")) == "1"
        and node.get(qn("w:start")) == "0"
        and node.get(qn("w:restart")) == "continuous"
        for node in line_number_nodes
    )
    audit.check("continuous line numbering", valid_line_numbers, f"sections={len(sections)}, visible start=1")

    field_instructions = " ".join(footer_xml.xpath("//w:instrText/text()", namespaces=NS))
    audit.check("dynamic PAGE field", "PAGE" in field_instructions, field_instructions.strip())
    update_fields = settings_xml.xpath("//w:updateFields", namespaces=NS)
    audit.check(
        "update fields on open",
        bool(update_fields) and update_fields[0].get(qn("w:val")) in {"true", "1"},
        "w:updateFields=true",
    )

    page_breaks = document_xml.xpath("//w:br[@w:type='page']", namespaces=NS)
    audit.check("front-matter page breaks", len(page_breaks) >= 2, f"page breaks={len(page_breaks)}")


def paragraph_style_id(paragraph_element) -> str | None:
    values = paragraph_element.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return values[0] if values else None


def audit_roles(document: Document, document_xml, audit: Audit) -> None:
    style_ids = {name: document.styles[name].style_id for name in (
        "SCI Body", "SCI Abstract Body", "SCI Statement Body", "图片标题", "表格标题", "Equation", "Reference"
    )}
    used = document_xml.xpath("//w:p/w:pPr/w:pStyle/@w:val", namespaces=NS)
    for name, style_id in style_ids.items():
        audit.check(f"{name} is assigned", style_id in used, style_id)

    body_ids = {
        document.styles["SCI Body"].style_id,
        document.styles["SCI Abstract Body"].style_id,
        document.styles["SCI Statement Body"].style_id,
    }
    body_paragraphs = [p for p in document.paragraphs if p.style.style_id in body_ids]
    audit.check("semantic body paragraphs", len(body_paragraphs) >= 4, f"count={len(body_paragraphs)}")

    body_children = list(document_xml.xpath("/w:document/w:body", namespaces=NS)[0])
    figure_caption_id = document.styles["图片标题"].style_id
    table_caption_id = document.styles["表格标题"].style_id
    figure_pair = False
    table_pair = False
    for index, child in enumerate(body_children):
        if child.tag == qn("w:p") and child.xpath(".//w:drawing", namespaces=NS):
            if index + 1 < len(body_children):
                figure_pair = paragraph_style_id(body_children[index + 1]) == figure_caption_id
        if child.tag == qn("w:tbl") and index > 0:
            table_pair = paragraph_style_id(body_children[index - 1]) == table_caption_id
    audit.check("figure caption placement", figure_pair, "caption immediately below inline figure")
    audit.check("table caption placement", table_pair, "caption immediately above table")


def audit_equation(document: Document, document_xml, audit: Audit) -> None:
    equation_style_id = document.styles["Equation"].style_id
    equation_paragraphs = document_xml.xpath(
        f"//w:p[w:pPr/w:pStyle[@w:val='{equation_style_id}']]", namespaces=NS
    )
    native_math = document_xml.xpath("//m:oMath", namespaces=NS)
    fractions = document_xml.xpath("//m:oMath//m:f", namespaces=NS)
    tab_nodes = sum(len(paragraph.xpath("./w:r/w:tab", namespaces=NS)) for paragraph in equation_paragraphs)
    tab_stops = sum(len(paragraph.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS)) for paragraph in equation_paragraphs)
    if equation_paragraphs and tab_stops < 2:
        style = document.styles["Equation"]
        tab_stops = len(style._element.xpath("./w:pPr/w:tabs/w:tab"))
    audit.check("native OMML equation", bool(native_math) and bool(fractions), f"oMath={len(native_math)}, fractions={len(fractions)}")
    audit.check("numbered equation tabs", tab_nodes >= 2 and tab_stops >= 2, f"TAB nodes={tab_nodes}, tab stops={tab_stops}")


def audit_table(document_xml, audit: Audit) -> None:
    tables = document_xml.xpath("//w:tbl", namespaces=NS)
    if not tables:
        audit.check("three-line table exists", False, "no table")
        return
    table = tables[0]
    borders = table.xpath("./w:tblPr/w:tblBorders", namespaces=NS)
    top = table.xpath("./w:tblPr/w:tblBorders/w:top", namespaces=NS)
    bottom = table.xpath("./w:tblPr/w:tblBorders/w:bottom", namespaces=NS)
    inside_h = table.xpath("./w:tblPr/w:tblBorders/w:insideH", namespaces=NS)
    inside_v = table.xpath("./w:tblPr/w:tblBorders/w:insideV", namespaces=NS)
    border_ok = (
        bool(borders)
        and bool(top)
        and bool(bottom)
        and top[0].get(qn("w:sz")) == "8"
        and bottom[0].get(qn("w:sz")) == "8"
        and all(node.get(qn("w:val")) == "nil" for node in inside_h + inside_v)
    )
    audit.check("three-line borders", border_ok, "1.0 pt top/bottom; no internal grid")

    header_cells = table.xpath("./w:tr[1]/w:tc", namespaces=NS)
    header_rule_ok = all(
        cell.xpath("./w:tcPr/w:tcBorders/w:bottom[@w:sz='4']", namespaces=NS) for cell in header_cells
    )
    repeat_header = bool(table.xpath("./w:tr[1]/w:trPr/w:tblHeader", namespaces=NS))
    audit.check("table header rule and repeat", header_rule_ok and repeat_header, "0.5 pt header separator")

    shading = table.xpath(".//w:shd", namespaces=NS)
    audit.check("table has no shading", not shading, f"shading nodes={len(shading)}")

    forbidden = ("firstLineChars", "hangingChars", "leftChars", "rightChars")
    bad_indents = []
    for ind in table.xpath(".//w:pPr/w:ind", namespaces=NS):
        if any(ind.get(qn(f"w:{name}")) is not None for name in forbidden):
            bad_indents.append(ind)
    audit.check("table character indents removed", not bad_indents, f"bad indents={len(bad_indents)}")

    grid_widths = [int(value) for value in table.xpath("./w:tblGrid/w:gridCol/@w:w", namespaces=NS)]
    table_widths = [int(value) for value in table.xpath("./w:tr[1]/w:tc/w:tcPr/w:tcW/@w:w", namespaces=NS)]
    tbl_width = table.xpath("./w:tblPr/w:tblW/@w:w", namespaces=NS)
    geometry_ok = grid_widths == table_widths and bool(tbl_width) and sum(grid_widths) == int(tbl_width[0])
    audit.check("fixed table geometry", geometry_ok, f"grid={grid_widths}, total={tbl_width[0] if tbl_width else 'missing'}")


def run_audit(path: Path) -> Audit:
    audit = Audit()
    audit.check("template exists", path.is_file(), str(path))
    if not path.is_file():
        return audit

    try:
        document = Document(path)
        audit.check("DOCX package opens", True, f"paragraphs={len(document.paragraphs)}, tables={len(document.tables)}")
    except Exception as exc:  # pragma: no cover - fatal package error
        audit.check("DOCX package opens", False, str(exc))
        return audit

    with zipfile.ZipFile(path) as archive:
        document_xml = parse_part(archive, "word/document.xml")
        settings_xml = parse_part(archive, "word/settings.xml")
        footer_names = [name for name in archive.namelist() if name.startswith("word/footer") and name.endswith(".xml")]
        footer_xml = parse_part(archive, footer_names[0]) if footer_names else etree.Element("missing")

    audit_styles(document, audit)
    audit_page(document, document_xml, footer_xml, settings_xml, audit)
    audit_roles(document, document_xml, audit)
    audit_equation(document, document_xml, audit)
    audit_table(document_xml, audit)
    return audit


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", nargs="?", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()
    audit = run_audit(args.template.resolve())
    if args.as_json:
        print(json.dumps({"failures": audit.failures, "results": audit.results}, indent=2, ensure_ascii=False))
    else:
        for item in audit.results:
            detail = f" - {item['detail']}" if item["detail"] else ""
            print(f"{item['status']}: {item['name']}{detail}")
        print(f"SUMMARY: {len(audit.results) - audit.failures} PASS, {audit.failures} FAIL")
    sys.exit(1 if audit.failures else 0)


if __name__ == "__main__":
    main()
