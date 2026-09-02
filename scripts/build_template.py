#!/usr/bin/env python3
"""Build the reviewed SCI manuscript DOCX template."""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image, ImageDraw
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "assets" / "SCI_manuscript_template.docx"
BLACK = RGBColor(0, 0, 0)
BLUE = RGBColor(5, 99, 193)


def set_font_name(font, name: str = "Times New Roman") -> None:
    font.name = name
    rpr = font._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), name)


def set_run_font(run, size: float, *, bold=None, italic=None, color=BLACK) -> None:
    set_font_name(run.font)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def remove_child(parent, tag: str) -> None:
    for child in list(parent):
        if child.tag == qn(tag):
            parent.remove(child)


def expose_style(style) -> None:
    element = style._element
    for tag in ("w:semiHidden", "w:unhideWhenUsed"):
        remove_child(element, tag)
    if element.find(qn("w:qFormat")) is None:
        element.append(OxmlElement("w:qFormat"))


def get_or_add_style(document: Document, name: str, style_type) -> object:
    try:
        style = document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, style_type)
    expose_style(style)
    return style


def configure_paragraph_style(
    document: Document,
    name: str,
    *,
    size: float,
    alignment,
    line_spacing: float,
    before: float,
    after: float,
    first_line: float = 0,
    left: float = 0,
    right: float = 0,
    bold: bool = False,
    italic: bool = False,
) -> object:
    style = get_or_add_style(document, name, WD_STYLE_TYPE.PARAGRAPH)
    set_font_name(style.font)
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    style.font.color.rgb = BLACK
    fmt = style.paragraph_format
    fmt.alignment = alignment
    fmt.line_spacing = line_spacing
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.first_line_indent = Pt(first_line)
    fmt.left_indent = Pt(left)
    fmt.right_indent = Pt(right)
    return style


def configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    set_font_name(normal.font)
    normal.font.size = Pt(14)
    normal.font.color.rgb = BLACK
    normal_fmt = normal.paragraph_format
    normal_fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal_fmt.line_spacing = 1.5
    normal_fmt.space_before = Pt(6)
    normal_fmt.space_after = Pt(0)
    normal_fmt.first_line_indent = Pt(0)
    normal_fmt.left_indent = Pt(0)
    normal_fmt.right_indent = Pt(0)

    for name in ("SCI Body", "SCI Abstract Body", "SCI Statement Body"):
        configure_paragraph_style(
            document,
            name,
            size=14,
            alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            line_spacing=1.5,
            before=6,
            after=0,
            first_line=28,
        )

    for index in (1, 2, 3):
        style = configure_paragraph_style(
            document,
            f"Heading {index}",
            size=14,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
            line_spacing=1.5,
            before=6,
            after=0,
            bold=True,
        )
        style.paragraph_format.keep_with_next = True
        ppr = style._element.get_or_add_pPr()
        outline = ppr.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            ppr.append(outline)
        outline.set(qn("w:val"), str(index - 1))

    configure_paragraph_style(
        document,
        "图片",
        size=14,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.0,
        before=6,
        after=0,
    )
    configure_paragraph_style(
        document,
        "图片标题",
        size=12,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        line_spacing=1.5,
        before=0,
        after=6,
        italic=True,
    )
    configure_paragraph_style(
        document,
        "表格标题",
        size=12,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.5,
        before=6,
        after=0,
        italic=True,
    )
    configure_paragraph_style(
        document,
        "Equation",
        size=14,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.0,
        before=0,
        after=0,
    )
    equation = document.styles["Equation"]
    equation.paragraph_format.tab_stops.clear_all()
    equation.paragraph_format.tab_stops.add_tab_stop(Cm(7.96), WD_TAB_ALIGNMENT.CENTER)
    equation.paragraph_format.tab_stops.add_tab_stop(Cm(15.92), WD_TAB_ALIGNMENT.RIGHT)

    configure_paragraph_style(
        document,
        "Reference",
        size=11,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.0,
        before=0,
        after=3,
        first_line=-36,
        left=36,
    )

    line_number = get_or_add_style(document, "Line Number", WD_STYLE_TYPE.CHARACTER)
    set_font_name(line_number.font)
    line_number.font.size = Pt(11)
    line_number.font.bold = False
    line_number.font.italic = False
    line_number.font.color.rgb = BLACK

    table_style = get_or_add_style(document, "SCI Three-Line Table", WD_STYLE_TYPE.TABLE)
    set_font_name(table_style.font)
    table_style.font.size = Pt(10)
    table_style.font.color.rgb = BLACK


def configure_section(section) -> None:
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)
    section.header_distance = Cm(1.27)
    section.footer_distance = Cm(1.27)

    sect_pr = section._sectPr
    for existing in sect_pr.findall(qn("w:lnNumType")):
        sect_pr.remove(existing)
    line_numbers = OxmlElement("w:lnNumType")
    line_numbers.set(qn("w:countBy"), "1")
    # Word exposes w:start="0" as visible starting line number 1.
    line_numbers.set(qn("w:start"), "0")
    line_numbers.set(qn("w:distance"), "360")
    line_numbers.set(qn("w:restart"), "continuous")
    sect_pr.insert_element_before(line_numbers, "w:pgNumType", "w:cols", "w:docGrid")

    header_paragraph = section.header.paragraphs[0]
    header_paragraph.clear()

    footer_paragraph = section.footer.paragraphs[0]
    footer_paragraph.clear()
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = footer_paragraph.add_run()
    set_run_font(run, 11)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    display = OxmlElement("w:t")
    display.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instruction, separate, display, end):
        run._r.append(element)


def set_update_fields(document: Document) -> None:
    settings = document.settings._element
    for existing in settings.findall(qn("w:updateFields")):
        settings.remove(existing)
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)


def set_zero_indent(paragraph) -> None:
    fmt = paragraph.paragraph_format
    fmt.first_line_indent = Pt(0)
    fmt.left_indent = Pt(0)
    fmt.right_indent = Pt(0)
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.get_or_add_ind()
    for name in ("firstLineChars", "hangingChars", "leftChars", "rightChars", "hanging"):
        ind.attrib.pop(qn(f"w:{name}"), None)
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")


def add_title_block(document: Document) -> None:
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.line_spacing = 1.5
    title.paragraph_format.space_before = Pt(6)
    title.paragraph_format.space_after = Pt(0)
    set_zero_indent(title)
    set_run_font(title.add_run("[Insert manuscript title]"), 14, bold=True)

    authors = document.add_paragraph()
    authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    authors.paragraph_format.line_spacing = 1.5
    authors.paragraph_format.space_before = Pt(6)
    authors.paragraph_format.space_after = Pt(0)
    set_zero_indent(authors)
    set_run_font(authors.add_run("[Author 1]"), 12)
    marker = authors.add_run("a")
    set_run_font(marker, 12)
    marker.font.superscript = True
    set_run_font(authors.add_run(", [Author 2]"), 12)
    marker = authors.add_run("a,b")
    set_run_font(marker, 12)
    marker.font.superscript = True
    set_run_font(authors.add_run(", [Author 3]"), 12)
    marker = authors.add_run("c,*")
    set_run_font(marker, 12)
    marker.font.superscript = True

    for marker_text, affiliation in (
        ("a", "[Affiliation 1, Department, Institution, City, Country]"),
        ("b", "[Affiliation 2, Department, Institution, City, Country]"),
        ("c", "[Affiliation 3, Department, Institution, City, Country]"),
    ):
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(0)
        set_zero_indent(paragraph)
        marker = paragraph.add_run(marker_text)
        set_run_font(marker, 12)
        marker.font.superscript = True
        set_run_font(paragraph.add_run(f" {affiliation}"), 12)

    corresponding = document.add_paragraph()
    corresponding.alignment = WD_ALIGN_PARAGRAPH.LEFT
    corresponding.paragraph_format.line_spacing = 1.5
    corresponding.paragraph_format.space_before = Pt(6)
    corresponding.paragraph_format.space_after = Pt(0)
    set_zero_indent(corresponding)
    marker = corresponding.add_run("*")
    set_run_font(marker, 12)
    marker.font.superscript = True
    set_run_font(corresponding.add_run(" Corresponding author: "), 12)
    email = corresponding.add_run("name@example.com")
    set_run_font(email, 12, color=BLUE)
    email.font.underline = True
    corresponding.add_run().add_break(WD_BREAK.PAGE)


def add_abstract(document: Document) -> None:
    abstract_heading = document.add_paragraph("Abstract", style="Heading 1")
    abstract_heading.paragraph_format.space_before = Pt(2.5)
    set_zero_indent(abstract_heading)
    document.add_paragraph(
        "[Insert the abstract. This paragraph uses the SCI Abstract Body style and the confirmed four-letter first-line indent.]",
        style="SCI Abstract Body",
    )
    keywords = document.add_paragraph()
    keywords.alignment = WD_ALIGN_PARAGRAPH.LEFT
    keywords.paragraph_format.line_spacing = 1.5
    keywords.paragraph_format.space_before = Pt(6)
    keywords.paragraph_format.space_after = Pt(0)
    set_zero_indent(keywords)
    set_run_font(keywords.add_run("Keywords: "), 14, bold=True)
    set_run_font(keywords.add_run("[keyword 1; keyword 2; keyword 3]"), 14)
    keywords.add_run().add_break(WD_BREAK.PAGE)


def create_placeholder_figure(path: Path) -> None:
    image = Image.new("RGB", (1500, 1000), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((90, 80, 1410, 900), outline=(40, 40, 40), width=4)
    draw.line((180, 790, 1320, 790), fill=(40, 40, 40), width=4)
    draw.line((180, 790, 180, 170), fill=(40, 40, 40), width=4)
    points = [(180, 720), (360, 650), (540, 610), (720, 500), (900, 390), (1080, 310), (1320, 220)]
    draw.line(points, fill=(55, 103, 168), width=12)
    for x, y in points:
        draw.ellipse((x - 10, y - 10, x + 10, y + 10), fill=(208, 87, 79))
    draw.text((520, 95), "REPLACE WITH FINAL FIGURE", fill=(60, 60, 60))
    draw.text((630, 915), "Independent variable", fill=(60, 60, 60))
    draw.text((100, 120), "Response", fill=(60, 60, 60))
    image.save(path, dpi=(300, 300))


def set_cell_width(cell, width_twips: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(width_twips))
    tc_w.set(qn("w:type"), "dxa")


def set_cell_margins(cell, top=60, left=80, bottom=60, right=80) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for name, value in (("top", top), ("start", left), ("bottom", bottom), ("end", right)):
        node = margins.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_bottom_border(cell, size_eighth_points: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    bottom = borders.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        borders.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size_eighth_points))
    bottom.set(qn("w:color"), "000000")


def configure_table(table, widths: list[int]) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    total = sum(widths)
    tbl_pr = table._tbl.tblPr

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "0")
    tbl_ind.set(qn("w:type"), "dxa")

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for child in list(borders):
        borders.remove(child)
    for edge in ("top", "bottom"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "8")
        node.set(qn("w:color"), "000000")
        borders.append(node)
    for edge in ("start", "end", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "nil")
        borders.append(node)

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)

    for row_index, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        tr_pr.append(cant_split)
        height = OxmlElement("w:trHeight")
        height.set(qn("w:val"), "397")
        height.set(qn("w:hRule"), "atLeast")
        tr_pr.append(height)
        if row_index == 0:
            tr_pr.append(OxmlElement("w:tblHeader"))

        for column_index, cell in enumerate(row.cells):
            set_cell_width(cell, widths[column_index])
            set_cell_margins(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if row_index == 0:
                set_cell_bottom_border(cell, 4)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if row_index > 0 and column_index == 1 else WD_ALIGN_PARAGRAPH.CENTER
                )
                set_zero_indent(paragraph)
                for run in paragraph.runs:
                    set_run_font(run, 10, bold=(row_index == 0))


def omath_run(text: str, *, italic: bool = True):
    run = OxmlElement("m:r")
    if italic:
        rpr = OxmlElement("m:rPr")
        style = OxmlElement("m:sty")
        style.set(qn("m:val"), "i")
        rpr.append(style)
        run.append(rpr)
    value = OxmlElement("m:t")
    value.text = text
    run.append(value)
    return run


def build_native_equation():
    equation = OxmlElement("m:oMath")
    equation.append(omath_run("v"))
    equation.append(omath_run("=", italic=False))
    fraction = OxmlElement("m:f")
    numerator = OxmlElement("m:num")
    numerator.append(omath_run("Δd"))
    denominator = OxmlElement("m:den")
    denominator.append(omath_run("Δt"))
    fraction.append(numerator)
    fraction.append(denominator)
    equation.append(fraction)
    return equation


def add_equation(document: Document) -> None:
    paragraph = document.add_paragraph(style="Equation")
    first_tab = paragraph.add_run("\t")
    set_run_font(first_tab, 14)
    paragraph._p.append(build_native_equation())
    second_tab = paragraph.add_run("\t")
    set_run_font(second_tab, 14)
    set_run_font(paragraph.add_run("(1)"), 14)


def add_manuscript_elements(document: Document, figure_path: Path) -> None:
    document.add_paragraph("1. Introduction", style="Heading 1")
    paragraph = document.add_paragraph(style="SCI Body")
    set_run_font(
        paragraph.add_run(
            "[Insert manuscript text here. Preserve the source citation system, whether it uses author-year citations such as Smith (2024) or numeric citations such as "
        ),
        14,
    )
    citation = paragraph.add_run("1")
    set_run_font(citation, 14)
    citation.font.superscript = True
    set_run_font(paragraph.add_run(".]"), 14)

    document.add_paragraph("1.1 Study context", style="Heading 2")
    document.add_paragraph(
        "[Use SCI Body for ordinary narrative paragraphs. Replace all bracketed guidance before submission.]",
        style="SCI Body",
    )

    document.add_paragraph("2. Figures", style="Heading 1")
    figure_paragraph = document.add_paragraph(style="图片")
    figure_paragraph.add_run().add_picture(str(figure_path), width=Cm(15.921))
    document.add_paragraph(
        "Figure 1. Replace this placeholder with the final inline figure and edit the caption.",
        style="图片标题",
    )

    document.add_paragraph("3. Tables", style="Heading 1")
    document.add_paragraph(
        "Table 1. Illustrative physical-model test conditions and observed responses.",
        style="表格标题",
    )
    table = document.add_table(rows=4, cols=4, style="SCI Three-Line Table")
    values = [
        ["Model ID", "Failure mode", "Peak displacement (mm)", "Failure time (min)"],
        ["PM-01", "No failure", "8.4", ">180"],
        ["PM-02", "Progressive sliding", "18.7", "146"],
        ["PM-03", "Toe failure", "36.2", "93"],
    ]
    for row, row_values in zip(table.rows, values):
        for cell, value in zip(row.cells, row_values):
            cell.text = value
    configure_table(table, [1500, 3000, 2300, 2226])

    document.add_paragraph("4. Equations", style="Heading 1")
    document.add_paragraph(
        "[Use native editable Word equations for mathematical expressions that require equation typesetting.]",
        style="SCI Body",
    )
    add_equation(document)

    document.add_paragraph("5. References", style="Heading 1")
    document.add_paragraph(
        "[Reference entry 1 - preserve its original position, numbering, punctuation, and citation mapping.]",
        style="Reference",
    )
    document.add_paragraph(
        "[Reference entry 2 - do not alphabetize, reorder, or renumber unless explicitly requested.]",
        style="Reference",
    )

    for heading, text in (
        ("Acknowledgments", "[Insert acknowledgments.]"),
        ("Funding", "[Insert funding information.]"),
        ("CRediT authorship contribution statement", "[Insert author contributions; author names may be bold.]"),
        ("Competing interests", "[Insert competing-interest statement.]"),
        ("Data availability", "[Insert data-availability statement.]"),
        ("Ethics statement", "[Insert ethics statement when applicable.]"),
    ):
        document.add_paragraph(heading, style="Heading 1")
        document.add_paragraph(text, style="SCI Statement Body")


def build_template(output: Path) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    configure_styles(document)
    for section in document.sections:
        configure_section(section)
    set_update_fields(document)

    properties = document.core_properties
    properties.title = "SCI Manuscript Template"
    properties.subject = "Reviewed English-language SCI manuscript formatting template"
    properties.author = "SCI Manuscript Format skill"
    properties.keywords = "SCI manuscript, DOCX template, three-line table, OMML"
    properties.created = datetime(2026, 9, 1, tzinfo=timezone.utc)
    properties.modified = datetime(2026, 9, 1, tzinfo=timezone.utc)

    add_title_block(document)
    add_abstract(document)
    with TemporaryDirectory(prefix="sci-template-") as temporary:
        figure_path = Path(temporary) / "figure-placeholder.png"
        create_placeholder_figure(figure_path)
        add_manuscript_elements(document, figure_path)

    document.save(output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    path = build_template(args.output.resolve())
    print(path)


if __name__ == "__main__":
    main()
