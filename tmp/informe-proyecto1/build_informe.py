from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(r"C:\Projects\Paralela")
REPO = ROOT / "Proyecto1_Paralela_Zip"
ASSETS = ROOT / "tmp" / "informe-proyecto1" / "assets"
OUTPUT = REPO / "docs" / "Informe_Proyecto_1_Editable.docx"

# Preset seleccionado: narrative_proposal.
# Overrides nombrados por continuidad con el PDF fuente institucional:
# A4, margenes de 2.5 cm y paleta UVG verde.
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
MARGIN_CM = 2.5
CONTENT_WIDTH_DXA = 9072
TABLE_INDENT_DXA = 120

FONT_BODY = "Calibri"
FONT_CODE = "Consolas"
FONT_COVER = "Times New Roman"
GREEN = "000000"
DARK_GREEN = "000000"
PALE_GREEN = "D9D9D9"
LIGHT_GREEN = "FFFFFF"
GOLD = "000000"
INK = "000000"
MUTED = "555555"
GRAY = "D9D9D9"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
RED = "9B1C1C"


def set_font(run, name=FONT_BODY, size=None, color=INK, bold=None, italic=None):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), name)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    if size is not None:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_width(cell, width_dxa):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(width_dxa))
    tc_w.set(qn("w:type"), "dxa")


def set_table_borders(table, color="AAB5AE", size=6):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), str(size))
        tag.set(qn("w:space"), "0")
        tag.set(qn("w:color"), color)


def configure_table_geometry(table, widths):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(TABLE_INDENT_DXA))
    tbl_ind.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell, width in zip(row.cells, widths):
            set_cell_width(cell, width)
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def format_cell_text(cell, bold=False, color=INK, size=9.2, align=WD_ALIGN_PARAGRAPH.LEFT):
    for paragraph in cell.paragraphs:
        paragraph.alignment = align
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(1.5)
        paragraph.paragraph_format.line_spacing = 1.0
        for run in paragraph.runs:
            set_font(run, size=size, color=color, bold=bold)


def add_table(doc, headers, rows, widths=None, font_size=9.2, header_fill=PALE_GREEN):
    if widths is None:
        base = CONTENT_WIDTH_DXA // len(headers)
        widths = [base] * len(headers)
        widths[-1] += CONTENT_WIDTH_DXA - sum(widths)
    table = doc.add_table(rows=1, cols=len(headers))
    configure_table_geometry(table, widths)
    set_table_borders(table)
    repeat_header(table.rows[0])
    for i, value in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = value
        shade_cell(cell, header_fill)
        format_cell_text(cell, bold=True, color=DARK_GREEN, size=font_size, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row_data in rows:
        row = table.add_row()
        for i, value in enumerate(row_data):
            cell = row.cells[i]
            cell.text = str(value)
            if len(table.rows) % 2 == 1:
                shade_cell(cell, LIGHT_GRAY)
            format_cell_text(cell, size=font_size)
        configure_table_geometry(table, widths)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    return table


def add_para(doc, text="", *, bold=False, italic=False, size=11, color=INK,
             align=WD_ALIGN_PARAGRAPH.JUSTIFY, before=0, after=8,
             line_spacing=1.333, keep=False, font_name=FONT_BODY):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.keep_together = keep
    r = p.add_run(text)
    set_font(r, name=font_name, size=size, color=color, bold=bold, italic=italic)
    return p


def create_list_numbering(doc, *, fmt, level_text):
    numbering = doc.part.numbering_part.element
    abstract_ids = [int(el.get(qn("w:abstractNumId"))) for el in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(el.get(qn("w:numId"))) for el in numbering.findall(qn("w:num"))]
    abstract_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    level.append(start)
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), fmt)
    level.append(num_fmt)
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), level_text)
    level.append(lvl_text)
    suff = OxmlElement("w:suff")
    suff.set(qn("w:val"), "tab")
    level.append(suff)
    p_pr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), "540")
    tabs.append(tab)
    p_pr.append(tabs)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "540")
    ind.set(qn("w:hanging"), "280")
    p_pr.append(ind)
    level.append(p_pr)
    abstract.append(level)
    # OOXML exige todos los abstractNum antes de los elementos num. Insertar
    # despues provoca que Word repare el archivo y mezcle listas distintas.
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        numbering.insert(list(numbering).index(first_num), abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_id = OxmlElement("w:abstractNumId")
    abs_id.set(qn("w:val"), str(abstract_id))
    num.append(abs_id)
    numbering.append(num)
    return num_id


def apply_list_numbering(paragraph, num_id):
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num_id_el = OxmlElement("w:numId")
    num_id_el.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl)
    num_pr.append(num_id_el)
    paragraph._p.get_or_add_pPr().append(num_pr)


def add_bullets(doc, items):
    num_id = create_list_numbering(doc, fmt="bullet", level_text="•")
    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.208
        apply_list_numbering(p, num_id)
        r = p.add_run(item)
        set_font(r, size=11)


def add_numbers(doc, items):
    num_id = create_list_numbering(doc, fmt="decimal", level_text="%1.")

    for item in items:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.208
        apply_list_numbering(p, num_id)
        r = p.add_run(item)
        set_font(r, size=11)


def add_code(doc, text):
    table = doc.add_table(rows=1, cols=1)
    configure_table_geometry(table, [CONTENT_WIDTH_DXA])
    repeat_header(table.rows[0])
    set_table_borders(table, color="C7CDC9", size=4)
    cell = table.cell(0, 0)
    shade_cell(cell, "F3F5F4")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    set_font(r, name=FONT_CODE, size=8.3, color="1B2720")
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_callout(doc, title, body, fill=PALE_GREEN):
    table = doc.add_table(rows=1, cols=1)
    configure_table_geometry(table, [CONTENT_WIDTH_DXA])
    repeat_header(table.rows[0])
    set_table_borders(table, color="8FA99A", size=6)
    cell = table.cell(0, 0)
    shade_cell(cell, fill)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(title)
    set_font(r, size=10.5, color=DARK_GREEN, bold=True)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.15
    r2 = p2.add_run(body)
    set_font(r2, size=9.6, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_placeholder_box(doc, title, prompt, blank_lines=4):
    table = doc.add_table(rows=1, cols=1)
    configure_table_geometry(table, [CONTENT_WIDTH_DXA])
    repeat_header(table.rows[0])
    set_table_borders(table, color="91A098", size=6)
    cell = table.cell(0, 0)
    shade_cell(cell, LIGHT_GREEN)
    p = cell.paragraphs[0]
    r = p.add_run(title)
    set_font(r, size=10, color=DARK_GREEN, bold=True)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(2)
    r2 = p2.add_run(prompt)
    set_font(r2, size=9, color=MUTED, italic=True)
    for _ in range(blank_lines):
        q = cell.add_paragraph(" ")
        q.paragraph_format.space_after = Pt(2)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font(r, size={1: 16, 2: 13, 3: 12}[level], color=GREEN if level < 3 else DARK_GREEN, bold=True)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph(style="Caption")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(text)
    set_font(r, size=9.5, color=MUTED, italic=True)
    return p


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    rel_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rel_id)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), GREEN)
    r_pr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), FONT_BODY)
    fonts.set(qn("w:hAnsi"), FONT_BODY)
    r_pr.append(fonts)
    new_run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_field(paragraph, instruction, placeholder=""):
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run = OxmlElement("w:r")
    run.append(begin)
    run.append(instr)
    run.append(separate)
    run.append(text)
    run.append(end)
    paragraph._p.append(run)


def page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def setup_styles(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_BODY
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT_BODY)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_BODY)
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string(INK)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.space_before = Pt(0)
    pf.space_after = Pt(8)
    pf.line_spacing = 1.333

    for level, size, before, after, color in (
        (1, 16, 18, 10, GREEN),
        (2, 13, 12, 6, GREEN),
        (3, 12, 8, 4, DARK_GREEN),
    ):
        style = styles[f"Heading {level}"]
        style.font.name = FONT_BODY
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT_BODY)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_BODY)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    for style_name in ("List Bullet", "List Number"):
        style = styles[style_name]
        style.font.name = FONT_BODY
        style.font.size = Pt(11)
        style.paragraph_format.left_indent = Inches(0.375)
        style.paragraph_format.first_line_indent = Inches(-0.194)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.208

    caption = styles["Caption"]
    caption.font.name = FONT_BODY
    caption.font.size = Pt(9.5)
    caption.font.italic = True
    caption.font.color.rgb = RGBColor.from_string(MUTED)

    if "Placeholder" not in styles:
        st = styles.add_style("Placeholder", WD_STYLE_TYPE.PARAGRAPH)
        st.font.name = FONT_BODY
        st.font.size = Pt(10)
        st.font.italic = True
        st.font.color.rgb = RGBColor.from_string(MUTED)


def setup_page(doc):
    section = doc.sections[0]
    section.page_width = Cm(PAGE_WIDTH_CM)
    section.page_height = Cm(PAGE_HEIGHT_CM)
    section.top_margin = Cm(MARGIN_CM)
    section.right_margin = Cm(MARGIN_CM)
    section.bottom_margin = Cm(MARGIN_CM)
    section.left_margin = Cm(MARGIN_CM)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)
    section.different_first_page_header_footer = True

    header = section.header
    header.paragraphs[0].text = ""

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(fp, "PAGE", "1")

    # Solicitar a Word/LibreOffice actualizar TOC y demás campos al abrir.
    settings = doc.settings._element
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")


def add_cover(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("UNIVERSIDAD DEL VALLE DE GUATEMALA")
    set_font(r, name=FONT_COVER, size=15, color=INK, bold=True)
    add_para(doc, "Facultad de Ingeniería", italic=True, size=12, color=INK,
             align=WD_ALIGN_PARAGRAPH.CENTER, after=2, line_spacing=1.0, font_name=FONT_COVER)
    add_para(doc, "Departamento de Ciencias de la Computación", italic=True, size=12,
             color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, after=2, line_spacing=1.0, font_name=FONT_COVER)
    add_para(doc, "CC3069 — Programación Paralela y Distribuida · Sección 20",
             size=11, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, after=2, line_spacing=1.0, font_name=FONT_COVER)
    add_para(doc, "Ing. Santiago Solórzano", size=11, color=INK,
             align=WD_ALIGN_PARAGRAPH.CENTER, after=10, line_spacing=1.0, font_name=FONT_COVER)

    logo = ASSETS / "logo_uvg.png"
    if logo.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        picture = p.add_run().add_picture(str(logo), width=Cm(3.0))
        picture._inline.docPr.set("descr", "Logotipo de la Universidad del Valle de Guatemala")

    add_para(doc, "Proyecto 1 - Screensaver", size=20, color=INK, bold=False,
             align=WD_ALIGN_PARAGRAPH.CENTER, after=5, line_spacing=1.0, font_name=FONT_COVER)
    add_para(doc, "Paralelización de una simulación gráfica con OpenMP",
             size=16, color=INK, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
             after=18, line_spacing=1.1, font_name=FONT_COVER)

    members = [
        ("Ricardo Arturo Godínez Sánchez", "23247"),
        ("Vianka Vanessa Castro Ordoñez", "23201"),
        ("Abby Sofia Donis Agreda", "22440"),
    ]
    for name, student_id in members:
        add_para(doc, f"{name} - {student_id}", size=11, color=INK,
                 align=WD_ALIGN_PARAGRAPH.CENTER, after=3, line_spacing=1.0, font_name=FONT_COVER)
    add_para(doc, "GUATEMALA, 31 de agosto de 2026",
             size=10.5, color=INK, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
             before=24, after=0, line_spacing=1.0, font_name=FONT_COVER)
    page_break(doc)


def add_editing_guide(doc):
    add_heading(doc, "Guía de edición para el equipo", 1)
    add_callout(
        doc,
        "Propósito de esta versión",
        "Este archivo convierte el informe PDF en un Word editable y lo alinea con el repositorio actual. Los textos marcados como POR COMPLETAR son espacios intencionales para mediciones, capturas y conclusiones reales; no deben rellenarse con datos inventados.",
    )
    add_numbers(doc, [
        "Cada integrante compila el mismo commit y anota su hardware, sistema operativo, compilador y comando exacto.",
        "Cada configuración se mide al menos diez veces con el modo --bench; no se usan FPS de la ventana como única medida del kernel paralelo.",
        "Se conserva el CSV crudo y se genera el resumen con promedio, desviación estándar, speedup y eficiencia.",
        "Las capturas deben mostrar la escena, la terminal con el comando y, cuando corresponda, la verificación determinista.",
        "Antes de entregar, se reemplazan todos los campos POR COMPLETAR, se actualiza la tabla de contenido con Ctrl+A y F9 y se revisa la numeración.",
    ])
    add_table(doc, ["Campo de control", "Valor"], [
        ("Documento base", "Informe Proyecto 1.pdf (18 páginas)"),
        ("Repositorio revisado", "https://github.com/Vann06/Proyecto1_Paralela_Zip"),
        ("Commit de referencia", "abb0305 — update paralelism"),
        ("Estado del repositorio al preparar este Word", "Árbol de trabajo limpio"),
        ("Responsable de consolidación", "[POR COMPLETAR]"),
        ("Fecha límite interna", "[POR COMPLETAR]"),
    ], [2600, 6472])
    add_heading(doc, "Convención de estados", 2)
    add_table(doc, ["Marca", "Significado"], [
        ("IMPLEMENTADO", "Existe en el código actual y fue contrastado con README, CMake y fuentes."),
        ("AUTOMATIZADO", "El repositorio incluye scripts, pero el equipo aún debe ejecutar y guardar evidencia."),
        ("POR COMPLETAR", "Requiere datos o juicio del equipo; se deja un espacio editable."),
        ("NO IMPLEMENTADO", "No debe describirse como función actual del programa."),
    ], [2100, 6972])
    page_break(doc)


def add_toc(doc):
    add_heading(doc, "Contenido", 1)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    add_field(p, 'TOC \\o "1-2" \\h \\z \\u', "Actualice este campo en Word: Ctrl+A, F9.")
    add_para(doc, "Nota de edición: después de agregar resultados o mover secciones, seleccione todo el documento (Ctrl+A) y presione F9 para actualizar títulos y números de página.", size=9.5, color=MUTED, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    page_break(doc)


def add_front_matter(doc):
    add_heading(doc, "Resumen", 1)
    add_para(doc, "ZipZip es un screensaver 3D en C++ que representa una granja suspendida en el espacio. La escena incluye vacas animadas, un campo de estrellas, estrellas fugaces, planetas con texturas procedurales y aros, una plataforma con textura de grama y un OVNI-gato cuyo cuerpo metálico y piloto verde neón se renderizan por grupos del archivo OBJ. SDL2 administra la ventana y los eventos; OpenGL realiza todo el dibujo en el hilo principal.")
    add_para(doc, "La evaluación académica se concentra en la simulación de las vacas. El mismo ejecutable ofrece modo serial y paralelo. OpenMP se aplica a dos pasadas: una interacción O(N²) entre vacas y una integración O(N). Cada iteración escribe posiciones auxiliares o estados disjuntos; una reducción min calcula la distancia mínima al OVNI y la barrera implícita entre pasadas conserva el orden de dependencias. El modo --bench excluye SDL/OpenGL, usa un dt fijo y registra tiempos comparables.")
    add_para(doc, "El repositorio ya automatiza diez repeticiones por configuración y calcula promedio, desviación estándar, speedup y eficiencia. Este informe deja sin completar deliberadamente las tablas de mediciones, porque deben llenarse con ejecuciones reales de los tres integrantes sobre el commit acordado.")

    add_heading(doc, "Introducción", 1)
    add_para(doc, "La programación paralela busca aprovechar varios núcleos para reducir el tiempo de una carga divisible. En memoria compartida, OpenMP permite expresar el reparto del trabajo mediante directivas sobre ciclos, pero el rendimiento depende de la fracción serial, el balance, la sincronización, la localidad y el costo de crear o coordinar hilos [1, 2, 4, 9].")
    add_para(doc, "Una simulación gráfica permite distinguir dos responsabilidades. La actualización del estado puede procesar entidades independientes en paralelo; en cambio, SDL2 y OpenGL mantienen la creación del contexto, los eventos y el renderizado en el hilo principal. Esta separación evita llamadas gráficas concurrentes y permite medir el kernel de simulación sin confundirlo con el costo de dibujo.")

    add_heading(doc, "Antecedentes y fundamentos", 2)
    add_bullets(doc, [
        "Speedup: S(T) = t_serial / t_paralelo. Un valor mayor que 1 indica mejora respecto de la referencia serial.",
        "Eficiencia: E(T) = S(T) / T. Expresa qué fracción del speedup ideal aporta cada hilo.",
        "Ley de Amdahl: la fracción serial limita el speedup cuando el tamaño del problema se mantiene fijo [1].",
        "Ley de Gustafson: al aumentar el problema, el trabajo paralelo puede dominar y justificar más procesadores [4].",
        "False sharing: hilos que escriben datos diferentes pero cercanos pueden invalidar la misma línea de caché; el experimento con chunk 1 busca hacerlo visible [9].",
        "Scheduling: static, dynamic y guided distribuyen iteraciones con costos y equilibrio distintos. La elección debe justificarse con mediciones.",
    ])

    add_heading(doc, "Objetivos", 1)
    add_heading(doc, "Objetivo general", 2)
    add_para(doc, "Implementar y evaluar una simulación gráfica serial y paralela con OpenMP, manteniendo el renderizado en el hilo principal y cuantificando el efecto del número de vacas, hilos, scheduling y distribución de memoria sobre el rendimiento.")
    add_heading(doc, "Objetivos específicos", 2)
    add_bullets(doc, [
        "Conservar una única base de código capaz de ejecutar el mismo algoritmo en modo serial o paralelo.",
        "Paralelizar primero la actualización de las vacas y comprobar que no introduce condiciones de carrera.",
        "Separar el tiempo de simulación del renderizado y ofrecer un benchmark reproducible sin ventana.",
        "Comparar tiempos, speedup y eficiencia con al menos diez repeticiones por configuración.",
        "Evaluar el efecto de scheduling y chunk, incluido un caso destinado a observar false sharing.",
        "Medir por separado si la paralelización de estrellas compensa su overhead.",
        "Documentar hardware, software, comandos, resultados y conclusiones verificables de cada integrante.",
    ])


def add_scope(doc):
    add_heading(doc, "Descripción y alcance del proyecto", 1)
    add_heading(doc, "Diseño de la simulación", 2)
    add_para(doc, "La escena final mantiene la idea de una granja espacial, pero su implementación actual difiere del concepto inicial: la plataforma es una cúpula semicircular vista desde arriba; las vacas rebotan, se separan y conservan cohesión; el OVNI recorre una trayectoria de Lissajous y no captura vacas todavía. No existe un sistema de partículas de captura, por lo que el informe no lo presenta como terminado.")
    concept = ASSETS / "concepto.jpg"
    if concept.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        picture = p.add_run().add_picture(str(concept), width=Cm(13.7))
        picture._inline.docPr.set("descr", "Concepto visual original de la granja espacial ZipZip")
        add_caption(doc, "Figura 1. Concepto visual original del proyecto; no representa exactamente todas las funciones actuales.")

    add_heading(doc, "Alcance implementado", 2)
    add_table(doc, ["Subsistema", "Estado actual", "Tratamiento paralelo"], [
        ("Ventana y eventos", "SDL2/OpenGL, 900×700 por defecto; 640×480 mínimo; cierre con Esc o X.", "Hilo principal."),
        ("Vacas", "N configurable; posición, velocidad, giro, escala, color, rebote, separación y cohesión.", "Dos pasadas seriales o OpenMP."),
        ("Estrellas", "M configurable; parpadeo, movimiento y estrellas fugaces.", "Serial por defecto; experimento opcional."),
        ("OVNI-gato", "Trayectoria de Lissajous; nave metálica y gato verde neón por grupos OBJ.", "Actualización O(1), serial."),
        ("Planetas", "Cuatro planetas, texturas procedurales, rotación y aros.", "Serial; fuera del kernel evaluado."),
        ("Fondo y grama", "BMP cargados con SDL2; fallback a colores sólidos si fallan.", "Carga y render en hilo principal."),
        ("Medición", "HUD y modo --bench; CSV y resumen estadístico.", "Cronometraje con steady_clock."),
        ("Captura/partículas", "No implementado en el commit de referencia.", "Trabajo futuro; no se evalúa."),
    ], [1900, 4540, 2632], font_size=8.6)

    add_heading(doc, "Requisitos del hito inicial", 2)
    add_table(doc, ["Requisito", "Evidencia actual"], [
        ("Ventana gráfica y API elegida", "SDL2 crea ventana/contexto; OpenGL renderiza la escena."),
        ("Variable en memoria", "Escena::vacas, Escena::ax/ay y CampoEstrellas::estrellas almacenan entidades y auxiliares."),
        ("Render de un elemento", "Renderer::dibujarVacas carga cow.obj y dibuja cada instancia."),
        ("Función de actualización", "actualizarEscena determina interacción, integración, rebote y giro."),
        ("Commit identificable", "El historial conserva hitos; para este informe se fija abb0305 como referencia."),
    ], [3150, 5922])


def add_architecture(doc):
    add_heading(doc, "Arquitectura del repositorio", 1)
    add_code(doc, "Proyecto1_Paralela_Zip/\n"
                  "├─ CMakeLists.txt\n"
                  "├─ assets/\n"
                  "│  ├─ models/       cow.obj, ufo_gato.obj, planet.obj\n"
                  "│  └─ textures/     space_background.bmp, grass.bmp\n"
                  "├─ include/zipzip/  interfaces de core, assets, simulation y rendering\n"
                  "├─ src/\n"
                  "│  ├─ main.cpp      CLI, eventos, ciclo gráfico y --bench\n"
                  "│  ├─ assets/       cargador OBJ\n"
                  "│  ├─ simulation/   escena, vacas, OVNI, planetas y estrellas\n"
                  "│  └─ rendering/    OpenGL, texturas y HUD\n"
                  "├─ scripts/         benchmark.sh y resumir_benchmark.py\n"
                  "└─ docs/            propuesta, diagramas, matemática y anexos")
    add_table(doc, ["Capa", "Responsabilidad", "Dependencias relevantes"], [
        ("Aplicación", "Validar argumentos, elegir modo, eventos, loop principal y benchmark.", "SDL2 solo en modo gráfico; OpenMP runtime."),
        ("Simulación", "Crear y actualizar estado determinista.", "C++ estándar y OpenMP; no depende de SDL/OpenGL."),
        ("Renderizado", "Dibujar fondo, estrellas, plataforma, vacas, OVNI, planetas y HUD.", "OpenGL y SDL2; siempre hilo principal."),
        ("Recursos", "Cargar y normalizar OBJ; conservar rangos g/o para materiales visuales.", "C++ estándar."),
        ("Núcleo", "RNG, cámara y cronómetro compartidos.", "Headers reutilizables."),
        ("Automatización", "Repetir experimentos y resumir CSV.", "Bash y Python 3."),
    ], [1700, 4450, 2922], font_size=8.7)

    add_heading(doc, "Flujo de ejecución", 2)
    add_table(doc, ["1. Configuración", "2. Estado", "3. Simulación", "4. Salida"], [
        ("Leer/validar CLI\nConfigurar OpenMP", "Reservar vacas, ax/ay y estrellas\nCrear escena con semilla", "Actualizar OVNI/planetas\nPasada A → barrera → Pasada B\nActualizar estrellas", "Gráfico: render + HUD\nBenchmark: fila CSV"),
    ], [2268, 2268, 2268, 2268], font_size=8.7)
    add_para(doc, "El diagrama detallado se mantiene en docs/anexo_01_diagrama_flujo.md. Ambas modalidades recorren el mismo flujo; la condición usarOpenMP decide si los pragmas ejecutan más de un hilo.", size=9.5, color=MUTED, italic=True)


def add_implementation(doc):
    add_heading(doc, "Implementación secuencial y paralela", 1)
    add_heading(doc, "Una sola base de código", 2)
    add_para(doc, "No existen dos proyectos duplicados. La opción --modo serial|paralelo controla el argumento if(usarOpenMP) de las directivas. Una compilación con ZIPZIP_OPENMP=OFF sirve como referencia adicional y rechaza de forma explícita la solicitud de modo paralelo.")

    add_heading(doc, "Pasada A: interacción entre vacas", 2)
    add_para(doc, "La pasada A evalúa la separación para cada par de vacas, por lo que su complejidad es O(N²). La iteración i solo lee el vector completo de vacas y escribe ax[i] y ay[i]. La distancia mínima al OVNI se combina con una reducción; no altera el movimiento.")
    add_code(doc, "#pragma omp parallel for if(usarOpenMP) schedule(runtime) \\\n"
                  "    reduction(min:distanciaMinimaOvni)\n"
                  "for (long i = 0; i < cantidadVacas; ++i) {\n"
                  "    // lee vacas[j]; escribe exclusivamente ax[i], ay[i]\n"
                  "}")

    add_heading(doc, "Pasada B: integración y rebote", 2)
    add_para(doc, "La barrera implícita al terminar la pasada A garantiza que todas las aceleraciones estén disponibles. En la pasada B, cada iteración modifica una Instancia diferente, aplica cohesión, limita rapidez, actualiza giro y resuelve el rebote dentro de la cúpula. Es O(N) y tiene menor intensidad computacional.")
    add_code(doc, "#pragma omp parallel for if(usarOpenMP) schedule(runtime)\n"
                  "for (long i = 0; i < cantidadVacas; ++i) {\n"
                  "    // lee ax[i], ay[i]; actualiza exclusivamente vacas[i]\n"
                  "}")

    add_heading(doc, "Sincronización y memoria compartida", 2)
    add_table(doc, ["Dato", "Acceso concurrente", "Protección/razón"], [
        ("Escena::vacas en pasada A", "Solo lectura.", "No requiere protección."),
        ("Escena::ax[i], ay[i]", "Un índice por iteración.", "Escrituras disjuntas."),
        ("distanciaMinimaOvni", "Mínimo global.", "reduction(min:...)."),
        ("Dependencia A → B", "B consume aceleraciones calculadas en A.", "Barrera implícita al final del primer parallel for."),
        ("Escena::vacas[i] en pasada B", "Una vaca por iteración.", "Escrituras disjuntas; no requiere critical ni mutex."),
        ("SDL/OpenGL/Renderer", "Contexto gráfico compartido.", "Se limita al hilo principal."),
    ], [2500, 2600, 3972], font_size=8.7)

    add_heading(doc, "Estrellas: experimento opcional", 2)
    add_para(doc, "Las estrellas se mantienen seriales por defecto. Con M=180, cada iteración hace poco trabajo y el overhead de OpenMP puede superar el ahorro. La bandera --estrellas-paralelas existe únicamente para medir esa hipótesis con M=180 y un M grande, por ejemplo 10 000; la decisión final debe basarse en las tablas del equipo.")

    add_heading(doc, "Programación defensiva", 2)
    add_bullets(doc, [
        "Rangos explícitos: vacas 1..100000, estrellas 1..1000000, hilos 1..1024, ancho 640..7680 y alto 480..4320.",
        "Rechazo de enteros parciales, semilla cero, opciones desconocidas, schedule inválido y argumentos incompletos.",
        "Mensaje claro si se pide OpenMP en una compilación serial o estrellas paralelas en modo serial.",
        "Carga de texturas con fallback; un BMP ausente no derriba toda la aplicación.",
        "Semilla reproducible y --dump-estado para contrastar resultados seriales/paralelos.",
    ])

    add_heading(doc, "Presentación gráfica actual", 2)
    add_table(doc, ["Elemento", "Implementación"], [
        ("Fondo", "space_background.bmp como quad ortográfico; estrellas encima."),
        ("Plataforma", "Cúpula semicircular con grass.bmp repetible."),
        ("Vacas", "cow.obj con transformaciones y colores por instancia."),
        ("OVNI-gato", "ufo_gato.obj conserva grupos g/o; nave metálica y gato verde neón."),
        ("Planetas", "planet.obj, cuatro texturas procedurales, rotación y aros."),
        ("HUD", "FPS, tiempo SIM, hilos y distancia mínima al OVNI."),
    ], [2300, 6772])


def add_run_guide(doc):
    add_heading(doc, "Entorno de ejecución y compilación", 2)
    add_heading(doc, "Requisitos", 3)
    add_bullets(doc, [
        "CMake 3.20 o posterior y compilador con C++17.",
        "SDL2 y OpenGL.",
        "OpenMP para el modo paralelo; puede deshabilitarse con ZIPZIP_OPENMP=OFF.",
        "Bash para benchmark.sh y Python 3 para resumir_benchmark.py.",
    ])

    add_heading(doc, "Windows — MSYS2 UCRT64", 3)
    add_code(doc, "pacman -S --needed mingw-w64-ucrt-x86_64-gcc \\\n"
                  "  mingw-w64-ucrt-x86_64-cmake mingw-w64-ucrt-x86_64-SDL2 \\\n"
                  "  mingw-w64-ucrt-x86_64-ninja\n\n"
                  "cmake -S . -B build -G Ninja -DCMAKE_BUILD_TYPE=Release\n"
                  "cmake --build build\n"
                  "./build/zipzip.exe")
    add_callout(doc, "Importante en UCRT64", "La barra invertida solo continúa una orden cuando queda al final de la línea. No escriba “\\ main.cpp”. Con CMake no es necesario memorizar la lista de archivos ni bibliotecas.", fill="FFF7E2")

    add_heading(doc, "Linux", 3)
    add_code(doc, "cmake -S . -B build -DCMAKE_BUILD_TYPE=Release\n"
                  "cmake --build build -j\n"
                  "./build/zipzip")

    add_heading(doc, "Compilación serial adicional", 3)
    add_code(doc, "cmake -S . -B build-serial -DZIPZIP_OPENMP=OFF -DCMAKE_BUILD_TYPE=Release\n"
                  "cmake --build build-serial -j")

    add_heading(doc, "Opciones relevantes", 3)
    add_table(doc, ["Opción", "Default", "Uso"], [
        ("N o --vacas N", "20", "Cantidad de vacas."),
        ("--estrellas N", "180", "Cantidad de estrellas."),
        ("--modo serial|paralelo", "paralelo*", "Activa/desactiva OpenMP en la simulación."),
        ("--hilos N", "runtime", "Número de hilos."),
        ("--schedule T", "static", "static, dynamic o guided."),
        ("--chunk N", "0", "Bloque automático o explícito."),
        ("--semilla N", "1234", "Escena reproducible."),
        ("--bench --frames N", "apagado / 300", "Medición sin ventana."),
        ("--dump-estado", "apagado", "Estado final para comparación exacta."),
        ("--estrellas-paralelas", "apagado", "Experimento dedicado de estrellas."),
        ("--ancho / --alto", "900 / 700", "Dimensiones del canvas."),
        ("--modelo ruta", "assets/models/cow.obj", "Modelo OBJ de vacas."),
    ], [2550, 1550, 4972], font_size=8.5)
    add_para(doc, "* En una compilación sin OpenMP, el modo predeterminado es serial.", size=9, color=MUTED, italic=True)

    add_heading(doc, "Verificación determinista", 3)
    add_code(doc, "./build/zipzip --bench --modo serial --vacas 3000 --frames 30 \\\n"
                  "  --semilla 1234 --dump-estado > serial.txt\n"
                  "./build/zipzip --bench --modo paralelo --vacas 3000 --frames 30 \\\n"
                  "  --semilla 1234 --hilos 8 --dump-estado > paralelo.txt\n"
                  "diff serial.txt paralelo.txt")
    add_para(doc, "Resultado esperado: diff no imprime diferencias. En Windows UCRT64, los mismos comandos funcionan con ./build/zipzip.exe.")


def add_methodology(doc):
    add_heading(doc, "Metodología", 1)
    add_callout(doc, "Regla académica", "Cada configuración debe tener al menos 10 mediciones. Una sola corrida sirve para diagnosticar, no para sostener conclusiones.")
    add_heading(doc, "Variables y métricas", 2)
    add_table(doc, ["Tipo", "Variables"], [
        ("Independientes", "N vacas; M estrellas; T hilos; modo; schedule; chunk; equipo."),
        ("Controladas", "Commit, Release, semilla 1234, frames, aplicaciones en segundo plano y perfil de energía."),
        ("Dependientes", "ms_sim_avg, ms_sim_p95, ms_pasada_a_avg, ms_pasada_b_avg y FPS derivado."),
        ("Calculadas", "Promedio, desviación estándar, speedup y eficiencia."),
    ], [2200, 6872])

    add_heading(doc, "Protocolo por integrante", 2)
    add_numbers(doc, [
        "Abrir una terminal en la raíz del repositorio y confirmar el commit con git rev-parse --short HEAD.",
        "Compilar en Release con OpenMP y ejecutar ./build/zipzip --ayuda.",
        "Ejecutar una prueba visual y guardar una captura con HUD visible.",
        "Comprobar equivalencia serial/paralela con --dump-estado y la misma semilla.",
        "Ejecutar benchmark.sh sin reducir REPETICIONES por debajo de 10.",
        "Generar resumen.csv con resumir_benchmark.py.",
        "Guardar CSV, capturas y observaciones en una carpeta identificada con el nombre del integrante.",
        "Copiar al informe los agregados; no transcribir a mano todas las filas crudas.",
    ])
    add_code(doc, "bash scripts/benchmark.sh build/zipzip > resultados.csv\n"
                  "python scripts/resumir_benchmark.py resultados.csv > resumen.csv")
    add_para(doc, "En Windows UCRT64, use build/zipzip.exe como argumento del script.")

    add_heading(doc, "Experimentos automatizados", 2)
    add_table(doc, ["Experimento", "Configuraciones predeterminadas", "Pregunta"], [
        ("Speedup", "N={1000,3000,5000}; T={1,2,4,8}; static.", "¿Cómo escala la actualización de vacas?"),
        ("Scheduling", "N=5000; static, dynamic, guided; T máximo.", "¿Compensa el balance dinámico su overhead?"),
        ("False sharing", "N=3000; static con chunk=0 y chunk=1.", "¿Las escrituras intercaladas degradan la pasada B?"),
        ("Estrellas", "M={180,10000}; serial, paralelo sin y con estrellas OpenMP.", "¿Cuándo conviene paralelizar estrellas?"),
    ], [1900, 3970, 3202], font_size=8.7)
    add_heading(doc, "Criterios de validez", 2)
    add_bullets(doc, [
        "Usar una build Release y el mismo commit para toda comparación dentro de un equipo.",
        "Ejecutar 30 frames de calentamiento internos antes de los frames medidos, como hace --bench.",
        "No mezclar equipos en una razón de speedup: la base serial debe provenir del mismo equipo y experimento.",
        "Reportar anomalías, thermal throttling, procesos en segundo plano y cambios de perfil de energía.",
        "Conservar resultados que contradigan la hipótesis; analizarlos en vez de descartarlos sin justificación.",
    ])


MEMBERS = [
    {
        "name": "Vianka Vanessa Castro Ordoñez",
        "id": "23201",
        "hardware": [
            ("Equipo", "HP EliteBook [VERIFICAR modelo exacto]"),
            ("CPU", "Intel Core i7 de 11.ª generación [VERIFICAR modelo]"),
            ("Núcleos / hilos lógicos", "4 / 8 [VERIFICAR]"),
            ("Sistema operativo", "Windows 11 Pro [AGREGAR versión]"),
        ],
    },
    {
        "name": "Ricardo Arturo Godínez Sánchez",
        "id": "23247",
        "hardware": [
            ("Equipo", "Lenovo Yoga 6 [VERIFICAR modelo exacto]"),
            ("CPU", "AMD Ryzen 7 5700U con Radeon Graphics"),
            ("Núcleos / hilos lógicos", "8 / 16 [VERIFICAR]"),
            ("Sistema operativo", "Arch Linux [AGREGAR kernel/versión]"),
        ],
    },
    {
        "name": "Abby Sofia Donis Agreda",
        "id": "22440",
        "hardware": [
            ("Equipo", "Lenovo ThinkPad [VERIFICAR modelo exacto]"),
            ("CPU", "Intel Core i7 de 8.ª generación [VERIFICAR modelo]"),
            ("Núcleos / hilos lógicos", "4 / 8 [VERIFICAR]"),
            ("Sistema operativo", "Windows 11 Pro [AGREGAR versión]"),
        ],
    },
]


def add_member_section(doc, member, index):
    page_break(doc)
    add_heading(doc, f"Resultados de {member['name']}", 2)
    add_para(doc, f"Carné: {member['id']}. Complete esta sección en el equipo donde se ejecutaron las pruebas; no copie los datos de otro integrante.", color=MUTED, italic=True)
    base_rows = member["hardware"] + [
        ("RAM", "[POR COMPLETAR: capacidad y velocidad si se conoce]"),
        ("GPU y controlador", "[POR COMPLETAR]"),
        ("Compilador", "[POR COMPLETAR: g++ --version]"),
        ("CMake / Ninja", "[POR COMPLETAR: versiones]"),
        ("OpenMP", "[POR COMPLETAR: _OPENMP/runtime]"),
        ("Commit", "[POR COMPLETAR; esperado abb0305 o commit final acordado]"),
        ("Perfil de energía", "[POR COMPLETAR]"),
        ("Fecha y hora", "[POR COMPLETAR]"),
    ]
    add_table(doc, ["Dato", "Valor observado"], base_rows, [2650, 6422], font_size=8.8)

    add_heading(doc, "Reproducibilidad y archivos", 2)
    add_table(doc, ["Evidencia", "Valor / ruta"], [
        ("Comando de compilación", "[POR COMPLETAR]"),
        ("Comando de benchmark", "[POR COMPLETAR]"),
        ("CSV crudo", "[POR COMPLETAR: ruta o enlace]"),
        ("CSV resumido", "[POR COMPLETAR: ruta o enlace]"),
        ("Cantidad de mediciones/configuración", "[POR COMPLETAR; mínimo 10]"),
        ("Resultado de diff", "[POR COMPLETAR: sin diferencias / explicar]"),
    ], [3000, 6072], font_size=8.8)

    add_heading(doc, "Resumen de mediciones", 2)
    rows = [("[Exp.]", "[N/M]", "[Modo/T/schedule/chunk]", "[ms]", "[ms]", "[×]", "[%]") for _ in range(8)]
    add_table(doc, ["Exp.", "N/M", "Configuración", "Prom.", "Desv.", "Speedup", "Efic."], rows,
              [700, 850, 2860, 1050, 1050, 1250, 1312], font_size=7.8)
    add_placeholder_box(doc, "Captura de ejecución visual", "Inserte una captura legible de la ventana con HUD y anote el comando utilizado.", blank_lines=7)
    add_placeholder_box(doc, "Observaciones del integrante", "Describa estabilidad, ruido, temperatura, anomalías, mejor configuración y cualquier limitación.", blank_lines=5)


def add_consolidated_results(doc):
    page_break(doc)
    add_heading(doc, "Resultados consolidados", 2)
    add_callout(doc, "Datos aún no incorporados", "Los valores obtenidos durante validaciones rápidas de desarrollo no se presentan como resultados académicos. Esta sección debe llenarse con las mediciones completas de los tres integrantes, cada una con al menos diez repeticiones y evidencia guardada.", fill="FFF7E2")
    add_heading(doc, "Speedup y eficiencia de vacas", 2)
    rows = []
    for member in ("Vianka", "Ricardo", "Abby"):
        for n in (1000, 3000, 5000):
            rows.append((member, n, "[T]", "[serial ms]", "[paralelo ms]", "[×]", "[%]"))
    add_table(doc, ["Equipo", "N", "T", "t serial", "t paralelo", "Speedup", "Eficiencia"], rows,
              [1300, 800, 650, 1400, 1500, 1600, 1822], font_size=7.8)

    add_heading(doc, "Scheduling", 2)
    add_table(doc, ["Equipo", "N", "T", "static", "dynamic", "guided", "Mejor"],
              [(m, 5000, "[T máx.]", "[ms]", "[ms]", "[ms]", "[tipo]") for m in ("Vianka", "Ricardo", "Abby")],
              [1300, 800, 900, 1350, 1350, 1350, 2022], font_size=8.2)

    add_heading(doc, "False sharing", 2)
    add_table(doc, ["Equipo", "N", "T", "chunk=0", "chunk=1", "Cambio %", "Interpretación"],
              [(m, 3000, "[T máx.]", "[ms]", "[ms]", "[%]", "[POR COMPLETAR]") for m in ("Vianka", "Ricardo", "Abby")],
              [1300, 750, 850, 1200, 1200, 1250, 2522], font_size=8.1)

    add_heading(doc, "Paralelización opcional de estrellas", 2)
    rows = []
    for m in ("Vianka", "Ricardo", "Abby"):
        for stars in (180, 10000):
            rows.append((m, stars, "[paralelo sin estrellas OMP]", "[paralelo con estrellas OMP]", "[conviene/no]"))
    add_table(doc, ["Equipo", "M", "Base", "Estrellas OpenMP", "Conclusión"], rows,
              [1300, 900, 2380, 2380, 2112], font_size=8.2)

    add_placeholder_box(doc, "Gráfica 1 — Speedup vs. número de hilos", "Inserte una gráfica por N o por equipo. Debe incluir título, ejes con unidades, leyenda y referencia en el texto.", blank_lines=8)
    add_placeholder_box(doc, "Gráfica 2 — Eficiencia vs. número de hilos", "Inserte una gráfica comparable con la anterior y explique la pérdida de eficiencia.", blank_lines=8)
    add_placeholder_box(doc, "Gráfica 3 — Scheduling / chunk / estrellas", "Use una o más gráficas solo si aportan una comparación clara; indique promedio y variabilidad.", blank_lines=8)


def add_analysis(doc):
    add_heading(doc, "Análisis y discusión", 1)
    add_para(doc, "Esta sección debe relacionar los valores consolidados con la arquitectura y no limitarse a repetir las tablas. Complete cada apartado con números concretos y mencione el equipo, N y T de la observación.")
    prompts = [
        ("Escalabilidad", "¿Para qué N aparece speedup? ¿En qué T se aplana o retrocede? Compare con Amdahl y con la intensidad O(N²) de la pasada A."),
        ("Eficiencia", "¿Cómo cambia E(T)? ¿La caída se explica por trabajo serial, overhead, memoria, frecuencia/temperatura o variabilidad?"),
        ("Pasadas A y B", "Compare ms_pasada_a_avg y ms_pasada_b_avg. ¿Cuál domina y cuál escala mejor?"),
        ("Scheduling", "¿Static fue suficiente por el costo uniforme? Si dynamic/guided cambia el tiempo, cuantifique el cambio."),
        ("False sharing", "¿chunk=1 empeoró la pasada B o el total? Explique la relación con Instancia de 40 bytes y líneas de caché de 64 bytes como hipótesis, no como hecho no medido."),
        ("Estrellas", "¿El overhead domina con M=180? ¿Existe un umbral donde --estrellas-paralelas mejora?"),
        ("Diferencias entre equipos", "Compare arquitectura de CPU, núcleos/hilos, sistema, compilador y variabilidad; no compare tiempos absolutos sin contexto."),
        ("Validez", "Identifique amenazas: procesos en segundo plano, turbo, throttling, perfiles de energía, versiones y precisión del cronómetro."),
    ]
    for title, prompt in prompts:
        add_placeholder_box(doc, title, prompt, blank_lines=4)

    add_heading(doc, "Conclusiones finales del equipo", 1)
    add_numbers(doc, [
        "[POR COMPLETAR: conclusión principal con el mejor speedup, N, T y equipo.]",
        "[POR COMPLETAR: conclusión sobre eficiencia y límite de escalamiento.]",
        "[POR COMPLETAR: conclusión sobre scheduling y false sharing.]",
        "[POR COMPLETAR: conclusión sobre paralelizar o no las estrellas.]",
        "[POR COMPLETAR: conclusión sobre corrección, sincronización y renderizado en hilo principal.]",
    ])
    add_heading(doc, "Recomendaciones", 1)
    add_bullets(doc, [
        "Mantener --modo serial y --modo paralelo dentro del mismo ejecutable para evitar divergencias funcionales.",
        "Conservar Release, semilla, frames y commit en toda comparación.",
        "No paralelizar OpenGL/SDL2 con este diseño; priorizar kernels de simulación medibles.",
        "Evaluar estructuras espaciales solo como trabajo futuro: reducirían O(N²), pero cambiarían el experimento académico actual.",
        "[POR COMPLETAR: recomendación basada en el mejor T observado por el equipo.]",
    ])


def add_rubric(doc):
    add_heading(doc, "Apéndices", 1)
    add_heading(doc, "Apéndice D. Lista de verificación para la entrega", 2)
    add_heading(doc, "Estado actual de cumplimiento", 3)
    add_table(doc, ["Requisito", "Estado", "Evidencia / acción"], [
        ("Versión secuencial funcional", "IMPLEMENTADO", "--modo serial en el mismo ejecutable."),
        ("Versión OpenMP funcional", "IMPLEMENTADO", "Dos parallel for en actualización de vacas."),
        ("Comparación sin duplicar proyecto", "IMPLEMENTADO", "Condición if(usarOpenMP)."),
        ("Número de threads", "IMPLEMENTADO", "--hilos y HUD/CSV."),
        ("Programación defensiva", "IMPLEMENTADO", "Rangos, errores, fallback y builds sin OpenMP."),
        ("Sincronización", "IMPLEMENTADO", "Escrituras disjuntas, reducción y barrera implícita."),
        ("Render en hilo principal", "IMPLEMENTADO", "SDL/OpenGL/Renderer no se llaman desde OpenMP."),
        ("Mediciones ≥10", "AUTOMATIZADO", "benchmark.sh obliga REPETICIONES≥10; falta ejecutar por integrante."),
        ("Speedup y eficiencia", "AUTOMATIZADO", "resumir_benchmark.py; falta consolidar resultados."),
        ("Schedules y chunk", "AUTOMATIZADO", "Experimentos dedicados; falta análisis."),
        ("Estrellas", "IMPLEMENTADO/EXPERIMENTO", "Serial por defecto, --estrellas-paralelas opcional."),
        ("Anexo 1", "IMPLEMENTADO", "docs/anexo_01_diagrama_flujo.md; resumen incorporado."),
        ("Anexo 2", "IMPLEMENTADO", "docs/anexo_02_catalogo_funciones.md; catálogo incorporado."),
        ("Anexo 3", "POR COMPLETAR", "Adjuntar bitácoras, CSV y capturas reales."),
        ("Conclusiones con datos", "POR COMPLETAR", "Rellenar después de consolidar mediciones."),
    ], [3300, 1900, 3872], font_size=8.2)

    add_heading(doc, "Lista de verificación antes de entregar", 3)
    add_bullets(doc, [
        "El código compila desde cero siguiendo el README.",
        "El commit evaluado es el mismo en los tres equipos o se explican las diferencias.",
        "Serial y paralelo producen el mismo --dump-estado con semilla idéntica.",
        "Cada configuración tiene al menos diez mediciones y unidades visibles.",
        "Las gráficas tienen título, ejes, leyenda y referencia en el texto.",
        "Las conclusiones citan valores concretos y no describen solamente la metodología.",
        "Se documentan CPU, núcleos, hilos, RAM, OS, compilador, flags, T, scheduling, N, M y seed.",
        "No quedan marcas POR COMPLETAR salvo que el docente permita explícitamente un borrador.",
        "La tabla de contenido y los números de página fueron actualizados con Ctrl+A, F9.",
    ])


def add_appendices(doc):
    page_break(doc)
    add_heading(doc, "Anexos", 1)
    add_heading(doc, "Anexo 1 — Diagrama de flujo final", 2)
    add_code(doc, "Inicio → validar argumentos → configurar OpenMP → crear estado\n"
                  "  ├─ --bench: calentamiento → medir frames → CSV\n"
                  "  └─ gráfico: SDL/OpenGL → eventos → simular → renderizar → HUD\n\n"
                  "Simular: OVNI/planetas serial → pasada A serial/OMP → barrera\n"
                  "          → pasada B serial/OMP → estrellas serial/OMP opcional")
    add_para(doc, "Fuente editable completa: docs/anexo_01_diagrama_flujo.md. Antes de entregar, puede exportarse el bloque Mermaid como imagen e insertarse debajo.", size=9.5, color=MUTED, italic=True)
    add_placeholder_box(doc, "Exportación del diagrama", "Inserte aquí la versión gráfica final del diagrama Mermaid si la rúbrica exige una figura.", blank_lines=8)

    add_heading(doc, "Anexo 2 — Catálogo de funciones", 2)
    function_rows = [
        ("leerArgumentos", "argc/argv", "bool", "Valida CLI y rangos."),
        ("configurarOpenMP", "Opciones", "hilos", "Configura threads y schedule."),
        ("ejecutarBench", "Opciones", "código", "Calienta, mide y emite CSV."),
        ("main", "argumentos", "código", "Coordina benchmark o modo gráfico."),
        ("Cronometro::*", "muestras", "estadísticas", "Promedio, mínimo, máximo y p95."),
        ("xorshift32 / Rng", "semilla/estado", "aleatorio", "Escena determinista."),
        ("cargarOBJ", "ruta/modelo", "bool", "Triangula, normaliza y conserva grupos."),
        ("crearEscena", "N, tamaño, seed", "—", "Reserva e inicializa entidades."),
        ("actualizarEscena", "Escena, dt, modo", "—", "Pasadas A y B; serial/OpenMP."),
        ("crearCampoEstrellas", "M, tamaño, seed", "—", "Inicializa estrellas y fugaces."),
        ("actualizarCampoEstrellas", "Campo, dt, modo", "—", "Actualiza brillo/movimiento."),
        ("Renderer::inicializar", "contexto", "—", "Configura OpenGL y recursos."),
        ("cargarTexturaBMP", "ruta/modo", "bool", "Sube BMP o permite fallback."),
        ("Renderer::dibujar", "modelos/estado/métricas", "—", "Compone el frame."),
        ("Renderer::dibujarVacas", "modelo/escena", "—", "Dibuja instancias."),
        ("Renderer::dibujarOvni", "modelo/escena", "—", "Aplica dos acabados por grupos."),
        ("Renderer::dibujarPlanetas", "escena", "—", "Texturas, rotación y aros."),
        ("Renderer::dibujarHUD", "métricas", "—", "Muestra FPS/SIM/hilos/mínimo."),
        ("benchmark.sh", "binario/variables", "CSV", "≥10 repeticiones/configuración."),
        ("resumir_benchmark.py", "CSV crudo", "CSV resumen", "Promedio, stdev, speedup, eficiencia."),
    ]
    add_table(doc, ["Función", "Entradas", "Salida", "Responsabilidad"], function_rows,
              [2200, 2000, 1300, 3572], font_size=7.9)
    add_para(doc, "El catálogo ampliado y mantenible está en docs/anexo_02_catalogo_funciones.md.", size=9.5, color=MUTED, italic=True)

    add_heading(doc, "Anexo 3 — Bitácora de pruebas", 2)
    add_table(doc, ["ID", "Integrante", "Fecha", "Commit", "Experimento", "Archivo/evidencia", "Estado"],
              [(str(i), "[Nombre]", "[AAAA-MM-DD]", "[hash]", "[tipo]", "[ruta/enlace]", "[OK/nota]") for i in range(1, 13)],
              [500, 1300, 1100, 900, 1400, 2600, 1272], font_size=7.4)
    add_heading(doc, "Anexo 4 — Repositorio", 2)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    add_hyperlink(p, "Repositorio GitHub de ZipZip", "https://github.com/Vann06/Proyecto1_Paralela_Zip")
    add_para(doc, "Commit documentado al preparar esta versión: abb0305 (update paralelism). Si el equipo entrega otro commit o tag, reemplácelo en el índice, la metodología y las bitácoras.")


def add_bibliography(doc):
    add_heading(doc, "Bibliografía", 1)
    refs = [
        "[1] Amdahl, G. M. (1967). Validity of the single processor approach to achieving large scale computing capabilities. AFIPS Spring Joint Computer Conference, 30, 483–485. https://doi.org/10.1145/1465482.1465560",
        "[2] Dagum, L., & Menon, R. (1998). OpenMP: An industry standard API for shared-memory programming. IEEE Computational Science & Engineering, 5(1), 46–55. https://doi.org/10.1109/99.660313",
        "[3] Flynn, M. J. (1972). Some computer organizations and their effectiveness. IEEE Transactions on Computers, C-21(9), 948–960. https://doi.org/10.1109/TC.1972.5009071",
        "[4] Gustafson, J. L. (1988). Reevaluating Amdahl’s Law. Communications of the ACM, 31(5), 532–533. https://doi.org/10.1145/42411.42415",
        "[5] Khronos Group. (2022). OpenGL 4.6 Core Profile Specification. https://registry.khronos.org/OpenGL/specs/gl/glspec46.core.pdf",
        "[6] OpenMP Architecture Review Board. (2021). OpenMP Application Programming Interface, Version 5.2. https://www.openmp.org/spec-html/5.2/openmp.html",
        "[7] Reynolds, C. W. (1987). Flocks, herds and schools: A distributed behavioral model. ACM SIGGRAPH Computer Graphics, 21(4), 25–34. https://doi.org/10.1145/37402.37406",
        "[8] SDL contributors. (s. f.). FAQ: Development. SDL2 Wiki. Recuperado el 20 de agosto de 2026 de https://wiki.libsdl.org/SDL2/FAQDevelopment",
        "[9] Williams, S., Waterman, A., & Patterson, D. (2009). Roofline: An insightful visual performance model for multicore architectures. Communications of the ACM, 52(4), 65–76. https://doi.org/10.1145/1498765.1498785",
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Cm(0.65)
        p.paragraph_format.first_line_indent = Cm(-0.65)
        p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(ref)
        set_font(r, size=9.5)


def build():
    doc = Document()
    setup_styles(doc)
    setup_page(doc)
    props = doc.core_properties
    props.title = "Proyecto 1 — Paralelización de una simulación gráfica con OpenMP"
    props.subject = "Informe técnico editable y plantilla de resultados"
    props.author = "Ricardo Godínez; Vianka Castro; Abby Donis"
    props.keywords = "OpenMP, SDL2, OpenGL, speedup, eficiencia, ZipZip"
    props.comments = "Actualizado a partir del PDF fuente y del repositorio en commit abb0305."

    add_cover(doc)
    add_toc(doc)
    add_front_matter(doc)
    add_scope(doc)
    add_architecture(doc)
    add_implementation(doc)
    add_run_guide(doc)
    add_methodology(doc)
    add_heading(doc, "Resultados", 1)
    add_para(doc, "Esta sección conserva el espacio de resultados del informe original. Cada integrante debe completar su equipo, comandos, archivos, al menos diez mediciones por configuración, capturas y observaciones; después se consolidan los valores para el análisis.")
    for idx, member in enumerate(MEMBERS, 1):
        add_member_section(doc, member, idx)
    add_consolidated_results(doc)
    add_analysis(doc)
    add_rubric(doc)
    add_appendices(doc)
    add_bibliography(doc)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build()
