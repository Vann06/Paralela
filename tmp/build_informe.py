from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


SOURCE = Path(sys.argv[1])
OUTPUT = Path(sys.argv[2])


def set_font(run, name: str | None = None, size: float | None = None, bold=None, italic=None):
    if name:
        run.font.name = name
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:ascii"), name)
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_outline_level(style, level: int):
    ppr = style.element.get_or_add_pPr()
    node = ppr.find(qn("w:outlineLvl"))
    if node is None:
        node = OxmlElement("w:outlineLvl")
        ppr.append(node)
    node.set(qn("w:val"), str(level))


def ensure_styles(doc: Document):
    styles = doc.styles

    if "Heading 2" not in [s.name for s in styles]:
        style = styles.add_style("Heading 2", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Aptos Display"
        style.font.size = Pt(15)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(5)
        style.paragraph_format.keep_with_next = True
        set_outline_level(style, 1)

    if "Heading 3" not in [s.name for s in styles]:
        style = styles.add_style("Heading 3", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Aptos Display"
        style.font.size = Pt(12)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(9)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.keep_with_next = True
        set_outline_level(style, 2)

    if "Texto del informe" not in [s.name for s in styles]:
        style = styles.add_style("Texto del informe", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Aptos"
        style.font.size = Pt(11)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.first_line_indent = Inches(0.3)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.space_after = Pt(6)

    if "Espacio reservado" not in [s.name for s in styles]:
        style = styles.add_style("Espacio reservado", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Aptos"
        style.font.size = Pt(10.5)
        style.font.italic = True
        style.font.color.rgb = RGBColor(70, 70, 70)
        style.paragraph_format.left_indent = Inches(0.25)
        style.paragraph_format.right_indent = Inches(0.25)
        style.paragraph_format.space_before = Pt(5)
        style.paragraph_format.space_after = Pt(8)

    if "Código" not in [s.name for s in styles]:
        style = styles.add_style("Código", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Consolas"
        style.font.size = Pt(9)
        style.paragraph_format.left_indent = Inches(0.25)
        style.paragraph_format.right_indent = Inches(0.1)
        style.paragraph_format.space_before = Pt(4)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.keep_together = True

    if "Pie de figura" not in [s.name for s in styles]:
        style = styles.add_style("Pie de figura", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Aptos"
        style.font.size = Pt(9.5)
        style.font.italic = True
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.space_before = Pt(3)
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.keep_with_next = True

    if "Bibliografía APA" not in [s.name for s in styles]:
        style = styles.add_style("Bibliografía APA", WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.font.name = "Aptos"
        style.font.size = Pt(10.5)
        style.paragraph_format.left_indent = Inches(0.3)
        style.paragraph_format.first_line_indent = Inches(-0.3)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.05


def shade_paragraph(paragraph, fill="F2F2F2"):
    ppr = paragraph._p.get_or_add_pPr()
    shd = ppr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        ppr.append(shd)
    shd.set(qn("w:fill"), fill)
    borders = ppr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        ppr.append(borders)
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "10")
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), "7F7F7F")
    borders.append(left)


def add_body(doc: Document, text: str, *, indent=True):
    p = doc.add_paragraph(style="Texto del informe")
    if not indent:
        p.paragraph_format.first_line_indent = Inches(0)
    p.add_run(text)
    return p


def add_labeled(doc: Document, label: str, text: str):
    p = doc.add_paragraph(style="Texto del informe")
    p.paragraph_format.first_line_indent = Inches(0)
    r = p.add_run(label)
    r.bold = True
    p.add_run(text)
    return p


def add_placeholder(doc: Document, title: str, text: str):
    p = doc.add_paragraph(style="Espacio reservado")
    shade_paragraph(p)
    r = p.add_run(title + ". ")
    r.bold = True
    r.italic = True
    p.add_run(text)
    return p


def add_numbering_definition(doc: Document, kind: str) -> int:
    numbering = doc.part.numbering_part.element
    abstract_ids = [int(x.get(qn("w:abstractNumId"))) for x in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(x.get(qn("w:numId"))) for x in numbering.findall(qn("w:num"))]
    abstract_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)

    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)
    numfmt = OxmlElement("w:numFmt")
    numfmt.set(qn("w:val"), "bullet" if kind == "bullet" else "decimal")
    lvl.append(numfmt)
    lvltext = OxmlElement("w:lvlText")
    lvltext.set(qn("w:val"), "•" if kind == "bullet" else "%1.")
    lvl.append(lvltext)
    suff = OxmlElement("w:suff")
    suff.set(qn("w:val"), "tab")
    lvl.append(suff)
    ppr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), "720")
    tabs.append(tab)
    ppr.append(tabs)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "720")
    ind.set(qn("w:hanging"), "360")
    ppr.append(ind)
    lvl.append(ppr)
    abstract.append(lvl)
    numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), "1")
    override.append(start_override)
    num.append(override)
    numbering.append(num)
    return num_id


def apply_numbering(paragraph, num_id: int):
    ppr = paragraph._p.get_or_add_pPr()
    numpr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    numpr.append(ilvl)
    numpr.append(num)
    ppr.append(numpr)


def add_numbering_instance(doc: Document, kind: str) -> int:
    """Create a fresh list instance from a valid numbering definition in the source."""
    numbering = doc.part.numbering_part.element
    wanted = "bullet" if kind == "bullet" else "decimal"
    abstract_id = None
    for abstract in numbering.findall(qn("w:abstractNum")):
        lvl = abstract.find(qn("w:lvl"))
        if lvl is None or lvl.get(qn("w:ilvl")) != "0":
            continue
        numfmt = lvl.find(qn("w:numFmt"))
        if numfmt is not None and numfmt.get(qn("w:val")) == wanted:
            abstract_id = int(abstract.get(qn("w:abstractNumId")))
            break
    if abstract_id is None:
        return add_numbering_definition(doc, kind)

    num_ids = [int(x.get(qn("w:numId"))) for x in numbering.findall(qn("w:num"))]
    num_id = max(num_ids, default=0) + 1
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start_override = OxmlElement("w:startOverride")
    start_override.set(qn("w:val"), "1")
    override.append(start_override)
    num.append(override)
    numbering.append(num)
    return num_id


def add_list(doc: Document, items: list[str], kind: str):
    num_id = add_numbering_instance(doc, kind)
    for item in items:
        p = doc.add_paragraph(style="Texto del informe")
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.left_indent = Inches(0.5)
        apply_numbering(p, num_id)
        p.add_run(item)


def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    mar = tcpr.first_child_found_in("w:tcMar")
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcpr.append(mar)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_width(cell, width_dxa: int):
    tcpr = cell._tc.get_or_add_tcPr()
    tcw = tcpr.find(qn("w:tcW"))
    if tcw is None:
        tcw = OxmlElement("w:tcW")
        tcpr.append(tcw)
    tcw.set(qn("w:w"), str(width_dxa))
    tcw.set(qn("w:type"), "dxa")


def add_table(doc: Document, headers: list[str], rows: list[list[str]], widths: list[float]):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    table.style = "Table Grid"
    total_dxa = int(round(sum(widths) * 1440))

    tblpr = table._tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), str(total_dxa))
    tblw.set(qn("w:type"), "dxa")
    tblind = tblpr.find(qn("w:tblInd"))
    if tblind is None:
        tblind = OxmlElement("w:tblInd")
        tblpr.append(tblind)
    tblind.set(qn("w:w"), "0")
    tblind.set(qn("w:type"), "dxa")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(round(width * 1440))))
        grid.append(col)

    for ci, text in enumerate(headers):
        cell = table.rows[0].cells[ci]
        set_cell_width(cell, int(round(widths[ci] * 1440)))
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "D9D9D9")
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        set_font(r, "Aptos", 10, bold=True)

    trpr = table.rows[0]._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    trpr.append(repeat)

    for row_values in rows:
        row = table.add_row()
        row_trpr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        cant_split.set(qn("w:val"), "true")
        row_trpr.append(cant_split)
        for ci, text in enumerate(row_values):
            cell = row.cells[ci]
            set_cell_width(cell, int(round(widths[ci] * 1440)))
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            if ci > 0 and len(text) < 20:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            set_font(r, "Aptos", 9.5)

    after = doc.add_paragraph()
    after.paragraph_format.space_after = Pt(3)
    return table


def add_equation(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_font(r, "Cambria Math", 11, italic=True)


def add_code(doc: Document, lines: list[str]):
    p = doc.add_paragraph(style="Código")
    shade_paragraph(p, "F7F7F7")
    for i, line in enumerate(lines):
        if i:
            p.add_run().add_break()
        p.add_run(line)


def add_hyperlink(paragraph, url: str, text: str | None = None):
    text = text or url
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(color)
    rpr.append(underline)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_reference(doc: Document, prefix: str, url: str):
    p = doc.add_paragraph(style="Bibliografía APA")
    p.add_run(prefix + " ")
    add_hyperlink(p, url)


def enable_field_updates(doc: Document):
    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")


doc = Document(SOURCE)
ensure_styles(doc)

# Conserva portada, logotipo, índice y pies de página; elimina únicamente el
# esqueleto vacío que se encuentra después del control de contenido del índice.
body = doc._element.body
children = list(body)
toc_index = next(i for i, element in enumerate(children) if element.tag == qn("w:sdt"))
for element in children[toc_index + 1 :]:
    if element.tag != qn("w:sectPr"):
        body.remove(element)

# Actualiza el título sin cambiar la composición ni la tipografía de la portada.
title = doc.paragraphs[8]
title.clear()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run("Proyecto 1 - ZipZip Espacial")
set_font(r, "Times New Roman", 28)

subtitle = doc.paragraphs[9]
subtitle.clear()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = subtitle.add_run("Paralelización de una simulación gráfica con OpenMP")
set_font(r, "Times New Roman", 24)

doc.add_page_break()

doc.add_heading("Resumen", level=1)
add_body(
    doc,
    "ZipZip Espacial es un screensaver interactivo desarrollado en C++ que representa una granja suspendida en el espacio. "
    "La versión base utiliza SDL2 para crear la ventana y administrar eventos, OpenGL para el renderizado y estructuras de datos "
    "propias para actualizar vacas y estrellas. El propósito del proyecto es construir primero una referencia secuencial correcta y, "
    "a partir de ella, paralelizar con OpenMP las actualizaciones independientes de los elementos de la escena. La evaluación propuesta "
    "separa el costo de simulación del costo de renderizado y compara tiempo de ejecución, cuadros por segundo, speedup y eficiencia para "
    "diferentes tamaños del problema y cantidades de hilos. El presente informe documenta el diseño, el estado actual y el protocolo de "
    "medición; los resultados cuantitativos y las conclusiones se incorporarán cuando finalice la implementación paralela y se ejecuten "
    "todas las configuraciones experimentales."
)

doc.add_heading("Introducción", level=1)
add_body(
    doc,
    "El aumento en la cantidad de núcleos disponibles en los procesadores actuales permite reducir el tiempo de ciertas tareas cuando el "
    "trabajo puede dividirse en operaciones independientes. Sin embargo, el simple hecho de agregar hilos no garantiza una mejora: la "
    "porción secuencial, la sincronización, el movimiento de datos y la granularidad del trabajo limitan la aceleración alcanzable. Por esta "
    "razón, un proyecto de computación paralela debe comparar una línea base bien definida con una versión paralela funcionalmente equivalente, "
    "y medir ambas bajo condiciones controladas."
)
add_body(
    doc,
    "En este proyecto se utiliza una simulación gráfica como caso de estudio. La escena contiene múltiples vacas con posición, velocidad, giro, "
    "escala y color propios, además de un campo de estrellas con movimiento y variación de brillo. Cada elemento actualiza principalmente su "
    "propio estado, de modo que los ciclos de actualización exhiben paralelismo de datos. OpenMP resulta apropiado porque permite expresar ese "
    "paralelismo sobre memoria compartida mediante directivas de compilador y rutinas de ejecución, sin reemplazar la organización general del "
    "programa en C++ (Dagum y Menon, 1998; OpenMP Architecture Review Board, 2021)."
)
add_body(
    doc,
    "La aplicación también plantea una separación importante entre simulación y presentación. Las transformaciones y reglas de movimiento se "
    "calculan en CPU, mientras que el dibujo se realiza por medio de OpenGL. Las operaciones de video de SDL2 deben permanecer en el hilo principal "
    "en la mayoría de los backends, por lo que la estrategia del proyecto consiste en paralelizar la actualización de datos y conservar el manejo "
    "de eventos y el renderizado en el hilo principal (SDL contributors, s. f.). Esta decisión reduce el riesgo de condiciones de carrera sobre el "
    "contexto gráfico y permite atribuir las mejoras al trabajo de simulación."
)
add_body(
    doc,
    "El alcance de este informe es presentar una estructura completa y reproducible antes de cerrar la fase experimental. Se describen los "
    "antecedentes, los objetivos, la arquitectura actual, las regiones candidatas para OpenMP, las variables compartidas y privadas, el método de "
    "medición y los formatos que se utilizarán para reportar resultados. No se incluyen cifras de rendimiento inventadas: las tablas y figuras de "
    "resultados quedan preparadas para llenarse con datos obtenidos en las máquinas del equipo."
)

doc.add_heading("Antecedentes", level=1)
doc.add_heading("Computación paralela y memoria compartida", level=2)
add_body(
    doc,
    "Flynn (1972) clasificó las arquitecturas de cómputo según la relación entre flujos de instrucciones y flujos de datos. Los procesadores "
    "multinúcleo utilizados en este proyecto permiten que varios hilos ejecuten instrucciones sobre distintos elementos de una misma colección, "
    "lo cual se aproxima al patrón MIMD a nivel de núcleos y puede complementarse con vectorización dentro de cada núcleo. En el modelo de memoria "
    "compartida, los hilos acceden al mismo espacio de direcciones. Esto simplifica el intercambio de información, pero exige definir con claridad "
    "qué datos se comparten, cuáles son privados y en qué puntos se requiere sincronización."
)
add_body(
    doc,
    "Las simulaciones con numerosas entidades suelen actualizar atributos equivalentes para cada elemento. Si una entidad no necesita leer ni "
    "modificar el estado de las demás durante una iteración, el conjunto puede particionarse entre hilos. Este patrón aparece en animaciones y "
    "modelos de agentes, donde cada entidad mantiene un estado local y aplica reglas de movimiento en pasos discretos (Reynolds, 1987). En ZipZip "
    "Espacial, las vacas rebotan contra límites globales y las estrellas se envuelven al cruzar el borde; ambas operaciones pueden evaluarse por "
    "elemento sin actualizar posiciones ajenas."
)

doc.add_heading("Modelo de ejecución de OpenMP", level=2)
add_body(
    doc,
    "OpenMP sigue un modelo fork-join. El programa inicia con un hilo principal; al entrar a una región paralela se crea un equipo de hilos y, al "
    "terminar la región, la ejecución vuelve a un solo hilo. El constructo de distribución de ciclos permite repartir iteraciones entre el equipo. "
    "Para un ciclo con costo uniforme por elemento, schedule(static) suele ser un punto de partida razonable porque asigna bloques de iteraciones con "
    "poco overhead. Si el costo varía de manera importante entre entidades, schedule(dynamic) puede mejorar el balance de carga, aunque introduce "
    "más coordinación. La decisión debe basarse en mediciones y no únicamente en intuición."
)
add_body(
    doc,
    "La especificación de OpenMP también define reglas de compartición de datos y rutinas como omp_get_wtime para medición de tiempo. En el diseño "
    "propuesto, los vectores que contienen las entidades son compartidos, pero cada iteración obtiene una referencia a un elemento distinto. Las "
    "variables temporales del cuerpo del ciclo son privadas; el paso de tiempo y los límites de la escena son valores de solo lectura. El uso de "
    "default(none) en las regiones nuevas se recomienda porque obliga a declarar la intención de cada variable y ayuda a detectar errores."
)

doc.add_heading("Rendimiento, speedup y escalabilidad", level=2)
add_body(
    doc,
    "La Ley de Amdahl establece que la aceleración de un programa está limitada por la fracción que permanece secuencial (Amdahl, 1967). En esta "
    "aplicación, el procesamiento de eventos, las llamadas de renderizado y el intercambio de buffers continuarán en el hilo principal; por tanto, "
    "aunque la actualización de miles de entidades se acelere, el tiempo total por cuadro no necesariamente disminuirá en la misma proporción. La "
    "Ley de Gustafson complementa esta perspectiva al considerar problemas cuyo tamaño aumenta con los recursos disponibles (Gustafson, 1988). "
    "Esto justifica evaluar escenas pequeñas y grandes: OpenMP puede no compensar su overhead con 20 vacas, pero sí con miles de elementos."
)
add_body(
    doc,
    "El rendimiento también depende del acceso a memoria y de la capacidad de cómputo del procesador. El modelo Roofline relaciona intensidad "
    "aritmética, ancho de banda de memoria y rendimiento máximo, y recuerda que una rutina con pocas operaciones por dato puede quedar limitada por "
    "memoria antes de ocupar todos los núcleos (Williams, Waterman y Patterson, 2009). Por ello, además del tiempo total, conviene medir únicamente "
    "las funciones de actualización y observar si el beneficio se estabiliza al aumentar los hilos."
)

doc.add_heading("SDL2, OpenGL y separación de responsabilidades", level=2)
add_body(
    doc,
    "SDL2 administra la ventana, el teclado y el contexto de OpenGL. OpenGL define el proceso de transformación y rasterización utilizado para "
    "dibujar los modelos, y su especificación constituye la referencia normativa de la API gráfica (Khronos Group, 2022). El código actual usa un "
    "contexto OpenGL asociado a la ventana de SDL2, arreglos de vértices y normales para el modelo OBJ, y llamadas de dibujo por cada vaca. Esta "
    "parte no será distribuida con OpenMP durante el proyecto, porque comparte estado del contexto y porque el objetivo central es analizar el "
    "paralelismo en CPU."
)
add_body(
    doc,
    "Mantener el renderizado en el hilo principal también mejora la interpretación de los resultados. El tiempo de actualización representa el "
    "trabajo que OpenMP sí puede acelerar, mientras que el tiempo de cuadro completo incluye actualización, dibujo, manejo de eventos y espera por "
    "sincronización vertical. Se reportarán ambas mediciones para no confundir una mejora local con una mejora visible en toda la aplicación."
)

doc.add_heading("Objetivos", level=1)
doc.add_heading("Objetivo general", level=2)
add_body(
    doc,
    "Diseñar, implementar y evaluar una versión paralela con OpenMP de la actualización de entidades de ZipZip Espacial, comparándola con una "
    "versión secuencial equivalente mediante métricas de tiempo, FPS, speedup y eficiencia bajo diferentes tamaños de escena y cantidades de hilos."
)

doc.add_heading("Objetivos específicos", level=2)
add_list(
    doc,
    [
        "Conservar una versión secuencial determinista que funcione como referencia de corrección y rendimiento.",
        "Identificar ciclos y tareas independientes en la actualización de vacas, estrellas y futuros sistemas de partículas.",
        "Aplicar directivas de OpenMP con una política explícita de variables shared, private y firstprivate, evitando condiciones de carrera.",
        "Mantener las operaciones de SDL2 y OpenGL en el hilo principal para proteger el contexto gráfico.",
        "Comparar al menos una estrategia de scheduling estático con alternativas justificadas por el costo de las iteraciones.",
        "Medir cada configuración en múltiples repeticiones y reportar variación, speedup, eficiencia y comportamiento al escalar el problema.",
        "Documentar el hardware, compilador, opciones de optimización, semillas y parámetros necesarios para reproducir los experimentos.",
    ],
    "bullet",
)

doc.add_heading("Descripción y alcance del proyecto", level=1)
doc.add_heading("Concepto de la simulación", level=2)
add_body(
    doc,
    "ZipZip Espacial representa una granja nocturna suspendida en un entorno espacial. Las vacas se desplazan dentro del área visible, giran y "
    "rebotan contra sus límites. Un fondo de estrellas se mueve lentamente y cambia de brillo. La propuesta artística contempla además un OVNI "
    "pilotado por un gato, capturas de vacas, partículas, planetas y eventos visuales; estos elementos se consideran alcance incremental y deben "
    "distinguirse de las funciones ya presentes en la versión base."
)

doc.add_heading("Estado funcional y alcance incremental", level=2)
add_table(
    doc,
    ["Subsistema", "Estado de la línea base", "Meta del proyecto"],
    [
        ["Ventana y eventos", "Ventana de 900 x 700, teclado, pausa y cierre", "Conservar en el hilo principal"],
        ["Vacas", "Cantidad configurable; posición, velocidad, giro, escala y color", "Paralelizar la actualización por entidad"],
        ["Estrellas", "180 por defecto; movimiento, envoltura y brillo", "Paralelizar la actualización por entidad"],
        ["Modelo 3D", "Carga de OBJ, triangulación, normalización y normales", "Mantener carga fuera del ciclo medido"],
        ["Renderizado", "OpenGL, iluminación, wireframe, culling y HUD", "Mantener secuencial; medir su impacto"],
        ["OVNI y partículas", "Previstos en la propuesta, no integrados en la línea base revisada", "Agregar solo si se completa y valida antes de medir"],
        ["Rendimiento", "FPS visible; VSync conmutable", "Agregar temporización reproducible y exportación de datos"],
    ],
    [1.25, 2.45, 2.55],
)

doc.add_heading("Parámetros del problema", level=2)
add_body(
    doc,
    "Para describir los experimentos se utilizarán símbolos consistentes. N representa la cantidad de vacas; M, la cantidad de estrellas; K, la "
    "capacidad o cantidad activa de partículas cuando exista ese sistema; P, la cantidad de elementos decorativos o planetas si su actualización "
    "forma parte del costo; T, el número de hilos de OpenMP; W y H, el ancho y alto de la ventana; y seed, la semilla de generación "
    "pseudoaleatoria. El informe final debe indicar cuáles parámetros afectan la simulación y cuáles afectan únicamente el dibujo."
)

doc.add_heading("Diseño de la solución", level=1)
doc.add_heading("Arquitectura del software", level=2)
add_body(
    doc,
    "El repositorio separa responsabilidades en módulos pequeños. main.cpp configura SDL2 y OpenGL, procesa eventos, calcula el paso de tiempo y "
    "coordina el ciclo principal. El módulo scene administra las instancias de vacas; starfield administra las estrellas; renderer concentra las "
    "llamadas de OpenGL y el HUD; y obj_loader carga la geometría Wavefront OBJ sin depender de SDL2 ni del renderizador. CMake configura el proyecto "
    "en C++17, enlaza SDL2 y OpenGL y copia los recursos al directorio del ejecutable."
)
add_list(
    doc,
    [
        "Aplicación: inicialización, argumentos, eventos, pausa, VSync y ciclo principal.",
        "Simulación: creación y actualización del estado de vacas y estrellas.",
        "Renderizado: proyección, iluminación, dibujo de estrellas y vacas, y HUD.",
        "Recursos: lectura, triangulación y normalización del modelo OBJ.",
        "Rendimiento: temporización de la línea base y de la variante OpenMP, aún por integrar como módulo explícito.",
    ],
    "bullet",
)

doc.add_heading("Estructuras de datos", level=2)
add_body(
    doc,
    "Cada Instancia de vaca almacena posición x-y, ángulo y velocidad de giro, escala, color y velocidad lineal. La estructura Escena contiene un "
    "std::vector<Instancia> y los límites visibles. Cada Estrella mantiene posición, velocidad, tamaño y parámetros de brillo; CampoEstrellas contiene "
    "el vector, los límites y un tiempo global. El modelo OBJ se transforma en arreglos contiguos de posiciones, normales y coordenadas de textura. "
    "El uso de vectores facilita distribuir rangos de índices entre hilos, aunque debe vigilarse que ninguna región paralela cambie el tamaño del "
    "vector mientras otros hilos conservan referencias a sus elementos."
)

doc.add_heading("Flujo secuencial de referencia", level=2)
add_list(
    doc,
    [
        "Leer argumentos, inicializar SDL2 y crear la ventana y el contexto OpenGL.",
        "Cargar y normalizar el modelo de vaca; preparar los arreglos de vértices y normales.",
        "Crear la escena y el campo de estrellas con semillas fijas.",
        "Procesar eventos y calcular dt en cada cuadro, limitándolo para evitar saltos extremos.",
        "Si la simulación no está pausada, actualizar todas las vacas y después todas las estrellas.",
        "Dibujar el fondo, las vacas y el HUD; finalmente intercambiar los buffers de la ventana.",
        "Al cerrar, liberar el estado de OpenGL y los recursos de SDL2.",
    ],
    "decimal",
)

doc.add_heading("Identificación de regiones paralelizables", level=2)
add_body(
    doc,
    "El candidato principal es el ciclo de actualizarEscena. En una iteración se modifica únicamente e.vacas[i]; los límites y dt son valores de "
    "solo lectura. El segundo candidato es el ciclo de actualizarCampoEstrellas, en el cual cada iteración modifica una estrella diferente. El "
    "incremento de campo.tiempo debe ejecutarse una sola vez antes del ciclo paralelo. Si se incorpora un sistema de partículas, su actualización "
    "puede seguir el mismo patrón siempre que la creación y eliminación de partículas no cambie concurrentemente el contenedor."
)
add_body(
    doc,
    "No se propone paralelizar la lectura de eventos, la modificación del título de la ventana, la carga inicial del modelo ni las llamadas de "
    "dibujo. Son tareas de baja frecuencia o asociadas a estado externo. Esta división mantiene una frontera clara: los hilos de trabajo producen "
    "un estado de simulación completo y el hilo principal consume ese estado para dibujar el cuadro."
)

doc.add_heading("Aplicación del método PCAM", level=2)
add_labeled(doc, "Partición. ", "Separar los arreglos de vacas, estrellas y partículas en rangos de índices. Cada entidad constituye una unidad de trabajo independiente durante un paso de simulación.")
add_labeled(doc, "Comunicación. ", "Las iteraciones no necesitan intercambiar datos entre sí. Comparten dt, límites y tiempo global como valores de lectura; la comunicación ocurre al finalizar la región, cuando el hilo principal utiliza los estados actualizados.")
add_labeled(doc, "Aglomeración. ", "Agrupar múltiples entidades por hilo para que el trabajo útil compense el costo de iniciar y sincronizar una región paralela. Para escenas pequeñas puede ser preferible ejecutar secuencialmente mediante una condición if en la directiva.")
add_labeled(doc, "Mapeo. ", "Asignar rangos contiguos con schedule(static) como configuración inicial. Las políticas dynamic o guided se evaluarán solamente si la carga por entidad deja de ser uniforme.")

doc.add_heading("Diseño propuesto con OpenMP", level=2)
add_body(
    doc,
    "La opción más directa es aplicar parallel for a cada ciclo de actualización. Para reducir el costo de crear equipos dos veces por cuadro, una "
    "segunda variante puede utilizar una sola región parallel y dos constructos for. El uso de nowait entre los ciclos es válido únicamente si las "
    "estrellas no dependen de las vacas y si ninguna tarea intermedia consume resultados parciales. Debe existir una barrera antes de dibujar, ya "
    "que el renderizador requiere que todos los elementos hayan terminado su actualización."
)
add_code(
    doc,
    [
        "// Patrón propuesto; se integrará y validará en el código fuente.",
        "#pragma omp parallel default(none) shared(escena, campo) firstprivate(dt)",
        "{",
        "    #pragma omp for schedule(static) nowait",
        "    for (size_t i = 0; i < escena.vacas.size(); ++i) { /* actualizar vaca i */ }",
        "",
        "    #pragma omp for schedule(static)",
        "    for (size_t i = 0; i < campo.estrellas.size(); ++i) { /* actualizar estrella i */ }",
        "}",
    ],
)
add_body(
    doc,
    "El fragmento anterior es una guía de diseño, no evidencia de que la directiva ya esté integrada. Antes de adoptarlo se debe mover la "
    "actualización de campo.tiempo fuera del segundo ciclo, comprobar la compartición de variables y verificar que la versión paralela produzca "
    "los mismos estados que la secuencial para una semilla y un dt fijos."
)

doc.add_heading("Seguridad de concurrencia", level=2)
add_table(
    doc,
    ["Dato o recurso", "Tratamiento", "Justificación"],
    [
        ["dt y límites", "firstprivate o shared de solo lectura", "No cambian dentro de la región"],
        ["Vectores de entidades", "shared", "Cada iteración modifica un índice distinto"],
        ["Referencia local a una entidad", "private", "Se crea dentro de cada iteración"],
        ["campo.tiempo", "Actualización única antes del for", "Evita una carrera sobre el acumulador global"],
        ["Generador pseudoaleatorio", "Secuencial al crear; estado por hilo si se usa en paralelo", "Un RNG compartido produciría carreras y resultados no deterministas"],
        ["SDL2 y contexto OpenGL", "Solo hilo principal", "El backend de video y el contexto gráfico mantienen estado externo"],
        ["Contadores globales", "reduction o atomic, si se agregan", "Evita actualizaciones perdidas"],
    ],
    [1.45, 1.95, 2.85],
)

doc.add_heading("Estado actual de implementación", level=1)
doc.add_heading("Funciones completadas en la línea base", level=2)
add_list(
    doc,
    [
        "Creación de una ventana OpenGL de 900 x 700 mediante SDL2.",
        "Carga de un modelo Wavefront OBJ, triangulación en abanico y generación de normales cuando faltan.",
        "Normalización del modelo y almacenamiento de posiciones y normales en arreglos contiguos.",
        "Creación determinista de vacas y estrellas con generadores xorshift32 y semillas separadas.",
        "Movimiento, giro, rebote de vacas, envoltura de estrellas y variación sinusoidal de brillo.",
        "Renderizado con iluminación, control de wireframe, face culling y VSync.",
        "HUD con FPS, cantidad de vacas y cantidad de estrellas.",
        "Organización del proyecto con CMake, C++17, SDL2 y OpenGL.",
    ],
    "bullet",
)

doc.add_heading("Pendientes técnicos antes de medir", level=2)
add_list(
    doc,
    [
        "Agregar OpenMP al sistema de construcción y producir binarios serial y paralelo comparables.",
        "Separar temporizadores de actualización, renderizado y cuadro completo.",
        "Permitir configurar N, M, T y seed sin recompilar; documentar valores predeterminados.",
        "Registrar resultados en CSV o en un formato tabular para evitar transcripción manual.",
        "Agregar un modo de benchmark sin renderizado para aislar el costo computacional de la simulación.",
        "Validar equivalencia entre variantes con semilla y paso de tiempo fijos.",
        "Decidir si OVNI, partículas y planetas se incorporarán antes de congelar el alcance experimental.",
    ],
    "bullet",
)

doc.add_heading("Metodología experimental", level=1)
doc.add_heading("Preguntas e hipótesis", level=2)
add_labeled(doc, "Pregunta 1. ", "¿A partir de qué tamaño de escena el trabajo útil compensa el overhead de OpenMP?")
add_labeled(doc, "Pregunta 2. ", "¿Cómo cambia el speedup cuando se aumenta T desde un hilo hasta la cantidad de hilos lógicos disponibles?")
add_labeled(doc, "Pregunta 3. ", "¿La mejora de la actualización se refleja en el tiempo total por cuadro o el renderizado se convierte en el cuello de botella?")
add_labeled(doc, "Hipótesis. ", "La versión paralela tendrá poco o ningún beneficio en escenas pequeñas, pero reducirá el tiempo de actualización en escenas grandes. El speedup del cuadro completo será menor que el speedup de la actualización porque SDL2 y OpenGL permanecen secuenciales. Para el estado actual, schedule(static) debería producir menor overhead que dynamic debido a que cada entidad ejecuta una cantidad similar de operaciones.")

doc.add_heading("Variantes a comparar", level=2)
add_table(
    doc,
    ["Variante", "Descripción", "Propósito"],
    [
        ["SEQ", "Compilación sin directivas OpenMP activas", "Línea base funcional y de tiempo"],
        ["OMP-1", "Mismo código OpenMP ejecutado con T = 1", "Estimar overhead del runtime"],
        ["OMP-p", "OpenMP con T = 2, 4, 8, ... hasta el límite del equipo", "Medir escalabilidad"],
        ["STATIC", "Distribución estática de iteraciones", "Caso principal para carga uniforme"],
        ["DYNAMIC", "Distribución dinámica con chunk documentado", "Comprobar costo y balance si la carga cambia"],
    ],
    [1.05, 3.05, 2.15],
)

doc.add_heading("Métricas", level=2)
add_body(
    doc,
    "Sea T1 el tiempo de la versión secuencial para una configuración y Tp el tiempo de la versión paralela con p hilos. El speedup indica cuántas "
    "veces es más rápida la ejecución paralela; la eficiencia expresa qué fracción del paralelismo ideal se aprovecha. La mejora porcentual facilita "
    "comunicar la reducción de tiempo. Las tres métricas se calcularán tanto para actualización como para el tiempo total cuando corresponda."
)
add_equation(doc, "Speedup: S(p) = T1 / Tp")
add_equation(doc, "Eficiencia: E(p) = [S(p) / p] x 100 %")
add_equation(doc, "Mejora: I(p) = [(T1 - Tp) / T1] x 100 %")
add_equation(doc, "FPS promedio = cuadros renderizados / tiempo observado")
add_body(
    doc,
    "Se reportarán media, mediana, mínimo, máximo y desviación estándar de las repeticiones. La mediana será la medida principal cuando existan "
    "valores atípicos causados por procesos del sistema. No se mezclará el FPS limitado por VSync con el FPS libre; para benchmarking se usará VSync "
    "desactivado y se registrará cualquier límite adicional del controlador gráfico."
)

doc.add_heading("Factores y configuraciones", level=2)
add_table(
    doc,
    ["Factor", "Valores iniciales propuestos", "Control"],
    [
        ["N: vacas", "20, 100, 500, 1 000, 5 000", "Misma semilla y modelo"],
        ["M: estrellas", "180, 1 000, 5 000, 20 000", "Separar pruebas de N y M"],
        ["T: hilos", "1, 2, 4, 8, ... hasta hilos lógicos", "No exceder el límite sin justificar"],
        ["Scheduling", "static; dynamic con chunk documentado", "Mismo trabajo y compilación"],
        ["Repeticiones", "Mínimo 10 por configuración", "Mismo número para todas"],
        ["VSync", "Desactivado en pruebas de FPS", "Confirmar valor en cada sesión"],
        ["Optimización", "Release, por ejemplo -O2 o /O2", "Mismas banderas en SEQ y OMP"],
    ],
    [1.25, 2.75, 2.25],
)
add_body(
    doc,
    "Los valores son una matriz inicial y pueden reducirse si el renderizado se vuelve impráctico. Para aislar OpenMP debe existir una prueba de "
    "actualización sin llamadas de dibujo, con una cantidad de pasos fija. La prueba interactiva se conserva para FPS y experiencia visual, pero no "
    "sustituye la medición aislada del kernel de simulación."
)

doc.add_heading("Procedimiento de medición", level=2)
add_list(
    doc,
    [
        "Registrar procesador, núcleos físicos, hilos lógicos, memoria, sistema operativo, compilador y versión de OpenMP.",
        "Cerrar aplicaciones no necesarias, conectar el equipo a energía y conservar el mismo perfil de rendimiento.",
        "Compilar SEQ y OMP en modo Release con las mismas opciones, excepto las necesarias para OpenMP.",
        "Ejecutar una fase de calentamiento que no se incluya en los resultados.",
        "Fijar N, M, seed, cantidad de pasos y T; desactivar VSync para las pruebas de FPS.",
        "Medir primero el kernel de actualización con omp_get_wtime y después el cuadro completo.",
        "Ejecutar al menos diez repeticiones por configuración, alternando el orden de variantes cuando sea posible.",
        "Guardar cada observación individual, no solo el promedio, y conservar fecha y equipo de origen.",
        "Calcular estadísticas, speedup y eficiencia mediante una hoja o script reproducible.",
        "Repetir cualquier dato atípico únicamente si existe una causa documentada; no descartar resultados de manera selectiva.",
    ],
    "decimal",
)

doc.add_heading("Validación de corrección", level=2)
add_body(
    doc,
    "La versión paralela debe evaluarse primero con un dt fijo y una semilla fija. Después de una cantidad conocida de pasos se compararán posiciones, "
    "velocidades, giros y brillo con la versión secuencial. Debido a operaciones de punto flotante, se utilizará una tolerancia documentada, por "
    "ejemplo 10^-5 para atributos que siguen el mismo orden de operaciones. También se comprobará que todas las entidades permanezcan dentro de los "
    "límites, que no existan valores NaN y que el programa cierre sin errores para cada T."
)
add_body(
    doc,
    "Si el sistema incorpora reducciones, decisiones globales o generación aleatoria dentro de regiones paralelas, la comparación debe ampliarse. "
    "Una ejecución visualmente correcta no demuestra ausencia de carreras; se recomienda activar advertencias del compilador y, si el entorno lo "
    "permite, utilizar un detector de carreras o pruebas repetidas con diferentes cantidades de hilos."
)

doc.add_heading("Amenazas a la validez", level=2)
add_list(
    doc,
    [
        "Sincronización vertical o límites del controlador que oculten diferencias reales de tiempo.",
        "Frecuencia dinámica, temperatura y ahorro de energía que cambien durante las pruebas.",
        "Procesos del sistema operativo y aplicaciones en segundo plano.",
        "Comparación de binarios con opciones de compilación distintas.",
        "Escenas demasiado pequeñas, en las que el overhead domina, o demasiado grandes, en las que el renderizado domina.",
        "Uso de una sola computadora, que limita la generalización a otras arquitecturas.",
        "Medición del cuadro completo sin separar simulación y presentación.",
        "Cambios funcionales entre SEQ y OMP que vuelvan injusta la comparación.",
    ],
    "bullet",
)

doc.add_heading("Resultados", level=1)
doc.add_heading("Entorno de ejecución", level=2)
add_placeholder(
    doc,
    "Espacio reservado para el entorno",
    "Completar una fila por computadora con la plantilla del Apéndice A. Indicar cuál equipo se utilizará para las gráficas principales y conservar "
    "los resultados de los demás como replicación o comparación secundaria."
)
add_table(
    doc,
    ["Equipo", "CPU / núcleos / hilos", "RAM", "SO", "Compilador", "OpenMP"],
    [["A", "Por registrar", "Por registrar", "Por registrar", "Por registrar", "Por registrar"],
     ["B", "Por registrar", "Por registrar", "Por registrar", "Por registrar", "Por registrar"],
     ["C", "Por registrar", "Por registrar", "Por registrar", "Por registrar", "Por registrar"]],
    [0.7, 1.6, 0.7, 1.0, 1.25, 1.0],
)

doc.add_heading("Corrección funcional", level=2)
add_placeholder(
    doc,
    "Espacio reservado para validación",
    "Reportar el número de pasos, dt, seed, tolerancia máxima observada y resultado de las pruebas de límites. Incluir cualquier diferencia entre SEQ "
    "y OMP antes de presentar rendimiento."
)

doc.add_heading("Tiempos de actualización", level=2)
add_table(
    doc,
    ["N / M", "T", "Mediana Tp (ms)", "Desv. est.", "Speedup", "Eficiencia"],
    [
        ["Por definir", "1", "Pendiente", "Pendiente", "1.00", "100 %"],
        ["Por definir", "2", "Pendiente", "Pendiente", "Pendiente", "Pendiente"],
        ["Por definir", "4", "Pendiente", "Pendiente", "Pendiente", "Pendiente"],
        ["Por definir", "8", "Pendiente", "Pendiente", "Pendiente", "Pendiente"],
    ],
    [1.0, 0.55, 1.35, 1.05, 1.1, 1.2],
)
add_placeholder(
    doc,
    "Figura 1 por insertar",
    "Gráfica de tiempo de actualización contra cantidad de hilos, con una serie por tamaño del problema. Incluir unidades, leyenda y barras de error."
)
add_placeholder(
    doc,
    "Figura 2 por insertar",
    "Gráfica de speedup contra cantidad de hilos. Agregar una línea ideal S(p) = p para facilitar la comparación."
)

doc.add_heading("Tiempo total y FPS", level=2)
add_table(
    doc,
    ["Configuración", "SEQ FPS", "OMP FPS", "Mejora FPS", "Tiempo cuadro SEQ", "Tiempo cuadro OMP"],
    [["Escena pequeña", "Pendiente", "Pendiente", "Pendiente", "Pendiente", "Pendiente"],
     ["Escena mediana", "Pendiente", "Pendiente", "Pendiente", "Pendiente", "Pendiente"],
     ["Escena grande", "Pendiente", "Pendiente", "Pendiente", "Pendiente", "Pendiente"]],
    [1.25, 0.85, 0.85, 1.05, 1.15, 1.15],
)
add_placeholder(
    doc,
    "Figura 3 por insertar",
    "Comparación de FPS o tiempo por cuadro entre SEQ y la mejor configuración OMP. Indicar explícitamente que VSync estuvo desactivado."
)

doc.add_heading("Comparación de scheduling", level=2)
add_table(
    doc,
    ["T", "static", "dynamic (chunk)", "guided", "Mejor opción", "Observación"],
    [["2", "Pendiente", "Pendiente", "Opcional", "Pendiente", ""],
     ["4", "Pendiente", "Pendiente", "Opcional", "Pendiente", ""],
     ["8", "Pendiente", "Pendiente", "Opcional", "Pendiente", ""]],
    [0.55, 1.0, 1.2, 0.85, 1.1, 1.55],
)

doc.add_heading("Análisis y discusión", level=1)
add_placeholder(
    doc,
    "Espacio reservado para la discusión",
    "Interpretar los resultados sin limitarse a repetir tablas. Explicar el punto en que OpenMP comienza a ser útil, el T con mejor compromiso, la "
    "caída de eficiencia, la diferencia entre tiempo de actualización y tiempo total, y el efecto del scheduling. Relacionar los hallazgos con la "
    "Ley de Amdahl, el tamaño del problema, el overhead y las características del hardware."
)
add_placeholder(
    doc,
    "Preguntas que debe responder el análisis",
    "¿La hipótesis se sostuvo? ¿El resultado fue consistente entre repeticiones y equipos? ¿El renderizado limitó el speedup visible? ¿Hubo una "
    "configuración con más hilos pero peor tiempo? ¿Qué parte del código debería optimizarse después?"
)

doc.add_heading("Conclusiones", level=1)
add_placeholder(
    doc,
    "Espacio reservado para conclusiones finales",
    "Redactar de tres a cinco conclusiones basadas directamente en datos. Cada conclusión debe indicar el tamaño del problema y la configuración que "
    "la respalda. Evitar afirmar que OpenMP siempre mejora el rendimiento o presentar como conclusión una característica de diseño que no fue medida."
)

doc.add_heading("Recomendaciones", level=1)
add_body(
    doc,
    "Antes de ejecutar la campaña definitiva se recomienda congelar el alcance funcional y conservar dos binarios comparables. Las mediciones de "
    "actualización deben ser independientes del renderizado, mientras que el FPS se utiliza como métrica complementaria de la experiencia completa. "
    "También se recomienda iniciar con schedule(static), usar default(none), mantener la generación aleatoria fuera de regiones compartidas y evitar "
    "cambiar el tamaño de los vectores mientras se actualizan en paralelo."
)
add_body(
    doc,
    "Si los resultados muestran que el renderizado domina, la siguiente etapa no debería consistir en agregar más hilos de CPU a los mismos ciclos, "
    "sino en reducir llamadas de dibujo, utilizar instancing o modernizar la transferencia de geometría. Esa optimización pertenece a una fase "
    "posterior y debe reportarse por separado para no atribuir a OpenMP mejoras que provengan de cambios en OpenGL."
)
add_placeholder(
    doc,
    "Espacio reservado para recomendaciones derivadas de resultados",
    "Agregar recomendaciones específicas sobre T, scheduling, tamaño mínimo útil y configuración de hardware una vez finalizado el análisis."
)

doc.add_heading("Apéndices", level=1)
doc.add_heading("Apéndice A. Registro de hardware y software", level=2)
add_table(
    doc,
    ["Campo", "Valor por completar"],
    [
        ["Integrante / identificador del equipo", ""],
        ["Procesador y frecuencia reportada", ""],
        ["Núcleos físicos / hilos lógicos", ""],
        ["Memoria RAM", ""],
        ["GPU y controlador", ""],
        ["Sistema operativo y versión", ""],
        ["Compilador y versión", ""],
        ["Versión o macro _OPENMP", ""],
        ["Opciones de compilación", ""],
        ["Perfil de energía", ""],
        ["Fecha y hora de la prueba", ""],
    ],
    [2.65, 3.6],
)

doc.add_heading("Apéndice B. Parámetros y controles", level=2)
add_table(
    doc,
    ["Entrada o control", "Uso"],
    [
        ["Primer argumento numérico", "Cantidad de vacas N; el valor predeterminado actual es 20"],
        ["Ruta no numérica", "Ruta al modelo OBJ; por defecto assets/models/cow.obj"],
        ["Espacio", "Pausar o continuar la simulación"],
        ["W", "Alternar entre relleno y wireframe"],
        ["C", "Activar o desactivar face culling"],
        ["V", "Activar o desactivar VSync"],
        ["Esc", "Cerrar la aplicación"],
        ["OMP_NUM_THREADS o parámetro T", "Fijar la cantidad de hilos de OpenMP"],
        ["seed", "Reproducir la distribución inicial; debe exponerse antes de medir"],
    ],
    [2.15, 4.1],
)

doc.add_heading("Apéndice C. Pseudocódigo de benchmark", level=2)
add_code(
    doc,
    [
        "crear_estado(N, M, seed)",
        "repetir calentamiento:",
        "    actualizar(dt)",
        "inicio = omp_get_wtime()",
        "repetir pasos_medidos:",
        "    actualizar(dt)",
        "fin = omp_get_wtime()",
        "tiempo_por_paso = (fin - inicio) / pasos_medidos",
        "guardar(N, M, T, scheduling, repetición, tiempo_por_paso)",
    ],
)
add_body(
    doc,
    "El estado debe reconstruirse con la misma semilla antes de cada repetición o restaurarse desde una copia equivalente. El temporizador no debe "
    "incluir carga del modelo, creación de ventana, escritura de archivos ni calentamiento. Para el benchmark de cuadro completo se utiliza un bloque "
    "separado que sí incorpora actualización y renderizado, pero mantiene fuera la inicialización."
)

doc.add_heading("Apéndice D. Lista de verificación para la entrega", level=2)
add_list(
    doc,
    [
        "El código fuente compila desde cero siguiendo las instrucciones del README.",
        "SEQ y OMP producen resultados equivalentes para una prueba determinista.",
        "El repositorio identifica claramente la directiva y las variables compartidas y privadas.",
        "Las tablas contienen unidades, repeticiones y estadística descriptiva.",
        "Las gráficas tienen título, ejes, leyenda y referencia en el texto.",
        "Las conclusiones citan valores concretos y no repiten únicamente la metodología.",
        "Se incluyen hardware, compilador, opciones de optimización, T, scheduling, N, M y seed.",
        "No se incluyen fuentes inventadas ni documentos proporcionados por el curso como bibliografía externa.",
        "El índice y la numeración de páginas fueron actualizados antes de entregar.",
    ],
    "bullet",
)

doc.add_heading("Bibliografía", level=1)
add_reference(
    doc,
    "Amdahl, G. M. (1967). Validity of the single processor approach to achieving large scale computing capabilities. En AFIPS Spring Joint Computer Conference, 30, 483-485.",
    "https://doi.org/10.1145/1465482.1465560",
)
add_reference(
    doc,
    "Dagum, L., y Menon, R. (1998). OpenMP: An industry standard API for shared-memory programming. IEEE Computational Science & Engineering, 5(1), 46-55.",
    "https://doi.org/10.1109/99.660313",
)
add_reference(
    doc,
    "Flynn, M. J. (1972). Some computer organizations and their effectiveness. IEEE Transactions on Computers, C-21(9), 948-960.",
    "https://doi.org/10.1109/TC.1972.5009071",
)
add_reference(
    doc,
    "Gustafson, J. L. (1988). Reevaluating Amdahl's Law. Communications of the ACM, 31(5), 532-533.",
    "https://doi.org/10.1145/42411.42415",
)
add_reference(
    doc,
    "Khronos Group. (2022). OpenGL 4.6 Core Profile Specification.",
    "https://registry.khronos.org/OpenGL/specs/gl/glspec46.core.pdf",
)
add_reference(
    doc,
    "OpenMP Architecture Review Board. (2021). OpenMP Application Programming Interface, Version 5.2.",
    "https://www.openmp.org/spec-html/5.2/openmp.html",
)
add_reference(
    doc,
    "Reynolds, C. W. (1987). Flocks, herds and schools: A distributed behavioral model. ACM SIGGRAPH Computer Graphics, 21(4), 25-34.",
    "https://doi.org/10.1145/37402.37406",
)
add_reference(
    doc,
    "SDL contributors. (s. f.). FAQ: Development. SDL2 Wiki. Recuperado el 20 de agosto de 2026 de",
    "https://wiki.libsdl.org/SDL2/FAQDevelopment",
)
add_reference(
    doc,
    "Williams, S., Waterman, A., y Patterson, D. (2009). Roofline: An insightful visual performance model for multicore architectures. Communications of the ACM, 52(4), 65-76.",
    "https://doi.org/10.1145/1498765.1498785",
)

# Evita encabezados demasiado pegados al texto anterior y fuerza saltos en
# tres puntos donde las tablas/listas largas producirían cortes poco naturales.
for paragraph in doc.paragraphs:
    if paragraph.style.name == "Heading 1" and paragraph.text.strip() != "Índice":
        paragraph.paragraph_format.space_before = Pt(12)
        paragraph.paragraph_format.keep_with_next = True
    if paragraph.text.strip() in {
        "Parámetros del problema",
        "Amenazas a la validez",
    }:
        paragraph.paragraph_format.page_break_before = True

enable_field_updates(doc)
doc.core_properties.title = "Informe Proyecto 1 - ZipZip Espacial"
doc.core_properties.subject = "Paralelización de una simulación gráfica con OpenMP"
doc.core_properties.keywords = "OpenMP, C++, SDL2, OpenGL, paralelismo, speedup, eficiencia"

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(OUTPUT)
print(OUTPUT)
