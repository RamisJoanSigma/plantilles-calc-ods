#!/usr/bin/env python3
"""
Plantilles ODS per a LibreOffice Calc:
1. Checklist amb recompte Sí/No
2. Factura amb fórmules
"""

from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell, TableColumn
from odf.text import P
from odf.style import (
    Style, TableCellProperties, ParagraphProperties, TextProperties,
    TableColumnProperties, TableRowProperties
)
import os

def make_cell(value=None, formula=None, style_name=None, valuetype=None, currency=None):
    """Crea una cel·la amb valor o fórmula."""
    cell = TableCell()
    if style_name:
        cell.setAttribute("stylename", style_name)
    if formula:
        cell.setAttribute("formula", f"of:={formula}")
        if valuetype:
            cell.setAttribute("valuetype", valuetype)
        if valuetype == "float":
            cell.setAttribute("value", "0")
        elif valuetype == "currency":
            cell.setAttribute("value", "0")
            if currency:
                cell.setAttribute("currency", currency)
    elif value is not None:
        if isinstance(value, (int, float)):
            cell.setAttribute("valuetype", "float")
            cell.setAttribute("value", str(value))
            cell.addElement(P(text=str(value)))
        elif isinstance(value, bool):
            cell.setAttribute("valuetype", "boolean")
            cell.setAttribute("booleanvalue", "true" if value else "false")
            cell.addElement(P(text="TRUE" if value else "FALSE"))
        else:
            cell.setAttribute("valuetype", "string")
            cell.addElement(P(text=str(value)))
    else:
        cell.setAttribute("valuetype", "string")
        cell.addElement(P(text=""))
    return cell

def make_text_cell(text, style_name=None):
    cell = TableCell(valuetype="string")
    if style_name:
        cell.setAttribute("stylename", style_name)
    cell.addElement(P(text=str(text) if text is not None else ""))
    return cell

def make_formula_cell(formula, style_name=None, valuetype="float"):
    cell = TableCell()
    if style_name:
        cell.setAttribute("stylename", style_name)
    cell.setAttribute("formula", f"of:={formula}")
    cell.setAttribute("valuetype", valuetype)
    if valuetype == "float":
        cell.setAttribute("value", "0")
    cell.addElement(P(text="0"))
    return cell

def create_checklist():
    doc = OpenDocumentSpreadsheet()

    # Estils
    style_title = Style(name="Title", family="table-cell")
    style_title.addElement(ParagraphProperties(textalign="center"))
    style_title.addElement(TextProperties(fontsize="16pt", fontweight="bold", color="#1a5276"))
    doc.styles.addElement(style_title)

    style_header = Style(name="Header", family="table-cell")
    style_header.addElement(ParagraphProperties(textalign="center"))
    style_header.addElement(TextProperties(fontsize="11pt", fontweight="bold", color="#ffffff"))
    style_header.addElement(TableCellProperties(backgroundcolor="#2e86ab", border="0.5pt solid #1a5276"))
    doc.styles.addElement(style_header)

    style_label = Style(name="Label", family="table-cell")
    style_label.addElement(TextProperties(fontsize="11pt", fontweight="bold"))
    style_label.addElement(TableCellProperties(backgroundcolor="#d6eaf8", border="0.5pt solid #aed6f1"))
    doc.styles.addElement(style_label)

    style_count = Style(name="Count", family="table-cell")
    style_count.addElement(ParagraphProperties(textalign="center"))
    style_count.addElement(TextProperties(fontsize="14pt", fontweight="bold", color="#1a5276"))
    style_count.addElement(TableCellProperties(backgroundcolor="#eaf2f8", border="1pt solid #2e86ab"))
    doc.styles.addElement(style_count)

    style_item = Style(name="Item", family="table-cell")
    style_item.addElement(TableCellProperties(border="0.5pt solid #d5d8dc"))
    doc.styles.addElement(style_item)

    style_yesno = Style(name="YesNo", family="table-cell")
    style_yesno.addElement(ParagraphProperties(textalign="center"))
    style_yesno.addElement(TableCellProperties(border="0.5pt solid #d5d8dc"))
    doc.styles.addElement(style_yesno)

    style_note = Style(name="Note", family="table-cell")
    style_note.addElement(TextProperties(fontsize="9pt", fontstyle="italic", color="#5d6d7e"))
    doc.styles.addElement(style_note)

    # Full
    table = Table(name="Checklist")

    # Columnes
    for _ in range(4):
        table.addElement(TableColumn())

    # Títol
    row = TableRow()
    cell = make_text_cell("CHECKLIST DE TASQUES", "Title")
    cell.setAttribute("numbercolumnsspanned", "4")
    row.addElement(cell)
    for _ in range(3):
        row.addElement(TableCell())
    table.addElement(row)

    # Espai
    table.addElement(TableRow())

    # Resum de recomptes
    row = TableRow()
    row.addElement(make_text_cell("RESUM", "Header"))
    row.addElement(make_text_cell("Total elements", "Header"))
    row.addElement(make_text_cell("Fets (Sí)", "Header"))
    row.addElement(make_text_cell("Pendents (No)", "Header"))
    table.addElement(row)

    # Fórmules de recompte (files 5 a 24 = 20 elements possibles)
    # TOTAL = COUNTA de les descripcions
    # Fets = COUNTIF Sí
    # Pendents = COUNTIF No
    row = TableRow()
    row.addElement(make_text_cell("", "Label"))
    # Total: compta cel·les no buides a B5:B24
    row.addElement(make_formula_cell('COUNTA(B5:B24)', "Count"))
    # Fets: COUNTIF C5:C24;"Sí"
    row.addElement(make_formula_cell('COUNTIF(C5:C24;"Sí")', "Count"))
    # Pendents
    row.addElement(make_formula_cell('COUNTIF(C5:C24;"No")', "Count"))
    table.addElement(row)

    # Espai + nota
    row = TableRow()
    cell = make_text_cell("Instruccions: a la columna «Fet?» selecciona «Sí» o «No» de la llista desplegable. Els recomptes s'actualitzen automàticament.", "Note")
    cell.setAttribute("numbercolumnsspanned", "4")
    row.addElement(cell)
    for _ in range(3):
        row.addElement(TableCell())
    table.addElement(row)

    # Capçalera de la llista
    row = TableRow()
    row.addElement(make_text_cell("#", "Header"))
    row.addElement(make_text_cell("Descripció de la tasca", "Header"))
    row.addElement(make_text_cell("Fet?", "Header"))
    row.addElement(make_text_cell("Notes", "Header"))
    table.addElement(row)

    # 20 files d'exemple (algunes amb dades de mostra)
    sample_tasks = [
        ("Revisar correu electrònic", "Sí", ""),
        ("Preparar presentació reunió", "No", "Pendents diapositives 5-8"),
        ("Enviar informe setmanal", "Sí", ""),
        ("Actualitzar documentació projecte", "No", ""),
        ("Trucar al proveïdor", "No", "Demanar pressupost"),
        ("Revisar pressupost Q4", "Sí", ""),
        ("Preparar material formació", "No", ""),
        ("Validar disseny web", "No", "Feedback dissenyador"),
        ("Backup de dades", "Sí", ""),
        ("Reunió equip dimecres", "No", ""),
    ]

    for i in range(20):
        row = TableRow()
        row.addElement(make_text_cell(str(i + 1), "Item"))
        if i < len(sample_tasks):
            desc, fet, notes = sample_tasks[i]
            row.addElement(make_text_cell(desc, "Item"))
            row.addElement(make_text_cell(fet, "YesNo"))
            row.addElement(make_text_cell(notes, "Item"))
        else:
            row.addElement(make_text_cell("", "Item"))
            row.addElement(make_text_cell("", "YesNo"))
            row.addElement(make_text_cell("", "Item"))
        table.addElement(row)

    # Nota final sobre validació de dades
    table.addElement(TableRow())
    row = TableRow()
    cell = make_text_cell(
        "Nota: Per activar la llista desplegable Sí/No a la columna C: selecciona el rang C5:C24 → Dades → Validesa → Criteri: Llista → Entrades: Sí;No",
        "Note"
    )
    cell.setAttribute("numbercolumnsspanned", "4")
    row.addElement(cell)
    for _ in range(3):
        row.addElement(TableCell())
    table.addElement(row)

    doc.spreadsheet.addElement(table)
    path = "/home/workdir/artifacts/checklist.ods"
    doc.save(path)
    print(f"Creat: {path}")
    return path

def create_factura():
    doc = OpenDocumentSpreadsheet()

    # Estils
    style_title = Style(name="Title", family="table-cell")
    style_title.addElement(ParagraphProperties(textalign="center"))
    style_title.addElement(TextProperties(fontsize="20pt", fontweight="bold", color="#1a5276"))
    doc.styles.addElement(style_title)

    style_header = Style(name="Header", family="table-cell")
    style_header.addElement(ParagraphProperties(textalign="center"))
    style_header.addElement(TextProperties(fontsize="9pt", fontweight="bold", color="#ffffff"))
    style_header.addElement(TableCellProperties(backgroundcolor="#1a5276", border="0.4pt solid #154360"))
    doc.styles.addElement(style_header)

    style_label = Style(name="Label", family="table-cell")
    style_label.addElement(TextProperties(fontsize="10pt", fontweight="bold"))
    style_label.addElement(TableCellProperties(backgroundcolor="#d4e6f1"))
    doc.styles.addElement(style_label)

    style_normal = Style(name="Normal", family="table-cell")
    style_normal.addElement(TableCellProperties(border="0.4pt solid #aeb6bf"))
    doc.styles.addElement(style_normal)

    style_money = Style(name="Money", family="table-cell")
    style_money.addElement(ParagraphProperties(textalign="right"))
    style_money.addElement(TableCellProperties(border="0.4pt solid #aeb6bf"))
    doc.styles.addElement(style_money)

    style_total = Style(name="Total", family="table-cell")
    style_total.addElement(ParagraphProperties(textalign="right"))
    style_total.addElement(TextProperties(fontsize="12pt", fontweight="bold", color="#1a5276"))
    style_total.addElement(TableCellProperties(backgroundcolor="#d5f5e3", border="1pt solid #1a5276"))
    doc.styles.addElement(style_total)

    style_section = Style(name="Section", family="table-cell")
    style_section.addElement(TextProperties(fontsize="11pt", fontweight="bold", color="#ffffff"))
    style_section.addElement(TableCellProperties(backgroundcolor="#2e86ab"))
    doc.styles.addElement(style_section)

    style_small = Style(name="Small", family="table-cell")
    style_small.addElement(TextProperties(fontsize="8pt", color="#5d6d7e"))
    doc.styles.addElement(style_small)

    table = Table(name="Factura")

    # Columnes: A Codi, B Article, C Unitats, D Preu, E Subtotal, F %Desc, G TotalDesc, H %IVA, I TotalIVA, J TotalIVA
    for _ in range(11):
        table.addElement(TableColumn())

    # === CAPÇALERA EMPRESA ===
    row = TableRow()
    cell = make_text_cell("NOM EMPRESA", "Title")
    cell.setAttribute("numbercolumnsspanned", "5")
    row.addElement(cell)
    for _ in range(4):
        row.addElement(TableCell())
    cell = make_text_cell("Adreça de l'empresa", "Normal")
    cell.setAttribute("numbercolumnsspanned", "3")
    row.addElement(cell)
    for _ in range(2):
        row.addElement(TableCell())
    table.addElement(row)

    row = TableRow()
    cell = make_text_cell("FACTURA", "Title")
    cell.setAttribute("numbercolumnsspanned", "5")
    row.addElement(cell)
    for _ in range(4):
        row.addElement(TableCell())
    cell = make_text_cell("Ciutat · CP", "Normal")
    cell.setAttribute("numbercolumnsspanned", "3")
    row.addElement(cell)
    for _ in range(2):
        row.addElement(TableCell())
    table.addElement(row)

    row = TableRow()
    row.addElement(make_text_cell("Número:", "Label"))
    row.addElement(make_text_cell("FAC-2026-001", "Normal"))
    row.addElement(make_text_cell("NIF:", "Label"))
    row.addElement(make_text_cell("B12345678", "Normal"))
    for _ in range(7):
        row.addElement(TableCell())
    table.addElement(row)

    row = TableRow()
    row.addElement(make_text_cell("Data:", "Label"))
    row.addElement(make_text_cell("22/11/2026", "Normal"))
    for _ in range(9):
        row.addElement(TableCell())
    table.addElement(row)

    table.addElement(TableRow())  # espai

    # === DADES CLIENT ===
    row = TableRow()
    row.addElement(make_text_cell("Client:", "Label"))
    cell = make_text_cell("Nom del client", "Normal")
    cell.setAttribute("numbercolumnsspanned", "3")
    row.addElement(cell)
    for _ in range(2):
        row.addElement(TableCell())
    row.addElement(make_text_cell("Comentaris:", "Label"))
    cell = make_text_cell("", "Normal")
    cell.setAttribute("numbercolumnsspanned", "3")
    row.addElement(cell)
    for _ in range(2):
        row.addElement(TableCell())
    table.addElement(row)

    row = TableRow()
    row.addElement(make_text_cell("Domicili:", "Label"))
    cell = make_text_cell("Carrer, número", "Normal")
    cell.setAttribute("numbercolumnsspanned", "3")
    row.addElement(cell)
    for _ in range(2):
        row.addElement(TableCell())
    for _ in range(5):
        row.addElement(TableCell())
    table.addElement(row)

    row = TableRow()
    row.addElement(make_text_cell("Ciutat:", "Label"))
    row.addElement(make_text_cell("Ciutat client", "Normal"))
    row.addElement(make_text_cell("NIF:", "Label"))
    row.addElement(make_text_cell("12345678A", "Normal"))
    for _ in range(7):
        row.addElement(TableCell())
    table.addElement(row)

    table.addElement(TableRow())  # espai

    # === CAPÇALERA LÍNIES ===
    # Files de dades començaran a la fila 12 (1-based: fila 12)
    headers = [
        "Codi", "Article", "Unitats", "Preu Unitari", "Subtotal [1]",
        "% Descompte", "Total descompte [2]", "% IVA", "Total IVA [3]", "Total amb IVA [4]"
    ]
    row = TableRow()
    for h in headers:
        row.addElement(make_text_cell(h, "Header"))
    # una columna extra buida
    row.addElement(TableCell())
    table.addElement(row)

    # Dades d'exemple (4 productes + 6 files buides per a fórmules)
    # Columnes: A=Codi, B=Article, C=Unitats, D=Preu, E=Subtotal, F=%Desc, G=TotalDesc, H=%IVA, I=TotalIVA, J=Total
    # Fórmules (fila r, 1-based):
    # [1] Subtotal E = C * D
    # [2] Total descompte G = E * F
    # [3] Total IVA I = (E - G) * H
    # [4] Total amb IVA J = E - G + I

    products = [
        ("123935", "Abric talla S", 2, 14.99, 0.05, 0.21),
        ("123936", "Sabates talla 36", 1, 29.15, 0.03, 0.21),
        ("123937", "Llibre de text", 3, 25.66, 0.10, 0.10),
        ("123938", "Patates", 5, 0.85, 0.00, 0.04),
    ]

    # Les files de productes seran files 12-21 (10 files)
    DATA_START = 12  # 1-based row number of first product
    NUM_ROWS = 10

    for i in range(NUM_ROWS):
        row = TableRow()
        r = DATA_START + i  # 1-based
        if i < len(products):
            codi, article, unitats, preu, desc, iva = products[i]
            row.addElement(make_text_cell(codi, "Normal"))
            row.addElement(make_text_cell(article, "Normal"))
            row.addElement(make_cell(unitats, style_name="Normal", valuetype="float"))
            row.addElement(make_cell(preu, style_name="Money", valuetype="float"))
        else:
            for _ in range(4):
                row.addElement(make_text_cell("", "Normal"))

        # E: Subtotal = C * D
        row.addElement(make_formula_cell(f"C{r}*D{r}", "Money"))
        # F: % Descompte (valor o buit)
        if i < len(products):
            row.addElement(make_cell(products[i][4], style_name="Money", valuetype="float"))
        else:
            row.addElement(make_text_cell("", "Normal"))
        # G: Total descompte = E * F
        row.addElement(make_formula_cell(f"E{r}*F{r}", "Money"))
        # H: % IVA
        if i < len(products):
            row.addElement(make_cell(products[i][5], style_name="Money", valuetype="float"))
        else:
            row.addElement(make_text_cell("", "Normal"))
        # I: Total IVA = (E - G) * H
        row.addElement(make_formula_cell(f"(E{r}-G{r})*H{r}", "Money"))
        # J: Total amb IVA = E - G + I
        row.addElement(make_formula_cell(f"E{r}-G{r}+I{r}", "Money"))
        table.addElement(row)

    table.addElement(TableRow())  # espai

    # === RESUMS ===
    # Fila de totes: Import brut, Total descomptes
    # Després taula per tipus d'IVA

    # Fila labels
    row = TableRow()
    row.addElement(make_text_cell("Import brut [5]", "Label"))
    row.addElement(make_text_cell("Total descomptes [6]", "Label"))
    row.addElement(make_text_cell("", "Normal"))
    row.addElement(make_text_cell("Tipus IVA", "Header"))
    row.addElement(make_text_cell("Base Imponible [7]", "Header"))
    row.addElement(make_text_cell("Import IVA [8]", "Header"))
    for _ in range(4):
        row.addElement(TableCell())
    table.addElement(row)

    # Fila de valors globals + primer tipus IVA (4%)
    # [5] = SUM(E12:E21)
    # [6] = SUM(G12:G21)
    SUMMARY_ROW = DATA_START + NUM_ROWS + 2  # ~24

    row = TableRow()
    # Import brut
    row.addElement(make_formula_cell(f"SUM(E{DATA_START}:E{DATA_START+NUM_ROWS-1})", "Money"))
    # Total descomptes
    row.addElement(make_formula_cell(f"SUM(G{DATA_START}:G{DATA_START+NUM_ROWS-1})", "Money"))
    row.addElement(make_text_cell("", "Normal"))
    # Tipus 4%
    row.addElement(make_cell(0.04, style_name="Money", valuetype="float"))
    # Base imponible per 4%: suma de (E-G) on H=4%
    # Utilitzem SUMIF simplificat: SUMIF(H12:H21;0,04;E12:E21) - SUMIF(H12:H21;0,04;G12:G21)
    r1, r2 = DATA_START, DATA_START + NUM_ROWS - 1
    row.addElement(make_formula_cell(
        f'SUMIF(H{r1}:H{r2};D{SUMMARY_ROW};E{r1}:E{r2})-SUMIF(H{r1}:H{r2};D{SUMMARY_ROW};G{r1}:G{r2})',
        "Money"
    ))
    # Import IVA 4% = Base * tipus
    row.addElement(make_formula_cell(f"E{SUMMARY_ROW}*D{SUMMARY_ROW}", "Money"))
    for _ in range(4):
        row.addElement(TableCell())
    table.addElement(row)

    # Tipus 10%
    row = TableRow()
    for _ in range(3):
        row.addElement(make_text_cell("", "Normal"))
    row.addElement(make_cell(0.10, style_name="Money", valuetype="float"))
    row.addElement(make_formula_cell(
        f'SUMIF(H{r1}:H{r2};D{SUMMARY_ROW+1};E{r1}:E{r2})-SUMIF(H{r1}:H{r2};D{SUMMARY_ROW+1};G{r1}:G{r2})',
        "Money"
    ))
    row.addElement(make_formula_cell(f"E{SUMMARY_ROW+1}*D{SUMMARY_ROW+1}", "Money"))
    for _ in range(4):
        row.addElement(TableCell())
    table.addElement(row)

    # Tipus 21%
    row = TableRow()
    for _ in range(3):
        row.addElement(make_text_cell("", "Normal"))
    row.addElement(make_cell(0.21, style_name="Money", valuetype="float"))
    row.addElement(make_formula_cell(
        f'SUMIF(H{r1}:H{r2};D{SUMMARY_ROW+2};E{r1}:E{r2})-SUMIF(H{r1}:H{r2};D{SUMMARY_ROW+2};G{r1}:G{r2})',
        "Money"
    ))
    row.addElement(make_formula_cell(f"E{SUMMARY_ROW+2}*D{SUMMARY_ROW+2}", "Money"))
    for _ in range(4):
        row.addElement(TableCell())
    table.addElement(row)

    table.addElement(TableRow())

    # Forma de pagament i TOTAL
    row = TableRow()
    row.addElement(make_text_cell("Forma de pagament:", "Label"))
    cell = make_text_cell("Transferència bancària", "Normal")
    cell.setAttribute("numbercolumnsspanned", "2")
    row.addElement(cell)
    row.addElement(TableCell())
    row.addElement(make_text_cell("TOTAL FACTURA", "Section"))
    # Total = Import brut - descomptes + suma IVAs  (o SUM de columna J)
    total_row = SUMMARY_ROW + 4
    row.addElement(make_formula_cell(f"SUM(J{r1}:J{r2})", "Total"))
    for _ in range(4):
        row.addElement(TableCell())
    table.addElement(row)

    table.addElement(TableRow())

    # Notes de fórmules
    notes = [
        "[1] Subtotal = Unitats × Preu Unitari          →  =C{r}*D{r}",
        "[2] Total descompte = Subtotal × % Descompte   →  =E{r}*F{r}",
        "[3] Total IVA = (Subtotal − Descompte) × % IVA →  =(E{r}-G{r})*H{r}",
        "[4] Total amb IVA = Subtotal − Descompte + IVA →  =E{r}-G{r}+I{r}",
        "[5] Import brut = SUM(Subtotals)",
        "[6] Total descomptes = SUM(Total descompte)",
        "[7] Base Imponible per tipus = SUMIF(%IVA; tipus; Subtotal) − SUMIF(%IVA; tipus; Descompte)",
        "[8] Import IVA = Base Imponible × Tipus IVA",
    ]
    for n in notes:
        row = TableRow()
        cell = make_text_cell(n, "Small")
        cell.setAttribute("numbercolumnsspanned", "10")
        row.addElement(cell)
        for _ in range(9):
            row.addElement(TableCell())
        table.addElement(row)

    doc.spreadsheet.addElement(table)
    path = "/home/workdir/artifacts/factura.ods"
    doc.save(path)
    print(f"Creat: {path}")
    return path

if __name__ == "__main__":
    create_checklist()
    create_factura()
    print("Fet!")
