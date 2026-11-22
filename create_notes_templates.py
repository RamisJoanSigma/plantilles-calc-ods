#!/usr/bin/env python3
"""
Plantilles XLSX per a WPS Office Sheet / Excel:
1. Quadern de notes de professorat (amb fórmules)
2. Llista de notes finals de l'alumnat
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Border, Side, Alignment, NamedStyle, Protection
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

# Colors
BLUE_DARK = "1A5276"
BLUE_MID = "2E86AB"
BLUE_LIGHT = "D4E6F1"
GREEN = "27AE60"
GREEN_LIGHT = "D5F5E3"
RED = "C0392B"
RED_LIGHT = "FADBD8"
ORANGE = "E67E22"
YELLOW = "F9E79F"
GRAY = "F4F6F7"
WHITE = "FFFFFF"
BLACK = "1C2833"

thin = Border(
    left=Side(style="thin", color="AEB6BF"),
    right=Side(style="thin", color="AEB6BF"),
    top=Side(style="thin", color="AEB6BF"),
    bottom=Side(style="thin", color="AEB6BF"),
)
thick = Border(
    left=Side(style="medium", color=BLUE_DARK),
    right=Side(style="medium", color=BLUE_DARK),
    top=Side(style="medium", color=BLUE_DARK),
    bottom=Side(style="medium", color=BLUE_DARK),
)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=11, color=BLACK, name="Calibri"):
    return Font(name=name, bold=bold, size=size, color=color)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_align():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def header_style(cell, text, bg=BLUE_DARK, fg=WHITE, size=11):
    cell.value = text
    cell.font = font(bold=True, size=size, color=fg)
    cell.fill = fill(bg)
    cell.alignment = center()
    cell.border = thin

def label_style(cell, text, bg=BLUE_LIGHT):
    cell.value = text
    cell.font = font(bold=True, size=10)
    cell.fill = fill(bg)
    cell.alignment = left_align()
    cell.border = thin

def normal_style(cell, value=None):
    if value is not None:
        cell.value = value
    cell.font = font(size=10)
    cell.border = thin
    cell.alignment = center()

def create_quadern():
    """Quadern de notes de professorat complet amb fórmules."""
    wb = Workbook()

    # ========== FULL 1: NOTES ==========
    ws = wb.active
    ws.title = "Notes"

    # Amplades
    widths = {
        "A": 5, "B": 28, "C": 26,
        "D": 10, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10, "J": 10,
        "K": 12, "L": 12, "M": 12, "N": 12, "O": 14, "P": 12
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # --- Títol ---
    ws.merge_cells("A1:P1")
    ws["A1"] = "QUADERN DE NOTES DEL PROFESSORAT"
    ws["A1"].font = font(bold=True, size=18, color=BLUE_DARK)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 30

    # --- Capçalera assignatura / grup ---
    ws.merge_cells("A2:B2")
    label_style(ws["A2"], "Assignatura / Mòdul:")
    ws.merge_cells("C2:E2")
    normal_style(ws["C2"], "Matemàtiques 2n ESO")
    ws["C2"].alignment = left_align()

    label_style(ws["F2"], "Grup:")
    normal_style(ws["G2"], "A")
    label_style(ws["H2"], "Curs:")
    normal_style(ws["I2"], "2026-2027")
    label_style(ws["J2"], "Professor/a:")
    ws.merge_cells("K2:M2")
    normal_style(ws["K2"], "Nom del professor/a")
    ws["K2"].alignment = left_align()

    # --- Filera de pesos (configurable) ---
    ws.row_dimensions[4].height = 22
    header_style(ws["A4"], "#", BLUE_MID)
    header_style(ws["B4"], "Cognoms i nom", BLUE_MID)
    header_style(ws["C4"], "Correu electrònic", BLUE_MID)

    # Activitats amb pesos
    activities = [
        ("Examen 1", 0.15),
        ("Examen 2", 0.15),
        ("Treball", 0.10),
        ("Projecte", 0.15),
        ("Actitud", 0.10),
        ("Oral", 0.10),
        ("Pràctiques", 0.15),
        ("Recuperació", 0.10),
    ]
    # Columnes D-K = activitats (8)
    for i, (name, weight) in enumerate(activities):
        col = get_column_letter(4 + i)  # D=4
        header_style(ws[f"{col}4"], name, BLUE_MID, size=9)

    header_style(ws["L4"], "Mitjana\nponderada", GREEN, size=9)
    header_style(ws["M4"], "Nota\nfinal", GREEN, size=9)
    header_style(ws["N4"], "Qualificació", GREEN, size=9)
    header_style(ws["O4"], "Observacions", BLUE_MID, size=9)
    header_style(ws["P4"], "Falten\nnotes?", ORANGE, size=9)

    # Filera de pesos (fila 5)
    ws["A5"] = ""
    ws["B5"] = "PESOS →"
    ws["B5"].font = font(bold=True, size=9, color=BLUE_DARK)
    ws["B5"].fill = fill(YELLOW)
    ws["B5"].alignment = center()
    ws["C5"] = ""
    for i, (name, weight) in enumerate(activities):
        col = get_column_letter(4 + i)
        cell = ws[f"{col}5"]
        cell.value = weight
        cell.number_format = "0%"
        cell.font = font(bold=True, size=9)
        cell.fill = fill(YELLOW)
        cell.alignment = center()
        cell.border = thin

    # Suma de pesos (verificació)
    ws["L5"] = '=SUM(D5:K5)'
    ws["L5"].number_format = "0%"
    ws["L5"].font = font(bold=True, size=9, color=GREEN)
    ws["L5"].fill = fill(GREEN_LIGHT)
    ws["L5"].alignment = center()
    ws["L5"].border = thin
    ws["M5"] = "← ha de ser 100%"
    ws["M5"].font = font(size=8, color=RED)

    # --- Alumnes d'exemple ---
    students = [
        ("García López, Anna", "anna.garcia@escola.cat", [8.5, 7.0, 9.0, 8.0, 9.5, 8.5, 7.5, None]),
        ("Martínez Ruiz, Joan", "joan.martinez@escola.cat", [5.0, 6.5, 7.0, 6.0, 8.0, 5.5, 6.5, 7.0]),
        ("Pérez Soler, Maria", "maria.perez@escola.cat", [9.0, 8.5, 9.5, 9.0, 10.0, 9.0, 8.5, None]),
        ("Sánchez Vidal, Pau", "pau.sanchez@escola.cat", [4.0, 5.0, 3.5, 4.5, 6.0, 4.0, 5.0, 5.5]),
        ("López Ferrer, Laura", "laura.lopez@escola.cat", [7.5, 8.0, 7.0, 8.5, 8.0, 7.5, 8.0, None]),
        ("Rodríguez Camps, Marc", "marc.rodriguez@escola.cat", [6.0, 5.5, 6.5, 7.0, 7.5, 6.0, 6.5, None]),
        ("Fernández Costa, Clara", "clara.fernandez@escola.cat", [9.5, 9.0, 10.0, 9.5, 9.0, 9.5, 9.0, None]),
        ("Gómez Riera, David", "david.gomez@escola.cat", [3.5, 4.0, 5.0, 4.5, 5.5, 3.0, 4.0, 5.0]),
        ("Torres Molina, Elena", "elena.torres@escola.cat", [8.0, 7.5, 8.5, 8.0, 9.0, 8.0, 7.5, None]),
        ("Navarro Pujol, Adrià", "adria.navarro@escola.cat", [6.5, 7.0, 6.0, 7.5, 7.0, 6.5, 7.0, None]),
        ("Vila Serra, Núria", "nuria.vila@escola.cat", [None, None, None, None, None, None, None, None]),
        ("Costa Amorós, Jordi", "jordi.costa@escola.cat", [5.5, 6.0, 5.0, 6.5, 7.0, 5.5, 6.0, None]),
    ]

    DATA_START = 6  # primera fila d'alumne
    NUM_STUDENTS = 25  # files preparades

    for i in range(NUM_STUDENTS):
        row = DATA_START + i
        ws.row_dimensions[row].height = 18

        # Número
        normal_style(ws[f"A{row}"], i + 1 if i < len(students) else None)

        if i < len(students):
            name, email, notes = students[i]
            normal_style(ws[f"B{row}"], name)
            ws[f"B{row}"].alignment = left_align()
            normal_style(ws[f"C{row}"], email)
            ws[f"C{row}"].alignment = left_align()
            for j, n in enumerate(notes):
                col = get_column_letter(4 + j)
                cell = ws[f"{col}{row}"]
                if n is not None:
                    cell.value = n
                    cell.number_format = "0.00"
                normal_style(cell)
        else:
            for col in range(2, 12):
                normal_style(ws.cell(row=row, column=col))

        # Mitjana ponderada: només amb notes existents
        formula_avg = (
            f'=IF(COUNT(D{row}:K{row})=0,"",'
            f'SUMPRODUCT(D{row}:K{row},$D$5:$K$5)/'
            f'SUMIF(D{row}:K{row},"<>",$D$5:$K$5))'
        )
        cell = ws[f"L{row}"]
        cell.value = formula_avg
        cell.number_format = "0.00"
        cell.font = font(bold=True, size=10)
        cell.fill = fill(GREEN_LIGHT)
        cell.border = thin
        cell.alignment = center()

        # Nota final (arrodonida a 1 decimal, o buida)
        cell = ws[f"M{row}"]
        cell.value = f'=IF(L{row}="","",ROUND(L{row},1))'
        cell.number_format = "0.0"
        cell.font = font(bold=True, size=11, color=BLUE_DARK)
        cell.fill = fill(GREEN_LIGHT)
        cell.border = thick
        cell.alignment = center()

        # Qualificació textual
        formula_qual = (
            f'=IF(M{row}="","",'
            f'IF(M{row}<5,"Insuficient",'
            f'IF(M{row}<6,"Suficient",'
            f'IF(M{row}<7,"Bé",'
            f'IF(M{row}<9,"Notable","Excel·lent")))))'
        )
        cell = ws[f"N{row}"]
        cell.value = formula_qual
        cell.font = font(bold=True, size=9)
        cell.border = thin
        cell.alignment = center()

        # Observacions (buida)
        normal_style(ws[f"O{row}"])
        ws[f"O{row}"].alignment = left_align()

        # Falten notes? (si hi ha cel·les buides a D-K)
        cell = ws[f"P{row}"]
        cell.value = f'=IF(B{row}="","",IF(COUNTBLANK(D{row}:K{row})>0,"Sí","No"))'
        cell.font = font(size=9)
        cell.border = thin
        cell.alignment = center()

    last_data = DATA_START + NUM_STUDENTS - 1

    # --- Format condicional notes (escala de color) ---
    ws.conditional_formatting.add(
        f"D{DATA_START}:K{last_data}",
        ColorScaleRule(
            start_type="num", start_value=0, start_color="F1948A",
            mid_type="num", mid_value=5, mid_color="F9E79F",
            end_type="num", end_value=10, end_color="82E0AA"
        )
    )
    # Qualificacions
    ws.conditional_formatting.add(
        f"N{DATA_START}:N{last_data}",
        FormulaRule(formula=[f'N{DATA_START}="Insuficient"'], fill=fill(RED_LIGHT))
    )
    ws.conditional_formatting.add(
        f"N{DATA_START}:N{last_data}",
        FormulaRule(formula=[f'N{DATA_START}="Excel·lent"'], fill=fill(GREEN_LIGHT))
    )
    ws.conditional_formatting.add(
        f"N{DATA_START}:N{last_data}",
        FormulaRule(formula=[f'N{DATA_START}="Notable"'], fill=fill("D5F5E3"))
    )

    # --- Resum estadístic a sota ---
    summary_row = last_data + 2
    ws[f"A{summary_row}"] = "RESUM DE LA CLASSE"
    ws[f"A{summary_row}"].font = font(bold=True, size=12, color=WHITE)
    ws[f"A{summary_row}"].fill = fill(BLUE_DARK)
    ws.merge_cells(f"A{summary_row}:D{summary_row}")

    stats = [
        ("Mitjana de la classe", f'=AVERAGEIF(M{DATA_START}:M{last_data},">=0")'),
        ("Nota màxima", f'=MAX(M{DATA_START}:M{last_data})'),
        ("Nota mínima", f'=MIN(M{DATA_START}:M{last_data})'),
        ("Aprovats (≥5)", f'=COUNTIF(M{DATA_START}:M{last_data},">=5")'),
        ("Suspesos (<5)", f'=COUNTIF(M{DATA_START}:M{last_data},"<5")'),
        ("% Aprovats", f'=IF(COUNTA(B{DATA_START}:B{last_data})=0,"",COUNTIF(M{DATA_START}:M{last_data},">=5")/COUNTA(B{DATA_START}:B{last_data}))'),
        ("Alumnes amb notes", f'=COUNTA(B{DATA_START}:B{last_data})'),
    ]

    for idx, (label, formula) in enumerate(stats):
        r = summary_row + 1 + idx
        label_style(ws[f"A{r}"], label)
        ws.merge_cells(f"A{r}:B{r}")
        cell = ws[f"C{r}"]
        cell.value = formula
        cell.font = font(bold=True, size=11, color=BLUE_DARK)
        cell.fill = fill(GRAY)
        cell.border = thin
        cell.alignment = center()
        if "Aprovats" in label and "%" in label:
            cell.number_format = "0.0%"
        elif "Mitjana" in label or "màxima" in label or "mínima" in label:
            cell.number_format = "0.00"

    # Comentari d'ajuda
    ws["D5"].comment = Comment(
        "Modifica els pesos d'aquesta fila. Han de sumar 100%. "
        "La mitjana ponderada només té en compte les notes ja introduïdes.",
        "Sistema"
    )

    # ========== FULL 2: ESTADÍSTIQUES ==========
    ws2 = wb.create_sheet("Estadístiques")
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 14

    ws2["A1"] = "ESTADÍSTIQUES DEL GRUP"
    ws2["A1"].font = font(bold=True, size=16, color=BLUE_DARK)
    ws2.merge_cells("A1:D1")

    ws2["A3"] = "Distribució de qualificacions"
    ws2["A3"].font = font(bold=True, size=12)

    headers2 = ["Qualificació", "Nombre", "Percentatge"]
    for i, h in enumerate(headers2):
        header_style(ws2.cell(row=4, column=i+1), h, BLUE_MID)

    quals = ["Insuficient", "Suficient", "Bé", "Notable", "Excel·lent"]
    for i, q in enumerate(quals):
        r = 5 + i
        normal_style(ws2[f"A{r}"], q)
        ws2[f"A{r}"].alignment = left_align()
        ws2[f"B{r}"] = f'=COUNTIF(Notes!N{DATA_START}:N{last_data},A{r})'
        normal_style(ws2[f"B{r}"])
        ws2[f"C{r}"] = f'=IF(SUM($B$5:$B$9)=0,0,B{r}/SUM($B$5:$B$9))'
        ws2[f"C{r}"].number_format = "0.0%"
        normal_style(ws2[f"C{r}"])

    # Gràfic de barres
    chart = BarChart()
    chart.type = "col"
    chart.title = "Distribució de qualificacions"
    chart.y_axis.title = "Nombre d'alumnes"
    data = Reference(ws2, min_col=2, min_row=4, max_row=9)
    cats = Reference(ws2, min_col=1, min_row=5, max_row=9)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 15
    chart.height = 10
    ws2.add_chart(chart, "E3")

    # ========== FULL 3: INSTRUCCIONS ==========
    ws3 = wb.create_sheet("Instruccions")
    ws3.column_dimensions["A"].width = 100
    ws3["A1"] = "INSTRUCCIONS D'ÚS DEL QUADERN DE NOTES"
    ws3["A1"].font = font(bold=True, size=14, color=BLUE_DARK)

    instructions = [
        "",
        "1. FULL «Notes»",
        "   • Modifica la fila 2 amb les dades de l'assignatura, grup i professor/a.",
        "   • A la fila 5 pots canviar els pesos de cada activitat (han de sumar 100%).",
        "   • Introdueix el nom, correu i notes de cada alumne a partir de la fila 6.",
        "   • Les columnes «Mitjana ponderada», «Nota final», «Qualificació» i «Falten notes?» es calculen automàticament.",
        "   • La mitjana ponderada només té en compte les notes ja introduïdes (ignora les cel·les buides).",
        "   • Format condicional: les notes es pinten de vermell (baix) a verd (alt).",
        "",
        "2. FÓRMULES PRINCIPALS",
        "   • Mitjana ponderada = SUMPRODUCT(notes; pesos) / SUMIF(notes no buides; pesos)",
        "   • Nota final = ROUND(mitjana; 1)",
        "   • Qualificació: <5 Insuficient | <6 Suficient | <7 Bé | <9 Notable | ≥9 Excel·lent",
        "   • Resum: AVERAGEIF, MAX, MIN, COUNTIF, percentatge d'aprovats",
        "",
        "3. FULL «Estadístiques»",
        "   • Mostra la distribució de qualificacions amb recompte i percentatge.",
        "   • Inclou un gràfic de barres automàtic.",
        "",
        "4. CONSELLS",
        "   • No esborris les fórmules de les columnes L, M, N i P.",
        "   • Pots afegir més files d'alumnes copiant el format i les fórmules.",
        "   • Els pesos es poden ajustar en qualsevol moment; les mitjanes es recalculen soles.",
        "   • Compatible amb WPS Office Sheet, LibreOffice Calc i Microsoft Excel.",
    ]
    for i, line in enumerate(instructions):
        ws3[f"A{i+2}"] = line
        ws3[f"A{i+2}"].font = font(size=10, bold=line.startswith(("1.", "2.", "3.", "4.")))

    path = "/home/workdir/artifacts/quadern_notes_professorat.xlsx"
    wb.save(path)
    print(f"Creat: {path}")
    return path


def create_llista_notes_finals():
    """Llista de notes finals de l'alumnat."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Notes finals"

    widths = {"A": 5, "B": 30, "C": 28, "D": 12, "E": 12, "F": 12, "G": 12, "H": 14, "I": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Títol
    ws.merge_cells("A1:I1")
    ws["A1"] = "LLISTA DE NOTES FINALS DE L'ALUMNAT"
    ws["A1"].font = font(bold=True, size=16, color=BLUE_DARK)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 28

    # Capçalera info
    label_style(ws["A2"], "Assignatura:")
    ws.merge_cells("B2:C2")
    normal_style(ws["B2"], "Matemàtiques 2n ESO")
    ws["B2"].alignment = left_align()
    label_style(ws["D2"], "Grup:")
    normal_style(ws["E2"], "A")
    label_style(ws["F2"], "Curs:")
    normal_style(ws["G2"], "2026-2027")

    # Capçaleres columnes
    headers = [
        "#", "Nom i cognoms", "Correu electrònic",
        "Nota 1a aval.", "Nota 2a aval.", "Nota 3a aval.",
        "Nota final", "Mitjana", "Qualificació"
    ]
    for i, h in enumerate(headers):
        header_style(ws.cell(row=4, column=i+1), h, BLUE_DARK, size=10)
    ws.row_dimensions[4].height = 30

    # Dades d'exemple
    students = [
        ("García López, Anna", "anna.garcia@escola.cat", 8.0, 7.5, 8.5),
        ("Martínez Ruiz, Joan", "joan.martinez@escola.cat", 5.5, 6.0, 6.5),
        ("Pérez Soler, Maria", "maria.perez@escola.cat", 9.0, 9.5, 9.0),
        ("Sánchez Vidal, Pau", "pau.sanchez@escola.cat", 4.0, 4.5, 5.0),
        ("López Ferrer, Laura", "laura.lopez@escola.cat", 7.5, 8.0, 7.5),
        ("Rodríguez Camps, Marc", "marc.rodriguez@escola.cat", 6.0, 6.5, 7.0),
        ("Fernández Costa, Clara", "clara.fernandez@escola.cat", 9.5, 9.0, 9.5),
        ("Gómez Riera, David", "david.gomez@escola.cat", 3.5, 4.0, 4.5),
        ("Torres Molina, Elena", "elena.torres@escola.cat", 8.0, 8.5, 8.0),
        ("Navarro Pujol, Adrià", "adria.navarro@escola.cat", 6.5, 7.0, 6.5),
        ("Vila Serra, Núria", "nuria.vila@escola.cat", 7.0, 7.5, 8.0),
        ("Costa Amorós, Jordi", "jordi.costa@escola.cat", 5.0, 5.5, 6.0),
        ("Moreno Gil, Sara", "sara.moreno@escola.cat", 8.5, 9.0, 8.5),
        ("Rubio Font, Alex", "alex.rubio@escola.cat", 4.5, 5.0, 5.5),
        ("Soler Roca, Martina", "martina.soler@escola.cat", 9.0, 8.5, 9.0),
    ]

    DATA_START = 5
    NUM = 30  # files preparades

    for i in range(NUM):
        row = DATA_START + i
        normal_style(ws[f"A{row}"], i + 1 if i < len(students) else None)

        if i < len(students):
            name, email, n1, n2, n3 = students[i]
            normal_style(ws[f"B{row}"], name)
            ws[f"B{row}"].alignment = left_align()
            normal_style(ws[f"C{row}"], email)
            ws[f"C{row}"].alignment = left_align()
            for col, val in [("D", n1), ("E", n2), ("F", n3)]:
                cell = ws[f"{col}{row}"]
                cell.value = val
                cell.number_format = "0.00"
                normal_style(cell)
        else:
            for col in ["B", "C", "D", "E", "F"]:
                normal_style(ws[f"{col}{row}"])

        # Nota final = mitjana de les 3 avaluacions (només si n'hi ha)
        cell = ws[f"G{row}"]
        cell.value = (
            f'=IF(COUNT(D{row}:F{row})=0,"",'
            f'ROUND(AVERAGE(D{row}:F{row}),1))'
        )
        cell.number_format = "0.0"
        cell.font = font(bold=True, size=11, color=BLUE_DARK)
        cell.fill = fill(GREEN_LIGHT)
        cell.border = thick
        cell.alignment = center()

        # Mitjana (mateixa, amb 2 decimals)
        cell = ws[f"H{row}"]
        cell.value = f'=IF(COUNT(D{row}:F{row})=0,"",AVERAGE(D{row}:F{row}))'
        cell.number_format = "0.00"
        cell.font = font(size=10)
        cell.fill = fill(GREEN_LIGHT)
        cell.border = thin
        cell.alignment = center()

        # Qualificació
        cell = ws[f"I{row}"]
        cell.value = (
            f'=IF(G{row}="","",'
            f'IF(G{row}<5,"Insuficient",'
            f'IF(G{row}<6,"Suficient",'
            f'IF(G{row}<7,"Bé",'
            f'IF(G{row}<9,"Notable","Excel·lent")))))'
        )
        cell.font = font(bold=True, size=9)
        cell.border = thin
        cell.alignment = center()

    last = DATA_START + NUM - 1

    # Format condicional
    ws.conditional_formatting.add(
        f"D{DATA_START}:F{last}",
        ColorScaleRule(
            start_type="num", start_value=0, start_color="F1948A",
            mid_type="num", mid_value=5, mid_color="F9E79F",
            end_type="num", end_value=10, end_color="82E0AA"
        )
    )
    ws.conditional_formatting.add(
        f"I{DATA_START}:I{last}",
        FormulaRule(formula=[f'I{DATA_START}="Insuficient"'], fill=fill(RED_LIGHT))
    )
    ws.conditional_formatting.add(
        f"I{DATA_START}:I{last}",
        FormulaRule(formula=[f'I{DATA_START}="Excel·lent"'], fill=fill(GREEN_LIGHT))
    )

    # Resum
    sr = last + 2
    ws[f"A{sr}"] = "RESUM"
    ws[f"A{sr}"].font = font(bold=True, size=12, color=WHITE)
    ws[f"A{sr}"].fill = fill(BLUE_DARK)
    ws.merge_cells(f"A{sr}:C{sr}")

    summaries = [
        ("Mitjana del grup", f'=AVERAGEIF(G{DATA_START}:G{last},">=0")', "0.00"),
        ("Aprovats", f'=COUNTIF(G{DATA_START}:G{last},">=5")', "0"),
        ("Suspesos", f'=COUNTIF(G{DATA_START}:G{last},"<5")', "0"),
        ("% Aprovats", f'=IF(COUNTA(B{DATA_START}:B{last})=0,"",COUNTIF(G{DATA_START}:G{last},">=5")/COUNTA(B{DATA_START}:B{last}))', "0.0%"),
        ("Nota màxima", f'=MAX(G{DATA_START}:G{last})', "0.0"),
        ("Nota mínima", f'=MIN(G{DATA_START}:G{last})', "0.0"),
    ]
    for idx, (label, formula, fmt) in enumerate(summaries):
        r = sr + 1 + idx
        label_style(ws[f"A{r}"], label)
        ws.merge_cells(f"A{r}:B{r}")
        cell = ws[f"C{r}"]
        cell.value = formula
        cell.number_format = fmt
        cell.font = font(bold=True, size=11, color=BLUE_DARK)
        cell.fill = fill(GRAY)
        cell.border = thin
        cell.alignment = center()

    # Nota al peu
    note_row = sr + 8
    ws[f"A{note_row}"] = (
        "Nota: La «Nota final» és la mitjana aritmètica de les 3 avaluacions (arrodonida a 1 decimal). "
        "La columna «Mitjana» mostra el mateix valor amb 2 decimals. "
        "Qualificació: <5 Insuficient | <6 Suficient | <7 Bé | <9 Notable | ≥9 Excel·lent."
    )
    ws[f"A{note_row}"].font = font(size=8, color="5D6D7E")
    ws.merge_cells(f"A{note_row}:I{note_row}")

    path = "/home/workdir/artifacts/llista_notes_finals.xlsx"
    wb.save(path)
    print(f"Creat: {path}")
    return path


if __name__ == "__main__":
    create_quadern()
    create_llista_notes_finals()
    print("Fet!")
