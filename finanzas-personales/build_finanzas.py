# -*- coding: utf-8 -*-
"""
Finanzas Personales (versión simple) · Excel 365, sin macros.

Uso:  python build_finanzas.py [salida.xlsx]

Registrar = Fecha, Monto, Categoría y Cuenta. Todo lo demás (tipo de movimiento,
grupo, fijo/variable, moneda, análisis de comida, gastos hormiga) se deduce de la
categoría y de la cuenta elegidas.
"""
import os
import random
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.comments import Comment
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = sys.argv[1] if len(sys.argv) > 1 else "Finanzas_Personales.xlsx"

# ----------------------------------------------------------------------------
# Estilo
# ----------------------------------------------------------------------------
FONT = "Aptos"
INK, TEXT, MUTED, LINE_C = "0F172A", "334155", "64748B", "E2E8F0"
PAGE, CARD, SOFT = "F1F5F9", "FFFFFF", "F8FAFC"
BLUE, ORANGE, AQUA, VIOLET = "2A78D6", "EB6834", "1BAF7A", "6250D6"
GOOD, GOOD_BG, WARN, WARN_BG, BAD, BAD_BG = "15803D", "DCFCE7", "B45309", "FEF3C7", "B91C1C", "FEE2E2"
INPUT_BG, INPUT_LINE, DEMO_BG = "FFF8E7", "F2C46D", "F3F0FF"

MONEY = '"RD$ "#,##0;[Red]-"RD$ "#,##0;"–"'
MONEY2 = '#,##0.00;[Red]-#,##0.00;"–"'
NUM = '#,##0;[Red]-#,##0;"–"'
PCT = '0%;[Red]-0%;"–"'
DATE = "dd/mm/yyyy"

thin = Side(style="thin", color=LINE_C)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
inp = Side(style="thin", color=INPUT_LINE)
INPUT_BOX = Border(left=inp, right=inp, top=inp, bottom=inp)
LEFT = Alignment(horizontal="left", vertical="center", indent=1)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", indent=1)


def F(size=10, bold=False, color=TEXT, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def put(ws, ref, value=None, font=None, fill_c=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    if font is not None:
        c.font = font
    if fill_c is not None:
        c.fill = fill(fill_c)
    if fmt is not None:
        c.number_format = fmt
    if align is not None:
        c.alignment = align
    if border is not None:
        c.border = border
    return c


def area(ws, rng, fill_c=None, border=None):
    for row in ws[rng]:
        for c in row:
            if fill_c is not None:
                c.fill = fill(fill_c)
            if border is not None:
                c.border = border


def merge(ws, rng, value=None, **kw):
    ws.merge_cells(rng)
    return put(ws, rng.split(":")[0], value, **kw)


def outline_box(ws, rng, color=LINE_C):
    s = Side(style="thin", color=color)
    rows = list(ws[rng])
    for i, row in enumerate(rows):
        for j, c in enumerate(row):
            c.border = Border(left=s if j == 0 else c.border.left, right=s if j == len(row) - 1 else c.border.right,
                              top=s if i == 0 else c.border.top, bottom=s if i == len(rows) - 1 else c.border.bottom)


def card(ws, rng):
    area(ws, rng, fill_c=CARD)
    outline_box(ws, rng)


def section(ws, rng, text, color=INK):
    ws.merge_cells(rng)
    first = rng.split(":")[0]
    put(ws, first, text, F(12, True, color), align=Alignment(vertical="bottom"))
    for row in ws[rng]:
        for c in row:
            c.border = Border(bottom=Side(style="medium", color=color))
    ws.row_dimensions[int("".join(ch for ch in first if ch.isdigit()))].height = 24


def note(ws, ref, text, wrap=False):
    return put(ws, ref, text, F(9, False, MUTED, True), align=Alignment(vertical="center", wrap_text=wrap))


def input_cell(ws, rng, value=None, fmt=None, align=None):
    if ":" in rng:
        ws.merge_cells(rng)
        for row in ws[rng]:
            for c in row:
                c.fill = fill(INPUT_BG)
                c.border = INPUT_BOX
    return put(ws, rng.split(":")[0], value, F(10, True, "1E3A8A"), INPUT_BG, fmt, align or CENTER, INPUT_BOX)


def add_name(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def dv_list(ws, src, rng, strict=True):
    d = DataValidation(type="list", formula1=src, allow_blank=True, showErrorMessage=strict)
    if strict:
        d.error, d.errorTitle = "Elige un valor de la lista.", "Valor no válido"
    d.add(rng)
    ws.add_data_validation(d)


def fecha_txt(x):
    return f'RIGHT("0"&DAY({x}),2)&"/"&RIGHT("0"&MONTH({x}),2)&"/"&YEAR({x})'


MESES = '"Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"'

SHEETS = [("Registro", BLUE), ("Resumen", BLUE), ("Metas", AQUA), ("Ajustes", MUTED), ("Motor", "94A3B8")]
wb = Workbook()
wb.remove(wb.active)
W = {}
for name, color in SHEETS:
    W[name] = wb.create_sheet(name)
    W[name].sheet_properties.tabColor = color
REG, RES, MET, AJU, MOT = (W[n] for n, _ in SHEETS)


def header(ws, title, subtitle, last_col):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 38
    ws.row_dimensions[2].height = 22
    area(ws, f"A1:{L(last_col)}1", fill_c=INK)
    area(ws, f"A2:{L(last_col)}2", fill_c="E8EEF6")
    put(ws, "B1", title, F(18, True, "FFFFFF"), align=Alignment(vertical="center"))
    sub = max(5, last_col - 5)
    merge(ws, f"{L(sub)}1:{L(last_col)}1", subtitle, font=F(9, False, "CBD5E1"),
          align=Alignment(horizontal="right", vertical="center"))


def nav(ws, cols):
    """Enlaces a las 4 hojas visibles."""
    for (sheet, _), col in zip(SHEETS[:4], cols):
        me = sheet == ws.title
        c = ws[f"{col}2"]
        c.value = ("● " if me else "") + sheet
        c.hyperlink = Hyperlink(ref=f"{col}2", location=f"'{sheet}'!A1", display=sheet)
        c.font = F(9, True, INK if me else BLUE, underline=None if me else "single")
        c.alignment = Alignment(horizontal="left", vertical="center")


def axis_text(size=900, color="475569"):
    cp = CharacterProperties(latin=DFont(typeface=FONT), sz=size, solidFill=color)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def style_chart(ch, legend=True, grid=True):
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.txPr = axis_text()
    ch.y_axis.txPr = axis_text()
    ch.x_axis.majorTickMark = ch.y_axis.majorTickMark = "none"
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=LINE_C))
    ch.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="EEF2F7"))) if grid else None
    if legend:
        ch.legend.position = "t"
        ch.legend.txPr = axis_text(900, TEXT)
    else:
        ch.legend = None
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))


def paint(series, color, labels=None):
    series.graphicalProperties.solidFill = color
    series.graphicalProperties.line.solidFill = "FFFFFF"
    if labels:
        series.dLbls = DataLabelList()
        series.dLbls.showVal = True
        for a in ("showSerName", "showCatName", "showLegendKey"):
            setattr(series.dLbls, a, False)
        series.dLbls.numFmt = labels
        series.dLbls.txPr = axis_text(800, TEXT)
        series.dLbls.position = "outEnd"


def kpi(ws, r, c1, c2, label, value, fmt, sub=None, accent=BLUE, size=20):
    a, b = L(c1), L(c2)
    card(ws, f"{a}{r}:{b}{r + 3}")
    for rr in range(r, r + 4):
        cur = ws[f"{a}{rr}"].border
        ws[f"{a}{rr}"].border = Border(left=Side(style="thick", color=accent), top=cur.top, bottom=cur.bottom)
    merge(ws, f"{a}{r}:{b}{r}", label.upper(), font=F(8, True, MUTED), align=Alignment(horizontal="left", vertical="bottom", indent=1))
    merge(ws, f"{a}{r + 1}:{b}{r + 2}", value, font=F(size, True, INK), fmt=fmt, align=Alignment(horizontal="left", vertical="center", indent=1))
    merge(ws, f"{a}{r + 3}:{b}{r + 3}", sub, font=F(8, False, MUTED), align=Alignment(horizontal="left", vertical="top", indent=1, wrap_text=True))
    for rr, h in zip(range(r, r + 4), (18, 18, 18, 22)):
        ws.row_dimensions[rr].height = h


# ----------------------------------------------------------------------------
# Datos base
# ----------------------------------------------------------------------------
CP, SI, BS = "Comida comprada", "Súper e ingredientes", "Snacks y bebidas"
# (categoría, grupo, tipo, fijo/variable, análisis de comida, comida del día)
CATS = [
    ("Desayuno comprado", "Alimentación", "Gasto", "Variable", CP, "Desayuno"),
    ("Almuerzo comprado", "Alimentación", "Gasto", "Variable", CP, "Almuerzo"),
    ("Cena comprada", "Alimentación", "Gasto", "Variable", CP, "Cena"),
    ("Restaurante", "Alimentación", "Gasto", "Variable", CP, None),
    ("Delivery", "Alimentación", "Gasto", "Variable", CP, None),
    ("Supermercado", "Alimentación", "Gasto", "Variable", SI, None),
    ("Ingredientes / colmado", "Alimentación", "Gasto", "Variable", SI, None),
    ("Snacks y bebidas", "Alimentación", "Gasto", "Variable", BS, None),
    ("Internet", "Hogar y servicios", "Gasto", "Fijo", None, None),
    ("Electricidad", "Hogar y servicios", "Gasto", "Fijo", None, None),
    ("Agua", "Hogar y servicios", "Gasto", "Fijo", None, None),
    ("Compras del hogar", "Hogar y servicios", "Gasto", "Variable", None, None),
    ("Mantenimiento del hogar", "Hogar y servicios", "Gasto", "Variable", None, None),
    ("Combustible", "Transporte y vehículo", "Gasto", "Variable", None, None),
    ("Mantenimiento / reparación vehículo", "Transporte y vehículo", "Gasto", "Variable", None, None),
    ("Seguro del vehículo", "Transporte y vehículo", "Gasto", "Fijo", None, None),
    ("Parqueos y peajes", "Transporte y vehículo", "Gasto", "Variable", None, None),
    ("Herramientas de IA", "Tecnología y suscripciones", "Gasto", "Fijo", None, None),
    ("Streaming y entretenimiento", "Tecnología y suscripciones", "Gasto", "Fijo", None, None),
    ("Software y servicios digitales", "Tecnología y suscripciones", "Gasto", "Fijo", None, None),
    ("Evaluación / cuenta de trading", "Trading", "Gasto", "Variable", None, None),
    ("Plataformas de trading", "Trading", "Gasto", "Fijo", None, None),
    ("Comisiones de trading", "Trading", "Gasto", "Variable", None, None),
    ("Salud y farmacia", "Salud", "Gasto", "Variable", None, None),
    ("Gimnasio", "Salud", "Gasto", "Fijo", None, None),
    ("Educación", "Educación", "Gasto", "Variable", None, None),
    ("Compras personales", "Compras personales", "Gasto", "Variable", None, None),
    ("Pareja y salidas", "Pareja y social", "Gasto", "Variable", None, None),
    ("Boda", "Boda", "Gasto", "Variable", None, None),
    ("Regalos", "Regalos", "Gasto", "Variable", None, None),
    ("Imprevistos", "Imprevistos", "Gasto", "Variable", None, None),
    ("Otros gastos", "Otros", "Gasto", "Variable", None, None),
    ("Salario", "Ingresos", "Ingreso", None, None, None),
    ("Ingreso extra", "Ingresos", "Ingreso", None, None, None),
    ("Transferencia entre cuentas", "Movimientos", "Transferencia", None, None, None),
    ("Pago de tarjeta", "Movimientos", "Pago tarjeta", None, None, None),
    ("Depósito / retiro de trading", "Movimientos", "Transferencia", None, None, None),
    ("Ganancia / pérdida de trading", "Trading", "Trading", None, None, None),
    ("Ajuste de saldo", "Movimientos", "Ajuste", None, None, None),
]
ACCOUNTS = [("Efectivo", "Efectivo", "DOP"), ("Banco", "Cuenta bancaria / débito", "DOP"),
            ("Tarjeta de crédito", "Tarjeta de crédito", "DOP"), ("Billetera digital", "Billetera digital", "DOP"),
            ("Cuenta de trading", "Cuenta de trading", "USD")]

R0, NROWS = 6, 4000
RL = R0 + NROWS - 1
CAT0, CATN = 40, 100          # filas de categorías en Ajustes
CATL = CAT0 + CATN - 1
CTA0, CTAN = 22, 12           # filas de cuentas en Ajustes
CTAL = CTA0 + CTAN - 1

# ----------------------------------------------------------------------------
# Nombres definidos
# ----------------------------------------------------------------------------
for nm, col in {"rFecha": "B", "rMonto": "C", "rCat": "D", "rCta": "E", "rCom": "F", "rHacia": "G", "rEstado": "H",
                "rTipo": "I", "rGrupo": "J", "rClas": "K", "rAlim": "L", "rComida": "M", "rMon": "N", "rDOP": "O",
                "rGasto": "P", "rIng": "Q", "rTrad": "R", "rHorm": "S", "rEfO": "T", "rEfD": "U", "rFlujo": "V",
                "rComIdx": "W"}.items():
    add_name(nm, f"'Registro'!${col}${R0}:${col}${RL}")
for nm, col in {"catNom": "B", "catGrupo": "C", "catTipo": "D", "catClas": "E", "catAlim": "F", "catComida": "G",
                "catGIdx": "I"}.items():
    add_name(nm, f"'Ajustes'!${col}${CAT0}:${col}${CATL}")
add_name("lstCat", f"OFFSET('Ajustes'!$B${CAT0},0,0,MAX(1,COUNTIF('Ajustes'!$B${CAT0}:$B${CATL},\"?*\")),1)")
for nm, col in {"ctaNom": "B", "ctaTipo": "C", "ctaMon": "D", "ctaIni": "E", "ctaSaldo": "G", "ctaDOP": "I", "ctaLiq": "J"}.items():
    add_name(nm, f"'Ajustes'!${col}${CTA0}:${col}${CTAL}")
add_name("lstCta", f"OFFSET('Ajustes'!$B${CTA0},0,0,MAX(1,COUNTIF('Ajustes'!$B${CTA0}:$B${CTAL},\"?*\")),1)")


# ----------------------------------------------------------------------------
# Datos DEMO (todas las filas llevan "DEMO" en Comercio)
# ----------------------------------------------------------------------------
def build_demo():
    rnd = random.Random(7)
    rows = []
    start, end = date(2026, 9, 1), date(2026, 10, 13)

    def add(d, monto, cat, cta, com="DEMO", hacia=None):
        if d <= end:
            rows.append((d, monto, cat, cta, com, hacia))

    d = start
    while d <= end:
        wd = d.weekday()
        if wd < 5:
            if rnd.random() < 0.35:
                add(d, rnd.randrange(180, 340, 10), "Desayuno comprado", rnd.choice(["Banco", "Efectivo"]), "DEMO Cafetería")
            if rnd.random() < 0.6:
                add(d, rnd.randrange(350, 600, 25), "Almuerzo comprado", rnd.choice(["Banco", "Billetera digital"]), "DEMO Comedor")
            if rnd.random() < 0.3:
                add(d, rnd.randrange(90, 220, 10), "Snacks y bebidas", "Efectivo", "DEMO Colmado")
        if rnd.random() < 0.12:
            add(d, rnd.randrange(400, 750, 25), "Cena comprada", "Tarjeta de crédito", "DEMO Pizzería")
        if wd == 4 and rnd.random() < 0.6:
            add(d, rnd.randrange(650, 1250, 50), "Delivery", "Tarjeta de crédito", "DEMO App delivery")
        if wd == 5 and rnd.random() < 0.5:
            add(d, rnd.randrange(1800, 3600, 100), "Restaurante", "Tarjeta de crédito", "DEMO Restaurante")
        if d.day in (3, 17):
            add(d, rnd.randrange(3500, 6500, 50), "Supermercado", "Tarjeta de crédito", "DEMO Supermercado")
        if wd == 6 and rnd.random() < 0.6:
            add(d, rnd.randrange(350, 900, 25), "Ingredientes / colmado", "Efectivo", "DEMO Mercado")
        if wd == 0:
            add(d, rnd.randrange(2000, 3000, 50), "Combustible", "Tarjeta de crédito", "DEMO Gasolinera")
        d += timedelta(days=1)
    for m in (9, 10):
        add(date(2026, m, 1), 5000, "Transferencia entre cuentas", "Banco", hacia="Efectivo")
        add(date(2026, m, 3), 1270, "Herramientas de IA", "Tarjeta de crédito", "DEMO Herramienta IA")
        add(date(2026, m, 5), 2050, "Internet", "Banco", "DEMO Internet")
        add(date(2026, m, 8), 2500, "Gimnasio", "Tarjeta de crédito", "DEMO Gimnasio")
        add(date(2026, m, 12), rnd.randrange(2700, 3400, 10), "Electricidad", "Banco", "DEMO Electricidad")
        add(date(2026, m, 10), rnd.randrange(1800, 3000, 100), "Pareja y salidas", "Tarjeta de crédito", "DEMO Salida")
    add(date(2026, 9, 15), 55000, "Salario", "Banco", "DEMO Salario (ficticio)")
    add(date(2026, 9, 30), 55000, "Salario", "Banco", "DEMO Salario (ficticio)")
    add(date(2026, 9, 25), 30000, "Pago de tarjeta", "Banco", hacia="Tarjeta de crédito")
    add(date(2026, 9, 16), 3000, "Transferencia entre cuentas", "Banco", hacia="Billetera digital")
    add(date(2026, 9, 18), 698, "Streaming y entretenimiento", "Tarjeta de crédito", "DEMO Streaming")
    add(date(2026, 9, 6), 15750, "Depósito / retiro de trading", "Banco", hacia="Cuenta de trading")
    add(date(2026, 9, 7), 150, "Evaluación / cuenta de trading", "Cuenta de trading", "DEMO Firma de evaluación")
    add(date(2026, 9, 20), -40, "Ganancia / pérdida de trading", "Cuenta de trading")
    add(date(2026, 10, 8), 95, "Ganancia / pérdida de trading", "Cuenta de trading")
    add(date(2026, 9, 12), 4200, "Restaurante", "Tarjeta de crédito", "DEMO Restaurante")
    add(date(2026, 9, 14), -1400, "Restaurante", "Billetera digital", "DEMO Amigo devolvió su parte")
    add(date(2026, 9, 20), 15000, "Boda", "Banco", "DEMO Proveedor boda")
    add(date(2026, 9, 26), 8000, "Ingreso extra", "Billetera digital", "DEMO Trabajo extra (ficticio)")
    add(date(2026, 10, 1), -150, "Ajuste de saldo", "Efectivo")
    add(date(2026, 10, 4), 1800, "Imprevistos", "Efectivo", "DEMO Reparación celular")
    rows.sort(key=lambda x: x[0])
    return rows


DEMO = build_demo()
DEMO_LAST = R0 + len(DEMO) - 1

# ============================================================================
# AJUSTES
# ============================================================================
ws = AJU
header(ws, "Ajustes", "Cuentas, categorías y parámetros", 11)
for col, w in {"A": 2, "B": 34, "C": 26, "D": 15, "E": 15, "F": 20, "G": 15, "H": 15, "I": 3, "J": 3, "K": 2}.items():
    ws.column_dimensions[col].width = w
ws.column_dimensions["I"].hidden = True
ws.column_dimensions["J"].hidden = True
nav(ws, "BCDE")

section(ws, "B4:H4", "Cómo registrar (10 segundos)")
howto = [
    "1. En «Registro» escribe Fecha (Ctrl+; = hoy), Monto, Categoría y Cuenta. Listo.",
    "2. Comercio es opcional (alimenta el ranking de lugares). «Hacia» solo se usa en transferencias y pagos de tarjeta.",
    "3. Compra con tarjeta de crédito = gasto normal con Cuenta «Tarjeta de crédito». Pagar la tarjeta = categoría «Pago de tarjeta» (no es gasto).",
    "4. ¿Te devolvieron dinero? Escribe el monto en NEGATIVO con la misma categoría del gasto.",
    "5. Trading: depositar/retirar = «Depósito / retiro de trading» (no es gasto). Ganancia (+) o pérdida (−) realizada = «Ganancia / pérdida de trading».",
    "6. El monto va en la moneda de la cuenta (si la cuenta es USD, escribe dólares).",
    f"Datos DEMO = filas lilas del Registro. Para borrarlos: selecciona B6:G{DEMO_LAST} en Registro y pulsa Suprimir (y las filas DEMO de Metas).",
]
for i, t in enumerate(howto):
    merge(ws, f"B{5 + i}:H{5 + i}", t, font=F(10, i == 6, TEXT if i < 6 else VIOLET), align=LEFT_WRAP)
    ws.row_dimensions[5 + i].height = 20

section(ws, "B13:H13", "Parámetros")
put(ws, "B14", "Tasa USD → DOP", F(10, True, INK), align=LEFT)
input_cell(ws, "C14", 63.50, "0.00")
note(ws, "D14", "Valor referencial: actualízalo con la tasa vigente.")
put(ws, "B15", "Umbral de gasto hormiga", F(10, True, INK), align=LEFT)
input_cell(ws, "C15", 500, MONEY)
note(ws, "D15", "Gastos variables iguales o menores a este monto. Es un parámetro de análisis, no un límite.")
add_name("TasaUSD", "'Ajustes'!$C$14")
add_name("Umbral", "'Ajustes'!$C$15")

section(ws, "B19:H19", "Cuentas y métodos de pago")
note(ws, "B20", "Sin saldo inicial no se calcula saldo (no se inventan saldos). Tarjeta de crédito en negativo = deuda.")
for col, h in zip("BCDEFGH", ["Cuenta", "Tipo", "Moneda", "Saldo inicial (opcional)", "Fecha del saldo", "Saldo actual", "Gasto este mes"]):
    put(ws, f"{col}21", h, F(9, True, "FFFFFF"), INK if col in "BCDEF" else "475569", align=CENTER)
ws.row_dimensions[21].height = 28
for i in range(CTAN):
    r = CTA0 + i
    acc = ACCOUNTS[i] if i < len(ACCOUNTS) else (None, None, None)
    for col, v in zip("BCD", acc):
        put(ws, f"{col}{r}", v, F(10, col == "B", INK), INPUT_BG, None, LEFT if col != "D" else CENTER, BOX)
    put(ws, f"E{r}", None, F(10, True, "1E3A8A"), INPUT_BG, MONEY2, RIGHT, BOX)
    put(ws, f"F{r}", None, F(10, False, "1E3A8A"), INPUT_BG, DATE, CENTER, BOX)
    crit = f'">="&IF(F{r}="",0,F{r})'
    put(ws, f"G{r}", f'=IF(B{r}="","",IF(E{r}="","—",E{r}+SUMIFS(rEfO,rCta,B{r},rFecha,{crit})+SUMIFS(rEfD,rHacia,B{r},rFecha,{crit})))',
        F(10, True, INK), SOFT, MONEY2, RIGHT, BOX)
    put(ws, f"H{r}", f'=IF(B{r}="","",SUMIFS(rGasto,rCta,B{r},rFecha,">="&mesHoy,rFecha,"<="&EOMONTH(diaHoy,0)))',
        F(10, False, INK), SOFT, MONEY2, RIGHT, BOX)
    ws[f"I{r}"] = f'=IF(ISNUMBER(G{r}),G{r}*IF(D{r}="USD",TasaUSD,1),"")'
    ws[f"J{r}"] = f'=IF(B{r}="","",IF(OR(C{r}="Tarjeta de crédito",C{r}="Cuenta de trading"),"No","Sí"))'

section(ws, "B37:H37", "Categorías (edita, agrega o borra filas)")
note(ws, "B38", "El Tipo decide cómo cuenta: Gasto · Ingreso · Transferencia / Pago tarjeta (no son gasto) · Trading (resultado aparte) · Ajuste (solo saldo).")
for col, h in zip("BCDEFG", ["Categoría", "Grupo", "Tipo", "Fijo / Variable", "Análisis de comida", "Comida del día"]):
    put(ws, f"{col}39", h, F(9, True, "FFFFFF"), INK, align=CENTER)
for i in range(CATN):
    r = CAT0 + i
    row = CATS[i] if i < len(CATS) else (None,) * 6
    for col, v in zip("BCDEFG", row):
        put(ws, f"{col}{r}", v, F(10, col == "B", INK), CARD, None, LEFT if col in "BCF" else CENTER, BOX)
    ws[f"I{r}"] = (f'=IF(OR(C{r}="",D{r}<>"Gasto"),"",IF(COUNTIFS(C${CAT0}:C{r},C{r},D${CAT0}:D{r},"Gasto")=1,'
                   f'MAX(I${CAT0 - 1}:I{r - 1})+1,""))')

# Listas para desplegables (columnas ocultas)
LISTS = {
    "M": ("lstTipoCta", ["Efectivo", "Cuenta bancaria / débito", "Tarjeta de crédito", "Billetera digital", "Cuenta de trading"]),
    "N": ("lstMon", ["DOP", "USD"]),
    "O": ("lstTipo", ["Gasto", "Ingreso", "Transferencia", "Pago tarjeta", "Trading", "Ajuste"]),
    "P": ("lstClas", ["Fijo", "Variable"]),
    "Q": ("lstAlim", [CP, SI, BS]),
    "R": ("lstComida", ["Desayuno", "Almuerzo", "Cena"]),
    "S": ("lstPeriodo", ["Mes actual", "Mes anterior", "Mes seleccionado", "Últimos 30 días", "Año actual", "Todo el historial"]),
    "U": ("lstFrec", ["Mensual", "Anual"]),
}
for col, (nm, vals) in LISTS.items():
    for i, v in enumerate(vals):
        ws[f"{col}{5 + i}"] = v
    add_name(nm, f"'Ajustes'!${col}$5:${col}${4 + len(vals)}")
    ws.column_dimensions[col].hidden = True
for k in range(24):
    ws[f"T{5 + k}"] = (f'=YEAR(DATE(YEAR(TODAY()),MONTH(TODAY())-{k},1))&"-"&'
                       f'RIGHT("0"&MONTH(DATE(YEAR(TODAY()),MONTH(TODAY())-{k},1)),2)')
add_name("lstMes", "'Ajustes'!$T$5:$T$28")
ws.column_dimensions["T"].hidden = True
add_name("gCP", "'Ajustes'!$Q$5")
add_name("gSI", "'Ajustes'!$Q$6")
add_name("gBS", "'Ajustes'!$Q$7")
for col, src in (("C", "=lstTipoCta"), ("D", "=lstMon")):
    dv_list(ws, src, f"{col}{CTA0}:{col}{CTAL}")
for col, src in (("D", "=lstTipo"), ("E", "=lstClas"), ("F", "=lstAlim"), ("G", "=lstComida")):
    dv_list(ws, src, f"{col}{CAT0}:{col}{CATL}")
ws.conditional_formatting.add(f"G{CTA0}:G{CTAL}", FormulaRule(formula=[f'ISTEXT(G{CTA0})'], font=Font(color="94A3B8", italic=True)))
ws.freeze_panes = "A3"

# ============================================================================
# REGISTRO
# ============================================================================
ws = REG
header(ws, "Registro", "Una fila = un movimiento · 4 datos obligatorios", 23)
for col, w in {"A": 2, "B": 12, "C": 13, "D": 30, "E": 20, "F": 26, "G": 18, "H": 30}.items():
    ws.column_dimensions[col].width = w
for i in range(9, 24):
    ws.column_dimensions[L(i)].width = 13
nav(ws, "BCDE")
ws.row_dimensions[3].height = 26
put(ws, "B3", f'=HYPERLINK("#\'Registro\'!B"&IFERROR(MATCH(TRUE,INDEX(rFecha="",0),0)+{R0 - 1},{RL}),"＋ Nuevo")',
    F(10, True, "FFFFFF"), BLUE, align=CENTER)
merge(ws, "C3:F3", "Ctrl+; = fecha de hoy · escribe las primeras letras en Categoría y Cuenta · devolución = monto negativo",
      font=F(9, False, MUTED, True), align=LEFT_WRAP)
put(ws, "G3", '=IF(COUNTIF(rEstado,"⚠*")>0,"⚠ Revisar: "&COUNTIF(rEstado,"⚠*"),"✓ Sin errores")', F(9, True, INK), align=LEFT)
put(ws, "H3", '="Movimientos: "&COUNTIF(rTipo,"?*")&IF(COUNTIF(rCom,"DEMO*")>0,"  ·  lilas = DEMO","")', F(9, False, MUTED), align=LEFT)
merge(ws, "B4:G4", "LLENA ESTO  ·  obligatorio: Fecha, Monto, Categoría, Cuenta", font=F(9, True, "1E3A8A"), fill_c="E8EEF6", align=LEFT)
merge(ws, "H4:W4", "AUTOMÁTICO · no editar (pulsa «+» sobre la columna I para ver el detalle)", font=F(9, True, MUTED), fill_c="EEF2F7", align=LEFT)
heads = ["Fecha", "Monto", "Categoría", "Cuenta", "Comercio (opcional)", "Hacia (solo transf.)", "Estado", "Tipo", "Grupo",
         "Fijo/Variable", "Análisis comida", "Comida del día", "Moneda", "Monto DOP", "Gasto real", "Ingreso", "Trading",
         "Hormiga", "Efecto saldo origen", "Efecto saldo destino", "Flujo efectivo", "Índice comercio"]
for j, h in enumerate(heads):
    put(ws, f"{L(2 + j)}5", h, F(10, True, "FFFFFF"), INK if j < 6 else "64748B", align=CENTER)
ws.row_dimensions[5].height = 30
ws["C5"].comment = Comment("En la moneda de la cuenta. Negativo = devolución/reembolso de un gasto.", "Sistema")
ws["G5"].comment = Comment("Solo para Transferencia entre cuentas, Pago de tarjeta y Depósito/retiro de trading: la cuenta que recibe.", "Sistema")


def reg_formulas(r):
    i = f"I{r}"

    def look(col):
        return f'IFERROR(INDEX({col},MATCH(D{r},catNom,0))&"","")'

    def liq(x):
        return f'IFERROR(INDEX(ctaLiq,MATCH({x},ctaNom,0))="Sí",FALSE)'

    sign = f'IF(OR({i}="Gasto",{i}="Transferencia",{i}="Pago tarjeta"),-1,1)'
    return {
        "H": (f'=IF(AND(B{r}="",C{r}="",D{r}="",E{r}=""),"",IF(B{r}="","⚠ Falta fecha",IF(C{r}="","⚠ Falta monto",'
              f'IF(D{r}="","⚠ Falta categoría",IF(ISNA(MATCH(D{r},catNom,0)),"⚠ Categoría no existe",'
              f'IF(E{r}="","⚠ Falta cuenta",IF(ISNA(MATCH(E{r},ctaNom,0)),"⚠ Cuenta no existe",'
              f'IF(AND(OR({i}="Transferencia",{i}="Pago tarjeta"),G{r}=""),"⚠ Indica la cuenta «Hacia»",'
              f'IF(AND(G{r}<>"",G{r}=E{r}),"⚠ «Hacia» igual a la cuenta",'
              f'IF(AND(G{r}<>"",ISNA(MATCH(G{r},ctaNom,0))),"⚠ «Hacia» no existe",'
              f'IF(AND({i}="Pago tarjeta",IFERROR(INDEX(ctaTipo,MATCH(G{r},ctaNom,0)),"")<>"Tarjeta de crédito"),"⚠ «Hacia» debe ser la tarjeta",'
              f'IF(AND({i}="Trading",IFERROR(INDEX(ctaTipo,MATCH(E{r},ctaNom,0)),"")<>"Cuenta de trading"),"⚠ Usa la cuenta de trading","✓"))))))))))))'),
        "I": f'=IF(OR(B{r}="",C{r}="",D{r}=""),"",{look("catTipo")})',
        "J": f'=IF({i}="","",{look("catGrupo")})',
        "K": f'=IF({i}="Gasto",IF({look("catClas")}="","Variable",{look("catClas")}),"")',
        "L": f'=IF({i}="","",{look("catAlim")})',
        "M": f'=IF({i}="","",{look("catComida")})',
        "N": f'=IF({i}="","",IFERROR(INDEX(ctaMon,MATCH(E{r},ctaNom,0))&"","DOP"))',
        "O": f'=IF({i}="","",C{r}*IF(N{r}="USD",TasaUSD,1))',
        "P": f'=IF({i}="","",IF({i}="Gasto",O{r},0))',
        "Q": f'=IF({i}="","",IF({i}="Ingreso",O{r},0))',
        "R": f'=IF({i}="","",IF({i}="Trading",O{r},0))',
        "S": f'=IF({i}="","",IF(AND({i}="Gasto",K{r}="Variable",J{r}<>"Trading",O{r}>0,O{r}<=Umbral),1,0))',
        "T": f'=IF(OR({i}="",E{r}=""),"",{sign}*C{r})',
        "U": (f'=IF(OR(G{r}="",AND({i}<>"Transferencia",{i}<>"Pago tarjeta")),"",'
              f'IF(IFERROR(INDEX(ctaMon,MATCH(G{r},ctaNom,0))&"","DOP")=N{r},C{r},IF(N{r}="USD",C{r}*TasaUSD,C{r}/TasaUSD)))'),
        "V": f'=IF({i}="","",IF({liq(f"E{r}")},{sign}*O{r},0)+IF(AND(U{r}<>"",{liq(f"G{r}")}),O{r},0))',
        "W": (f'=IF(OR(F{r}="",{i}=""),"",IF(COUNTIFS(F${R0}:F{r},F{r},I${R0}:I{r},"?*")=1,'
              f'MAX(W${R0 - 1}:W{r - 1})+1,""))'),
    }


f_in, f_calc = F(10, False, INK), F(9, False, MUTED)
for k in range(NROWS):
    r = R0 + k
    if k < len(DEMO):
        for col, v in zip("BCDEFG", DEMO[k]):
            ws[f"{col}{r}"] = v
    for col, fml in reg_formulas(r).items():
        ws[f"{col}{r}"] = fml
    ws[f"B{r}"].number_format = DATE
    ws[f"C{r}"].number_format = "#,##0.00"
    for col in "BCDEFG":
        ws[f"{col}{r}"].font = f_in
    ws[f"H{r}"].font = F(9, True, GOOD)
    for c in range(9, 24):
        ws.cell(r, c).font = f_calc
    for col in "OPQRTUV":
        ws[f"{col}{r}"].number_format = "#,##0.00"
tbl = Table(displayName="tblRegistro", ref=f"B5:W{RL}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
ws.add_table(tbl)
ws.column_dimensions.group("I", "W", hidden=True, outline_level=1)
ws.freeze_panes = f"C{R0}"
dvd = DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)", formula2="DATE(2100,12,31)",
                     allow_blank=True, showErrorMessage=True, error="Escribe una fecha válida.", errorTitle="Fecha")
dvd.add(f"B{R0}:B{RL}")
ws.add_data_validation(dvd)
dvm = DataValidation(type="decimal", operator="between", formula1="-1000000000", formula2="1000000000",
                     allow_blank=True, showErrorMessage=True, error="Escribe un número.", errorTitle="Monto")
dvm.add(f"C{R0}:C{RL}")
ws.add_data_validation(dvm)
dv_list(ws, "=lstCat", f"D{R0}:D{RL}")
dv_list(ws, "=lstCta", f"E{R0}:E{RL}")
dv_list(ws, "=lstCta", f"G{R0}:G{RL}")
dv_list(ws, "=lstComercios", f"F{R0}:F{RL}", strict=False)
ws.conditional_formatting.add(f"B{R0}:G{RL}", FormulaRule(formula=[f'LEFT($F{R0},4)="DEMO"'], fill=fill(DEMO_BG)))
ws.conditional_formatting.add(f"H{R0}:H{RL}", FormulaRule(formula=[f'LEFT($H{R0},1)="⚠"'], fill=fill(BAD_BG), font=Font(color=BAD, bold=True)))
ws.conditional_formatting.add(f"C{R0}:C{RL}", FormulaRule(formula=[f'$I{R0}="Ingreso"'], font=Font(color=GOOD, bold=True)))
ws.conditional_formatting.add(f"C{R0}:C{RL}", FormulaRule(formula=[f'OR($I{R0}="Transferencia",$I{R0}="Pago tarjeta",$I{R0}="Ajuste")'], font=Font(color="64748B", italic=True)))
ws.conditional_formatting.add(f"C{R0}:C{RL}", FormulaRule(formula=[f'$I{R0}="Trading"'], font=Font(color=VIOLET, bold=True)))

# ============================================================================
# MOTOR (oculta)
# ============================================================================
ws = MOT
header(ws, "Motor", "Cálculos intermedios · no editar", 44)
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 14
P = [
    ("diaHoy", "Hoy", "=TODAY()", DATE),
    ("mesHoy", "Inicio de este mes", "=DATE(YEAR(diaHoy),MONTH(diaHoy),1)", DATE),
    ("semHoy", "Lunes de esta semana", "=diaHoy-WEEKDAY(diaHoy,3)", DATE),
    ("minFecha", "Primera fecha", "=IF(COUNT(rFecha)=0,diaHoy,MIN(rFecha))", DATE),
    ("maxFecha", "Última fecha", "=IF(COUNT(rFecha)=0,diaHoy,MAX(rFecha))", DATE),
    ("mesSel", "Mes elegido", '=IFERROR(DATE(VALUE(LEFT(fMes,4)),VALUE(RIGHT(fMes,2)),1),mesHoy)', DATE),
    ("pK", "Índice período", "=IFERROR(MATCH(fPeriodo,lstPeriodo,0),1)", "0"),
    ("pIni", "Inicio", "=CHOOSE(pK,mesHoy,EOMONTH(diaHoy,-2)+1,mesSel,diaHoy-29,DATE(YEAR(diaHoy),1,1),minFecha)", DATE),
    ("pFin", "Fin", "=CHOOSE(pK,EOMONTH(diaHoy,0),EOMONTH(diaHoy,-1),EOMONTH(mesSel,0),diaHoy,DATE(YEAR(diaHoy),12,31),MAX(maxFecha,diaHoy))", DATE),
    ("pDias", "Días transcurridos", "=MAX(1,MIN(pFin,MAX(diaHoy,pIni))-pIni+1)", "0"),
    ("refMes", "Mes de referencia", "=DATE(YEAR(pIni+pDias-1),MONTH(pIni+pDias-1),1)", DATE),
    ("m3", "Meses con datos (3 previos)", '=MAX(1,(COUNTIFS(rTipo,"?*",rFecha,">="&EDATE(mesHoy,-3),rFecha,"<"&EDATE(mesHoy,-2))>0)'
                                         '+(COUNTIFS(rTipo,"?*",rFecha,">="&EDATE(mesHoy,-2),rFecha,"<"&EDATE(mesHoy,-1))>0)'
                                         '+(COUNTIFS(rTipo,"?*",rFecha,">="&EDATE(mesHoy,-1),rFecha,"<"&mesHoy)>0))', "0"),
    ("pTexto", "Texto", "=" + fecha_txt("pIni") + '&"  →  "&' + fecha_txt("pFin"), None),
]
for k, (nm, lab, fml, fmt) in enumerate(P):
    r = 4 + k
    put(ws, f"B{r}", lab, F(9))
    put(ws, f"C{r}", fml, F(9, True, INK), fmt=fmt)
    add_name(nm, f"'Motor'!$C${r}")
DD = 'rFecha,">="&pIni,rFecha,"<="&pFin'


def block(c0, title, n, name_fml, val_fml, cnt_fml):
    """Lista con monto, conteo y clave de orden única (para LARGE + MATCH)."""
    a, b, c, d = (L(c0 + j) for j in range(4))
    put(ws, f"{a}3", title, F(10, True, INK))
    for j, h in enumerate(["Nombre", "Monto", "Nº", "Clave"]):
        put(ws, f"{L(c0 + j)}4", h, F(9, True, "FFFFFF"), "475569", align=CENTER)
    for k in range(1, n + 1):
        r = 4 + k
        ws[f"{a}{r}"] = "=" + name_fml(k)
        ws[f"{b}{r}"] = f'=IF({a}{r}="","",{val_fml(f"{a}{r}")})'
        ws[f"{c}{r}"] = f'=IF({a}{r}="","",{cnt_fml(f"{a}{r}")})'
        ws[f"{d}{r}"] = f'=IF({a}{r}="",-1E+12-{k},N({b}{r})+({n + 1}-{k})/1E+7)'

    def rng(col):
        return f"'Motor'!${col}$5:${col}${4 + n}"
    return {"name": rng(a), "val": rng(b), "cnt": rng(c), "key": rng(d)}


B_GRP = block(5, "Gasto por grupo", 20, lambda k: f'IFERROR(INDEX(catGrupo,MATCH({k},catGIdx,0))&"","")',
              lambda x: f"SUMIFS(rGasto,rGrupo,{x},{DD})", lambda x: f'COUNTIFS(rGrupo,{x},rTipo,"Gasto",{DD})')
B_CAT = block(10, "Gasto por categoría", CATN, lambda k: f'IF(INDEX(catTipo,{k})&""="Gasto",INDEX(catNom,{k})&"","")',
              lambda x: f"SUMIFS(rGasto,rCat,{x},{DD})", lambda x: f'COUNTIFS(rCat,{x},rMonto,">0",{DD})')
B_CTA = block(15, "Gasto por cuenta", CTAN, lambda k: f'INDEX(ctaNom,{k})&""',
              lambda x: f"SUMIFS(rGasto,rCta,{x},{DD})", lambda x: f'COUNTIFS(rCta,{x},rTipo,"Gasto",{DD})')
B_COM = block(20, "Gasto por comercio", 200, lambda k: f'IFERROR(INDEX(rCom,MATCH({k},rComIdx,0))&"","")',
              lambda x: f"SUMIFS(rGasto,rCom,{x},{DD})", lambda x: f'COUNTIFS(rCom,{x},rTipo,"Gasto",{DD})')
add_name("lstComercios", "OFFSET('Motor'!$T$5,0,0,MAX(1,COUNTIF('Motor'!$T$5:$T$204,\"?*\")),1)")

put(ws, "Y3", "Últimos 6 meses", F(10, True, INK))
for j, h in enumerate(["Mes", "Etiqueta", "Ingresos", "Gasto real", "Comida comprada", "Súper e ingredientes"]):
    put(ws, f"{L(25 + j)}4", h, F(9, True, "FFFFFF"), "475569", align=CENTER)
for j in range(1, 7):
    r = 4 + j
    dd = f'rFecha,">="&Y{r},rFecha,"<="&EOMONTH(Y{r},0)'
    ws[f"Y{r}"] = f"=EDATE(refMes,{j - 6})"
    ws[f"Y{r}"].number_format = DATE
    ws[f"Z{r}"] = f'=CHOOSE(MONTH(Y{r}),{MESES})&" "&RIGHT(YEAR(Y{r}),2)'
    ws[f"AA{r}"] = f"=SUMIFS(rIng,{dd})"
    ws[f"AB{r}"] = f"=SUMIFS(rGasto,{dd})"
    ws[f"AC{r}"] = f"=SUMIFS(rGasto,rAlim,gCP,{dd})"
    ws[f"AD{r}"] = f"=SUMIFS(rGasto,rAlim,gSI,{dd})"


def top(c0, src, n):
    """Top-N ordenado (sin funciones de matriz dinámica)."""
    a, b, c = (L(c0 + j) for j in range(3))
    for j, h in enumerate(["Top", "Monto", "Nº"]):
        put(ws, f"{L(c0 + j)}4", h, F(9, True, "FFFFFF"), "475569", align=CENTER)
    for j in range(1, n + 1):
        r = 4 + j
        pos = f"MATCH(LARGE({src['key']},{j}),{src['key']},0)"
        ws[f"{a}{r}"] = f'=IFERROR(IF(N(INDEX({src["val"]},{pos}))>0,INDEX({src["name"]},{pos})&"",""),"")'
        ws[f"{b}{r}"] = f'=IF({a}{r}="","",INDEX({src["val"]},{pos}))'
        ws[f"{c}{r}"] = f'=IF({a}{r}="","",INDEX({src["cnt"]},{pos}))'
    return a, b, c


T_GRP = top(32, B_GRP, 12)   # AF:AH
T_CAT = top(35, B_CAT, 8)    # AI:AK
T_CTA = top(38, B_CTA, 6)    # AL:AN
T_COM = top(41, B_COM, 5)    # AO:AQ

# ============================================================================
# RESUMEN (dashboard)
# ============================================================================
ws = RES
LAST = 18
header(ws, "Resumen", "Cifras en DOP · se actualiza solo", LAST)
ws.column_dimensions["A"].width = 2
for c in range(2, LAST):
    ws.column_dimensions[L(c)].width = 11
ws.column_dimensions[L(LAST)].width = 2
nav(ws, "BCEG")
area(ws, f"A3:{L(LAST)}86", fill_c=PAGE)

merge(ws, "B4:D4", "PERÍODO", font=F(8, True, MUTED), align=Alignment(vertical="bottom"))
merge(ws, "E4:G4", "MES (si eliges «Mes seleccionado»)", font=F(8, True, MUTED), align=Alignment(vertical="bottom"))
input_cell(ws, "B5:D5", "Todo el historial")
input_cell(ws, "E5:F5", "2026-10")
add_name("fPeriodo", "'Resumen'!$B$5")
add_name("fMes", "'Resumen'!$E$5")
dv_list(ws, "=lstPeriodo", "B5")
dv_list(ws, "=lstMes", "E5")
ws.row_dimensions[5].height = 26
merge(ws, "H5:Q5", '="Mostrando "&pTexto&"  ·  "&pDias&" días"', font=F(9, False, MUTED, True), align=LEFT)

GP = f"SUMIFS(rGasto,{DD})"
IP = f"SUMIFS(rIng,{DD})"
kpi(ws, 7, 2, 5, "Ingresos", f"={IP}", MONEY, "Salario e ingresos extra", AQUA)
kpi(ws, 7, 6, 9, "Gasto real", f"={GP}", MONEY,
    f'="Fijos RD$ "&FIXED(SUMIFS(rGasto,rClas,"Fijo",{DD}),0)&"  ·  variables RD$ "&FIXED(SUMIFS(rGasto,rClas,"Variable",{DD}),0)', ORANGE)
kpi(ws, 7, 10, 13, "Ahorro (ingresos − gasto)", f"={IP}-{GP}", MONEY,
    f'=IF({IP}>0,"Tasa de ahorro: "&FIXED(({IP}-{GP})/{IP}*100,0)&"%","Sin ingresos en el período")', BLUE)
NSAL = 'SUMPRODUCT((ctaLiq="Sí")*ISNUMBER(ctaDOP))'
kpi(ws, 7, 14, 17, "Dinero disponible", f'=IF({NSAL}=0,"Sin saldos",SUMIFS(ctaDOP,ctaLiq,"Sí"))', MONEY,
    f'=IF({NSAL}=0,"Escribe saldos iniciales en Ajustes para verlo","Deuda tarjeta: RD$ "&FIXED(-SUMIFS(ctaDOP,ctaTipo,"Tarjeta de crédito"),0))', VIOLET)
MES_G = 'SUMIFS(rGasto,rFecha,">="&mesHoy,rFecha,"<="&EOMONTH(diaHoy,0))'
kpi(ws, 12, 2, 5, "Hoy", "=SUMIFS(rGasto,rFecha,diaHoy)", MONEY, '=COUNTIFS(rFecha,diaHoy,rTipo,"Gasto")&" gastos"', ORANGE, 16)
kpi(ws, 12, 6, 9, "Esta semana", '=SUMIFS(rGasto,rFecha,">="&semHoy,rFecha,"<="&semHoy+6)', MONEY, "De lunes a domingo", ORANGE, 16)
kpi(ws, 12, 10, 13, "Este mes", f"={MES_G}", MONEY,
    f'=IF(N(limMensual)=0,"Sin límite mensual","Límite: "&FIXED({MES_G}/limMensual*100,0)&"% usado")', ORANGE, 16)
kpi(ws, 12, 14, 17, "Promedio diario", f"={GP}/pDias", MONEY, f'="Por semana: RD$ "&FIXED({GP}/pDias*7,0)', BLUE, 16)

merge(ws, "B17:Q17", "=budTexto", font=F(10, True, INK), fill_c=CARD, align=LEFT)
outline_box(ws, "B17:Q17")
ws.row_dimensions[17].height = 26
for lvl, bg, fg in ((3, BAD_BG, BAD), (2, WARN_BG, WARN), (1, GOOD_BG, GOOD)):
    ws.conditional_formatting.add("B17:Q17", FormulaRule(formula=[f"budNivel={lvl}"], fill=fill(bg), font=Font(color=fg, bold=True)))

section(ws, "B19:Q19", "¿En qué se va el dinero?")
card(ws, "B20:Q34")
put(ws, "B21", "Grupo", F(8, True, MUTED), align=LEFT)
put(ws, "E21", "Monto", F(8, True, MUTED), align=RIGHT)
put(ws, "F21", "%", F(8, True, MUTED), align=RIGHT)
for j in range(12):
    r = 22 + j
    merge(ws, f"B{r}:D{r}", f"='Motor'!{T_GRP[0]}{5 + j}", font=F(9, False, INK), align=LEFT)
    put(ws, f"E{r}", f"='Motor'!{T_GRP[1]}{5 + j}", F(9, True, INK), fmt=NUM, align=RIGHT)
    put(ws, f"F{r}", f'=IF(OR(E{r}="",{GP}=0),"",E{r}/{GP})', F(9, False, MUTED), fmt=PCT, align=RIGHT)
ch = BarChart()
ch.type = "bar"
ch.gapWidth = 45
ch.add_data(Reference(MOT, min_col=33, min_row=5, max_row=16), titles_from_data=False)
ch.set_categories(Reference(MOT, min_col=32, min_row=5, max_row=16))
style_chart(ch, legend=False, grid=False)
ch.x_axis.scaling.orientation = "maxMin"
ch.y_axis.delete = True
paint(ch.series[0], ORANGE, '#,##0;-#,##0;""')
ch.width, ch.height = 18, 7.4
ws.add_chart(ch, "H20")

section(ws, "B36:Q36", "Evolución  ·  cómo pago  ·  dónde gasto más")
card(ws, "B37:I52")
card(ws, "J37:Q52")
ch = BarChart()
ch.type = "col"
ch.gapWidth = 60
ch.add_data(Reference(MOT, min_col=27, min_row=4, max_row=10), titles_from_data=True)
ch.add_data(Reference(MOT, min_col=28, min_row=4, max_row=10), titles_from_data=True)
ch.set_categories(Reference(MOT, min_col=26, min_row=5, max_row=10))
style_chart(ch)
paint(ch.series[0], BLUE)
paint(ch.series[1], ORANGE)
ch.y_axis.numFmt = "#,##0"
ch.width, ch.height = 15.8, 7.8
ws.add_chart(ch, "B37")
put(ws, "J38", "CUENTA / MÉTODO", F(8, True, MUTED), align=LEFT)
put(ws, "O38", "Gasto", F(8, True, MUTED), align=RIGHT)
put(ws, "P38", "Veces", F(8, True, MUTED), align=RIGHT)
put(ws, "Q38", "%", F(8, True, MUTED), align=RIGHT)
for j in range(6):
    r = 39 + j
    merge(ws, f"J{r}:N{r}", f"='Motor'!{T_CTA[0]}{5 + j}", font=F(9, False, INK), align=LEFT)
    put(ws, f"O{r}", f"='Motor'!{T_CTA[1]}{5 + j}", F(9, True, INK), fmt=NUM, align=RIGHT)
    put(ws, f"P{r}", f"='Motor'!{T_CTA[2]}{5 + j}", F(9, False, MUTED), fmt=NUM, align=RIGHT)
    put(ws, f"Q{r}", f'=IF(OR(O{r}="",{GP}=0),"",O{r}/{GP})', F(9, False, MUTED), fmt=PCT, align=RIGHT)
put(ws, "J46", "COMERCIO / LUGAR (TOP 5)", F(8, True, MUTED), align=LEFT)
for j in range(5):
    r = 47 + j
    merge(ws, f"J{r}:N{r}", f"='Motor'!{T_COM[0]}{5 + j}", font=F(9, False, INK), align=LEFT)
    put(ws, f"O{r}", f"='Motor'!{T_COM[1]}{5 + j}", F(9, True, INK), fmt=NUM, align=RIGHT)
    put(ws, f"P{r}", f"='Motor'!{T_COM[2]}{5 + j}", F(9, False, MUTED), fmt=NUM, align=RIGHT)
ws.conditional_formatting.add("O39:O44", DataBarRule(start_type="num", start_value=0, end_type="max", color="9EC5F4", showValue=True))
ws.conditional_formatting.add("O47:O51", DataBarRule(start_type="num", start_value=0, end_type="max", color="F4B795", showValue=True))

section(ws, "B54:Q54", "Comida: comprar hecho vs. cocinar")
CPT, CPN = f"SUMIFS(rGasto,rAlim,gCP,{DD})", f'COUNTIFS(rAlim,gCP,rTipo,"Gasto",rMonto,">0",{DD})'
SIT = f"SUMIFS(rGasto,rAlim,gSI,{DD})"
ALT = f'SUMIFS(rGasto,rAlim,"?*",{DD})'
kpi(ws, 55, 2, 5, "Comida comprada", f"={CPT}", MONEY, f'={CPN}&" compras  ·  "&FIXED({CPN}/pDias*7,1)&" por semana"', ORANGE, 16)
kpi(ws, 55, 6, 9, "Súper e ingredientes", f"={SIT}", MONEY,
    f'="Comprada = "&IF({ALT}=0,0,FIXED({CPT}/{ALT}*100,0))&"% de lo gastado en comida"', BLUE, 16)
card(ws, "J55:Q58")
for k, m in enumerate(("Desayuno", "Almuerzo", "Cena")):
    a, b = L(10 + k * 3), L(11 + k * 3)
    s = f'SUMIFS(rGasto,rComida,"{m}",{DD})'
    n = f'COUNTIFS(rComida,"{m}",rMonto,">0",{DD})'
    put(ws, f"{a}55", f"{m.upper()} PROMEDIO", F(8, True, MUTED), align=Alignment(horizontal="left", vertical="bottom", indent=1))
    merge(ws, f"{a}56:{b}57", f'=IF({n}=0,"–",{s}/{n})', font=F(15, True, INK), fmt=MONEY, align=LEFT)
    put(ws, f"{a}58", f'={n}&" comprados"', F(8, False, MUTED), align=LEFT)
card(ws, "B60:Q64")
put(ws, "B61", "Comidas que cocinaste en casa en este período (opcional):", F(9, True, INK), align=LEFT)
input_cell(ws, "H61:I61", None, "0")
put(ws, "B62", "Simular reducir las comidas compradas en:", F(9, True, INK), align=LEFT)
input_cell(ws, "H62:I62", 0.25, "0%")
add_name("casaN", "'Resumen'!$H$61")
add_name("redPct", "'Resumen'!$H$62")
fuera = f"IF({CPN}=0,0,{CPT}/{CPN})"
casa = f"IF(N(casaN)=0,0,{SIT}/casaN)"
merge(ws, "K61:Q63",
      f'=IF({CPN}=0,"Aún no hay comidas compradas en el período.",'
      f'"Comida comprada: RD$ "&FIXED({fuera},0)&" c/u"&IF(N(casaN)>0,"  ·  en casa: RD$ "&FIXED({casa},0)&" c/u ("&IF({casa}<{fuera},"más barato","más caro")&" según tus datos)","")'
      f'&CHAR(10)&"Reducir "&FIXED(redPct*100,0)&"% ≈ "&FIXED({CPN}/pDias*30.4*redPct,0)&" comidas/mes  →  ahorro ≈ RD$ "'
      f'&FIXED({CPN}/pDias*30.4*redPct*MAX(0,{fuera}-{casa}),0)&" al mes"&IF(N(casaN)=0," (sin restar lo que cuesta cocinarlas)",""))',
      font=F(10, True, INK), fill_c=SOFT, align=TOP_WRAP)
note(ws, "B64", "Costo en casa = súper e ingredientes del período ÷ comidas que cocinaste. No se asume que cocinar sea más barato.")

section(ws, "B66:Q66", "Gastos hormiga  ·  trading  ·  flujo de efectivo")
HT, HN = f"SUMIFS(rGasto,rHorm,1,{DD})", f"COUNTIFS(rHorm,1,{DD})"
kpi(ws, 67, 2, 5, "Gastos hormiga", f"={HT}", MONEY,
    f'={HN}&" compras ≤ RD$ "&FIXED(Umbral,0)&"  ·  "&IF({GP}=0,0,FIXED({HT}/{GP}*100,0))&"% del gasto"', ORANGE, 16)
kpi(ws, 67, 6, 9, "Hormiga: proyección mensual", f"={HT}/pDias*30.4", MONEY, f'="Al año: RD$ "&FIXED({HT}/pDias*365,0)', ORANGE, 16)
TR = f"SUMIFS(rTrad,{DD})"
TC = f'SUMIFS(rGasto,rGrupo,"Trading",{DD})'
kpi(ws, 67, 10, 13, "Trading: resultado neto", f"={TR}-{TC}", MONEY,
    f'="Ganancia/pérdida RD$ "&FIXED({TR},0)&" − costos RD$ "&FIXED({TC},0)', VIOLET, 16)
kpi(ws, 67, 14, 17, "Flujo de efectivo neto", f"=SUMIFS(rFlujo,{DD})", MONEY,
    "Efectivo, banco y billetera. Incluye pagos de tarjeta y depósitos a trading.", BLUE, 16)

section(ws, "B73:Q73", "Top categorías del período")
card(ws, "B74:Q83")
for j in range(8):
    col0 = 2 if j < 4 else 10
    rr = 75 + (j % 4) * 2
    put(ws, f"{L(col0)}{rr}", j + 1, F(9, True, "94A3B8"), align=CENTER)
    merge(ws, f"{L(col0 + 1)}{rr}:{L(col0 + 4)}{rr}", f"='Motor'!{T_CAT[0]}{5 + j}", font=F(10, False, INK), align=LEFT)
    merge(ws, f"{L(col0 + 5)}{rr}:{L(col0 + 6)}{rr}", f"='Motor'!{T_CAT[1]}{5 + j}", font=F(10, True, INK), fmt=NUM, align=RIGHT)
    put(ws, f"{L(col0 + 7)}{rr}", f'=IF(\'Motor\'!{T_CAT[2]}{5 + j}="","",\'Motor\'!{T_CAT[2]}{5 + j}&"×")', F(9, False, MUTED), align=CENTER)
ws.freeze_panes = "A6"

# ============================================================================
# METAS (límites, objetivos, suscripciones)
# ============================================================================
ws = MET
header(ws, "Metas", "Etapa 2 · todo vacío hasta que lo necesites", 10)
for col, w in {"A": 2, "B": 32, "C": 16, "D": 16, "E": 12, "F": 18, "G": 18, "H": 18, "I": 2}.items():
    ws.column_dimensions[col].width = w
nav(ws, "BCDE")
merge(ws, "B3:H3", "Primero registra 1–2 meses. Luego usa «Promedio mensual» como referencia y escribe límites o metas en las celdas amarillas. Vacío = sin límite.",
      font=F(9, False, MUTED, True), align=LEFT_WRAP)
ws.row_dimensions[3].height = 30

section(ws, "B5:H5", "Límites de gasto (este mes)")
for j, h in enumerate(["Grupo", "Límite mensual", "Gasto este mes", "% usado", "Estado", "Promedio mensual", "Gasto esta semana"]):
    put(ws, f"{L(2 + j)}6", h, F(9, True, "FFFFFF"), INK if j < 2 else "475569", align=CENTER)
ws.row_dimensions[6].height = 28
groups = []
for c in CATS:
    if c[2] == "Gasto" and c[1] not in groups:
        groups.append(c[1])
rows = ["TOTAL (todos los gastos)"] + groups + [None] * 4
B0 = 7
for k, g in enumerate(rows):
    r = B0 + k
    total = k == 0
    put(ws, f"B{r}", g, F(10, True, INK), CARD if total else INPUT_BG, None, LEFT, BOX)
    input_cell(ws, f"C{r}", None, MONEY2)
    crit = "" if total else f"rGrupo,B{r},"
    put(ws, f"D{r}", f'=IF(B{r}="","",SUMIFS(rGasto,{crit}rFecha,">="&mesHoy,rFecha,"<="&EOMONTH(diaHoy,0)))', F(10, False, INK), SOFT, MONEY2, RIGHT, BOX)
    put(ws, f"E{r}", f'=IF(OR(B{r}="",N(C{r})=0),"",D{r}/C{r})', F(10, True, INK), SOFT, "0%", CENTER, BOX)
    put(ws, f"F{r}", f'=IF(B{r}="","",IF(N(C{r})=0,"—",IF(E{r}>1,"✕ Excedido",IF(E{r}>=0.8,"⚠ Cerca del límite","✓ En rango"))))', F(10, True, MUTED), SOFT, None, CENTER, BOX)
    put(ws, f"G{r}", f'=IF(B{r}="","",SUMIFS(rGasto,{crit}rFecha,">="&EDATE(mesHoy,-3),rFecha,"<"&mesHoy)/m3)', F(10, False, INK), SOFT, MONEY2, RIGHT, BOX)
    put(ws, f"H{r}", f'=IF(B{r}="","",SUMIFS(rGasto,{crit}rFecha,">="&semHoy,rFecha,"<="&semHoy+6))', F(10, False, INK), SOFT, MONEY2, RIGHT, BOX)
BL = B0 + len(rows) - 1
add_name("limMensual", f"'Metas'!$C${B0}")
note(ws, f"B{BL + 1}", "Promedio mensual = últimos 3 meses completos con datos. Alerta al llegar al 80 %.")
ws.conditional_formatting.add(f"E{B0}:E{BL}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9EC5F4", showValue=True))
for sym, bg, fg in (("✕", BAD_BG, BAD), ("⚠", WARN_BG, WARN), ("✓", GOOD_BG, GOOD)):
    ws.conditional_formatting.add(f"F{B0}:F{BL}", FormulaRule(formula=[f'LEFT(F{B0},1)="{sym}"'], fill=fill(bg), font=Font(color=fg, bold=True)))
grp_dv = DataValidation(type="list", formula1="=OFFSET('Motor'!$E$5,0,0,MAX(1,COUNTIF('Motor'!$E$5:$E$24,\"?*\")),1)", allow_blank=True)
grp_dv.add(f"B{B0 + 1}:B{BL}")
ws.add_data_validation(grp_dv)
# Estado del presupuesto para el Resumen (columna oculta)
put(ws, "K7", f"=COUNT(C{B0}:C{BL})")
put(ws, "K8", f'=COUNTIF(F{B0}:F{BL},"✕*")')
put(ws, "K9", f'=COUNTIF(F{B0}:F{BL},"⚠*")')
put(ws, "K10", "=IF(K7=0,0,IF(K8>0,3,IF(K9>0,2,1)))")
put(ws, "K11", '=IF(K7=0,"Etapa 1 · sin límites configurados: solo registra. Cuando conozcas tus patrones, define límites en «Metas».",'
               '"Presupuesto del mes  ·  excedidos: "&K8&"  ·  cerca del límite: "&K9&IF(N(limMensual)>0,"  ·  total usado: "&FIXED(E7*100,0)&"%",""))')
ws.column_dimensions["K"].hidden = True
add_name("budNivel", "'Metas'!$K$10")
add_name("budTexto", "'Metas'!$K$11")

ro = BL + 3
section(ws, f"B{ro}:H{ro}", "Objetivos de ahorro")
for j, h in enumerate(["Objetivo", "Meta (DOP)", "Ahorrado", "% avance", "Falta", "Fecha objetivo", "Aporte mensual necesario"]):
    put(ws, f"{L(2 + j)}{ro + 1}", h, F(9, True, "FFFFFF"), INK if j in (0, 1, 2, 5) else "475569", align=CENTER)
ws.row_dimensions[ro + 1].height = 28
O0 = ro + 2
for k, g in enumerate(["Fondo de emergencia", "Boda", "Inversión", None, None]):
    r = O0 + k
    put(ws, f"B{r}", g, F(10, True, INK), INPUT_BG, None, LEFT, BOX)
    input_cell(ws, f"C{r}", None, MONEY2)
    input_cell(ws, f"D{r}", None, MONEY2)
    input_cell(ws, f"G{r}", None, DATE)
    put(ws, f"E{r}", f'=IF(N(C{r})=0,"",MIN(1,N(D{r})/C{r}))', F(10, True, INK), SOFT, "0%", CENTER, BOX)
    put(ws, f"F{r}", f'=IF(N(C{r})=0,"",MAX(0,C{r}-N(D{r})))', F(10, False, INK), SOFT, MONEY2, RIGHT, BOX)
    put(ws, f"H{r}", f'=IF(OR(F{r}="",NOT(ISNUMBER(G{r}))),"",F{r}/MAX(1,DATEDIF(MIN(TODAY(),G{r}),G{r},"m")))', F(10, False, INK), SOFT, MONEY2, RIGHT, BOX)
ws.conditional_formatting.add(f"E{O0}:E{O0 + 4}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="8FD9B6", showValue=True))
note(ws, f"B{O0 + 5}", "Referencia fondo de emergencia: 3 a 6 meses de gasto. Tu gasto mensual promedio:")
put(ws, f"G{O0 + 5}", '=SUMIFS(rGasto,rFecha,">="&EDATE(mesHoy,-3),rFecha,"<"&mesHoy)/m3', F(10, True, INK), fmt=MONEY2, align=RIGHT)

rs = O0 + 8
section(ws, f"B{rs}:H{rs}", "Suscripciones y pagos recurrentes (agenda: no suma en reportes)")
for j, h in enumerate(["Servicio", "Monto", "Moneda", "Frecuencia", "Próximo cobro", "Costo mensual DOP", "Días para el cobro"]):
    put(ws, f"{L(2 + j)}{rs + 1}", h, F(9, True, "FFFFFF"), INK if j < 5 else "475569", align=CENTER)
ws.row_dimensions[rs + 1].height = 28
S0 = rs + 2
subs = [("DEMO Herramienta IA", 20, "USD", "Mensual", date(2026, 11, 3)), ("DEMO Internet", 2050, "DOP", "Mensual", date(2026, 11, 5)),
        ("DEMO Gimnasio", 2500, "DOP", "Mensual", date(2026, 11, 8))]
for k in range(12):
    r = S0 + k
    vals = subs[k] if k < len(subs) else (None,) * 5
    put(ws, f"B{r}", vals[0], F(10, True, INK), INPUT_BG, None, LEFT, BOX)
    input_cell(ws, f"C{r}", vals[1], MONEY2)
    input_cell(ws, f"D{r}", vals[2])
    input_cell(ws, f"E{r}", vals[3])
    input_cell(ws, f"F{r}", vals[4], DATE)
    m = f'IF(E{r}="Anual",12,1)'
    put(ws, f"G{r}", f'=IF(OR(B{r}="",C{r}=""),"",C{r}*IF(D{r}="USD",TasaUSD,1)/{m})', F(10, True, INK), SOFT, MONEY2, RIGHT, BOX)
    nxt = f'IF(F{r}>=TODAY(),F{r},EDATE(F{r},{m}*CEILING((DATEDIF(F{r},TODAY(),"m")+1)/{m},1)))'
    put(ws, f"H{r}", f'=IF(NOT(ISNUMBER(F{r})),"",{nxt}-TODAY())', F(10, True, INK), SOFT, '0" días"', CENTER, BOX)
SL = S0 + 11
put(ws, f"F{SL + 1}", "Total mensual:", F(9, True, MUTED), align=RIGHT)
put(ws, f"G{SL + 1}", f"=SUM(G{S0}:G{SL})", F(11, True, INK), fmt=MONEY2, align=RIGHT)
note(ws, f"B{SL + 1}", "Cada cobro real se anota en Registro. Si la fecha ya pasó, se calcula el siguiente cobro.")
ws.conditional_formatting.add(f"H{S0}:H{SL}", FormulaRule(formula=[f'AND(ISNUMBER(H{S0}),H{S0}<=7)'], fill=fill(WARN_BG), font=Font(color=WARN, bold=True)))
ws.conditional_formatting.add(f"B{S0}:F{SL}", FormulaRule(formula=[f'LEFT($B{S0},4)="DEMO"'], fill=fill(DEMO_BG)))
dv_list(ws, "=lstMon", f"D{S0}:D{SL}")
dv_list(ws, "=lstFrec", f"E{S0}:E{SL}")
ws.freeze_panes = "A3"

# ----------------------------------------------------------------------------
MOT.sheet_state = "hidden"
if os.environ.get("PREVIEW"):
    REG.sheet_state = "hidden"
for w in (RES, MET, AJU):
    w.page_setup.orientation = "landscape"
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
AJU.print_area = f"A1:H{CAT0 + len(CATS)}"
wb.active = 0
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("OK ->", OUT, "| filas DEMO:", len(DEMO), "| última fila DEMO:", DEMO_LAST)
