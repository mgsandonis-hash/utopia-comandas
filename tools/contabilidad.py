# Genera la hoja de contabilidad de Utopía a partir del libro guardado en Supabase (docs: col=contabilidad, id=libro).
# Uso: python3 tools/contabilidad.py libro.json Utopia-contabilidad.xlsx
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from datetime import date
import json, sys
DATA=json.load(open(sys.argv[1]))
OUT=sys.argv[2] if len(sys.argv)>2 else "Utopia-contabilidad.xlsx"

WINE="5E0F1C"; CREAM="F1E9DB"; SURF="FAF5EC"; SOFT="ECDCD6"; LINE="DCC6BD"; MUTED="8A4A52"
EUR='#,##0.00 "€";-#,##0.00 "€";"-"'
F=lambda **k: Font(name="Arial", **k)
thin=Side(style="thin", color=LINE)
wb=Workbook()

# ---------- Movimientos ----------
mv=wb.active; mv.title="Movimientos"
mv.sheet_view.showGridLines=False
mv.merge_cells("A1:G1"); mv["A1"]="UTOPÍA COFFEE CLUB · CONTABILIDAD · MOVIMIENTOS"
mv["A1"].font=F(bold=True, size=16, color=CREAM); mv["A1"].fill=PatternFill("solid", fgColor=WINE); mv.row_dimensions[1].height=30
mv["A1"].alignment=Alignment(vertical="center", indent=1)
mv.merge_cells("A2:G2"); mv["A2"]="Apunta una fila por movimiento. Tipo: elige de la lista. Entrada = dinero que entra; Salida = dinero que sale. Texto azul = datos apuntados a mano."
mv["A2"].font=F(italic=True, size=9, color=MUTED)
heads=["Fecha","Concepto","Tipo","Forma de pago","Entrada (€)","Salida (€)","Notas"]
for i,h in enumerate(heads,1):
    c=mv.cell(row=4,column=i,value=h); c.font=F(bold=True,color=CREAM); c.fill=PatternFill("solid",fgColor=WINE)
    c.alignment=Alignment(horizontal="right" if i in (5,6) else "left", vertical="center", indent=0 if i in (5,6) else 1)
mv.row_dimensions[4].height=22
def _d(x):
    if not x: return None
    y,m,dd=map(int,x.split("-")); return date(y,m,dd)
rows=[(_d(m.get("fecha")), m.get("concepto"), m.get("tipo"), m.get("forma",""), m.get("entrada"), m.get("salida"), m.get("nota","")) for m in DATA["movimientos"]]
N=300; first=5; last=first+N-1
for i in range(N):
    r=first+i; fill=PatternFill("solid",fgColor=SURF if i%2 else "FFFFFF")
    vals=rows[i] if i<len(rows) else (None,)*7
    for j in range(7):
        c=mv.cell(row=r,column=j+1,value=vals[j]); c.fill=fill; c.border=Border(bottom=thin)
        c.font=F(color="0000FF" if (i<len(rows) and j in (0,1,4,5)) else "000000", size=10)
        if j==0: c.number_format="DD/MM/YYYY"; c.alignment=Alignment(indent=1)
        if j in (4,5): c.number_format=EUR
        if j==6: c.font=F(size=9, italic=True, color=MUTED)
for col,w in zip("ABCDEFG",[13,34,18,15,14,14,58]): mv.column_dimensions[col].width=w
mv.freeze_panes="A5"
dv=DataValidation(type="list", formula1='"Inversión,Recuperación,Venta,Otro ingreso,Gasto,Retirada de caja,Aporte a caja"', allow_blank=True)
dv2=DataValidation(type="list", formula1='"Efectivo,Bizum,Banco,Tarjeta"', allow_blank=True)
mv.add_data_validation(dv); mv.add_data_validation(dv2); dv.add(f"C{first}:C{last}"); dv2.add(f"D{first}:D{last}")
mv.auto_filter.ref=f"A4:G{last}"
RNG=lambda col: f"Movimientos!${col}${first}:${col}${last}"

# ---------- Resumen ----------
rs=wb.create_sheet("Resumen", 0); rs.sheet_view.showGridLines=False
rs.merge_cells("A1:C1"); rs["A1"]="UTOPÍA COFFEE CLUB · CONTABILIDAD"
rs["A1"].font=F(bold=True,size=18,color=CREAM); rs["A1"].fill=PatternFill("solid",fgColor=WINE); rs["A1"].alignment=Alignment(vertical="center",indent=1); rs.row_dimensions[1].height=34
rs.merge_cells("A2:C2"); rs["A2"]="Todo se calcula solo a partir de la hoja Movimientos."; rs["A2"].font=F(italic=True,size=9,color=MUTED)
rs.column_dimensions["A"].width=46; rs.column_dimensions["B"].width=16; rs.column_dimensions["C"].width=60
def head(r,t):
    rs.merge_cells(f"A{r}:C{r}"); c=rs[f"A{r}"]; c.value=t; c.font=F(bold=True,size=12,color=WINE); c.border=Border(bottom=Side(style="medium",color=WINE)); rs.row_dimensions[r].height=22
def line(r,label,formula,note="",bold=False,fmt=EUR,fill=None,color="000000"):
    rs[f"A{r}"]=label; rs[f"B{r}"]=formula; rs[f"C{r}"]=note
    rs[f"A{r}"].font=F(bold=bold,size=10); rs[f"B{r}"].font=F(bold=bold,size=10,color=color); rs[f"C{r}"].font=F(italic=True,size=9,color=MUTED)
    rs[f"B{r}"].number_format=fmt; rs[f"A{r}"].alignment=Alignment(indent=1)
    for col in "ABC":
        rs[f"{col}{r}"].border=Border(bottom=thin)
        if fill: rs[f"{col}{r}"].fill=PatternFill("solid",fgColor=fill)
S=lambda tipo, col: f'=SUMIF({RNG("C")},"{tipo}",{RNG(col)})'
head(4,"La inversión")
line(5,"Inversión inicial",S("Inversión","F"),"Lo que pusisteis para montar la cafetería")
line(6,"Recuperado fuera de la cafetería",S("Recuperación","E"),"Venta de la antigua cafetera, cata de café y evento")
line(7,"Pendiente antes de contar la cafetería","=B5-B6","",bold=True)
head(9,"La cafetería")
line(10,"Ventas",S("Venta","E"),"Efectivo + Bizum, de los cierres de caja")
line(11,"Otros ingresos (propinas, bote…)",S("Otro ingreso","E"),"Revisa el «Bote café» en Movimientos")
line(12,"Gastos",S("Gasto","F"),"Compras y gastos apuntados")
line(13,"Beneficio de la cafetería","=B10+B11-B12","Ventas + otros ingresos − gastos",bold=True,fill=SOFT,color=WINE)
head(15,"Recuperación de la inversión")
line(16,"Recuperado en total","=B6+B13","Fuera de la cafetería + beneficio de la cafetería",bold=True)
line(17,"Falta por recuperar","=MAX(0,B5-B16)","",bold=True,fill=SOFT,color=WINE)
line(18,"% recuperado","=IF(B5=0,0,MIN(1,B16/B5))","",bold=True,fmt="0.0%")
head(20,"Dónde está el dinero de la cafetería")
line(21,"Fondo de cambio en el cajón",DATA.get("fondo",150),"Fondo de caja acordado (dato de Mario). Cámbialo si cambia el fondo",color="0000FF")
line(22,"Sacado del cajón",S("Retirada de caja","F"),"Retiradas de caja apuntadas en Movimientos")
line(23,"Cobrado por Bizum",f'=SUMIFS({RNG("E")},{RNG("C")},"Venta",{RNG("D")},"Bizum")',"Está en la cuenta del Bizum")
head(25,"Objetivo: entrada del piso")
line(26,"Dinero para la entrada del piso",DATA.get("objetivo") or None,"Pon aquí cuánto necesitáis",color="0000FF",fill="FFFF00" if not DATA.get("objetivo") else None)
line(27,"Dinero reunido",DATA.get("reunido_formula","=B6+B22+B23"),"Recuperado fuera de la cafetería + sacado del cajón + Bizum (dinero que ya no está en la cafetería)",bold=True)
line(28,"Falta","=IF(B26>0,MAX(0,B26-B27),0)","",bold=True,fill=SOFT,color=WINE)
line(29,"% conseguido","=IF(B26>0,MIN(1,B27/B26),0)","",bold=True,fmt="0.0%")
r=31
rs.merge_cells(f"A{r}:C{r}"); rs[f"A{r}"]="Cómo usarla: añade cada movimiento en una fila nueva de Movimientos (fecha, concepto, tipo de la lista, forma de pago y el importe en Entrada o en Salida). Este resumen se actualiza solo. Ejemplo: «03/11/2026 · Ventas de la cafetería · Venta · Efectivo · 48,20»."
rs[f"A{r}"].font=F(italic=True,size=9,color=MUTED); rs[f"A{r}"].alignment=Alignment(wrap_text=True,vertical="top"); rs.row_dimensions[r].height=40
wb.save(OUT)
