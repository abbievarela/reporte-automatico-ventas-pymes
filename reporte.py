import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os

# ─────────────────────────────────────────
# PASO A: Leer el Excel sucio
# ─────────────────────────────────────────
print("📂 Leyendo archivo...")
df = pd.read_excel("ventas_raw.xlsx")

# ─────────────────────────────────────────
# PASO B: Limpiar los datos
# ─────────────────────────────────────────
print("🧹 Limpiando datos...")

# Eliminar filas donde faltan datos clave
df = df.dropna(subset=["cliente", "producto", "cantidad", "precio_unitario"])

# Convertir fechas (acepta múltiples formatos)
df["fecha"] = pd.to_datetime(df["fecha"], dayfirst=True, errors="coerce")
df = df.dropna(subset=["fecha"])

# Estandarizar nombres (todo en minúscula y sin espacios extra)
df["cliente"] = df["cliente"].str.strip().str.title()
df["producto"] = df["producto"].str.strip().str.title()

# Calcular total por fila
df["total"] = df["cantidad"] * df["precio_unitario"]

# Agregar columna de mes
df["mes"] = df["fecha"].dt.to_period("M").astype(str)

# ─────────────────────────────────────────
# PASO C: Calcular los KPIs
# ─────────────────────────────────────────
print("📊 Calculando métricas...")

total_ventas = df["total"].sum()
ticket_promedio = df["total"].mean()
total_transacciones = len(df)

top_productos = (
    df.groupby("producto")["total"]
    .sum()
    .sort_values(ascending=False)
    .head(5)
    .reset_index()
)
top_productos.columns = ["Producto", "Total Vendido"]

ventas_por_mes = (
    df.groupby("mes")["total"]
    .sum()
    .reset_index()
)
ventas_por_mes.columns = ["Mes", "Total"]

clientes_frecuentes = (
    df.groupby("cliente")["total"]
    .agg(["sum", "count"])
    .sort_values("sum", ascending=False)
    .reset_index()
)
clientes_frecuentes.columns = ["Cliente", "Total Comprado", "Cantidad de Órdenes"]

# ─────────────────────────────────────────
# PASO D: Crear gráfico
# ─────────────────────────────────────────
print("📈 Generando gráfico...")

fig, ax = plt.subplots(figsize=(8, 4))
ax.bar(ventas_por_mes["Mes"], ventas_por_mes["Total"], color="#4F81BD")
ax.set_title("Ventas por Mes", fontsize=14, fontweight="bold")
ax.set_xlabel("Mes")
ax.set_ylabel("Total ($)")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("grafico_ventas.png", dpi=150)
plt.close()

# ─────────────────────────────────────────
# PASO E: Crear el Excel de salida
# ─────────────────────────────────────────
print("📝 Creando reporte Excel...")

wb = Workbook()

# — Hoja 1: Resumen ejecutivo —
ws1 = wb.active
ws1.title = "Resumen"

color_header = "1F3864"
color_azul_claro = "D6E4F0"

def escribir_titulo(ws, fila, texto):
    ws.cell(row=fila, column=1, value=texto).font = Font(bold=True, size=13, color="FFFFFF")
    ws.cell(row=fila, column=1).fill = PatternFill("solid", fgColor=color_header)
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)

def escribir_kpi(ws, fila, label, valor):
    ws.cell(row=fila, column=1, value=label).font = Font(bold=True)
    ws.cell(row=fila, column=1).fill = PatternFill("solid", fgColor=color_azul_claro)
    ws.cell(row=fila, column=2, value=valor)

escribir_titulo(ws1, 1, "📊 REPORTE AUTOMÁTICO DE VENTAS")
escribir_kpi(ws1, 3, "Total de ventas", f"$ {total_ventas:,.0f}")
escribir_kpi(ws1, 4, "Ticket promedio", f"$ {ticket_promedio:,.0f}")
escribir_kpi(ws1, 5, "Transacciones registradas", total_transacciones)

ws1.column_dimensions["A"].width = 30
ws1.column_dimensions["B"].width = 25

# — Hoja 2: Top Productos —
ws2 = wb.create_sheet("Top Productos")
escribir_titulo(ws2, 1, "🏆 TOP PRODUCTOS POR VENTA")
for col_num, col_name in enumerate(top_productos.columns, 1):
    ws2.cell(row=2, column=col_num, value=col_name).font = Font(bold=True)
for row_num, row in top_productos.iterrows():
    ws2.cell(row=row_num+3, column=1, value=row["Producto"])
    ws2.cell(row=row_num+3, column=2, value=row["Total Vendido"])
ws2.column_dimensions["A"].width = 25
ws2.column_dimensions["B"].width = 20

# — Hoja 3: Ventas por Mes —
ws3 = wb.create_sheet("Ventas por Mes")
escribir_titulo(ws3, 1, "📅 VENTAS MENSUALES")
for col_num, col_name in enumerate(ventas_por_mes.columns, 1):
    ws3.cell(row=2, column=col_num, value=col_name).font = Font(bold=True)
for row_num, row in ventas_por_mes.iterrows():
    ws3.cell(row=row_num+3, column=1, value=row["Mes"])
    ws3.cell(row=row_num+3, column=2, value=row["Total"])
ws3.column_dimensions["A"].width = 15
ws3.column_dimensions["B"].width = 20

# — Hoja 4: Clientes —
ws4 = wb.create_sheet("Clientes")
escribir_titulo(ws4, 1, "👥 CLIENTES FRECUENTES")
for col_num, col_name in enumerate(clientes_frecuentes.columns, 1):
    ws4.cell(row=2, column=col_num, value=col_name).font = Font(bold=True)
for row_num, row in clientes_frecuentes.iterrows():
    ws4.cell(row=row_num+3, column=1, value=row["Cliente"])
    ws4.cell(row=row_num+3, column=2, value=row["Total Comprado"])
    ws4.cell(row=row_num+3, column=3, value=row["Cantidad de Órdenes"])
ws4.column_dimensions["A"].width = 25
ws4.column_dimensions["B"].width = 22
ws4.column_dimensions["C"].width = 22

# — Insertar gráfico en Resumen —
from openpyxl.drawing.image import Image as XLImage
img = XLImage("grafico_ventas.png")
img.width = 480
img.height = 240
ws1.add_image(img, "D3")

wb.save("reporte_ventas_final.xlsx")
print("✅ ¡Listo! Se generó el archivo 'reporte_ventas_final.xlsx'")