# generar_reporte.py
import os
import tempfile
import shutil
from datetime import time, datetime, timedelta
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------- Configuración ----------------
DESAYUNO_START = time(8, 30)
DESAYUNO_END = time(12, 15)
COMIDA_START = time(12, 25)
COMIDA_END = time(16, 30)
PRECIO_DESAYUNO = 84 #120 real
PRECIO_COMIDA = 98 #140 real

# Tolerancia en minutos para aceptar marcas cercanas al inicio/fin
TOLERANCIA_MIN = 5  # cambia a 0 si no quieres tolerancia

# Archivos de entrada (debe haberlos en la misma carpeta que el script)
REGISTROS_TXT = 'registros.txt'
EXCEL_PADRON = 'BaseDeDatos_2026_2.xlsx'
SHEET_PADRON = 'Base de datos (nueva)'

# Carpeta de salida (por defecto la carpeta actual)
OUT_DIR = os.path.abspath('.')

# Nombre de archivo de fuente Roboto Slab (colócala en la carpeta del script si la tienes)
ROBOTO_SLAB_TTF = 'RobotoSlab-Regular.ttf'

# ---------------- Funciones auxiliares ----------------
def dentro_con_tolerancia(hora, start, end, tol_min=TOLERANCIA_MIN):
    if pd.isna(hora):
        return False
    dt = datetime.combine(datetime.today(), hora)
    s = datetime.combine(datetime.today(), start) - timedelta(minutes=tol_min)
    e = datetime.combine(datetime.today(), end) + timedelta(minutes=tol_min)
    return s <= dt <= e

def asignar_servicio(h):
    if pd.isna(h):
        return 'Otro'
    if dentro_con_tolerancia(h, DESAYUNO_START, DESAYUNO_END):
        return 'Desayuno'
    if dentro_con_tolerancia(h, COMIDA_START, COMIDA_END):
        return 'Comida'
    return 'Otro'

def safe_save(func_save, target_path):
    """
    Guarda usando un archivo temporal en la misma carpeta y luego mueve al destino.
    Crea el temporal con la misma extensión que target_path para que pandas/otros detecten el engine.
    func_save debe aceptar una ruta (path) y escribir el archivo en esa ruta.
    """
    try:
        dirpath = os.path.dirname(os.path.abspath(target_path)) or '.'
        _, ext = os.path.splitext(target_path)
        fd, tmp = tempfile.mkstemp(suffix=ext, dir=dirpath)
        os.close(fd)
        try:
            func_save(tmp)
            if os.path.exists(target_path):
                os.remove(target_path)
            shutil.move(tmp, target_path)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)
    except PermissionError as e:
        raise PermissionError(f"No se pudo escribir {target_path}. Cierra el archivo si está abierto y vuelve a intentar. Detalle: {e}")
    except Exception:
        raise

# ---------------- Registrar fuente (Roboto Slab) y estilos ----------------
FONT_NAME = 'RobotoSlab'
FALLBACK_FONT = 'Helvetica'
if os.path.exists(os.path.join(OUT_DIR, ROBOTO_SLAB_TTF)):
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, os.path.join(OUT_DIR, ROBOTO_SLAB_TTF)))
        base_font = FONT_NAME
    except Exception:
        base_font = FALLBACK_FONT
else:
    base_font = FALLBACK_FONT

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TitleCustom', parent=styles['Title'], fontName=base_font, fontSize=18, leading=22, textColor=HexColor('#1F4E79')))
styles.add(ParagraphStyle(name='Heading2Custom', parent=styles['Heading2'], fontName=base_font, fontSize=14, leading=18, textColor=HexColor('#0B6E4F')))
styles.add(ParagraphStyle(name='Heading3Custom', parent=styles['Heading3'], fontName=base_font, fontSize=12, leading=15, textColor=HexColor('#333333')))
styles.add(ParagraphStyle(name='NormalCustom', parent=styles['Normal'], fontName=base_font, fontSize=10, leading=13, textColor=HexColor('#222222')))

# ---------------- Cargar registros del checador ----------------
df = pd.read_csv(REGISTROS_TXT, sep=r'\t+', engine='python', header=0,
                 names=['ID','Nombre','Depart','Tiempo','ID_dispositivo'], skipinitialspace=True)
df['Tiempo'] = df['Tiempo'].astype(str).str.strip()
df['FechaHora'] = pd.to_datetime(df['Tiempo'], format='%d/%m/%Y %H:%M:%S', dayfirst=True, errors='coerce')
df = df.dropna(subset=['FechaHora']).copy()
df['Fecha'] = df['FechaHora'].dt.date
df['Hora'] = df['FechaHora'].dt.time
df['ID'] = pd.to_numeric(df['ID'], errors='coerce').astype('Int64')

# Asignar servicio con tolerancia
df['Servicio'] = df['Hora'].apply(asignar_servicio)

# ---------------- Cargar padrón y normalizar columnas ----------------
padron = pd.read_excel(EXCEL_PADRON, sheet_name=SHEET_PADRON, engine='openpyxl', header=0)
padron.columns = [str(c).strip().replace('.', '').replace(' ', '_').lower() for c in padron.columns]

# Detectar columna ID (variantes) y renombrar a 'ID'
id_cols = [c for c in padron.columns if c.startswith('id')]
if not id_cols:
    raise SystemExit("No se encontró ninguna columna que parezca 'ID' en la hoja. Columnas disponibles: " + ", ".join(padron.columns))
id_col = id_cols[0]
padron = padron.rename(columns={id_col: 'ID'})

# Detectar columna aportación (variantes) y renombrar a 'aportacion'
aport_cols = [c for c in padron.columns if 'aport' in c]
if aport_cols:
    padron = padron.rename(columns={aport_cols[0]: 'aportacion'})
else:
    padron['aportacion'] = 0

# Asegurar columnas mínimas
padron['ID'] = pd.to_numeric(padron['ID'], errors='coerce').astype('Int64')
padron['aportacion'] = pd.to_numeric(padron['aportacion'], errors='coerce').fillna(0)

# Normalizar columna de nombre si existe
name_cols = [c for c in padron.columns if 'nombre' in c]
if name_cols:
    padron = padron.rename(columns={name_cols[0]: 'nombre'})
else:
    padron['nombre'] = padron.get('nombre', pd.NA)

# Preparar padron para merge (columnas con nombres en mayúscula como en el txt)
padron_for_merge = padron.rename(columns={'nombre': 'Nombre', 'aportacion': 'Aportación'})

# ---------------- Unir aportaciones por ID o por Nombre ----------------
df = df.merge(padron_for_merge[['ID','Nombre','Aportación']], on='ID', how='left', suffixes=('','_pad'))

mask_noaport = df['Aportación'].isna()
if mask_noaport.any():
    tmp = df[mask_noaport].merge(padron_for_merge[['Nombre','Aportación']], left_on='Nombre', right_on='Nombre', how='left')
    df.loc[mask_noaport, 'Aportación'] = tmp['Aportación'].fillna(0).values

df['Aportación'] = df['Aportación'].fillna(0)

# ---------------- Agregados y cálculos ----------------
df['ID_para_contar'] = df['ID'].fillna(df['Nombre'])
resumen = df.groupby(['Fecha','Servicio']).agg(
    Asistentes=('ID_para_contar','nunique'),
    Registros=('ID_para_contar','count')
).reset_index()

def precio_por_servicio(s):
    return PRECIO_DESAYUNO if s=='Desayuno' else (PRECIO_COMIDA if s=='Comida' else 0)

resumen['PrecioUnitario'] = resumen['Servicio'].apply(precio_por_servicio)
resumen['TotalBruto'] = resumen['Asistentes'] * resumen['PrecioUnitario']

aportado_por_fecha_servicio = df.groupby(['Fecha','Servicio']).agg(
    SumaAportaciones=('Aportación','sum')
).reset_index()
resumen = resumen.merge(aportado_por_fecha_servicio, on=['Fecha','Servicio'], how='left')
resumen['SumaAportaciones'] = resumen['SumaAportaciones'].fillna(0)
resumen['TotalNeto'] = resumen['TotalBruto'] - resumen['SumaAportaciones']

tot_desayunos = resumen[resumen['Servicio']=='Desayuno']['Asistentes'].sum()
tot_comidas = resumen[resumen['Servicio']=='Comida']['Asistentes'].sum()
costo_desayunos = tot_desayunos * PRECIO_DESAYUNO
costo_comidas = tot_comidas * PRECIO_COMIDA
total_bruto = costo_desayunos + costo_comidas
total_aportes = df['Aportación'].sum()
# total_neto (para pago) = total bruto - aportaciones (según tu instrucción)
total_neto = total_bruto - total_aportes

asistencias_por_becado = df.groupby(['ID','Nombre','Servicio']).size().unstack(fill_value=0).reset_index()
if 'Desayuno' not in asistencias_por_becado.columns:
    asistencias_por_becado['Desayuno'] = 0
if 'Comida' not in asistencias_por_becado.columns:
    asistencias_por_becado['Comida'] = 0
asistencias_por_becado['TotalAsistencias'] = asistencias_por_becado['Desayuno'] + asistencias_por_becado['Comida']

# ---------------- Desglose fiscal y retenciones (base = total_neto) ----------------
# Partidas y porcentajes
PORC_IVA = 0.16
PORC_RET_IVA = 2/3        # retención sobre el IVA
PORC_RET_ISR = 0.0125     # 1.25% sobre la base

base_factura = float(total_neto)  # ahora la base es el total neto restando aportaciones
iva = round(base_factura * PORC_IVA, 2)
ret_iva = round(iva * PORC_RET_IVA, 2)
ret_isr = round(base_factura * PORC_RET_ISR, 2)
neto_factura = round(base_factura + iva - ret_iva - ret_isr, 2)

desglose_factura = pd.DataFrame([{
    'Base': base_factura,
    'IVA_16%': iva,
    'Retencion_IVA_2_3': ret_iva,
    'Retencion_ISR_1_25%': ret_isr,
    'Neto_Factura': neto_factura
}])

# ---------------- Preparar nombres de archivo con periodo en formato español DD-MM-YYYY ----------------
fechas = pd.to_datetime(df['Fecha']).dt.date
if fechas.empty:
    inicio = fin = datetime.now().date()
else:
    inicio = fechas.min()
    fin = fechas.max()

inicio_str = inicio.strftime('%d-%m-%Y')
fin_str = fin.strftime('%d-%m-%Y')
if inicio == fin:
    periodo_str = inicio_str
else:
    periodo_str = f"{inicio_str}_a_{fin_str}"

pdf_name = f"reporte_periodo_{periodo_str}.pdf"
resumen_xlsx = f"resumen_diario_{periodo_str}.xlsx"
detalle_csv = f"detalle_asistencias_{periodo_str}.csv"
asist_xlsx = f"asistencias_por_becado_{periodo_str}.xlsx"
desglose_xlsx = f"desglose_factura_{periodo_str}.xlsx"

pdf_path = os.path.join(OUT_DIR, pdf_name)
resumen_path = os.path.join(OUT_DIR, resumen_xlsx)
detalle_path = os.path.join(OUT_DIR, detalle_csv)
asist_path = os.path.join(OUT_DIR, asist_xlsx)
desglose_path = os.path.join(OUT_DIR, desglose_xlsx)

# ---------------- Exportar auxiliares con safe_save ----------------
def save_resumen_excel(path):
    # Escribe resumen y añade hoja 'Desglose' con el desglose fiscal
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
        # Añadir hoja con desglose
        desglose_factura.to_excel(writer, sheet_name='Desglose', index=False)

def save_detalle_csv(path):
    df.to_csv(path, index=False)

def save_asistencias_excel(path):
    asistencias_por_becado.to_excel(path, index=False)

def save_desglose_excel(path):
    desglose_factura.to_excel(path, index=False)

safe_save(save_resumen_excel, resumen_path)
safe_save(save_detalle_csv, detalle_path)
safe_save(save_asistencias_excel, asist_path)
safe_save(save_desglose_excel, desglose_path)

# ---------------- Generar PDF (estilizado) ----------------
styles_pdf = styles
flow = []

# Header block with color bar
header_color = HexColor('#0B6E4F')  # deep green
accent_color = HexColor('#1F4E79')  # deep blue
muted_gray = HexColor('#F3F6F8')

flow.append(Spacer(1, 6))

# Title
flow.append(Paragraph("Reporte de asistencia y costos", styles_pdf['TitleCustom']))
flow.append(Spacer(1, 6))
flow.append(Paragraph(f"<b>Periodo:</b> {inicio_str} a {fin_str}", styles_pdf['NormalCustom']))
flow.append(Spacer(1, 8))

# Summary card (boxed)
summary_table = [
    ['Concepto', 'Cantidad', 'Costo (MXN)'],
    ['Total desayunos servidos', str(tot_desayunos), f"${costo_desayunos:.2f}"],
    ['Total comidas servidas', str(tot_comidas), f"${costo_comidas:.2f}"],
    ['Total bruto (costos)', '', f"${total_bruto:.2f}"],
    ['Total aportaciones recibidas', '', f"-${total_aportes:.2f}"],
    ['Total neto (bruto - aportaciones)', '', f"${total_neto:.2f}"]
]
tbl = Table(summary_table, colWidths=[140*mm/3, 50*mm/3, 70*mm/3], hAlign='LEFT')
tbl.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), accent_color),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('FONTNAME', (0,0), (-1,0), base_font),
    ('FONTSIZE', (0,0), (-1,0), 10),
    ('ALIGN', (1,1), (-1,-1), 'CENTER'),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#D6E3EE')),
    ('BACKGROUND', (0,1), (-1,-1), muted_gray),
    ('FONTNAME', (0,1), (-1,-1), base_font),
]))
flow.append(tbl)
flow.append(Spacer(1, 12))

# Desglose fiscal card
flow.append(Paragraph("Desglose fiscal de la factura", styles_pdf['Heading2Custom']))
desglose_table = [
    ['Partida', 'Importe (MXN)'],
    ['Base (total neto)', f"${base_factura:.2f}"],
    ['IVA (16%)', f"${iva:.2f}"],
    ['Retención IVA (2/3 del IVA)', f"-${ret_iva:.2f}"],
    ['Retención ISR (1.25% sobre la base)', f"-${ret_isr:.2f}"],
    ['Neto factura', f"${neto_factura:.2f}"]
]
dt = Table(desglose_table, colWidths=[120*mm/2, 80*mm/2], hAlign='LEFT')
dt.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), header_color),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('FONTNAME', (0,0), (-1,0), base_font),
    ('FONTSIZE', (0,0), (-1,0), 10),
    ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#E6EEF6')),
    ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
    ('FONTNAME', (0,1), (-1,-1), base_font),
]))
flow.append(dt)
flow.append(Spacer(1, 12))

# Tabla resumen diario (con estilo alternado)
flow.append(Paragraph("Resumen diario (servicio)", styles_pdf['Heading2Custom']))
table_data = [['Fecha','Servicio','Asistentes','PrecioUnitario','TotalBruto','SumaAportaciones','TotalNeto']]
for _, r in resumen.sort_values(['Fecha','Servicio']).iterrows():
    table_data.append([
        str(r['Fecha']),
        r['Servicio'],
        int(r['Asistentes']),
        f"${r['PrecioUnitario']}",
        f"${r['TotalBruto']:.2f}",
        f"${r['SumaAportaciones']:.2f}",
        f"${r['TotalNeto']:.2f}"
    ])
col_widths = [30*mm, 30*mm, 25*mm, 30*mm, 30*mm, 35*mm, 35*mm]
t = Table(table_data, colWidths=col_widths, repeatRows=1)
# Alternating row colors
row_count = len(table_data)
table_style = [
    ('BACKGROUND',(0,0),(-1,0), accent_color),
    ('TEXTCOLOR',(0,0),(-1,0), colors.white),
    ('FONTNAME',(0,0),(-1,0), base_font),
    ('FONTSIZE',(0,0),(-1,0),10),
    ('GRID',(0,0),(-1,-1),0.4,HexColor('#D6E3EE')),
    ('ALIGN',(2,1),(-1,-1),'CENTER'),
    ('FONTNAME',(0,1),(-1,-1), base_font),
]
for i in range(1, row_count):
    bg = HexColor('#FFFFFF') if i % 2 == 0 else HexColor('#F7FBFF')
    table_style.append(('BACKGROUND', (0,i), (-1,i), bg))
t.setStyle(TableStyle(table_style))
flow.append(t)
flow.append(Spacer(1, 12))

# Tabla por becado (muestra top 40; Excel contiene todo)
flow.append(Paragraph("Asistencias por becado (resumen)", styles_pdf['Heading2Custom']))
table_b = [['ID','Nombre','Desayuno','Comida','TotalAsistencias']]
for _, r in asistencias_por_becado.sort_values('TotalAsistencias', ascending=False).head(40).iterrows():
    table_b.append([str(r['ID']), r['Nombre'], int(r['Desayuno']), int(r['Comida']), int(r['TotalAsistencias'])])
tb = Table(table_b, colWidths=[20*mm, 80*mm, 25*mm, 25*mm, 30*mm], repeatRows=1)
tb.setStyle(TableStyle([
    ('BACKGROUND',(0,0),(-1,0), accent_color),
    ('TEXTCOLOR',(0,0),(-1,0), colors.white),
    ('FONTNAME',(0,0),(-1,0), base_font),
    ('GRID',(0,0),(-1,-1),0.4,HexColor('#D6E3EE')),
    ('FONTNAME',(0,1),(-1,-1), base_font),
]))
flow.append(tb)
flow.append(Spacer(1, 12))

def save_pdf(path):
    doc = SimpleDocTemplate(path, pagesize=A4, rightMargin=18*mm, leftMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    doc.build(flow)

safe_save(save_pdf, pdf_path)

print("Reportes generados:")
print(" PDF:", pdf_path)
print(" Resumen Excel (con hoja 'Desglose'):", resumen_path)
print(" Detalle CSV:", detalle_path)
print(" Asistencias por becado Excel:", asist_path)
print(" Desglose fiscal (archivo separado):", desglose_path)
