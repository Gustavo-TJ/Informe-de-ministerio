import streamlit as st
import pandas as pd
import datetime
import io
import json
import os

# Importaciones necesarias para generar el PDF listo para imprimir
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# Configuración de la página
st.set_page_config(
    page_title="Informe de Predicación",
    page_icon="📋",
    layout="wide"
)

# ARCHIVO LOCAL PARA GUARDAR DATOS PERSISTENTES
DATA_FILE = "datos_predicacion.json"

# --- FUNCIONES PARA GUARDAR Y CARGAR DATOS EN DISCO ---
def guardar_datos_disco():
    """Guarda st.session_state en un archivo JSON local."""
    data_to_save = {
        "nombre_publicador": st.session_state.get("nombre_publicador", ""),
        "carryover": st.session_state.get("carryover", {}),
        "goals": st.session_state.get("goals", {}),
        "tables": {}
    }
    for mes, df in st.session_state.get("data", {}).items():
        data_to_save["tables"][mes] = df.to_dict(orient="records")
    
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data_to_save, f, ensure_ascii=False, indent=2)

def cargar_datos_disco(meses):
    """Carga los datos guardados en disco o inicializa estructuras vacías."""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                saved_data = json.load(f)
            
            st.session_state.nombre_publicador = saved_data.get("nombre_publicador", "")
            st.session_state.carryover = saved_data.get("carryover", {mes: 0.0 for mes in meses})
            st.session_state.goals = saved_data.get("goals", {mes: 50.0 for mes in meses})
            
            df_dict = {}
            for mes in meses:
                if mes in saved_data.get("tables", {}):
                    df_dict[mes] = pd.DataFrame(saved_data["tables"][mes])
                else:
                    # Estructura por defecto si falta algún mes
                    df_dict[mes] = crear_df_mes_vacio(mes, meses.index(mes))
            st.session_state.data = df_dict
            return
        except Exception as e:
            st.warning("No se pudo leer el archivo de guardado previo, creando uno nuevo.")

    # Si no existe archivo guardado, crear inicialización por defecto
    st.session_state.nombre_publicador = ""
    st.session_state.carryover = {mes: 0.0 for mes in meses}
    st.session_state.goals = {mes: 50.0 for mes in meses}
    df_dict = {}
    for i, mes in enumerate(meses):
        df_dict[mes] = crear_df_mes_vacio(mes, i)
    st.session_state.data = df_dict

def crear_df_mes_vacio(mes, index_mes):
    year = 2025 if index_mes < 4 else 2026
    month_idx = ((index_mes + 8) % 12) + 1
    num_days = pd.Period(f'{year}-{month_idx:02d}').days_in_month
    dates = [f"{year}-{month_idx:02d}-{d:02d}" for d in range(1, num_days + 1)]
    return pd.DataFrame({
        'Fecha': dates,
        'Horas': [0] * num_days,
        'Minutos': [0] * num_days,
        'Estudios / Personas': [''] * num_days,
        'Notas': [''] * num_days
    })

# --- ESTILOS VISUALES PERSONALIZADOS (CSS) ---
st.markdown("""
    <style>
    /* Estilo de la barra lateral */
    [data-testid="stSidebar"] {
        background-color: #F8FAFC;
        border-right: 2px solid #DBEAFE;
    }
    
    /* Títulos principales */
    h1 {
        color: #1E3A8A !important;
        font-weight: 700 !important;
    }
    
    /* Subtítulos de secciones */
    h2, h3 {
        color: #1D4ED8 !important;
        font-weight: 600 !important;
    }
    
    /* Contenedores de tarjetas de métricas */
    [data-testid="stMetricValue"] {
        color: #1E40AF !important;
        font-weight: 700 !important;
    }
    
    /* Resaltar tabla editable */
    [data-testid="stDataEditor"] {
        background-color: #F0F4F8;
        border-radius: 10px;
        padding: 8px;
        border: 1px solid #BFDBFE;
    }
    
    /* Botón de descarga */
    div.stDownloadButton > button {
        background-color: #2563EB !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #1D4ED8 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Mapeo de meses del año de servicio (Septiembre a Agosto)
meses = ['Septiembre', 'Octubre', 'Noviembre', 'Diciembre', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto']

# Detección automática del mes actual
hoy = datetime.date.today()
meses_nombres = {
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto'
}
mes_actual_default = meses_nombres.get(hoy.month, 'Septiembre')

# Cargar datos desde disco si no están en session_state
if 'data' not in st.session_state:
    cargar_datos_disco(meses)

# --- BARRA LATERAL ---
st.sidebar.header("👤 Perfil del Publicador")

nombre_input = st.sidebar.text_input(
    "Nombre del Publicador", 
    value=st.session_state.nombre_publicador, 
    placeholder="Ej. Juan Pérez"
)
if nombre_input != st.session_state.nombre_publicador:
    st.session_state.nombre_publicador = nombre_input
    guardar_datos_disco()

uploaded_file = st.sidebar.file_uploader("Cargar archivo Excel de respaldo", type=["xlsx"])
if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        for sheet in xls.sheet_names:
            if sheet in meses:
                df_cargado = pd.read_excel(uploaded_file, sheet_name=sheet)
                df_cargado['Horas'] = pd.to_numeric(df_cargado['Horas'], errors='coerce').fillna(0).astype(int)
                df_cargado['Minutos'] = pd.to_numeric(df_cargado['Minutos'], errors='coerce').fillna(0).astype(int)
                df_cargado['Estudios / Personas'] = df_cargado['Estudios / Personas'].fillna('').astype(str)
                df_cargado['Notas'] = df_cargado['Notas'].fillna('').astype(str)
                st.session_state.data[sheet] = df_cargado
        guardar_datos_disco()
        st.sidebar.success("¡Tus datos se cargaron y guardaron correctamente!")
    except Exception as e:
        st.sidebar.error("Error al leer el archivo Excel.")

st.sidebar.markdown("---")

# Selector de Mes
idx_default = meses.index(mes_actual_default) if mes_actual_default in meses else 0
mes_activo = st.sidebar.selectbox("Selecciona el Mes", meses, index=idx_default)

st.sidebar.header("⚙️ Configuración del Mes")

val_carry = st.sidebar.number_input(
    "Horas del mes anterior", 
    min_value=0.0, 
    value=float(st.session_state.carryover.get(mes_activo, 0.0)), 
    step=0.25
)
if val_carry != st.session_state.carryover.get(mes_activo, 0.0):
    st.session_state.carryover[mes_activo] = val_carry
    guardar_datos_disco()

val_goal = st.sidebar.number_input(
    "Meta del mes (Horas)", 
    min_value=0.0, 
    value=float(st.session_state.goals.get(mes_activo, 50.0)), 
    step=1.0
)
if val_goal != st.session_state.goals.get(mes_activo, 50.0):
    st.session_state.goals[mes_activo] = val_goal
    guardar_datos_disco()

# --- CABECERA PRINCIPAL ---
st.title("📋 Informe de Predicación")
if st.session_state.nombre_publicador.strip():
    st.subheader(f"Publicador: {st.session_state.nombre_publicador}")
st.caption("Año de Servicio 2025 - 2026 (Septiembre 2025 – Agosto 2026)")

# --- EDITOR DE DATOS DEL MES ACTIVO ---
st.subheader(f"📅 Registro Diario de Actividad: {mes_activo}")

df_mes_original = st.session_state.data[mes_activo]
df_edited = st.data_editor(
    df_mes_original,
    num_rows="fixed",
    column_config={
        "Fecha": st.column_config.TextColumn("Fecha", disabled=True),
        "Horas": st.column_config.NumberColumn("Horas", min_value=0, max_value=24, step=1, format="%d h"),
        "Minutos": st.column_config.NumberColumn("Minutos", min_value=0, max_value=59, step=5, format="%d m"),
        "Estudios / Personas": st.column_config.TextColumn("Estudios Bíblicos / Nombres", width="medium"),
        "Notas": st.column_config.TextColumn("Notas Adicionales", width="large")
    },
    hide_index=True,
    use_container_width=True,
    key=f"editor_{mes_activo}"
)

# Verificar si hubo cambios en la tabla para guardar en disco
if not df_edited.equals(st.session_state.data[mes_activo]):
    st.session_state.data[mes_activo] = df_edited
    guardar_datos_disco()

# --- CÁLCULO Y DASHBOARD DEL MES ACTIVO ---
h_series = pd.to_numeric(df_edited['Horas'], errors='coerce').fillna(0)
m_series = pd.to_numeric(df_edited['Minutos'], errors='coerce').fillna(0)
total_minutos_mes = int((h_series * 60 + m_series).sum())
tot_h, tot_m = divmod(total_minutos_mes, 60)
horas_mes_decimal = total_minutos_mes / 60.0

horas_acumuladas_mes = horas_mes_decimal + st.session_state.carryover[mes_activo]
meta_mes = st.session_state.goals[mes_activo]
diferencia_mes = horas_acumuladas_mes - meta_mes

estudios_serie = df_edited['Estudios / Personas'].astype(str).str.strip()
total_estudiantes = (estudios_serie != '').sum()

# Resumen del Mes
st.markdown("---")
st.subheader(f"📊 Resumen del Mes: {mes_activo}")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Horas del Mes", f"{tot_h}h {tot_m}m")
col2.metric("Suma con Mes Ant.", f"{horas_acumuladas_mes:.2f} hrs")
col3.metric("Meta del Mes", f"{meta_mes:.0f} hrs")
col4.metric("Diferencia Mes", f"{diferencia_mes:+.2f} hrs", delta_color="normal" if diferencia_mes >= 0 else "inverse")
col5.metric("Estudios Bíblicos", f"{total_estudiantes}")

st.progress(min(max(horas_acumuladas_mes / meta_mes if meta_mes > 0 else 0.0, 0.0), 1.0))

# --- CÁLCULO GENERAL DEL AÑO Y META ANUAL ---
total_minutos_anuales = 0
for m, df in st.session_state.data.items():
    h_s = pd.to_numeric(df['Horas'], errors='coerce').fillna(0)
    m_s = pd.to_numeric(df['Minutos'], errors='coerce').fillna(0)
    total_minutos_anuales += int((h_s * 60 + m_s).sum())

total_horas_anuales = total_minutos_anuales / 60.0
meta_anual = 600.0
faltante_anual = meta_anual - total_horas_anuales
porcentaje_anual = min(max(total_horas_anuales / meta_anual, 0.0), 1.0)

# Panel de Meta Anual
st.markdown("---")
st.subheader("🎯 Progreso Anual del Precursor (Meta: 600 hrs)")
col_a1, col_a2, col_a3 = st.columns(3)
tot_h_anual, tot_m_anual = divmod(total_minutos_anuales, 60)

col_a1.metric("Horas Totales Acumuladas", f"{tot_h_anual}h {tot_m_anual}m")
col_a2.metric("Restantes para las 600h", f"{max(0.0, faltante_anual):.2f} hrs" if faltante_anual > 0 else "¡Meta Anual Alcanzada!")
col_a3.metric("Porcentaje del Año", f"{porcentaje_anual * 100:.1f}%")

st.progress(porcentaje_anual)

# --- TABLA DE RESUMEN ANUAL CON SUMA ACUMULADA ---
st.markdown("---")
st.subheader("📊 Resumen del Año de Servicio (Septiembre - Agosto)")

resumen_anual = []
acumulado_progresivo = 0.0

for m in meses:
    df = st.session_state.data[m]
    h_s = pd.to_numeric(df['Horas'], errors='coerce').fillna(0)
    m_s = pd.to_numeric(df['Minutos'], errors='coerce').fillna(0)
    m_tot = int((h_s * 60 + m_s).sum())
    h_m = m_tot / 60.0
    th, tm = divmod(m_tot, 60)
    
    c_o = st.session_state.carryover[m]
    tot = h_m + c_o
    g = st.session_state.goals[m]
    
    acumulado_progresivo += h_m
    
    e_ser = df['Estudios / Personas'].astype(str).str.strip()
    n_est = (e_ser != '').sum()
    
    resumen_anual.append({
        "Mes": m,
        "Horas Mes": f"{th}h {tm}m",
        "Viene Mes Ant.": c_o,
        "Total Mes": round(tot, 2),
        "Meta Mes": g,
        "Diferencia": round(tot - g, 2),
        "Horas Acum. Año": f"{int(acumulado_progresivo)}h {int((acumulado_progresivo % 1) * 60)}m",
        "Estudios": n_est
    })

df_resumen = pd.DataFrame(resumen_anual)
st.dataframe(df_resumen, hide_index=True, use_container_width=True)

# --- GUARDAR Y EXPORTAR INFORME ---
st.markdown("---")
st.subheader("💾 Guardar / Exportar Informe")

def to_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for m in meses:
            st.session_state.data[m].to_excel(writer, sheet_name=m, index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen_Anual", index=False)
    return output.getvalue()

def generar_pdf_planilla(nombre, df_resumen_data, tot_h_a, tot_m_a):
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm
    )
    
    styles = getSampleStyleSheet()
    
    titulo_style = ParagraphStyle(
        'TituloPDF',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1,
        spaceAfter=6
    )
    
    subtitulo_style = ParagraphStyle(
        'SubtituloPDF',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#475569'),
        alignment=1,
        spaceAfter=15
    )
    
    meta_style = ParagraphStyle(
        'MetaPDF',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor('#1E293B')
    )
    
    cell_style = ParagraphStyle(
        'CellPDF',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        alignment=1
    )
    
    header_cell_style = ParagraphStyle(
        'HeaderCellPDF',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        fontName='Helvetica-Bold',
        textColor=colors.white,
        alignment=1
    )

    story = []
    
    story.append(Paragraph("<b>INFORME DE SERVICIO DEL PRECURSOR</b>", titulo_style))
    story.append(Paragraph("Año de Servicio 2025 – 2026 (Septiembre 2025 – Agosto 2026)", subtitulo_style))
    
    nombre_txt = nombre.strip() if nombre.strip() else "No especificado"
    info_text = f"<b>Publicador:</b> {nombre_txt} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Total Horas Acumuladas:</b> {tot_h_a}h {tot_m_a}m / 600h"
    story.append(Paragraph(info_text, meta_style))
    story.append(Spacer(1, 0.4 * cm))
    
    headers = ["Mes", "Horas Mes", "Viene Ant.", "Total Mes", "Meta", "Dif.", "Horas Acum. Año", "Estudios"]
    table_data = [[Paragraph(h, header_cell_style) for h in headers]]
    
    for row in df_resumen_data.to_dict(orient='records'):
        table_data.append([
            Paragraph(str(row["Mes"]), cell_style),
            Paragraph(str(row["Horas Mes"]), cell_style),
            Paragraph(f"{row['Viene Mes Ant.']:.2f}", cell_style),
            Paragraph(f"{row['Total Mes']:.2f}", cell_style),
            Paragraph(f"{row['Meta Mes']:.0f}", cell_style),
            Paragraph(f"{row['Diferencia']:+.2f}", cell_style),
            Paragraph(str(row["Horas Acum. Año"]), cell_style),
            Paragraph(str(row["Estudios"]), cell_style),
        ])
    
    col_widths = [2.6 * cm, 2.2 * cm, 2.0 * cm, 2.0 * cm, 1.8 * cm, 1.8 * cm, 2.8 * cm, 1.8 * cm]
    
    tabla = Table(table_data, colWidths=col_widths, repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    
    story.append(tabla)
    doc.build(story)
    
    buffer.seek(0)
    return buffer.getvalue()

nombre_limpio = st.session_state.nombre_publicador.strip().replace(" ", "_") if st.session_state.nombre_publicador.strip() else "Publicador"
file_name_excel = f"Informe_Predicacion_{nombre_limpio}_2025_2026.xlsx"
file_name_pdf = f"Planilla_Impresion_{nombre_limpio}_2025_2026.pdf"

col_exp1, col_exp2 = st.columns(2)

with col_exp1:
    st.download_button(
        label="📥 Descargar Informe en Excel",
        data=to_excel(),
        file_name=file_name_excel,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col_exp2:
    pdf_bytes = generar_pdf_planilla(st.session_state.nombre_publicador, df_resumen, tot_h_anual, tot_m_anual)
    st.download_button(
        label="📄 Descargar Planilla PDF Listo para Imprimir",
        data=pdf_bytes,
        file_name=file_name_pdf,
        mime="application/pdf",
        use_container_width=True
    )
