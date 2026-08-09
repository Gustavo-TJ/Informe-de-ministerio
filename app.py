import streamlit as st
import pandas as pd
import datetime
import io

st.set_page_config(
    page_title="Informe de Servicio - Precursor Regular",
    page_icon="📖",
    layout="wide"
)

# Mapeo de meses del año de servicio
meses = ['Septiembre', 'Octubre', 'Noviembre', 'Diciembre', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto']

# Detección automática del mes actual
hoy = datetime.date.today()
meses_nombres = {
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre',
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto'
}
mes_actual_default = meses_nombres.get(hoy.month, 'Septiembre')

# Inicialización de la sesión personal
if 'data' not in st.session_state:
    df_dict = {}
    for i, mes in enumerate(meses):
        year = 2025 if i < 4 else 2026
        month_idx = ((i + 8) % 12) + 1
        num_days = pd.Period(f'{year}-{month_idx:02d}').days_in_month
        
        dates = [f"{year}-{month_idx:02d}-{d:02d}" for d in range(1, num_days + 1)]
        df_dict[mes] = pd.DataFrame({
            'Fecha': dates,
            'Horas': [0.0] * num_days,
            'Estudios / Personas': [''] * num_days,
            'Visitas': [0] * num_days,
            'Notas': [''] * num_days
        })
    st.session_state.data = df_dict
    st.session_state.carryover = {mes: 0.0 for mes in meses}
    st.session_state.goals = {mes: 50.0 for mes in meses}

st.title("📖 Registro e Informe de Servicio del Campo")
st.caption("Año de Servicio 2025 - 2026")

# Barra lateral: Gestión de usuarios e importación
st.sidebar.header("👤 Mi Registro Personal")

uploaded_file = st.sidebar.file_uploader("Cargar mi archivo Excel de respaldo", type=["xlsx"])
if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        for sheet in xls.sheet_names:
            if sheet in meses:
                st.session_state.data[sheet] = pd.read_excel(uploaded_file, sheet_name=sheet)
        st.sidebar.success("¡Tus datos se cargaron correctamente!")
    except Exception as e:
        st.sidebar.error("Error al leer el archivo Excel.")

st.sidebar.markdown("---")

# Selector de Mes
idx_default = meses.index(mes_actual_default) if mes_actual_default in meses else 0
mes_activo = st.sidebar.selectbox("Selecciona el Mes", meses, index=idx_default)

st.sidebar.header("⚙️ Configuración del Mes")
st.session_state.carryover[mes_activo] = st.sidebar.number_input(
    "Horas del mes anterior", 
    min_value=0.0, 
    value=float(st.session_state.carryover[mes_activo]), 
    step=0.25
)
st.session_state.goals[mes_activo] = st.sidebar.number_input(
    "Meta del mes (Horas)", 
    min_value=0.0, 
    value=float(st.session_state.goals[mes_activo]), 
    step=1.0
)

# Cálculo de Totales
df_mes = st.session_state.data[mes_activo]
horas_mes = pd.to_numeric(df_mes['Horas'], errors='coerce').fillna(0).sum()
horas_acumuladas = horas_mes + st.session_state.carryover[mes_activo]
meta = st.session_state.goals[mes_activo]
diferencia = horas_acumuladas - meta

# Dashboard Resumen
col1, col2, col3, col4 = st.columns(4)
col1.metric("Horas del Mes", f"{horas_mes:.2f} hrs")
col2.metric("Suma con Mes Anterior", f"{horas_acumuladas:.2f} hrs")
col3.metric("Meta del Mes", f"{meta:.2f} hrs")
col4.metric(
    "Diferencia", 
    f"{diferencia:+.2f} hrs", 
    delta_color="normal" if diferencia >= 0 else "inverse"
)

st.progress(min(max(horas_acumuladas / meta if meta > 0 else 0.0, 0.0), 1.0))

# Editor de Datos
st.subheader(f"📅 Registro Diario: {mes_activo}")
df_edited = st.data_editor(
    df_mes,
    num_rows="fixed",
    column_config={
        "Fecha": st.column_config.TextColumn("Fecha", disabled=True),
        "Horas": st.column_config.NumberColumn("Horas Realizadas", min_value=0.0, max_value=24.0, step=0.25, format="%.2f hrs"),
        "Estudios / Personas": st.column_config.TextColumn("Estudios Bíblicos / Nombres", width="medium"),
        "Visitas": st.column_config.NumberColumn("Revisitas", min_value=0, step=1),
        "Notas": st.column_config.TextColumn("Notas Adicionales", width="large")
    },
    hide_index=True,
    use_container_width=True
)

st.session_state.data[mes_activo] = df_edited

# Resumen Anual
st.markdown("---")
st.subheader("📊 Resumen Anual 2025-2026")

resumen_anual = []
for m, df in st.session_state.data.items():
    h_m = pd.to_numeric(df['Horas'], errors='coerce').fillna(0).sum()
    c_o = st.session_state.carryover[m]
    tot = h_m + c_o
    g = st.session_state.goals[m]
    vis = pd.to_numeric(df['Visitas'], errors='coerce').fillna(0).sum() if 'Visitas' in df.columns else 0
    resumen_anual.append({
        "Mes": m,
        "Horas Mes": h_m,
        "Viene Mes Ant.": c_o,
        "Total": tot,
        "Meta": g,
        "Diferencia": tot - g,
        "Revisitas Totales": vis
    })

df_resumen = pd.DataFrame(resumen_anual)
st.dataframe(df_resumen, hide_index=True, use_container_width=True)

# Exportación a Excel
st.markdown("---")
st.subheader("💾 Guardar / Exportar Informe")

def to_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for m, df in st.session_state.data.items():
            df.to_excel(writer, sheet_name=m, index=False)
    return output.getvalue()

st.download_button(
    label="📥 Descargar mi Informe en Excel",
    data=to_excel(),
    file_name="Mi_Informe_Precursor_2025_2026.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
