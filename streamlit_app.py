import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import requests

st.set_page_config(page_title="Gestor de Pagos y Lotes", page_icon="🏢", layout="wide")

SHEET_ID = "1CdRBWSW9QDh63s7-8lnZvGu8rwz65z2duPnk2Nqfzqc"

@st.cache_data(ttl=15)
def cargar_datos(sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    res = requests.get(url)
    if res.status_code != 200:
        return None
    df = pd.read_csv(io.StringIO(res.content.decode('utf-8')), header=None)
    return df

st.title("🏢 Gestor de Pagos y Lotes Inmobiliarios")

col_btn, _ = st.columns([2, 8])
with col_btn:
    if st.button("🔄 Refrescar Datos"):
        st.cache_data.clear()
        st.rerun()

df_ct = cargar_datos("CT")
df_ventas = cargar_datos("VENTAS")

if df_ct is None or df_ventas is None:
    st.error("No se pudo conectar con las pestañas 'CT' o 'VENTAS'.")
    st.stop()

# Lista de palabras clave, meses y encabezados que NO son clientes
PALABRAS_IGNORAR = {
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SETIEMBRE", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
    "NOMBRE", "NOMBRES", "NOMBRRE", "CLIENTE", "ID", "TOTAL", "SUBTOTAL"
}

def es_cliente_valido(val, fila):
    if pd.isna(val):
        return False
    texto = str(val).strip().upper()
    if not texto or texto in PALABRAS_IGNORAR:
        return False
    # Una fila de cliente real tiene Lote o Etapa asignada en columna B o C
    etapa = str(fila.iloc[1]).strip() if len(fila) > 1 and pd.notna(fila.iloc[1]) else ""
    lote = str(fila.iloc[2]).strip() if len(fila) > 2 and pd.notna(fila.iloc[2]) else ""
    return bool(etapa or lote)

# Filtrar únicamente clientes reales
filas_validas = []
for idx, fila in df_ct.iterrows():
    if es_cliente_valido(fila.iloc[0], fila):
        filas_validas.append(fila)

if not filas_validas:
    st.warning("No se encontraron registros de clientes en la pestaña CT.")
    st.stop()

df_clientes = pd.DataFrame(filas_validas)
nombres_clientes = sorted(df_clientes.iloc[:, 0].astype(str).str.strip().unique().tolist())

cliente_sel = st.selectbox(
    "🔍 Buscar cliente (Columna A):",
    options=nombres_clientes,
    index=None,
    placeholder="Escribe o selecciona un cliente..."
)

if cliente_sel:
    fila_ct = df_clientes[df_clientes.iloc[:, 0].astype(str).str.strip() == cliente_sel].iloc[0]

    # Cruzar datos con pestaña VENTAS
    fila_v = None
    matches_v = df_ventas[df_ventas.iloc[:, 1].astype(str).str.strip().str.upper() == cliente_sel.upper()]
    if not matches_v.empty:
        fila_v = matches_v.iloc[0]

    def to_float(val):
        if pd.isna(val): return 0.0
        s = str(val).replace("S/.", "").replace("S/", "").replace(",", "").replace(" ", "").strip()
        try: return float(s)
        except: return 0.0

    etapa = str(fila_ct.iloc[1]).strip() if len(fila_ct) > 1 and pd.notna(fila_ct.iloc[1]) else "-"
    lote = str(fila_ct.iloc[2]).strip() if len(fila_ct) > 2 and pd.notna(fila_ct.iloc[2]) else "-"
    m2 = str(fila_ct.iloc[3]).strip() if len(fila_ct) > 3 and pd.notna(fila_ct.iloc[3]) else "-"

    # Precios desde VENTAS (Cols G, L, O) o de respaldo en CT
    valor_total = to_float(fila_v.iloc[6]) if fila_v is not None and len(fila_v) > 6 else to_float(fila_ct.iloc[4])
    monto_inicial = to_float(fila_v.iloc[11]) if fila_v is not None and len(fila_v) > 11 else to_float(fila_ct.iloc[5])
    modalidad = str(fila_v.iloc[14]).strip().upper() if fila_v is not None and len(fila_v) > 14 and pd.notna(fila_v.iloc[14]) else "FINANCIADO"
    valor_cuota = to_float(fila_ct.iloc[8])

    # Contar cuotas pagadas (casillas con valor o notas desde Col M / índice 12 a la 47)
    cuotas_vals = fila_ct.iloc[12:48].tolist() if len(fila_ct) > 12 else []
    cuotas_totales = 36
    cuotas_estados = []

    for c in cuotas_vals:
        s_val = str(c).strip()
        if pd.notna(c) and s_val != "" and s_val != "nan" and s_val != "None":
            cuotas_estados.append(True)
        else:
            cuotas_estados.append(False)

    # Rellenar hasta 36 si la hoja tiene menos columnas
    while len(cuotas_estados) < cuotas_totales:
        cuotas_estados.append(False)

    cuotas_pagadas = sum(cuotas_estados)

    es_contado = "CONTADO" in modalidad
    if es_contado:
        total_pagado = valor_total
        saldo_pendiente = 0.0
    else:
        total_pagado = monto_inicial + (cuotas_pagadas * valor_cuota)
        saldo_pendiente = max(0.0, valor_total - total_pagado)

    st.markdown("---")
    st.subheader(f"Ficha de: {cliente_sel}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📍 Lote / Etapa", f"{etapa} - {lote}", f"{m2} m²")
    c2.metric("💳 Modalidad", modalidad)
    c3.metric("💰 Inicial", f"S/. {monto_inicial:,.2f}")
    c4.metric("📅 Valor Cuota", f"S/. {valor_cuota:,.2f}")

    c5, c6, c7 = st.columns(3)
    c5.metric("💵 Total Pagado", f"S/. {total_pagado:,.2f}")
    c6.metric("⏳ Saldo Pendiente", f"S/. {saldo_pendiente:,.2f}")
    c7.metric("📊 Cuotas Pagadas", f"{cuotas_pagadas} de {cuotas_totales}")

    col_chart, col_matrix = st.columns([1, 1])

    with col_chart:
        st.subheader("Estado de Amortización")
        fig = go.Figure(data=[go.Pie(
            labels=['Total Pagado', 'Saldo Pendiente'],
            values=[total_pagado, saldo_pendiente],
            hole=.5,
            marker_colors=['#00E676', '#FFA500'] if not es_contado else ['#00E676', '#E0E0E0']
        )])
        fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=300)
        st.plotly_chart(fig, use_container_width=True)

    with col_matrix:
        st.subheader("Matriz de las 36 Cuotas")
        cuotas_cols = st.columns(6)
        for i in range(1, cuotas_totales + 1):
            with cuotas_cols[(i - 1) % 6]:
                if es_contado or cuotas_estados[i - 1]:
                    st.success(f"C{i}")
                else:
                    st.info(f"C{i}")
