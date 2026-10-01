import streamlit as st
import openpyxl
import io
import requests
import plotly.graph_objects as go

st.set_page_config(page_title="Gestor de Pagos y Lotes", page_icon="🏢", layout="wide")

SHEET_ID = "1CdRBWSW9QDh63s7-8lnZvGu8rwz65z2duPnk2Nqfzqc"

@st.cache_data(ttl=15)
def descargar_workbook():
    # Descarga el archivo XLSX conservando los estilos de color y formatos de celda
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    res = requests.get(url)
    if res.status_code != 200:
        return None
    wb = openpyxl.load_workbook(io.BytesIO(res.content), data_only=True)
    return wb

st.title("🏢 Gestor de Pagos y Lotes Inmobiliarios")

col_btn, _ = st.columns([2, 8])
with col_btn:
    if st.button("🔄 Refrescar Datos"):
        st.cache_data.clear()
        st.rerun()

wb = descargar_workbook()
if not wb or "CT" not in wb.sheetnames:
    st.error("No se pudo conectar con la hoja de cálculo o no existe la pestaña 'CT'.")
    st.stop()

ws_ct = wb["CT"]
ws_ventas = wb["VENTAS"] if "VENTAS" in wb.sheetnames else None

def get_color_estado(cell):
    """
    Detecta el estado de la cuota según color de celda (RGB/ARGB)
    o por contenido de la anotación registrada.
    """
    color_detectado = None

    if cell and cell.fill and cell.fill.start_color:
        color = cell.fill.start_color
        rgb = getattr(color, 'rgb', None)
        if rgb and isinstance(rgb, str):
            rgb = str(rgb).upper().strip()
            # Si tiene 8 caracteres (ARGB con canal alfa de Google), tomamos los últimos 6
            if len(rgb) == 8:
                rgb = rgb[2:]

            # Tonos verdes comunes en Google Sheets
            if any(g in rgb for g in ["00FF00", "57BB8A", "6AA84F", "00E676", "38761D", "85E89D", "B7E1CD", "274E13", "81C784", "4CAF50", "A8DAB5"]):
                color_detectado = "VERDE"
            # Tonos naranja / ámbar (cuotas por vencer / vencidas)
            elif any(o in rgb for o in ["FFA500", "FF9900", "F6B26B", "F9AB00", "E69138", "FB8C00", "FF5722", "FFB74D", "FF9800", "FCE8E6"]):
                color_detectado = "NARANJA"
            # Tonos rojos (lotes caídos / mora crítica)
            elif any(r in rgb for r in ["FF0000", "CC0000", "E06666", "EA4335", "990000", "E57373", "F44336"]):
                color_detectado = "ROJO"

    if color_detectado in ["VERDE", "NARANJA", "ROJO"]:
        return color_detectado

    # Respaldo de seguridad por contenido si el tema de Sheets no expuso el RGB
    val_str = str(cell.value or "").strip().upper() if cell else ""
    if val_str and val_str not in ["NONE", "NAN", ""]:
        if any(w in val_str for w in ["DEBE", "MORA", "FALTA"]):
            return "NARANJA"
        if any(w in val_str for w in ["BAJA", "PERDIDO", "CANCELADO"]):
            return "ROJO"
        # Si tiene texto de pago o monto (ej: '636 //30-5' o 'S/.795')
        return "VERDE"

    return "BLANCO"

def clean_number(val):
    if val is None:
        return 0.0
    s = str(val).replace("S/.", "").replace("S/", "").replace(",", "").replace(" ", "").strip()
    try:
        return float(s)
    except:
        return 0.0

def normalizar(txt):
    if not txt:
        return ""
    return str(txt).replace(" ", "").replace("-", "").replace("_", "").upper().strip()

PALABRAS_IGNORAR = {
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SETIEMBRE", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
    "NOMBRE", "NOMBRES", "NOMBRRE", "CLIENTE", "ID", "TOTAL", "SUBTOTAL"
}

# 1. Encontrar la columna donde empieza la cuota 1 en la fila de encabezados
fila_encabezado_idx = 1
col_cuota_1 = 13  # Por defecto columna M

for r in range(1, 10):
    for c in range(8, 20):
        val = str(ws_ct.cell(row=r, column=c).value).strip()
        if val == "1":
            fila_encabezado_idx = r
            col_cuota_1 = c
            break
    if col_cuota_1 != 13:
        break

# 2. Extraer lotes y clientes permitiendo duplicados mediante identificación única
registros_lotes = []
for r in range(fila_encabezado_idx + 1, ws_ct.max_row + 1):
    val_nom = ws_ct.cell(row=r, column=1).value
    if val_nom is None:
        continue
    nom_str = str(val_nom).strip()
    if not nom_str or nom_str.upper() in PALABRAS_IGNORAR:
        continue
        
    etapa = str(ws_ct.cell(row=r, column=2).value or "").strip()
    lote = str(ws_ct.cell(row=r, column=3).value or "").strip()
    if not etapa and not lote:
        continue

    # Etiqueta visible única para no pisar clientes repetidos
    label = f"{nom_str} — {etapa} Lote {lote}"
    registros_lotes.append({
        "id_compuesto": label,
        "nombre": nom_str,
        "etapa": etapa,
        "lote": lote,
        "row_idx": r
    })

if not registros_lotes:
    st.warning("No se encontraron registros válidos de clientes en la pestaña CT.")
    st.stop()

opciones_busqueda = [item["id_compuesto"] for item in registros_lotes]

cliente_sel_label = st.selectbox(
    "🔍 Buscar por Cliente o Lote (Columna A):",
    options=opciones_busqueda,
    index=None,
    placeholder="Escribe el nombre del cliente o el lote..."
)

if cliente_sel_label:
    data_sel = next(item for item in registros_lotes if item["id_compuesto"] == cliente_sel_label)
    row_idx = data_sel["row_idx"]
    nom_cliente = data_sel["nombre"]
    etapa = data_sel["etapa"]
    lote = data_sel["lote"]

    m2 = str(ws_ct.cell(row=row_idx, column=4).value or "-").strip()
    valor_cuota = clean_number(ws_ct.cell(row=row_idx, column=9).value) # Columna I en CT

    # Valores base tomados de CT como respaldo
    valor_total = clean_number(ws_ct.cell(row=row_idx, column=5).value) # Columna E en CT
    monto_inicial = clean_number(ws_ct.cell(row=row_idx, column=6).value) # Columna F en CT
    modalidad = "FINANCIADO"

    # Cruce con la pestaña VENTAS haciendo coincidir CLIENTE Y LOTE
    if ws_ventas:
        for r_v in range(2, ws_ventas.max_row + 1):
            nom_v = str(ws_ventas.cell(row=r_v, column=2).value or "").strip()
            lote_v = str(ws_ventas.cell(row=r_v, column=5).value or "").strip()
            
            coincide_nom = normalizar(nom_v) == normalizar(nom_cliente)
            coincide_lote = normalizar(lote_v) == normalizar(lote) if lote and lote_v else True
            
            if coincide_nom and coincide_lote:
                v_total = clean_number(ws_ventas.cell(row=r_v, column=7).value)   # Columna G en VENTAS
                v_inic = clean_number(ws_ventas.cell(row=r_v, column=12).value)   # Columna L en VENTAS
                mod_v = str(ws_ventas.cell(row=r_v, column=15).value or "").strip().upper() # Columna O en VENTAS
                
                if v_total > 0: valor_total = v_total
                if v_inic > 0: monto_inicial = v_inic
                if mod_v: modalidad = mod_v
                break

    # 3. Procesar las 36 cuotas evaluando verde, naranja, rojo o blanco
    cuotas_estados = []
    total_verdes = 0
    total_naranjas = 0
    total_rojas = 0

    for i in range(36):
        c_idx = col_cuota_1 + i
        cell = ws_ct.cell(row=row_idx, column=c_idx)
        est = get_color_estado(cell)
        
        if est == "VERDE":
            total_verdes += 1
            cuotas_estados.append("VERDE")
        elif est == "NARANJA":
            total_naranjas += 1
            cuotas_estados.append("NARANJA")
        elif est == "ROJO":
            total_rojas += 1
            cuotas_estados.append("ROJO")
        else:
            cuotas_estados.append("BLANCO")

    # Reglas financieras según la modalidad de pago
    es_contado = "CONTADO" in modalidad
    if es_contado:
        total_pagado = valor_total
        saldo_pendiente = 0.0
    else:
        total_pagado = monto_inicial + (total_verdes * valor_cuota)
        saldo_pendiente = max(0.0, valor_total - total_pagado)

    st.markdown("---")
    st.subheader(f"Ficha de: {nom_cliente}")

    # Tarjetas principales
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📍 Lote / Etapa", f"{etapa} - {lote}", f"{m2} m²")
    c2.metric("💳 Modalidad", modalidad)
    c3.metric("💰 Inicial", f"S/. {monto_inicial:,.2f}")
    c4.metric("📅 Valor Cuota", f"S/. {valor_cuota:,.2f}")

    c5, c6, c7 = st.columns(3)
    c5.metric("💵 Total Pagado", f"S/. {total_pagado:,.2f}")
    c6.metric("⏳ Saldo Pendiente", f"S/. {saldo_pendiente:,.2f}")
    c7.metric("📊 Cuotas Pagadas", f"{total_verdes} de 36", f"{total_naranjas} vencidas" if total_naranjas > 0 else None)

    col_chart, col_matrix = st.columns([1, 1])

    with col_chart:
        st.subheader("Estado de Amortización")
        fig = go.Figure(data=[go.Pie(
            labels=['Total Pagado', 'Saldo Pendiente'],
            values=[total_pagado, saldo_pendiente],
            hole=.5,
            marker_colors=['#00E676', '#FFA500'] if not es_contado else ['#00E676', '#CCCCCC']
        )])
        fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=300)
        st.plotly_chart(fig, use_container_width=True)

    with col_matrix:
        st.subheader("Matriz de las 36 Cuotas")
        st.caption("🟢 Verde: Pagada | 🟠 Naranja: Vencida | 🔴 Rojo: Pérdida | ⚪ Neutro: Pendiente")
        cuotas_cols = st.columns(6)
        for i in range(1, 37):
            with cuotas_cols[(i - 1) % 6]:
                est = cuotas_estados[i - 1]
                if es_contado or est == "VERDE":
                    st.success(f"C{i}")
                elif est == "NARANJA":
                    st.warning(f"C{i}")
                elif est == "ROJO":
                    st.error(f"C{i}")
                else:
                    st.info(f"C{i}")
