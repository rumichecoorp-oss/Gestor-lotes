import streamlit as st
import openpyxl
import io
import requests
import plotly.graph_objects as go
import streamlit.components.v1 as components

st.set_page_config(
    page_title="GESTOR PAGOS VALLE HERMOSO",
    page_icon="🏡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilos CSS personalizados para móvil y reducción de elementos
st.markdown("""
<style>
    /* Estructura general compacta */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 5rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 580px !important;
    }
    
    /* Encabezado principal */
    .app-header {
        text-align: center;
        margin-bottom: 1rem;
    }
    .app-title {
        font-size: 1.35rem;
        font-weight: 800;
        color: #1A365D;
        letter-spacing: -0.5px;
        margin: 0;
    }
    
    /* Tarjeta contenedor del cliente (área capturable) */
    #reporte-cliente-card {
        background: #ffffff;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 16px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.04);
        margin-top: 10px;
    }

    /* Grid móvil para las métricas */
    .metrics-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 8px;
        margin-bottom: 12px;
    }
    .metric-card {
        background: #F8FAFC;
        border: 1px solid #EDF2F7;
        border-radius: 10px;
        padding: 8px 10px;
    }
    .metric-label {
        font-size: 0.72rem;
        color: #718096;
        text-transform: uppercase;
        font-weight: 600;
        margin-bottom: 2px;
    }
    .metric-val {
        font-size: 1.05rem;
        font-weight: 700;
        color: #2D3748;
    }
    .metric-val.green { color: #2E7D32; }
    .metric-val.orange { color: #DD6B20; }

    /* Cuadrícula compacta de las 36 cuotas */
    .cuotas-container {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 5px;
        margin-top: 8px;
    }
    .cuota-badge {
        font-size: 0.75rem;
        font-weight: 700;
        text-align: center;
        padding: 5px 0;
        border-radius: 6px;
        color: #ffffff;
    }
    .badge-verde { background-color: #00C853; }
    .badge-naranja { background-color: #FF9100; }
    .badge-rojo { background-color: #E53935; }
    .badge-blanco { background-color: #E2E8F0; color: #4A5568 !important; }

    /* Botón flotante inferior derecho para compartir */
    .floating-share-btn {
        position: fixed;
        bottom: 20px;
        right: 20px;
        z-index: 999999;
        background: #1A365D;
        color: #ffffff;
        border: none;
        border-radius: 50px;
        padding: 12px 18px;
        font-size: 0.9rem;
        font-weight: 700;
        box-shadow: 0 6px 18px rgba(0,0,0,0.25);
        display: flex;
        align-items: center;
        gap: 8px;
        cursor: pointer;
    }
</style>
""", unsafe_allow_html=True)

SHEET_ID = "1CdRBWSW9QDh63s7-8lnZvGu8rwz65z2duPnk2Nqfzqc"

@st.cache_data(ttl=15)
def descargar_workbook():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    res = requests.get(url)
    if res.status_code != 200:
        return None
    wb = openpyxl.load_workbook(io.BytesIO(res.content), data_only=True)
    return wb

# Encabezado
st.markdown('<div class="app-header"><h1 class="app-title">🏡 GESTOR PAGOS VALLE HERMOSO</h1></div>', unsafe_allow_html=True)

if st.button("🔄 Refrescar Datos", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

wb = descargar_workbook()
if not wb or "CT" not in wb.sheetnames:
    st.error("No se pudo conectar con la hoja de cálculo o no existe la pestaña 'CT'.")
    st.stop()

ws_ct = wb["CT"]
ws_ventas = wb["VENTAS"] if "VENTAS" in wb.sheetnames else None

def get_color_estado(cell):
    if not cell:
        return "BLANCO"

    val_str = str(cell.value or "").strip().upper()
    tiene_contenido = bool(val_str and val_str not in ["NONE", "NAN", ""])

    rgb_str = ""
    indexed_val = None

    if cell.fill and (cell.fill.fill_type or cell.fill.start_color or cell.fill.fgColor):
        color_obj = getattr(cell.fill, 'fgColor', None) or getattr(cell.fill, 'start_color', None)
        if color_obj:
            v_rgb = getattr(color_obj, 'rgb', None) or getattr(color_obj, 'value', None)
            if v_rgb and isinstance(v_rgb, str):
                rgb_str = v_rgb.upper().strip()
                if len(rgb_str) == 8:
                    rgb_str = rgb_str[2:]
            indexed_val = getattr(color_obj, 'indexed', None) or getattr(color_obj, 'index', None)

    verdes_rgb = [
        "00FF00", "57BB8A", "6AA84F", "00E676", "38761D", "85E89D", 
        "B7E1CD", "274E13", "81C784", "4CAF50", "2E7D32", "00C853", 
        "34A853", "137333", "0F9D58", "A8DAB5", "D9EAD3"
    ]
    if any(v in rgb_str for v in verdes_rgb) or indexed_val in [11, 17, 42, 43]:
        return "VERDE"

    naranjas_rgb = [
        "FF9900", "FF6D01", "FFA500", "F6B26B", "F9AB00", "E69138", 
        "FB8C00", "FF5722", "FF7043", "FFB74D", "FF9800", "E55100", 
        "B45F06", "783F04", "F9CB9C", "FCE5CD", "FF8A65", "FFAB40"
    ]
    if any(n in rgb_str for n in naranjas_rgb) or indexed_val in [51, 52, 53, 54]:
        return "NARANJA"

    rojos_rgb = ["FF0000", "CC0000", "E06666", "EA4335", "990000", "E57373", "F44336", "C62828"]
    if any(r in rgb_str for r in rojos_rgb) or indexed_val in [10, 16]:
        return "ROJO"

    # Regla por contenido si no es verde
    if tiene_contenido:
        if any(w in val_str for w in ["BAJA", "PERDIDO", "CANCELADO"]):
            return "ROJO"
        return "NARANJA"

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

fila_encabezado_idx = 1
col_cuota_1 = 13

for r in range(1, 10):
    for c in range(8, 20):
        val = str(ws_ct.cell(row=r, column=c).value).strip()
        if val == "1":
            fila_encabezado_idx = r
            col_cuota_1 = c
            break
    if col_cuota_1 != 13:
        break

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

    label = f"{nom_str} — {etapa} Lote {lote}"
    registros_lotes.append({
        "id_compuesto": label,
        "nombre": nom_str,
        "etapa": etapa,
        "lote": lote,
        "row_idx": r
    })

if not registros_lotes:
    st.warning("No se encontraron clientes.")
    st.stop()

opciones_busqueda = [item["id_compuesto"] for item in registros_lotes]

cliente_sel_label = st.selectbox(
    "🔍 Seleccionar o buscar cliente:",
    options=opciones_busqueda,
    index=None,
    placeholder="Escribe el nombre o lote..."
)

if cliente_sel_label:
    data_sel = next(item for item in registros_lotes if item["id_compuesto"] == cliente_sel_label)
    row_idx = data_sel["row_idx"]
    nom_cliente = data_sel["nombre"]
    etapa = data_sel["etapa"]
    lote = data_sel["lote"]

    m2 = str(ws_ct.cell(row=row_idx, column=4).value or "-").strip()
    valor_cuota = clean_number(ws_ct.cell(row=row_idx, column=9).value)

    valor_total = clean_number(ws_ct.cell(row=row_idx, column=5).value)
    monto_inicial = clean_number(ws_ct.cell(row=row_idx, column=6).value)
    modalidad = "FINANCIADO"

    if ws_ventas:
        for r_v in range(2, ws_ventas.max_row + 1):
            nom_v = str(ws_ventas.cell(row=r_v, column=2).value or "").strip()
            lote_v = str(ws_ventas.cell(row=r_v, column=5).value or "").strip()
            
            coincide_nom = normalizar(nom_v) == normalizar(nom_cliente)
            coincide_lote = normalizar(lote_v) == normalizar(lote) if lote and lote_v else True
            
            if coincide_nom and coincide_lote:
                v_total = clean_number(ws_ventas.cell(row=r_v, column=7).value)
                v_inic = clean_number(ws_ventas.cell(row=r_v, column=12).value)
                mod_v = str(ws_ventas.cell(row=r_v, column=15).value or "").strip().upper()
                
                if v_total > 0: valor_total = v_total
                if v_inic > 0: monto_inicial = v_inic
                if mod_v: modalidad = mod_v
                break

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

    es_contado = "CONTADO" in modalidad
    if es_contado:
        total_pagado = valor_total
        saldo_pendiente = 0.0
    else:
        total_pagado = monto_inicial + (total_verdes * valor_cuota)
        saldo_pendiente = max(0.0, valor_total - total_pagado)

    # Generación de la cuadrícula de cuotas compactas
    badges_html = ""
    for i in range(1, 37):
        est = cuotas_estados[i - 1]
        clase = "badge-verde" if (es_contado or est == "VERDE") else ("badge-naranja" if est == "NARANJA" else ("badge-rojo" if est == "ROJO" else "badge-blanco"))
        badges_html += f'<div class="cuota-badge {clase}">C{i}</div>'

    # Render de la tarjeta de reporte capturable
    st.markdown(f"""
    <div id="reporte-cliente-card">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px; border-bottom:1px solid #EDF2F7; padding-bottom:8px;">
            <div>
                <h3 style="margin:0; font-size:1.2rem; color:#1A365D;">{nom_cliente}</h3>
                <span style="font-size:0.82rem; color:#718096; font-weight:600;">{etapa} — Lote {lote} ({m2} m²)</span>
            </div>
            <div style="text-align:right;">
                <span style="background:#EBF8FF; color:#2B6CB0; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700;">{modalidad}</span>
            </div>
        </div>

        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-label">Valor Total</div>
                <div class="metric-val">S/. {valor_total:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Monto Inicial</div>
                <div class="metric-val">S/. {monto_inicial:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Pagado</div>
                <div class="metric-val green">S/. {total_pagado:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Saldo Pendiente</div>
                <div class="metric-val orange">S/. {saldo_pendiente:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Valor de Cuota</div>
                <div class="metric-val">S/. {valor_cuota:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Cuotas Pagadas</div>
                <div class="metric-val">{total_verdes} de 36 {"(" + str(total_naranjas) + " venc)" if total_naranjas > 0 else ""}</div>
            </div>
        </div>

        <div style="margin-top:12px;">
            <div style="font-size:0.78rem; font-weight:700; color:#4A5568; margin-bottom:4px; display:flex; justify-content:space-between;">
                <span>Matriz de 36 Cuotas</span>
                <span style="font-size:0.7rem; color:#718096;">🟢 Pagada | 🟠 Vencida | ⚪ Pend</span>
            </div>
            <div class="cuotas-container">
                {badges_html}
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Gráfico de dona compacto debajo de la tarjeta
    st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
    fig = go.Figure(data=[go.Pie(
        labels=['Total Pagado', 'Saldo Pendiente'],
        values=[total_pagado, saldo_pendiente],
        hole=.55,
        marker_colors=['#00E676', '#FF9100'] if not es_contado else ['#00E676', '#CCCCCC'],
        textinfo='percent',
        hoverinfo='label+value'
    )])
    fig.update_layout(
        height=210,
        margin=dict(t=5, b=5, l=5, r=5),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5, font=dict(size=11))
    )
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

    # Botón Flotante Inferior Derecho con html2canvas y Web Share API
    components.html(f"""
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <button id="btnShareReporte" class="floating-share-btn" onclick="capturarYCompartir()">
        📸 Compartir Ficha
    </button>

    <script>
    async function capturarYCompartir() {{
        const card = window.parent.document.getElementById('reporte-cliente-card');
        if (!card) {{
            alert("No se encontró la ficha del cliente.");
            return;
        }}

        const btn = document.getElementById('btnShareReporte');
        btn.innerText = "⏳ Generando...";
        btn.disabled = true;

        try {{
            const canvas = await html2canvas(card, {{
                scale: 2,
                useCORS: true,
                backgroundColor: "#ffffff"
            }});

            canvas.toBlob(async (blob) => {{
                const file = new File([blob], "Estado_Cuenta_{nom_cliente.replace(' ', '_')}.png", {{ type: "image/png" }});

                // Si el navegador soporta compartir archivos nativamente (móviles / WhatsApp)
                if (navigator.canShare && navigator.canShare({{ files: [file] }})) {{
                    try {{
                        await navigator.share({{
                            title: 'Estado de Cuenta - Valle Hermoso',
                            text: 'Reporte de lote para {nom_cliente} ({etapa} Lote {lote})',
                            files: [file]
                        }});
                    }} catch (e) {{
                        console.log("Compartir cancelado o no soportado:", e);
                    }}
                }} else {{
                    // Descarga directa si es computadora
                    const link = document.createElement('a');
                    link.download = "Estado_Cuenta_{nom_cliente.replace(' ', '_')}.png";
                    link.href = canvas.toDataURL('image/png');
                    link.click();
                }}
                btn.innerText = "📸 Compartir Ficha";
                btn.disabled = false;
            }}, 'image/png');
        }} catch (err) {{
            alert("Error al capturar la imagen: " + err);
            btn.innerText = "📸 Compartir Ficha";
            btn.disabled = false;
        }}
    }}
    </script>
    """, height=70)
