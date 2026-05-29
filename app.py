"""
╔══════════════════════════════════════════════════════════════════╗
║         APP DE PEDIDOS — AUTOPLANET                             ║
║  Uso: streamlit run app.py                                      ║
╚══════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import io
import os
import tempfile
import numpy as np
from datetime import datetime

# ── Importar toda la lógica del procesador ────────────────────────
from procesar_pedidos import leer_archivo, limpiar, calcular_pivots, escribir_excel, CONFIG

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE LA PÁGINA
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Reporte de Pedidos",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1B3A6B 0%, #2563EB 100%);
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        color: white;
    }
    .kpi-card {
        background: white;
        padding: 1.2rem;
        border-radius: 10px;
        border-left: 4px solid #2563EB;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 1rem;
    }
    .kpi-card.red   { border-left-color: #DC2626; }
    .kpi-card.green { border-left-color: #16A34A; }
    .kpi-card.amber { border-left-color: #D97706; }
    .kpi-card.orange{ border-left-color: #EA580C; }
    .kpi-value { font-size: 1.8rem; font-weight: 700; color: #1B3A6B; }
    .kpi-label { font-size: 0.8rem; color: #64748B; text-transform: uppercase; letter-spacing: 0.05em; }
    .alert-box {
        background: #FEE2E2; border: 1px solid #DC2626;
        border-radius: 8px; padding: 1rem; margin: 0.5rem 0;
    }
    .success-box {
        background: #DCFCE7; border: 1px solid #16A34A;
        border-radius: 8px; padding: 1rem; margin: 0.5rem 0;
    }
    .warning-box {
        background: #FEF3C7; border: 1px solid #D97706;
        border-radius: 8px; padding: 1rem; margin: 0.5rem 0;
    }
    .info-box {
        background: #DBEAFE; border: 1px solid #2563EB;
        border-radius: 8px; padding: 1rem; margin: 0.5rem 0;
    }
    .stDataFrame { border-radius: 8px; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# DETECCIÓN AUTOMÁTICA DE COLUMNAS
# ══════════════════════════════════════════════════════════════════
COLUMN_ALIASES = {
    "col_cliente":     ["Nombre del solicitante", "Nmbr Solic", "Nombre Solic", "Nmbr.Solic"],
    "col_nombre_vend": ["Nombre vendedor", "Nombre Ven", "Nmbr Vend", "Nombre Vendedor"],
    "col_canal":       ["Den.Of.Vta", "Den Of Vta"],
    "col_venta":       ["Mon.Sol.$", "Mon Sol $"],
    "col_factura":     ["Mon.Fac.$", "Mon Fac $"],
    "col_rechaz":      ["Mnt.Rechaz", "Mnt Rechaz"],
    "col_cant":        ["Cant. Ori.", "Cant.Ori.", "Uni.Solici"],
    "col_entrega":     ["Entrega"],
    "col_doc":         ["Doc.vta.", "Doc.Vta."],
    "col_sku":         ["SKU SAP"],
    "col_material":    ["Texto breve material"],
    "col_estatus_sm":  ["Estatus SM"],
    "col_estatus_fa":  ["Estatus FA"],
    "col_tipo_doc":    ["ClDocVenta"],
    "col_vendedor":    ["Vendedor"],
    "col_bloq":        ["BloqEntreg"],
    "col_den_bloq":    ["Den.Bl.Ent"],
    "col_mot_rech":    ["Mot. Rech.", "Mot.Rech."],
}

def detectar_columnas(df_cols):
    """Detecta automáticamente las columnas del archivo y actualiza CONFIG."""
    cambios = {}
    alertas = []
    for key, aliases in COLUMN_ALIASES.items():
        valor_actual = CONFIG[key]
        if valor_actual not in df_cols:
            encontrado = False
            for alias in aliases:
                if alias in df_cols:
                    CONFIG[key] = alias
                    cambios[key] = (valor_actual, alias)
                    encontrado = True
                    break
            if not encontrado:
                alertas.append(f"⚠️ No se encontró columna para **{key}** (esperaba: {valor_actual})")
    return cambios, alertas

# ══════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📦 Reporte de Pedidos")
    st.markdown("---")
    st.markdown("### 📂 Cargar archivo")
    archivo = st.file_uploader(
        "Arrastra o selecciona el archivo .xls diario",
        type=["xls", "xlsx"],
        help="El archivo de pedidos exportado del sistema"
    )

    st.markdown("---")
    st.markdown("### ⚙️ Opciones")
    top_clientes = st.slider("Top clientes", 10, 200, 100)
    top_skus     = st.slider("Top SKUs",     10, 200, 100)
    CONFIG["top_clientes"] = top_clientes
    CONFIG["top_skus"]     = top_skus

    st.markdown("---")
    st.markdown("### ℹ️ Instrucciones")
    st.markdown("""
    1. Sube el archivo `.xls` del día
    2. Revisa el resumen en pantalla
    3. Descarga el Excel formateado
    """)
    st.markdown("---")
    st.caption("Detecta cambios de columnas automáticamente")

# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════

# Header
st.markdown("""
<div class="main-header">
    <h1 style="margin:0; font-size:2rem;">📦 Reporte Diario de Pedidos</h1>
    <p style="margin:0.5rem 0 0; opacity:0.85;">Procesa el archivo del día, detecta cambios y genera el Excel automáticamente</p>
</div>
""", unsafe_allow_html=True)

if archivo is None:
    st.markdown("""
    <div class="info-box">
        <b>👈 Sube el archivo de pedidos del día en el panel izquierdo para comenzar.</b>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="kpi-card">
            <div class="kpi-label">Paso 1</div>
            <div style="font-size:2rem;">📂</div>
            <b>Sube el .xls</b><br>
            <small>Arrastra el archivo diario en el panel izquierdo</small>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="kpi-card green">
            <div class="kpi-label">Paso 2</div>
            <div style="font-size:2rem;">🔍</div>
            <b>Revisión automática</b><br>
            <small>Detecta columnas, limpia datos y calcula KPIs</small>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="kpi-card amber">
            <div class="kpi-label">Paso 3</div>
            <div style="font-size:2rem;">⬇️</div>
            <b>Descarga el Excel</b><br>
            <small>10 hojas formateadas listas para usar</small>
        </div>""", unsafe_allow_html=True)
    st.stop()

# ── Procesar archivo ──────────────────────────────────────────────
with st.spinner("⏳ Leyendo y procesando el archivo..."):
    try:
        # Guardar archivo temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            tmp.write(archivo.read())
            tmp_path = tmp.name

        # Leer
        df_raw, fecha_archivo = leer_archivo(tmp_path)

        # ── DETECCIÓN AUTOMÁTICA DE COLUMNAS ──────────────────────
        df_cols = list(df_raw.columns)
        cambios, alertas_col = detectar_columnas(df_cols)

        # Limpiar y calcular
        df = limpiar(df_raw)
        kpis, canal_df, cliente_df, fecha_df, sku_df, loc_df, vend_df, se_res, se_det, nc_det, base, sin_e, cenf_res, cenf_det = calcular_pivots(df)

        # Generar Excel en memoria
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
            ruta_salida = tmp_out.name

        nombre_base = os.path.splitext(archivo.name)[0]
        hoy = datetime.today().strftime("%Y-%m-%d")
        nombre_descarga = f"Reporte_{nombre_base}_{hoy}.xlsx"

        escribir_excel(df, fecha_archivo, kpis, canal_df, cliente_df, fecha_df,
                       sku_df, loc_df, vend_df, se_res, se_det, nc_det, base, ruta_salida,
                       cenf_res=cenf_res, cenf_det=cenf_det)

        with open(ruta_salida, "rb") as f:
            excel_bytes = f.read()

        os.unlink(tmp_path)
        os.unlink(ruta_salida)

        st.success(f"✅ Archivo procesado correctamente — {fecha_archivo}")

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
        st.exception(e)
        st.stop()

# ── Alertas de columnas detectadas ───────────────────────────────
if cambios:
    with st.expander(f"🔄 {len(cambios)} columna(s) con nombre diferente al esperado — detectadas automáticamente", expanded=True):
        for key, (antes, despues) in cambios.items():
            st.markdown(f"- **{key}**: `{antes}` → `{despues}` ✅")

if alertas_col:
    for a in alertas_col:
        st.warning(a)

# ── BOTÓN DE DESCARGA ─────────────────────────────────────────────
st.markdown("### ⬇️ Descargar reporte")
st.download_button(
    label=f"📥 Descargar {nombre_descarga}",
    data=excel_bytes,
    file_name=nombre_descarga,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    type="primary"
)

st.markdown("---")

# ══════════════════════════════════════════════════════════════════
# KPIs
# ══════════════════════════════════════════════════════════════════
st.markdown("### 📊 Resumen del día")

col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
with col1:
    st.metric("Pedidos", f"{kpis['total_pedidos']:,.0f}")
with col2:
    st.metric("Venta Sol.", f"${kpis['total_venta']/1e6:.1f}M")
with col3:
    st.metric("Fac. Bruto", f"${kpis['total_fac_bruto']/1e6:.1f}M")
with col4:
    st.metric("NC ZDEV", f"-${kpis['total_nc']/1e6:.1f}M", delta=f"-${kpis['total_nc']/1e6:.1f}M", delta_color="inverse")
with col5:
    st.metric("Fac. Neto", f"${kpis['total_fac_neto']/1e6:.1f}M")
with col6:
    st.metric("Rechazado", f"${kpis['total_rechaz']/1e6:.1f}M",
              delta=f"${kpis['total_rechaz']/1e6:.1f}M", delta_color="inverse")
with col7:
    color = "normal" if kpis['tasa_sm'] >= 90 else "inverse"
    st.metric("Tasa WMS", f"{kpis['tasa_sm']:.1f}%", delta=f"{'✅' if kpis['tasa_sm']>=90 else '⚠️'}")

st.markdown("---")

# ══════════════════════════════════════════════════════════════════
# ALERTAS AUTOMÁTICAS
# ══════════════════════════════════════════════════════════════════
st.markdown("### 🚨 Alertas del día")

col_a, col_b = st.columns(2)
with col_a:
    if kpis['tasa_sm'] < 90:
        st.markdown(f"""<div class="alert-box">⚠️ <b>Tasa WMS bajo el 90%:</b> {kpis['tasa_sm']:.1f}% — Revisar pedidos pendientes</div>""", unsafe_allow_html=True)
    else:
        st.markdown(f"""<div class="success-box">✅ <b>Tasa WMS OK:</b> {kpis['tasa_sm']:.1f}%</div>""", unsafe_allow_html=True)

    if kpis['total_se_ped'] > 0:
        st.markdown(f"""<div class="warning-box">📋 <b>{kpis['total_se_ped']} pedidos sin entrega</b> — ${kpis['total_se_val']/1e6:.1f}M en riesgo</div>""", unsafe_allow_html=True)

with col_b:
    if kpis['total_nc'] > 0:
        st.markdown(f"""<div class="alert-box">🔴 <b>Notas de crédito ZDEV:</b> -${kpis['total_nc']/1e6:.1f}M restados del facturado</div>""", unsafe_allow_html=True)
    if kpis['total_rechaz'] > 0:
        st.markdown(f"""<div class="warning-box">🟠 <b>Monto rechazado:</b> ${kpis['total_rechaz']/1e6:.1f}M</div>""", unsafe_allow_html=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════════════
# TABLAS PREVIEW
# ══════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs(["📦 Por Canal", "🚨 Sin Entrega", "👥 Top Clientes", "🧑‍💼 Vendedores"])

with tab1:
    st.markdown("#### Ventas por Canal")
    canal_show = canal_df[['Canal','Pedidos','Lineas','Venta_Sol','Fac_Bruto','NC','Fac_Neto','Rechazado','Part_%']].copy()
    canal_show.columns = ['Canal','Pedidos','Líneas','Venta Sol.$','Fac. Bruto $','NC ZDEV $','Fac. Neto $','Rechazado $','Part. %']
    canal_show['Venta Sol.$']  = canal_show['Venta Sol.$'].apply(lambda x: f"${x:,.0f}")
    canal_show['Fac. Bruto $'] = canal_show['Fac. Bruto $'].apply(lambda x: f"${x:,.0f}")
    canal_show['NC ZDEV $']    = canal_show['NC ZDEV $'].apply(lambda x: f"-${x:,.0f}" if x>0 else "-")
    canal_show['Fac. Neto $']  = canal_show['Fac. Neto $'].apply(lambda x: f"${x:,.0f}")
    canal_show['Rechazado $']  = canal_show['Rechazado $'].apply(lambda x: f"${x:,.0f}" if x>0 else "-")
    canal_show['Part. %']      = canal_show['Part. %'].apply(lambda x: f"{x*100:.1f}%")
    st.dataframe(canal_show, use_container_width=True, hide_index=True)

with tab2:
    st.markdown(f"#### {len(se_res)} clientes con pedidos sin entrega")
    if len(se_res) > 0:
        se_show = se_res[['Canal','Cliente',CONFIG['col_vendedor'],CONFIG['col_nombre_vend'],'Pedidos','Lineas','Venta_Sol','Bloq','Motivo']].copy()
        se_show.columns = ['Canal','Cliente','Cód. Vendedor','Nombre Vendedor','Pedidos','Líneas','Venta Sol.$','Bloqueados','Motivo']
        se_show['Venta Sol.$'] = se_show['Venta Sol.$'].apply(lambda x: f"${x:,.0f}")
        se_show['Bloqueados']  = se_show['Bloqueados'].astype(int)
        st.dataframe(se_show, use_container_width=True, hide_index=True)
    else:
        st.success("✅ No hay pedidos sin entrega")

with tab3:
    st.markdown(f"#### Top {min(20, len(cliente_df))} clientes (preview)")
    cli_show = cliente_df.head(20)[['Canal','Cliente',CONFIG['col_nombre_vend'],'Pedidos','Venta_Sol','Fac_Neto','Rechazado']].copy()
    cli_show.columns = ['Canal','Cliente','Vendedor','Pedidos','Venta Sol.$','Fac. Neto $','Rechazado $']
    cli_show['Venta Sol.$']  = cli_show['Venta Sol.$'].apply(lambda x: f"${x:,.0f}")
    cli_show['Fac. Neto $']  = cli_show['Fac. Neto $'].apply(lambda x: f"${x:,.0f}")
    cli_show['Rechazado $']  = cli_show['Rechazado $'].apply(lambda x: f"${x:,.0f}" if x>0 else "-")
    st.dataframe(cli_show, use_container_width=True, hide_index=True)

with tab4:
    st.markdown(f"#### Top {min(20, len(vend_df))} vendedores (preview)")
    vend_show = vend_df.head(20)[['Canal',CONFIG['col_nombre_vend'],'Pedidos','Clientes','Venta_Sol','Fac_Neto','Rechazado','SE_Ped']].copy()
    vend_show.columns = ['Canal','Vendedor','Pedidos','Clientes','Venta Sol.$','Fac. Neto $','Rechazado $','Sin Entrega']
    vend_show['Venta Sol.$']  = vend_show['Venta Sol.$'].apply(lambda x: f"${x:,.0f}")
    vend_show['Fac. Neto $']  = vend_show['Fac. Neto $'].apply(lambda x: f"${x:,.0f}")
    vend_show['Rechazado $']  = vend_show['Rechazado $'].apply(lambda x: f"${x:,.0f}" if x>0 else "-")
    vend_show['Sin Entrega']  = vend_show['Sin Entrega'].astype(int)
    st.dataframe(vend_show, use_container_width=True, hide_index=True)

st.markdown("---")
st.caption(f"Generado el {datetime.today().strftime('%d/%m/%Y %H:%M')} · Datos: {fecha_archivo}")
