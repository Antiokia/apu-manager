import streamlit as st
import pandas as pd
from engine import procesar_archivo
from drive_helper import (
    subir_a_drive, drive_configurado,
    generar_url_autorizacion, obtener_refresh_token,
    _tiene_refresh_token
)

# ── CONFIG ────────────────────────────────────────────────────
st.set_page_config(
    page_title="APU Manager · IENEL",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── ESTILOS ───────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Barlow:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif;
}

/* Fondo */
.stApp { background: #0d1117; }

/* Ocultar menú hamburguesa y footer */
#MainMenu, footer, header { visibility: hidden; }

/* ── HERO ── */
.hero {
    background: linear-gradient(135deg, #00467F 0%, #003560 50%, #001f3d 100%);
    border-radius: 16px;
    padding: 40px 48px 36px;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: "";
    position: absolute;
    top: -60px; right: -60px;
    width: 300px; height: 300px;
    background: radial-gradient(circle, rgba(255,180,0,0.15) 0%, transparent 70%);
    border-radius: 50%;
}
.hero-logo {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 4px;
    color: #FFB400;
    text-transform: uppercase;
    margin-bottom: 10px;
}
.hero-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 52px;
    font-weight: 800;
    color: #ffffff;
    line-height: 1;
    margin: 0 0 8px;
}
.hero-sub {
    font-size: 16px;
    font-weight: 300;
    color: rgba(255,255,255,0.6);
    margin: 0;
}
.hero-badge {
    display: inline-block;
    background: rgba(255,180,0,0.15);
    border: 1px solid rgba(255,180,0,0.4);
    color: #FFB400;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 2px;
    padding: 4px 12px;
    border-radius: 20px;
    margin-top: 16px;
}

/* ── UPLOAD ZONE ── */
.upload-card {
    background: #161b22;
    border: 2px dashed #30363d;
    border-radius: 12px;
    padding: 40px;
    text-align: center;
    transition: border-color .2s;
    margin-bottom: 24px;
}
.upload-card:hover { border-color: #FFB400; }
.upload-label {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 20px;
    font-weight: 600;
    color: #8b949e;
    margin-bottom: 6px;
}
.upload-hint { font-size: 13px; color: #484f58; }

/* ── KPI CARDS ── */
.kpi-row { display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }
.kpi-card {
    flex: 1;
    min-width: 160px;
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 12px;
    padding: 20px 24px;
    position: relative;
    overflow: hidden;
}
.kpi-card::after {
    content: "";
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
    background: #FFB400;
}
.kpi-label {
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 2px;
    color: #8b949e;
    text-transform: uppercase;
    margin-bottom: 8px;
}
.kpi-value {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 32px;
    font-weight: 700;
    color: #ffffff;
    line-height: 1;
}
.kpi-sub { font-size: 12px; color: #484f58; margin-top: 4px; }

/* ── CATEGORÍA HEADER ── */
.cat-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 14px 20px;
    border-radius: 8px 8px 0 0;
    margin-top: 28px;
    margin-bottom: 0;
}
.cat-dot {
    width: 10px; height: 10px;
    border-radius: 50%;
    background: white;
    flex-shrink: 0;
}
.cat-name {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 17px;
    font-weight: 700;
    letter-spacing: 1px;
    color: white;
    text-transform: uppercase;
}
.cat-total {
    margin-left: auto;
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 17px;
    font-weight: 700;
    color: white;
}

/* ── DATAFRAME ── */
.stDataFrame { border-radius: 0 0 8px 8px; overflow: hidden; }
[data-testid="stDataFrameContainer"] { border: none !important; }

/* ── TOTAL GENERAL ── */
.total-bar {
    background: linear-gradient(90deg, #00467F 0%, #003560 100%);
    border-radius: 10px;
    padding: 20px 32px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-top: 32px;
}
.total-label {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 18px;
    font-weight: 700;
    letter-spacing: 3px;
    color: rgba(255,255,255,0.7);
    text-transform: uppercase;
}
.total-value {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 42px;
    font-weight: 800;
    color: #FFB400;
}

/* ── BOTÓN DESCARGA ── */
.stDownloadButton button {
    background: #FFB400 !important;
    color: #001f3d !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 15px !important;
    font-weight: 700 !important;
    letter-spacing: 1px !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 28px !important;
    transition: all .2s !important;
}
.stDownloadButton button:hover {
    background: #ffc933 !important;
    transform: translateY(-1px);
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    background: #161b22;
    border: 2px dashed #30363d;
    border-radius: 12px;
    padding: 8px;
}
[data-testid="stFileUploader"]:hover { border-color: #FFB400; }
[data-testid="stFileUploaderDropzone"] { background: transparent !important; }
[data-testid="stFileUploaderDropzone"] label { color: #8b949e !important; }
[data-testid="stFileUploaderDropzoneInstructions"] { color: #484f58 !important; }

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: #161b22;
    border-radius: 8px;
    padding: 4px;
    gap: 4px;
    margin-bottom: 20px;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 14px;
    font-weight: 600;
    letter-spacing: 1px;
    color: #8b949e;
    border-radius: 6px;
    padding: 8px 20px;
}
.stTabs [aria-selected="true"] {
    background: #00467F !important;
    color: white !important;
}

/* ── ITEMS TABLE ── */
.items-row {
    background: #161b22;
    border: 1px solid #21262d;
    border-radius: 8px;
    padding: 12px 20px;
    display: flex;
    gap: 24px;
    margin-bottom: 8px;
    align-items: center;
}
.items-num {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 20px;
    font-weight: 700;
    color: #FFB400;
    width: 40px;
}
.items-info { flex: 1; }
.items-cant { font-size: 13px; color: #8b949e; }
.items-base {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 16px;
    font-weight: 600;
    color: #ffffff;
}

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #0d1117; }
::-webkit-scrollbar-thumb { background: #30363d; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)


# ── HERO HEADER ──────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-logo">⚡ IENEL · Sistema de Presupuestos</div>
    <div class="hero-title">APU Manager</div>
    <p class="hero-sub">Análisis de Precios Unitarios · Explosión de Insumos · Consolidado</p>
    <div class="hero-badge">VERSIÓN PYTHON · Beta</div>
</div>
""", unsafe_allow_html=True)


# ── UPLOAD ───────────────────────────────────
col_up, col_info = st.columns([2, 1])

with col_up:
    archivo = st.file_uploader(
        "Sube tu formato de presupuesto (.xlsm / .xlsx / .xls)",
        type=["xlsm", "xlsx", "xls"],
        label_visibility="visible",
    )

with col_info:
    st.markdown("""
    <div style="background:#161b22; border:1px solid #21262d; border-radius:12px; padding:20px; margin-top:8px;">
        <div style="font-family:'Barlow Condensed',sans-serif; font-size:15px; font-weight:700; color:#FFB400; margin-bottom:12px; letter-spacing:1px;">REQUISITOS DEL ARCHIVO</div>
        <div style="font-size:13px; color:#8b949e; line-height:1.8;">
            ✓ &nbsp;Hoja <b style="color:#c9d1d9">PPTO</b> con ítems y cantidades<br>
            ✓ &nbsp;Hojas <b style="color:#c9d1d9">item 1</b>, <b style="color:#c9d1d9">item 2</b>… con APU<br>
            ✓ &nbsp;Hoja <b style="color:#c9d1d9">AUX</b> con tabla de categorías<br>
            ✓ &nbsp;Formatos <b style="color:#c9d1d9">.xlsm · .xlsx · .xls</b>
        </div>
    </div>
    """, unsafe_allow_html=True)


# ── PANEL OAUTH (solo si Drive configurado pero sin refresh_token) ──
if drive_configurado() and not _tiene_refresh_token():
    st.markdown("""
    <div style="background:#1c2128; border:1px solid #FFB400; border-radius:10px;
                padding:20px 24px; margin-bottom:24px;">
        <div style="font-family:'Barlow Condensed',sans-serif; font-size:16px;
                    font-weight:700; color:#FFB400; margin-bottom:8px; letter-spacing:1px;">
            🔑 AUTORIZACIÓN DE GOOGLE DRIVE PENDIENTE
        </div>
        <div style="font-size:13px; color:#8b949e; margin-bottom:16px;">
            Solo debes hacer esto <b style="color:#c9d1d9">una vez</b>.
            Sigue los pasos para conectar tu Drive personal.
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_auth1, col_auth2 = st.columns([1, 1])

    with col_auth1:
        if st.button("1️⃣  Generar link de autorización", use_container_width=True):
            try:
                url = generar_url_autorizacion()
                st.session_state["auth_url"] = url
            except Exception as e:
                st.error(f"Error generando URL: {e}")

        if "auth_url" in st.session_state:
            st.markdown(f"""
            <div style="background:#0d1117; border:1px solid #30363d; border-radius:8px;
                        padding:12px; margin-top:8px; word-break:break-all;">
                <div style="font-size:11px; color:#8b949e; margin-bottom:6px;">Abre este link en tu navegador:</div>
                <a href="{st.session_state['auth_url']}" target="_blank"
                   style="color:#58a6ff; font-size:12px;">
                    🔗 Abrir autorización de Google
                </a>
            </div>
            """, unsafe_allow_html=True)

    with col_auth2:
        codigo = st.text_input(
            "2️⃣  Pega aquí el código que te dio Google",
            placeholder="4/0AX4XfWi...",
        )
        if st.button("Conectar Drive ✓", use_container_width=True):
            if codigo:
                try:
                    rt = obtener_refresh_token(codigo)
                    st.success("✅ ¡Conectado! Copia este refresh_token en tus Secrets de Streamlit:")
                    st.code(f'[oauth]\nrefresh_token = "{rt}"', language="toml")
                    st.info("Después de pegarlo en Secrets, la app se reconecta sola.")
                except Exception as e:
                    st.error(f"Error: {e}")
            else:
                st.warning("Pega el código primero.")

    st.divider()


# ── PROCESAMIENTO ─────────────────────────────
if archivo is not None:
    try:
        file_bytes = archivo.read()

        with st.spinner("Procesando APUs..."):
            tabla, subtotales, total_general, excel_bytes, items_info, cats = procesar_archivo(file_bytes)

        # ── SUBIDA A DRIVE (silenciosa) ────────
        if drive_configurado():
            try:
                subir_a_drive(file_bytes, archivo.name)
            except Exception:
                pass

        # ── KPIs ──────────────────────────────
        n_items     = len(items_info)
        n_cats      = len(tabla)
        n_insumos   = sum(len(v) for v in tabla.values())

        st.markdown(f"""
        <div class="kpi-row">
            <div class="kpi-card">
                <div class="kpi-label">Total General</div>
                <div class="kpi-value">$ {total_general:,.0f}</div>
                <div class="kpi-sub">Costo directo consolidado</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Ítems APU</div>
                <div class="kpi-value">{n_items}</div>
                <div class="kpi-sub">Hojas procesadas</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Categorías</div>
                <div class="kpi-value">{n_cats}</div>
                <div class="kpi-sub">Con insumos activos</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">Insumos únicos</div>
                <div class="kpi-value">{n_insumos}</div>
                <div class="kpi-sub">Consolidados</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ── TABS ──────────────────────────────
        tab1, tab2, tab3 = st.tabs(["📋  EXPLOSIÓN DE INSUMOS", "📊  RESUMEN POR CATEGORÍA", "🔢  DETALLE ÍTEMS"])

        # Mapa de colores de cats
        cat_color_map = {c["categoria"]: c["hex"] for c in cats}

        # ── TAB 1: Explosión ──────────────────
        with tab1:
            st.markdown("""
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
                <div style="font-family:'Barlow Condensed',sans-serif; font-size:22px; font-weight:700; color:#c9d1d9;">
                    Consolidado de insumos por categoría
                </div>
            </div>
            """, unsafe_allow_html=True)

            col_dl, _ = st.columns([1, 3])
            with col_dl:
                st.download_button(
                    label="⬇  Descargar Excel",
                    data=excel_bytes,
                    file_name=f"{archivo.name.rsplit('.',1)[0]}_explosion.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            for cat_nombre, filas in tabla.items():
                subtotal = subtotales[cat_nombre]
                hex_color = cat_color_map.get(cat_nombre, "444444")
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)

                st.markdown(f"""
                <div class="cat-header" style="background:rgba({r},{g},{b},0.85);">
                    <div class="cat-dot"></div>
                    <div class="cat-name">{cat_nombre}</div>
                    <div class="cat-total">$ {subtotal:,.0f}</div>
                </div>
                """, unsafe_allow_html=True)

                # Construir DataFrame de la categoría
                es_pct = filas[0]["es_pct"] if filas else False
                rows_df = []
                for f in filas:
                    cant_display = f"{f['cantidad']:.2%}" if f["es_pct"] else f"{f['cantidad']:,.2f}"
                    rows_df.append({
                        "Clave":          f["clave"],
                        "Descripción":    f["descripcion"],
                        "Unidad":         f["unidad"],
                        "Cantidad":       cant_display,
                        "Costo Unit.":    f"$ {f['costo']:,.0f}",
                        "Valor Total":    f"$ {f['valor_total']:,.0f}",
                    })

                df = pd.DataFrame(rows_df)
                st.dataframe(
                    df,
                    use_container_width=True,
                    hide_index=True,
                    height=min(38 * len(df) + 38, 400),
                )

            # Total general
            st.markdown(f"""
            <div class="total-bar">
                <div class="total-label">Total General Costo Directo</div>
                <div class="total-value">$ {total_general:,.0f}</div>
            </div>
            """, unsafe_allow_html=True)

        # ── TAB 2: Resumen ────────────────────
        with tab2:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

            # Tabla resumen
            resumen_rows = []
            for cat_nombre, subtotal in subtotales.items():
                pct = subtotal / total_general * 100 if total_general > 0 else 0
                resumen_rows.append({
                    "Categoría":   cat_nombre,
                    "Subtotal":    f"$ {subtotal:,.0f}",
                    "% del Total": f"{pct:.1f}%",
                    "Insumos":     len(tabla[cat_nombre]),
                })

            df_res = pd.DataFrame(resumen_rows)
            st.dataframe(df_res, use_container_width=True, hide_index=True)

            # Barras simples en HTML
            st.markdown("<div style='margin-top:24px; margin-bottom:8px; font-family:Barlow Condensed,sans-serif; font-size:16px; font-weight:700; color:#8b949e; letter-spacing:2px;'>DISTRIBUCIÓN</div>", unsafe_allow_html=True)

            for cat_nombre, subtotal in subtotales.items():
                pct = subtotal / total_general * 100 if total_general > 0 else 0
                hex_color = cat_color_map.get(cat_nombre, "444444")
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                st.markdown(f"""
                <div style="margin-bottom:10px;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
                        <span style="font-size:13px; color:#c9d1d9;">{cat_nombre}</span>
                        <span style="font-size:13px; font-weight:600; color:#FFB400;">{pct:.1f}%</span>
                    </div>
                    <div style="background:#21262d; border-radius:4px; height:8px; overflow:hidden;">
                        <div style="width:{pct:.1f}%; height:100%; background:rgb({r},{g},{b}); border-radius:4px; transition:width .5s;"></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # ── TAB 3: Detalle ítems ──────────────
        with tab3:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            st.markdown("""
            <div style="font-family:'Barlow Condensed',sans-serif; font-size:16px; color:#8b949e; letter-spacing:2px; margin-bottom:16px;">
                ÍTEMS PROCESADOS DEL PRESUPUESTO
            </div>
            """, unsafe_allow_html=True)

            for it in items_info:
                st.markdown(f"""
                <div class="items-row">
                    <div class="items-num">#{it['item']:02d}</div>
                    <div class="items-info">
                        <div class="items-cant">Cantidad en PPTO: {it['cantidad']:g} und</div>
                        <div class="items-base">Base directa: $ {it['base']:,.0f}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error procesando el archivo: {e}")
        st.exception(e)

else:
    # Estado vacío
    st.markdown("""
    <div style="text-align:center; padding:80px 0; color:#30363d;">
        <div style="font-size:72px; margin-bottom:16px;">📂</div>
        <div style="font-family:'Barlow Condensed',sans-serif; font-size:22px; font-weight:600; letter-spacing:2px;">
            Sube un archivo para comenzar
        </div>
        <div style="font-size:14px; margin-top:8px;">
            Compatible con el formato IENEL de APU con macro VBA
        </div>
    </div>
    """, unsafe_allow_html=True)
