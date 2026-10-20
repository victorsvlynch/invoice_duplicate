import os
import datetime
import pandas as pd
import streamlit as st

from src.matcher import check_invoices_batch, parse_invoice_entry
from src.excel_loader import (
    DEFAULT_SHAREPOINT_URL,
    DEFAULT_LOCAL_FILE,
    LOCAL_DATA_DIR,
    download_sharepoint_file,
    load_excel_data,
    get_excel_sheet_names,
    get_column_p_name,
    ensure_data_dir,
)
from src.report_generator import generate_excel_report

# Page Configuration
st.set_page_config(
    page_title="Verificador de Facturas Duplicadas | Logística",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Example sample provided by user
SAMPLE_INVOICES = """F85S-0000064535
F85S-0000064619
F72P-0000047299
F73P-0000016233
F75P-0000097864
F71P-0000013193
F116-0000101026
F118-0000075094
F124-0000088458
F123-0000042122
F020-0000003752
FP02-0000069670
FP02-0000069989
FP02-0000070357
FP01-0000066926
FP03-0000034670
F001-0000178498
F001-0000178573
F001-0000178775
FF02-0000106137
FF02-0000106659
FF02-0000107266
FF01-0000047971"""

ensure_data_dir()

# Custom CSS for styling
st.markdown(
    """
    <style>
    .main-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1F497D;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #555555;
        margin-bottom: 1.2rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 12px;
        border-left: 5px solid #1F497D;
    }
    .badge-dup {
        background-color: #ffcccc;
        color: #990000;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    .badge-ok {
        background-color: #d4edda;
        color: #155724;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    /* Ocultar botón Deploy de la esquina superior derecha */
    .stAppDeployButton,
    [data-testid="stAppDeployButton"],
    [data-testid="stToolbarActions"],
    header [data-testid="stToolbarActions"],
    .stDeployButton {
        display: none !important;
        visibility: hidden !important;
    }
    /* Ocultar menú hamburguesa y decoración superior */
    #MainMenu {
        visibility: hidden !important;
    }
    [data-testid="stDecoration"] {
        display: none !important;
    }
    /* Ocultar 'Made with Streamlit' y footer */
    footer,
    [data-testid="stFooter"] {
        display: none !important;
        visibility: hidden !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def init_session_state():
    if "pasted_text" not in st.session_state:
        st.session_state["pasted_text"] = ""
    if "last_results" not in st.session_state:
        st.session_state["last_results"] = None


init_session_state()

# -----------------------------------------------------------------------------
# SIDEBAR: FUENTE DEL EXCEL & CONFIGURACIÓN
# -----------------------------------------------------------------------------
st.sidebar.image("https://img.icons8.com/color/96/microsoft-excel-2019--v1.png", width=64)
st.sidebar.title("Configuración Excel")

source_option = st.sidebar.radio(
    "Origen del archivo Excel:",
    [
        "📁 Subir Archivo Excel",
        "💾 Archivo Local en Disco (data/)",
        "🌐 Descargar de SharePoint",
    ],
    index=0,
)

excel_file_source = None
file_label = ""

if source_option == "📁 Subir Archivo Excel":
    uploaded = st.sidebar.file_uploader(
        "Subir archivo 'Requ & OC.xlsx':",
        type=["xlsx", "xls", "xlsm"],
        help="Sube directamente el archivo Excel descargado de SharePoint.",
    )
    if uploaded is not None:
        excel_file_source = uploaded
        file_label = uploaded.name
        # Persist to local data folder as well
        try:
            with open(DEFAULT_LOCAL_FILE, "wb") as f:
                f.write(uploaded.getbuffer())
        except Exception:
            pass
    elif os.path.exists(DEFAULT_LOCAL_FILE):
        st.sidebar.info("💡 Se encontró una versión previa guardada en `data/Requ & OC.xlsx`.")
        if st.sidebar.button("Cargar versión guardada"):
            excel_file_source = DEFAULT_LOCAL_FILE
            file_label = "data/Requ & OC.xlsx (Guardado previo)"

elif source_option == "💾 Archivo Local en Disco (data/)":
    if os.path.exists(DEFAULT_LOCAL_FILE):
        file_size_mb = os.path.getsize(DEFAULT_LOCAL_FILE) / (1024 * 1024)
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(DEFAULT_LOCAL_FILE))
        st.sidebar.success(f"Archivo encontrado: `{os.path.basename(DEFAULT_LOCAL_FILE)}`")
        st.sidebar.caption(f"Tamaño: {file_size_mb:.2f} MB | Modificado: {mtime.strftime('%d/%m/%Y %H:%M')}")
        excel_file_source = DEFAULT_LOCAL_FILE
        file_label = f"data/Requ & OC.xlsx ({file_size_mb:.2f} MB)"
    else:
        st.sidebar.warning(f"No existe el archivo `{DEFAULT_LOCAL_FILE}`.")
        st.sidebar.write("Coloca el archivo en la carpeta `data/` o utiliza la opción de Subir Archivo.")

elif source_option == "🌐 Descargar de SharePoint":
    sp_url = st.sidebar.text_input(
        "Enlace de SharePoint:",
        value=DEFAULT_SHAREPOINT_URL,
        help="Enlace completo a SharePoint o OneDrive.",
    )
    sp_cookies = st.sidebar.text_input(
        "Cookies de sesión (Opcional, si requiere login):",
        value="",
        type="password",
        help="Si el SharePoint es privado corporativo, pega aquí las cookies rtFa y FedAuth.",
    )

    if st.sidebar.button("Descargar desde SharePoint", type="secondary"):
        with st.sidebar.status("Descargando archivo..."):
            success, msg = download_sharepoint_file(sp_url, sp_cookies, DEFAULT_LOCAL_FILE)
            if success:
                st.sidebar.success(msg)
                excel_file_source = DEFAULT_LOCAL_FILE
                file_label = "SharePoint (Descargado)"
            else:
                st.sidebar.error(msg)

    if os.path.exists(DEFAULT_LOCAL_FILE) and excel_file_source is None:
        excel_file_source = DEFAULT_LOCAL_FILE
        file_label = "data/Requ & OC.xlsx (Caché local)"

# -----------------------------------------------------------------------------
# CARGA DEL EXCEL Y DETECCIÓN DE COLUMNA P
# -----------------------------------------------------------------------------
df_excel = None
selected_col = None
header_row = 0

if excel_file_source is not None:
    sheets = get_excel_sheet_names(excel_file_source)
    selected_sheet = st.sidebar.selectbox("Hoja de Excel:", sheets, index=0)
    header_row_choice = st.sidebar.number_input("Fila de Encabezados (0 = fila 1 de Excel):", min_value=0, max_value=20, value=0)
    header_row = int(header_row_choice)

    with st.spinner("Leyendo estructura del Excel..."):
        df_excel, _, err = load_excel_data(excel_file_source, selected_sheet, header_row)

    if err:
        st.sidebar.error(err)
    elif df_excel is not None:
        st.sidebar.success(f"Excel cargado: {len(df_excel):,} filas | {len(df_excel.columns)} columnas")

        # Auto detect Column P (index 15)
        default_col_p, p_idx = get_column_p_name(df_excel)
        col_list = list(df_excel.columns)
        default_idx = col_list.index(default_col_p) if default_col_p in col_list else 0

        # Format column display: "Columna P (Índice 15): [Nombre]"
        def format_col_name(c):
            idx = col_list.index(c)
            letter = chr(65 + idx) if idx < 26 else f"Col{idx+1}"
            return f"Columna {letter} (idx {idx}): {c}"

        selected_col = st.sidebar.selectbox(
            "Columna a verificar (Por defecto Columna P):",
            col_list,
            index=default_idx,
            format_func=format_col_name,
            help="Selecciona la columna donde están los números de comprobante/factura en el Excel.",
        )

# Matching Configuration
st.sidebar.markdown("---")
st.sidebar.subheader("Criterios de Coincidencia")
match_exact = st.sidebar.checkbox(
    "1. Coincidencia Exacta (Serie + Número)",
    value=True,
    help="Ejemplo: F85S-0000064535 coincide exactamente con F85S-0000064535.",
)
match_zeros = st.sidebar.checkbox(
    "2. Por Correlativo con ceros",
    value=True,
    help="Ejemplo: 0000064535 coincide con cualquier registro que contenga este correlativo con ceros.",
)
match_no_zeros = st.sidebar.checkbox(
    "3. Por Número sin ceros",
    value=True,
    help="Ejemplo: 64535 coincide con 0000064535 o F85S-0000064535.",
)
match_diff_series = st.sidebar.checkbox(
    "⚠️ Alertar si coincide el número pero con Serie distinta",
    value=True,
    help="Ejemplo: si buscas F85S-0000064535 y en el Excel existe F001-0000064535, te alertará para revisar si hubo error en la serie.",
)


# -----------------------------------------------------------------------------
# MAIN VIEW
# -----------------------------------------------------------------------------
st.markdown('<div class="main-title">📑 Sistema de Verificación de Facturas Duplicadas</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Verifica comprobantes pegados contra la <b>Columna P</b> del archivo Excel de Requerimientos & Órdenes de Compra.</div>',
    unsafe_allow_html=True,
)

if df_excel is not None and selected_col is not None:
    st.info(
        f"📊 **Archivo activo:** `{file_label}` | **Hoja:** `{selected_sheet}` | "
        f"**Columna de búsqueda:** `{selected_col}` ({len(df_excel):,} registros indexados)"
    )
else:
    st.warning("⚠️ Carga o sube primero el archivo Excel `Requ & OC.xlsx` en el panel izquierdo para comenzar la verificación.")

# Query text input area
col_top1, col_top2 = st.columns([3, 1])

with col_top1:
    st.subheader("Pega aquí la lista de facturas a verificar:")

with col_top2:
    st.write("")
    btn_sample = st.button("📋 Cargar Ejemplo de Prueba", use_container_width=True)
    if btn_sample:
        st.session_state["pasted_text"] = SAMPLE_INVOICES

user_text = st.text_area(
    "Ingresa un código por línea (puedes pegar series completas, con ceros o sin ceros):",
    value=st.session_state["pasted_text"],
    height=220,
    placeholder="F85S-0000064535\n0000064535\n64535\nFP02-0000069670\n...",
)
st.session_state["pasted_text"] = user_text

col_btn1, col_btn2, _ = st.columns([2, 1, 3])
with col_btn1:
    btn_verify = st.button("🔍 Verificar Duplicados en Excel", type="primary", use_container_width=True)
with col_btn2:
    if st.button("🗑️ Limpiar", use_container_width=True):
        st.session_state["pasted_text"] = ""
        st.session_state["last_results"] = None
        st.rerun()

# -----------------------------------------------------------------------------
# EXECUTION & RESULTS
# -----------------------------------------------------------------------------
if btn_verify:
    if df_excel is None or selected_col is None:
        st.error("No se ha cargado ningún archivo Excel. Por favor, sube el archivo 'Requ & OC.xlsx' en la barra lateral.")
    elif not user_text.strip():
        st.error("Por favor, ingresa o pega al menos una factura para verificar.")
    else:
        with st.spinner("Analizando y verificando contra la base de datos de Excel..."):
            summary_df, details_df, stats = check_invoices_batch(
                df=df_excel,
                col_name=selected_col,
                raw_query_text=user_text,
                match_exact=match_exact,
                match_zeros=match_zeros,
                match_no_zeros=match_no_zeros,
                match_diff_series=match_diff_series,
                header_row_offset=header_row + 2,
            )
            st.session_state["last_results"] = {
                "summary_df": summary_df,
                "details_df": details_df,
                "stats": stats,
            }

# Render Results
if st.session_state.get("last_results"):
    results = st.session_state["last_results"]
    summary_df = results["summary_df"]
    details_df = results["details_df"]
    stats = results["stats"]

    st.markdown("---")
    st.subheader("Resultados de la Verificación")

    # Metrics
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Total Consultados", f"{stats['total_queries']} comprobantes")
    with m2:
        st.metric(
            "Duplicados en Excel",
            f"{stats['found_count']}",
            delta=f"-{stats['found_count']} ya registrados" if stats["found_count"] > 0 else "0",
            delta_color="inverse",
        )
    with m3:
        st.metric(
            "No Encontrados (Limpios)",
            f"{stats['not_found_count']}",
            delta=f"+{stats['not_found_count']} listos para registrar" if stats["not_found_count"] > 0 else "0",
            delta_color="normal",
        )
    with m4:
        st.metric("Repetidos en Lista", f"{stats['internal_duplicates']}")

    # Alert banners
    if stats["found_count"] > 0:
        st.warning(
            f"⚠️ Se detectaron **{stats['found_count']} comprobantes** que ya existen en el Excel "
            f"({stats['multiple_matches_count']} tienen múltiples coincidencias)."
        )
    else:
        st.success("🎉 ¡Excelente! Ninguna de las facturas consultadas se encuentra duplicada en el archivo Excel.")

    # Export Buttons
    col_dl1, col_dl2, _ = st.columns([2, 2, 4])
    with col_dl1:
        report_bytes = generate_excel_report(summary_df, details_df, stats)
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="📥 Descargar Reporte Completo (.xlsx)",
            data=report_bytes,
            file_name=f"Reporte_Duplicados_{now_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col_dl2:
        csv_data = summary_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📄 Descargar Resumen (.csv)",
            data=csv_data,
            file_name=f"Resumen_Duplicados_{now_str}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # Tabs for detailed exploration
    tab_resumen, tab_detalles, tab_no_encontrados, tab_preview = st.tabs(
        [
            "📋 Resumen General",
            f"🔍 Detalle de Coincidencias en Excel ({len(details_df)})",
            f"✅ Solo No Encontrados ({stats['not_found_count']})",
            "👀 Vista Previa de Columna P en Excel",
        ]
    )

    with tab_resumen:
        col_f1, col_f2 = st.columns([2, 2])
        with col_f1:
            filter_status = st.selectbox(
                "Filtrar por estado:",
                ["Todos los registros", "Solo Duplicados", "Solo No Encontrados"],
                index=0,
            )
        with col_f2:
            search_box = st.text_input("Filtrar por texto:", placeholder="Ej. F85S, 64535...")

        filtered_summary = summary_df.copy()
        if filter_status == "Solo Duplicados":
            filtered_summary = filtered_summary[filtered_summary["Estado en Excel"].str.contains("DUPLICADO")]
        elif filter_status == "Solo No Encontrados":
            filtered_summary = filtered_summary[filtered_summary["Estado en Excel"] == "NO ENCONTRADO"]

        if search_box:
            filtered_summary = filtered_summary[
                filtered_summary.astype(str).apply(lambda row: row.str.contains(search_box, case=False).any(), axis=1)
            ]

        # Style dataframe for Streamlit
        def highlight_status(val):
            if "DUPLICADO" in str(val):
                return "background-color: #ffcccc; color: #800000; font-weight: bold;"
            elif val == "NO ENCONTRADO":
                return "background-color: #d4edda; color: #155724; font-weight: bold;"
            return ""

        styled_summary = filtered_summary.style.map(highlight_status, subset=["Estado en Excel"])
        st.dataframe(styled_summary, use_container_width=True, height=450)

    with tab_detalles:
        if not details_df.empty:
            st.markdown(
                "Esta tabla muestra exactamente en **qué fila del Excel** se encontró cada comprobante, "
                "el tipo de coincidencia detectada y los datos adyacentes de esa fila:"
            )
            st.dataframe(details_df, use_container_width=True, height=500)
        else:
            st.info("No hay coincidencias que mostrar.")

    with tab_no_encontrados:
        not_found_items = summary_df[summary_df["Estado en Excel"] == "NO ENCONTRADO"]
        if not not_found_items.empty:
            st.success(f"Se encontraron **{len(not_found_items)} comprobantes nuevos** listos para registrar:")
            st.dataframe(
                not_found_items[
                    ["N°", "Texto Consultado", "Serie Detectada", "Correlativo Ceros", "Número Sin Ceros"]
                ],
                use_container_width=True,
                height=350,
            )

            # Clean copy-paste text box
            plain_text_clean = "\n".join(not_found_items["Texto Consultado"].tolist())
            st.text_area("Copiar lista limpia de comprobantes nuevos:", value=plain_text_clean, height=150)
        else:
            st.info("Todos los comprobantes de la lista están duplicados en el Excel.")

    with tab_preview:
        if df_excel is not None:
            st.markdown(f"**Primeras 15 filas de la hoja `{selected_sheet}`:**")
            st.dataframe(df_excel.head(15), use_container_width=True)
