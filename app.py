"""
Grilla de Precios / PricePack Architecture
Herramienta visual para posicionamiento competitivo de precio.

Ejecutar con:  streamlit run app.py

Backend de datos:
- Si existen los secrets "gcp_service_account" y "sheet_id" (Streamlit Cloud
  -> Settings -> Secrets), la app lee y escribe contra ese Google Sheet.
- Si no existen (por ejemplo corriendo local sin configurar nada), usa
  data/productos.csv como respaldo, para poder probar sin credenciales.
"""

import os
import math
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st
from PIL import Image

# ----------------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(APP_DIR, "data", "productos.csv")
IMAGES_DIR = os.path.join(APP_DIR, "images")
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(os.path.dirname(DATA_PATH), exist_ok=True)

PLACEHOLDER_IMG = os.path.join(IMAGES_DIR, "_placeholder.png")
if not os.path.exists(PLACEHOLDER_IMG):
    img = Image.new("RGB", (200, 200), color=(230, 230, 230))
    img.save(PLACEHOLDER_IMG)

COLUMNS = [
    "sku", "nombre", "tipo", "empresa", "categoria", "canal",
    "presentacion", "unidad_medida", "imagen_url",
    "market_share_pct", "precio_distribuidor", "precio_gondola",
    "ppk", "respeto_precio_pct", "margen_pct",
]
NUMERIC_COLS = [
    "market_share_pct", "precio_distribuidor", "precio_gondola",
    "ppk", "respeto_precio_pct", "margen_pct",
]

st.set_page_config(
    page_title="Grilla de Precios / PricePack",
    page_icon="📊",
    layout="wide",
)

# ----------------------------------------------------------------------------
# Candado de acceso (contraseña compartida)
# ----------------------------------------------------------------------------
def check_password() -> bool:
    """Pide una contraseña si se configuró st.secrets['app_password'].
    Si no hay contraseña configurada (uso local de prueba), no pide nada."""
    required = st.secrets.get("app_password", None)
    if not required:
        return True

    if st.session_state.get("_authed", False):
        return True

    st.title("📊 Grilla de Precios / PricePack")
    pwd = st.text_input("Contraseña de acceso", type="password")
    if st.button("Entrar"):
        if pwd == required:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()
    return False


check_password()

# ----------------------------------------------------------------------------
# Backend de datos: Google Sheets si hay credenciales, si no CSV local
# ----------------------------------------------------------------------------
USE_SHEETS = "gcp_service_account" in st.secrets and "sheet_id" in st.secrets


def _get_worksheet():
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["sheet_id"])
    try:
        ws = sh.worksheet("productos")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="productos", rows=1000, cols=len(COLUMNS))
        ws.append_row(COLUMNS)
    return ws


TEXT_COLS = [c for c in COLUMNS if c not in NUMERIC_COLS]


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    for c in NUMERIC_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in TEXT_COLS:
        df[c] = df[c].fillna("").astype(str)
        df.loc[df[c].str.lower().isin(["nan", "none"]), c] = ""
    return df[COLUMNS]


def load_data() -> pd.DataFrame:
    if USE_SHEETS:
        ws = _get_worksheet()
        records = ws.get_all_records()
        df = pd.DataFrame(records) if records else pd.DataFrame(columns=COLUMNS)
    elif os.path.exists(DATA_PATH):
        df = pd.read_csv(DATA_PATH)
        if "imagen" in df.columns and "imagen_url" not in df.columns:
            df = df.rename(columns={"imagen": "imagen_url"})
    else:
        df = pd.DataFrame(columns=COLUMNS)
    return _normalize(df)


def save_data(df: pd.DataFrame) -> None:
    df = _normalize(df.copy())
    if USE_SHEETS:
        ws = _get_worksheet()
        ws.clear()
        ws.append_row(COLUMNS)
        values = df.fillna("").astype(str).values.tolist()
        if values:
            ws.append_rows(values)
    else:
        df.to_csv(DATA_PATH, index=False)


def resolve_image(url: str) -> str:
    if isinstance(url, str) and url.strip().lower().startswith(("http://", "https://")):
        return url.strip()
    return PLACEHOLDER_IMG


# ----------------------------------------------------------------------------
# Grilla de precios estilo "price ladder" (filas = marca/empresa, columnas = precio)
# ----------------------------------------------------------------------------
import html as _html


def _price_step(vals: pd.Series) -> float:
    """Elige un ancho de columna 'redondo' (1, 2, 2.5, 5, 10 x potencia de 10)
    apuntando a unas 7-9 columnas."""
    vals = vals.dropna()
    if vals.empty:
        return 10.0
    rng = float(vals.max() - vals.min())
    if rng <= 0:
        rng = float(vals.max()) or 10.0
    raw_step = rng / 8.0
    magnitude = 10 ** math.floor(math.log10(raw_step)) if raw_step > 0 else 1
    for mult in (1, 2, 2.5, 5, 10):
        step = magnitude * mult
        if step >= raw_step:
            return step
    return magnitude * 10


def _fmt_num(v, decimals=0):
    if pd.isna(v):
        return "—"
    return f"{v:,.{decimals}f}"


def build_price_grid_html(cat_df: pd.DataFrame, price_col: str = "precio_gondola") -> str:
    """Arma una grilla HTML: filas = empresa/marca, columnas = rangos de precio.
    Cada producto se ubica en la columna que corresponde a su precio."""
    vals = cat_df[price_col].dropna()
    if vals.empty:
        return "<p style='color:#888'>No hay precios para graficar con los filtros actuales.</p>"

    step = _price_step(vals)
    min_v = math.floor(vals.min() / step) * step
    max_v = math.ceil(vals.max() / step) * step
    if max_v <= min_v:
        max_v = min_v + step
    n_cols = int(round((max_v - min_v) / step))
    col_labels = [_fmt_num(min_v + i * step) for i in range(n_cols)]

    def col_index(v):
        idx = int((v - min_v) // step)
        return min(max(idx, 0), n_cols - 1)

    # Orden de filas: propios primero, luego competencia, alfabético dentro de cada grupo
    rows_meta = (
        cat_df[["empresa", "tipo"]]
        .drop_duplicates()
        .sort_values(["tipo", "empresa"], ascending=[False, True])
    )
    rows_order = rows_meta["empresa"].tolist()

    cells = {emp: {i: [] for i in range(n_cols)} for emp in rows_order}
    for _, r in cat_df.dropna(subset=[price_col]).iterrows():
        idx = col_index(r[price_col])
        cells.setdefault(r["empresa"], {i: [] for i in range(n_cols)})
        cells[r["empresa"]][idx].append(r)

    # ---- construir HTML ----
    col_w = 150
    row_head_w = 150

    head_cells = "".join(
        f"<th style='min-width:{col_w}px;max-width:{col_w}px;background:#1f2937;"
        f"color:#fff;padding:8px 4px;font-size:12px;text-align:center;"
        f"border:1px solid #33415580;'>{_html.escape(lbl)}</th>"
        for lbl in col_labels
    )

    body_rows = ""
    for emp in rows_order:
        tipo = rows_meta.loc[rows_meta["empresa"] == emp, "tipo"].iloc[0]
        is_own = tipo == "Propio"
        row_color = "#e8f1ff" if is_own else "#f6f6f6"
        border_color = "#2563eb" if is_own else "#9ca3af"

        row_head = (
            f"<td style='min-width:{row_head_w}px;max-width:{row_head_w}px;"
            f"background:{row_color};border-left:5px solid {border_color};"
            f"border-top:1px solid #ddd;border-bottom:1px solid #ddd;"
            f"padding:8px;font-weight:700;font-size:13px;vertical-align:middle;'>"
            f"{_html.escape(str(emp))}"
            f"<div style='font-weight:400;font-size:10px;color:#666;'>"
            f"{'Propio' if is_own else 'Competencia'}</div></td>"
        )

        col_cells = ""
        for i in range(n_cols):
            products = cells[emp][i]
            if not products:
                col_cells += (
                    f"<td style='background:{row_color}22;border:1px solid #eee;'></td>"
                )
                continue
            boxes = ""
            for r in products:
                img = resolve_image(r.get("imagen_url"))
                pres = _html.escape(str(r.get("presentacion", "") or ""))
                unidad = _html.escape(str(r.get("unidad_medida", "") or ""))
                precio = _fmt_num(r.get(price_col), 0)
                ppk = _fmt_num(r.get("ppk"), 2)
                nombre = _html.escape(str(r.get("nombre", "") or ""))
                boxes += (
                    f"<div title='{nombre}' style='background:#fff;border:1px solid {border_color};"
                    f"border-radius:6px;padding:4px;margin:3px auto;width:92%;font-size:10.5px;"
                    f"box-shadow:0 1px 2px rgba(0,0,0,.08);'>"
                    f"<img src='{img}' style='width:100%;height:46px;object-fit:contain;"
                    f"border-radius:4px;margin-bottom:3px;'/>"
                    f"<div style='font-weight:600;color:#111;line-height:1.2;'>{pres}{unidad}</div>"
                    f"<div style='color:#111;'>Precio: <b>{precio}</b></div>"
                    f"<div style='color:#b91c1c;'>PPK: <b>{ppk}</b></div>"
                    f"</div>"
                )
            col_cells += (
                f"<td style='background:{row_color}55;border:1px solid #eee;"
                f"vertical-align:top;text-align:center;'>{boxes}</td>"
            )

        body_rows += f"<tr>{row_head}{col_cells}</tr>"

    table_html = f"""
    <div style="overflow-x:auto;border:1px solid #ccc;border-radius:8px;">
      <table style="border-collapse:collapse;width:100%;">
        <thead>
          <tr>
            <th style="min-width:{row_head_w}px;background:#111827;color:#fff;
                       padding:8px;font-size:12px;text-align:left;">Marca / Empresa</th>
            {head_cells}
          </tr>
        </thead>
        <tbody>
          {body_rows}
        </tbody>
      </table>
    </div>
    """
    return table_html


# ----------------------------------------------------------------------------
# Estado
# ----------------------------------------------------------------------------
if "df" not in st.session_state:
    st.session_state.df = load_data()

df = st.session_state.df

# ----------------------------------------------------------------------------
# Sidebar: filtros globales
# ----------------------------------------------------------------------------
st.sidebar.title("📊 Grilla de Precios")
st.sidebar.caption("Posicionamiento competitivo por SKU")
st.sidebar.caption(f"Fuente de datos: {'Google Sheets' if USE_SHEETS else 'CSV local (modo prueba)'}")

if st.sidebar.button("🔄 Recargar datos"):
    st.session_state.df = load_data()
    st.rerun()

canales = sorted([c for c in df["canal"].dropna().unique()])
categorias = sorted([c for c in df["categoria"].dropna().unique()])
empresas = sorted([c for c in df["empresa"].dropna().unique()])

f_canal = st.sidebar.multiselect("Canal", canales, default=canales)
f_categoria = st.sidebar.multiselect("Categoría", categorias, default=categorias)
f_empresa = st.sidebar.multiselect("Empresa", empresas, default=empresas)

df_f = df[
    df["canal"].isin(f_canal)
    & df["categoria"].isin(f_categoria)
    & df["empresa"].isin(f_empresa)
].copy()

st.sidebar.markdown("---")
st.sidebar.caption(f"{len(df_f)} artículos en vista actual de {len(df)} totales")

# ----------------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------------
tab_mapa, tab_tarjetas, tab_tabla, tab_mant = st.tabs(
    ["🗺️ Mapa de posicionamiento", "🗂️ Tarjetas", "📋 Tabla comparativa", "🛠️ Mantenimiento"]
)

# ----------------------------------------------------------------------------
# TAB 1: Mapa / Grilla de posicionamiento
# ----------------------------------------------------------------------------
with tab_mapa:
    st.subheader("Grilla de precios — Price Ladder por categoría")
    st.caption(
        "Cada fila es una marca/empresa y cada columna es un rango de precio. "
        "Los productos se ubican automáticamente en la columna que corresponde a su precio "
        "(igual que una grilla de arquitectura de precios). Azul = productos propios, gris = competencia."
    )

    if df_f.empty:
        st.info("No hay artículos con los filtros actuales.")
    else:
        cats_disponibles = sorted(df_f["categoria"].dropna().unique())
        if not cats_disponibles:
            st.info("No hay categorías con los filtros actuales.")
        else:
            colf1, colf2 = st.columns([2, 1])
            with colf1:
                cat_sel = st.selectbox("Categoría a graficar", cats_disponibles, key="cat_grilla")
            with colf2:
                precio_base = st.radio(
                    "Ordenar columnas por",
                    ["precio_gondola", "ppk"],
                    format_func=lambda v: "Precio de góndola" if v == "precio_gondola" else "PPK",
                    horizontal=True,
                )

            cat_df = df_f[df_f["categoria"] == cat_sel].copy()

            st.markdown(
                f"""
                <div style="background:#111827;color:#fff;padding:10px 16px;border-radius:8px 8px 0 0;
                            display:flex;justify-content:space-between;align-items:center;">
                    <div style="font-size:16px;font-weight:700;letter-spacing:.5px;">
                        {_html.escape(str(cat_sel)).upper()}
                    </div>
                    <div style="font-size:11px;color:#cbd5e1;">
                        Actualizada al {pd.Timestamp.today().strftime('%d/%m/%Y')}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown(build_price_grid_html(cat_df, price_col=precio_base), unsafe_allow_html=True)

            st.markdown("")
            c1, c2, c3, c4 = st.columns(4)
            own = cat_df[cat_df.tipo == "Propio"]
            comp = cat_df[cat_df.tipo == "Competencia"]
            c1.metric("PPK mediana (propios)", f"{own['ppk'].median():.2f}" if not own.empty else "—")
            c2.metric("PPK mediana (competencia)", f"{comp['ppk'].median():.2f}" if not comp.empty else "—")
            c3.metric("Share propio (vista)", f"{own['market_share_pct'].fillna(0).sum():.1f}%")
            c4.metric("Share competencia (vista)", f"{comp['market_share_pct'].fillna(0).sum():.1f}%")

# ----------------------------------------------------------------------------
# TAB 2: Tarjetas (imagen + métricas)
# ----------------------------------------------------------------------------
with tab_tarjetas:
    st.subheader("Tarjetas por artículo")
    if df_f.empty:
        st.info("No hay artículos con los filtros actuales.")
    else:
        cards_per_row = 4
        rows = math.ceil(len(df_f) / cards_per_row)
        df_f_sorted = df_f.sort_values(["categoria", "tipo", "nombre"])
        records = df_f_sorted.to_dict("records")

        for r in range(rows):
            cols = st.columns(cards_per_row)
            chunk = records[r * cards_per_row:(r + 1) * cards_per_row]
            for col, rec in zip(cols, chunk):
                with col:
                    st.image(resolve_image(rec.get("imagen_url")), use_container_width=True)
                    badge = "🟦 Propio" if rec["tipo"] == "Propio" else "🟥 Competencia"
                    st.markdown(f"**{rec['nombre']}**")
                    st.caption(f"{badge} · {rec.get('empresa','')} · {rec.get('canal','')}")
                    st.write(f"Share: **{rec.get('market_share_pct', float('nan')):.1f}%**" if pd.notna(rec.get('market_share_pct')) else "Share: —")
                    st.write(f"P. distribuidor: {rec.get('precio_distribuidor', float('nan')):.2f}" if pd.notna(rec.get('precio_distribuidor')) else "P. distribuidor: —")
                    st.write(f"P. góndola: {rec.get('precio_gondola', float('nan')):.2f}" if pd.notna(rec.get('precio_gondola')) else "P. góndola: —")
                    st.write(f"PPK: {rec.get('ppk', float('nan')):.2f}" if pd.notna(rec.get('ppk')) else "PPK: —")
                    if rec["tipo"] == "Propio":
                        st.write(f"% Respeto precio: {rec.get('respeto_precio_pct', float('nan')):.0f}%" if pd.notna(rec.get('respeto_precio_pct')) else "% Respeto precio: —")
                        st.write(f"Margen: {rec.get('margen_pct', float('nan')):.1f}%" if pd.notna(rec.get('margen_pct')) else "Margen: —")
                    st.markdown("---")

# ----------------------------------------------------------------------------
# TAB 3: Tabla comparativa con índice de precio
# ----------------------------------------------------------------------------
with tab_tabla:
    st.subheader("Tabla comparativa — Índice de precio vs. benchmark de categoría")
    if df_f.empty:
        st.info("No hay artículos con los filtros actuales.")
    else:
        bench = (
            df_f[df_f["tipo"] == "Competencia"]
            .groupby("categoria")["ppk"]
            .mean()
            .rename("ppk_benchmark_categoria")
        )
        tabla = df_f.merge(bench, on="categoria", how="left")
        tabla["indice_precio_vs_categoria"] = (
            tabla["ppk"] / tabla["ppk_benchmark_categoria"] * 100
        ).round(1)

        def posicion(idx):
            if pd.isna(idx):
                return "—"
            if idx > 105:
                return "Premium"
            if idx < 95:
                return "Descuento"
            return "Paridad"

        tabla["posicionamiento"] = tabla["indice_precio_vs_categoria"].apply(posicion)

        show_cols = [
            "sku", "nombre", "tipo", "empresa", "categoria", "canal",
            "market_share_pct", "precio_distribuidor", "precio_gondola",
            "ppk", "ppk_benchmark_categoria", "indice_precio_vs_categoria",
            "posicionamiento", "respeto_precio_pct", "margen_pct",
        ]
        st.dataframe(
            tabla[show_cols].sort_values(["categoria", "indice_precio_vs_categoria"], ascending=[True, False]),
            use_container_width=True,
            hide_index=True,
        )
        st.caption(
            "Índice de precio = PPK del artículo / PPK promedio de la competencia en su categoría × 100. "
            ">105 = Premium · 95–105 = Paridad · <95 = Descuento."
        )

# ----------------------------------------------------------------------------
# TAB 4: Mantenimiento (CRUD)
# ----------------------------------------------------------------------------
with tab_mant:
    st.subheader("Mantenimiento de artículos")
    st.caption(
        "Edita, agrega o elimina filas directamente en la tabla. "
        "Los cambios se guardan al presionar 'Guardar cambios' y quedan disponibles "
        "para todo el equipo de inmediato."
    )
    st.caption(
        "Columna **imagen_url**: pega el enlace directo a la foto del producto "
        "(por ejemplo, un link de Google Drive compartido como 'cualquiera con el enlace' "
        "convertido a formato de imagen, o un link de SharePoint/intranet público)."
    )

    edited = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        height=420,
        column_config={
            "tipo": st.column_config.SelectboxColumn(options=["Propio", "Competencia"]),
            "imagen_url": st.column_config.TextColumn(),
            "market_share_pct": st.column_config.NumberColumn(format="%.1f"),
            "precio_distribuidor": st.column_config.NumberColumn(format="%.2f"),
            "precio_gondola": st.column_config.NumberColumn(format="%.2f"),
            "ppk": st.column_config.NumberColumn(format="%.2f"),
            "respeto_precio_pct": st.column_config.NumberColumn(format="%.0f"),
            "margen_pct": st.column_config.NumberColumn(format="%.1f"),
        },
        key="editor",
    )

    col_a, col_b = st.columns([1, 3])
    with col_a:
        if st.button("💾 Guardar cambios", type="primary"):
            save_data(edited)
            st.session_state.df = load_data()
            st.success("Guardado. Los filtros y gráficos ya reflejan los cambios para todo el equipo.")
            st.rerun()

    st.markdown("---")
    st.download_button(
        "⬇️ Descargar respaldo en CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="productos_respaldo.csv",
        mime="text/csv",
    )
