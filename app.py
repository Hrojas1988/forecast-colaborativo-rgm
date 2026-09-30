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

# Versión en base64 del placeholder: la grilla de precios se dibuja como HTML
# crudo, y un <img src="..."> ahí NO puede apuntar a un archivo del servidor
# (el navegador de quien ve la app no tiene acceso a ese disco). Por eso el
# placeholder se embebe directo en el HTML como data URI.
import base64 as _base64
with open(PLACEHOLDER_IMG, "rb") as _f:
    PLACEHOLDER_IMG_DATA_URI = "data:image/png;base64," + _base64.b64encode(_f.read()).decode("ascii")

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

# Cadena de valor: costo y márgenes por etapa para armar el precio "de abajo hacia arriba"
CADENA_COLUMNS = [
    "sku", "costo_produccion", "margen_fabricante_pct",
    "margen_distribuidor_pct", "margen_detallista_pct",
]
CADENA_NUMERIC_COLS = [
    "costo_produccion", "margen_fabricante_pct",
    "margen_distribuidor_pct", "margen_detallista_pct",
]
CADENA_DATA_PATH = os.path.join(os.path.dirname(DATA_PATH), "cadena_valor.csv")

# Levantamiento de precios en campo (fuerza de ventas)
LEVANT_COLUMNS = ["fecha", "vendedor", "sku", "nombre", "canal", "cliente", "precio_gondola"]
LEVANT_NUMERIC_COLS = ["precio_gondola"]
LEVANT_DATA_PATH = os.path.join(os.path.dirname(DATA_PATH), "levantamientos.csv")

st.set_page_config(
    page_title="Grilla de Precios / PricePack",
    page_icon="📊",
    layout="wide",
)

# ----------------------------------------------------------------------------
# Candado de acceso (contraseña compartida)
# ----------------------------------------------------------------------------
def check_password() -> str:
    """Pide una contraseña y devuelve el rol asociado: 'analista' o 'ventas'.
    - app_password           -> acceso completo (analista)
    - app_password_ventas    -> acceso limitado, solo a 'Levantar precios'
    Si no hay ninguna contraseña configurada (uso local de prueba), no pide nada
    y da acceso completo."""
    pwd_analista = st.secrets.get("app_password", None)
    pwd_ventas = st.secrets.get("app_password_ventas", None)

    if not pwd_analista and not pwd_ventas:
        return "analista"

    if st.session_state.get("_role"):
        return st.session_state["_role"]

    st.title("📊 Grilla de Precios / PricePack")
    pwd = st.text_input("Contraseña de acceso", type="password")
    if st.button("Entrar"):
        if pwd_analista and pwd == pwd_analista:
            st.session_state["_role"] = "analista"
            st.rerun()
        elif pwd_ventas and pwd == pwd_ventas:
            st.session_state["_role"] = "ventas"
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()
    return ""


ROLE = check_password()

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


def _get_worksheet_cadena():
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
        ws = sh.worksheet("cadena_valor")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="cadena_valor", rows=1000, cols=len(CADENA_COLUMNS))
        ws.append_row(CADENA_COLUMNS)
    return ws


def _normalize_cadena(df: pd.DataFrame) -> pd.DataFrame:
    for c in CADENA_COLUMNS:
        if c not in df.columns:
            df[c] = "" if c == "sku" else np.nan
    for c in CADENA_NUMERIC_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["sku"] = df["sku"].fillna("").astype(str)
    df.loc[df["sku"].str.lower().isin(["nan", "none"]), "sku"] = ""
    return df[CADENA_COLUMNS]


def load_cadena_data() -> pd.DataFrame:
    if USE_SHEETS:
        ws = _get_worksheet_cadena()
        records = ws.get_all_records()
        df = pd.DataFrame(records) if records else pd.DataFrame(columns=CADENA_COLUMNS)
    elif os.path.exists(CADENA_DATA_PATH):
        df = pd.read_csv(CADENA_DATA_PATH)
    else:
        df = pd.DataFrame(columns=CADENA_COLUMNS)
    return _normalize_cadena(df)


def save_cadena_data(df: pd.DataFrame) -> None:
    df = _normalize_cadena(df.copy())
    if USE_SHEETS:
        ws = _get_worksheet_cadena()
        ws.clear()
        ws.append_row(CADENA_COLUMNS)
        values = df.fillna("").astype(str).values.tolist()
        if values:
            ws.append_rows(values)
    else:
        df.to_csv(CADENA_DATA_PATH, index=False)


def calc_cadena(row: pd.Series) -> dict:
    """Arma el precio 'de abajo hacia arriba': costo -> precio fábrica ->
    precio distribuidor -> precio góndola, aplicando cada margen como % sobre
    el precio de venta de esa etapa (margen bruto, no markup sobre costo)."""
    costo = row.get("costo_produccion")
    m_fab = row.get("margen_fabricante_pct")
    m_dist = row.get("margen_distribuidor_pct")
    m_det = row.get("margen_detallista_pct")

    out = {"costo_produccion": costo, "precio_fabrica": np.nan,
           "precio_distribuidor_calc": np.nan, "precio_gondola_calc": np.nan}

    if pd.isna(costo):
        return out

    precio_fabrica = costo
    if pd.notna(m_fab) and m_fab < 100:
        precio_fabrica = costo / (1 - m_fab / 100)
    out["precio_fabrica"] = precio_fabrica

    precio_dist = precio_fabrica
    if pd.notna(m_dist) and m_dist < 100:
        precio_dist = precio_fabrica / (1 - m_dist / 100)
    out["precio_distribuidor_calc"] = precio_dist

    precio_gondola = precio_dist
    if pd.notna(m_det) and m_det < 100:
        precio_gondola = precio_dist / (1 - m_det / 100)
    out["precio_gondola_calc"] = precio_gondola

    return out


def _get_worksheet_levant():
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
        ws = sh.worksheet("levantamientos")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="levantamientos", rows=2000, cols=len(LEVANT_COLUMNS))
        ws.append_row(LEVANT_COLUMNS)
    return ws


def _normalize_levant(df: pd.DataFrame) -> pd.DataFrame:
    for c in LEVANT_COLUMNS:
        if c not in df.columns:
            df[c] = np.nan if c in LEVANT_NUMERIC_COLS else ""
    for c in LEVANT_NUMERIC_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in [c for c in LEVANT_COLUMNS if c not in LEVANT_NUMERIC_COLS]:
        df[c] = df[c].fillna("").astype(str)
        df.loc[df[c].str.lower().isin(["nan", "none"]), c] = ""
    return df[LEVANT_COLUMNS]


def load_levant_data() -> pd.DataFrame:
    if USE_SHEETS:
        ws = _get_worksheet_levant()
        records = ws.get_all_records()
        df = pd.DataFrame(records) if records else pd.DataFrame(columns=LEVANT_COLUMNS)
    elif os.path.exists(LEVANT_DATA_PATH):
        df = pd.read_csv(LEVANT_DATA_PATH)
    else:
        df = pd.DataFrame(columns=LEVANT_COLUMNS)
    return _normalize_levant(df)


def append_levant_rows(rows: list[dict]) -> None:
    """Agrega filas nuevas de levantamiento sin borrar las existentes."""
    nuevas = _normalize_levant(pd.DataFrame(rows))
    if USE_SHEETS:
        ws = _get_worksheet_levant()
        values = nuevas.fillna("").astype(str).values.tolist()
        if values:
            ws.append_rows(values)
    else:
        existentes = load_levant_data()
        combinado = pd.concat([existentes, nuevas], ignore_index=True)
        combinado.to_csv(LEVANT_DATA_PATH, index=False)


def compute_moda(precios: pd.Series):
    """Devuelve (moda, n_muestras). Si no hay precios repetidos exactos,
    intenta con precios redondeados al entero más cercano; si tampoco hay
    repetidos, cae de vuelta a la mediana como aproximación."""
    s = pd.to_numeric(precios, errors="coerce").dropna()
    n = len(s)
    if n == 0:
        return np.nan, 0, False
    counts = s.value_counts()
    if counts.max() > 1:
        moda = counts[counts == counts.max()].index.min()
        return float(moda), n, True
    redond = s.round(0)
    counts_r = redond.value_counts()
    if counts_r.max() > 1:
        moda = counts_r[counts_r == counts_r.max()].index.min()
        return float(moda), n, True
    return float(s.median()), n, False


def render_levantamiento_form(productos_df: pd.DataFrame, key_prefix: str = "") -> None:
    """Formulario simple para que un vendedor capture precios de góndola de
    varios productos a la vez, para un canal/cliente/fecha dados."""
    import datetime

    if productos_df.empty:
        st.info("Todavía no hay productos cargados para levantar precios.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        vendedor = st.text_input("Tu nombre", key=f"{key_prefix}_vendedor")
    with c2:
        canales_disp = sorted([c for c in productos_df["canal"].dropna().unique() if c])
        canal_sel = st.selectbox("Canal", canales_disp or ["General"], key=f"{key_prefix}_canal")
    with c3:
        cliente = st.text_input("Cliente / punto de venta", key=f"{key_prefix}_cliente",
                                 placeholder="Ej. Walmart San Pedro")

    fecha = st.date_input("Fecha de la visita", value=datetime.date.today(), key=f"{key_prefix}_fecha")

    cats_disp = ["Todas"] + sorted([c for c in productos_df["categoria"].dropna().unique() if c])
    cat_sel = st.selectbox("Filtrar por categoría (opcional)", cats_disp, key=f"{key_prefix}_cat")

    base = productos_df.copy()
    if cat_sel != "Todas":
        base = base[base["categoria"] == cat_sel]

    tabla = base[["sku", "nombre", "empresa", "tipo"]].copy().sort_values(["tipo", "empresa", "nombre"])
    tabla["precio_gondola"] = np.nan

    st.caption("Escribe el precio que ves en el anaquel. Deja en blanco los que no viste.")
    capturado = st.data_editor(
        tabla,
        use_container_width=True,
        height=420,
        hide_index=True,
        disabled=["sku", "nombre", "empresa", "tipo"],
        column_config={
            "precio_gondola": st.column_config.NumberColumn("Precio góndola", format="%.2f"),
        },
        key=f"{key_prefix}_editor_levant",
    )

    if st.button("✅ Enviar precios", type="primary", key=f"{key_prefix}_btn_enviar"):
        llenos = capturado.dropna(subset=["precio_gondola"])
        if not vendedor.strip():
            st.error("Escribe tu nombre antes de enviar.")
        elif not cliente.strip():
            st.error("Escribe el cliente / punto de venta antes de enviar.")
        elif llenos.empty:
            st.warning("No has puesto ningún precio todavía.")
        else:
            filas = [
                {
                    "fecha": fecha.isoformat(),
                    "vendedor": vendedor.strip(),
                    "sku": r["sku"],
                    "nombre": r["nombre"],
                    "canal": canal_sel,
                    "cliente": cliente.strip(),
                    "precio_gondola": r["precio_gondola"],
                }
                for _, r in llenos.iterrows()
            ]
            append_levant_rows(filas)
            st.session_state.df_levant = load_levant_data()
            st.success(f"¡Listo! Se guardaron {len(filas)} precios para {cliente.strip()} ({canal_sel}).")
            st.rerun()


import re as _re

_DRIVE_ID_PATTERNS = [
    r"/file/d/([a-zA-Z0-9_-]{10,})",   # .../file/d/FILE_ID/view...
    r"[?&]id=([a-zA-Z0-9_-]{10,})",    # ...?id=FILE_ID / uc?export=view&id=FILE_ID
    r"/d/([a-zA-Z0-9_-]{10,})",        # .../d/FILE_ID/...
]


def resolve_image(url: str) -> str:
    if not isinstance(url, str) or not url.strip():
        return PLACEHOLDER_IMG_DATA_URI
    u = url.strip()
    if not u.lower().startswith(("http://", "https://")):
        return PLACEHOLDER_IMG_DATA_URI

    # Cualquier link de Google Drive (visor, compartir, o ya convertido) se
    # normaliza al formato de miniatura, que es el que Google permite
    # embeber de forma confiable en otras páginas.
    if "drive.google.com" in u.lower():
        for pat in _DRIVE_ID_PATTERNS:
            m = _re.search(pat, u)
            if m:
                return f"https://drive.google.com/thumbnail?id={m.group(1)}&sz=w400"
        return u  # no se pudo extraer el id, se intenta tal cual

    return u


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
                share = _fmt_num(r.get("market_share_pct"), 1)
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
                    f"<div style='color:#0369a1;'>Share: <b>{share}%</b></div>"
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
if "df_cadena" not in st.session_state:
    st.session_state.df_cadena = load_cadena_data()
if "df_levant" not in st.session_state:
    st.session_state.df_levant = load_levant_data()

df = st.session_state.df
df_cadena = st.session_state.df_cadena
df_levant = st.session_state.df_levant

# ----------------------------------------------------------------------------
# Vista de VENTAS: solo el formulario de levantamiento, sin análisis ni filtros
# ----------------------------------------------------------------------------
if ROLE == "ventas":
    st.title("📝 Levantar precios de góndola")
    st.caption("Ingresa el precio que ves en el anaquel para cada producto. No hace falta llenarlos todos.")
    render_levantamiento_form(df, key_prefix="ventas")
    st.stop()

# ----------------------------------------------------------------------------
# Sidebar: filtros globales (solo analista)
# ----------------------------------------------------------------------------
st.sidebar.title("📊 Grilla de Precios")
st.sidebar.caption("Posicionamiento competitivo por SKU")
st.sidebar.caption(f"Fuente de datos: {'Google Sheets' if USE_SHEETS else 'CSV local (modo prueba)'}")

if st.sidebar.button("🔄 Recargar datos"):
    st.session_state.df = load_data()
    st.session_state.df_cadena = load_cadena_data()
    st.session_state.df_levant = load_levant_data()
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
tab_mapa, tab_burbujas, tab_tabla, tab_cadena, tab_levant, tab_moda, tab_mant = st.tabs(
    [
        "🗺️ Grilla de precios",
        "🔵 Mapa de burbujas",
        "📋 Tabla comparativa",
        "🧩 Cadena de valor",
        "📝 Levantar precios",
        "📈 Moda de precios",
        "🛠️ Mantenimiento",
    ]
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
# TAB 2: Mapa de burbujas (PPK vs. Precio de góndola)
# ----------------------------------------------------------------------------
with tab_burbujas:
    st.subheader("Mapa de burbujas — PPK vs. Precio de góndola")
    st.caption(
        "Cada burbuja es un SKU. Eje X = precio por unidad de medida (PPK). "
        "Eje Y = precio de góndola. Tamaño de burbuja = market share. "
        "Las líneas punteadas marcan las medianas de la vista actual."
    )

    if df_f.empty or df_f[["ppk", "precio_gondola"]].dropna().empty:
        st.info("No hay datos suficientes (PPK y precio de góndola) para graficar con los filtros actuales.")
    else:
        plot_df = df_f.dropna(subset=["ppk", "precio_gondola"]).copy()
        plot_df["market_share_pct"] = plot_df["market_share_pct"].fillna(1)

        x_med = plot_df["ppk"].median()
        y_med = plot_df["precio_gondola"].median()

        color_map = {"Propio": "#1f77b4", "Competencia": "#d62728"}
        fig = go.Figure()

        for tipo, sub in plot_df.groupby("tipo"):
            fig.add_trace(
                go.Scatter(
                    x=sub["ppk"],
                    y=sub["precio_gondola"],
                    mode="markers+text",
                    name=tipo,
                    text=sub["nombre"].str.slice(0, 18),
                    textposition="top center",
                    textfont=dict(size=9),
                    marker=dict(
                        size=sub["market_share_pct"].clip(lower=3) * 2.2,
                        color=color_map.get(tipo, "#888888"),
                        line=dict(width=1, color="white"),
                        opacity=0.85,
                    ),
                    customdata=sub[
                        ["sku", "empresa", "canal", "categoria", "market_share_pct",
                         "precio_distribuidor", "respeto_precio_pct", "margen_pct"]
                    ],
                    hovertemplate=(
                        "<b>%{text}</b><br>"
                        "SKU: %{customdata[0]}<br>"
                        "Empresa: %{customdata[1]}<br>"
                        "Canal: %{customdata[2]} | Categoría: %{customdata[3]}<br>"
                        "PPK: %{x:.2f} | Precio góndola: %{y:.2f}<br>"
                        "Market share: %{customdata[4]:.1f}%<br>"
                        "Precio distribuidor: %{customdata[5]:.2f}<br>"
                        "% Respeto de precio: %{customdata[6]}<br>"
                        "Margen: %{customdata[7]:.1f}%"
                        "<extra></extra>"
                    ),
                )
            )

        fig.add_vline(x=x_med, line_dash="dot", line_color="gray")
        fig.add_hline(y=y_med, line_dash="dot", line_color="gray")

        fig.update_xaxes(showgrid=True, gridwidth=1, gridcolor="rgba(200,200,200,0.4)", title="PPK (precio por unidad de medida)")
        fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor="rgba(200,200,200,0.4)", title="Precio de góndola")

        fig.update_layout(
            height=620,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            margin=dict(l=10, r=10, t=30, b=10),
            plot_bgcolor="white",
        )

        st.plotly_chart(fig, use_container_width=True)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("PPK mediana (propios)", f"{plot_df[plot_df.tipo=='Propio']['ppk'].median():.2f}" if not plot_df[plot_df.tipo=='Propio'].empty else "—")
        c2.metric("PPK mediana (competencia)", f"{plot_df[plot_df.tipo=='Competencia']['ppk'].median():.2f}" if not plot_df[plot_df.tipo=='Competencia'].empty else "—")
        own_share = plot_df[plot_df.tipo == "Propio"]["market_share_pct"].sum()
        comp_share = plot_df[plot_df.tipo == "Competencia"]["market_share_pct"].sum()
        c3.metric("Share propio (vista)", f"{own_share:.1f}%")
        c4.metric("Share competencia (vista)", f"{comp_share:.1f}%")

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
# TAB 4: Cadena de valor — costo y márgenes vs. precio final del competidor
# ----------------------------------------------------------------------------
with tab_cadena:
    st.subheader("Cadena de valor — de costo a precio de góndola")
    st.caption(
        "Arma el precio 'de abajo hacia arriba': costo de producción → margen del fabricante "
        "→ margen del distribuidor → margen del detallista → precio de góndola. "
        "Compáralo contra tu precio real y contra el precio final de un competidor."
    )

    propios = df[df["tipo"] == "Propio"].copy()
    competencia = df[df["tipo"] == "Competencia"].copy()

    if propios.empty:
        st.info("No hay artículos propios (tipo = 'Propio') para modelar la cadena de valor.")
    else:
        colsel1, colsel2 = st.columns(2)
        with colsel1:
            sku_propio = st.selectbox(
                "Tu SKU",
                propios["sku"],
                format_func=lambda s: f"{s} — {propios.loc[propios.sku==s, 'nombre'].values[0]}",
            )
        with colsel2:
            comp_opts = ["(ninguno)"] + competencia["sku"].tolist()
            sku_comp = st.selectbox(
                "Comparar contra el precio final de",
                comp_opts,
                format_func=lambda s: "(ninguno)" if s == "(ninguno)" else
                    f"{s} — {competencia.loc[competencia.sku==s, 'nombre'].values[0]}",
            )

        fila_cadena = df_cadena[df_cadena["sku"] == sku_propio]
        if fila_cadena.empty:
            st.warning(
                "Este SKU todavía no tiene costo y márgenes cargados. "
                "Baja a la sección **'Editar costos y márgenes'** para agregarlos."
            )
            costo_row = pd.Series({c: np.nan for c in CADENA_COLUMNS})
        else:
            costo_row = fila_cadena.iloc[0]

        calc = calc_cadena(costo_row)
        producto_row = propios[propios.sku == sku_propio].iloc[0]
        precio_real = producto_row.get("precio_gondola")
        precio_competidor = (
            competencia.loc[competencia.sku == sku_comp, "precio_gondola"].values[0]
            if sku_comp != "(ninguno)" and not competencia.loc[competencia.sku == sku_comp].empty
            else np.nan
        )

        if pd.notna(calc["precio_gondola_calc"]):
            # ---- Waterfall: costo -> margen fabricante -> margen distribuidor -> margen detallista
            costo = calc["costo_produccion"]
            add_fab = calc["precio_fabrica"] - costo
            add_dist = calc["precio_distribuidor_calc"] - calc["precio_fabrica"]
            add_det = calc["precio_gondola_calc"] - calc["precio_distribuidor_calc"]

            fig_w = go.Figure(go.Waterfall(
                orientation="v",
                measure=["absolute", "relative", "relative", "relative", "total"],
                x=["Costo producción", "+ Margen fabricante", "+ Margen distribuidor",
                   "+ Margen detallista", "Precio góndola (calc.)"],
                y=[costo, add_fab, add_dist, add_det, 0],
                text=[f"{v:,.0f}" for v in [costo, add_fab, add_dist, add_det, calc["precio_gondola_calc"]]],
                textposition="outside",
                connector=dict(line=dict(color="rgba(150,150,150,0.5)")),
                increasing=dict(marker=dict(color="#2563eb")),
                totals=dict(marker=dict(color="#111827")),
            ))

            shapes = []
            annotations = []
            if pd.notna(precio_real):
                shapes.append(dict(
                    type="line", xref="paper", x0=0, x1=1,
                    yref="y", y0=precio_real, y1=precio_real,
                    line=dict(color="#16a34a", width=2, dash="dash"),
                ))
                annotations.append(dict(
                    xref="paper", x=1.0, yref="y", y=precio_real,
                    text=f"Tu precio real: {precio_real:,.0f}",
                    showarrow=False, font=dict(color="#16a34a", size=11),
                    xanchor="right", yanchor="bottom",
                ))
            if pd.notna(precio_competidor):
                shapes.append(dict(
                    type="line", xref="paper", x0=0, x1=1,
                    yref="y", y0=precio_competidor, y1=precio_competidor,
                    line=dict(color="#dc2626", width=2, dash="dash"),
                ))
                annotations.append(dict(
                    xref="paper", x=1.0, yref="y", y=precio_competidor,
                    text=f"Precio competidor: {precio_competidor:,.0f}",
                    showarrow=False, font=dict(color="#dc2626", size=11),
                    xanchor="right", yanchor="top",
                ))

            fig_w.update_layout(
                height=480,
                shapes=shapes,
                annotations=annotations,
                margin=dict(l=10, r=10, t=30, b=10),
                plot_bgcolor="white",
                showlegend=False,
            )
            fig_w.update_yaxes(showgrid=True, gridcolor="rgba(200,200,200,0.4)", title="Precio")
            st.plotly_chart(fig_w, use_container_width=True)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Precio calculado (cadena)", f"{calc['precio_gondola_calc']:,.0f}")
            c2.metric("Tu precio real", f"{precio_real:,.0f}" if pd.notna(precio_real) else "—")
            if pd.notna(precio_competidor):
                brecha = (calc["precio_gondola_calc"] / precio_competidor - 1) * 100
                c3.metric("Precio del competidor", f"{precio_competidor:,.0f}")
                c4.metric("Brecha vs. competidor", f"{brecha:+.1f}%")
                if brecha > 2:
                    st.warning(
                        f"Con este costo y estos márgenes, tu precio de góndola quedaría "
                        f"**{brecha:.1f}% por encima** del competidor seleccionado."
                    )
                elif brecha < -2:
                    st.success(
                        f"Con este costo y estos márgenes, tu precio de góndola quedaría "
                        f"**{abs(brecha):.1f}% por debajo** del competidor seleccionado — hay espacio."
                    )
                else:
                    st.info("Tu precio calculado queda prácticamente en paridad con el competidor.")
            else:
                c3.metric("Precio del competidor", "—")
                c4.metric("Brecha vs. competidor", "—")
        else:
            st.info(
                "Agrega al menos el costo de producción para este SKU en "
                "'Editar costos y márgenes' para ver el puente de precio."
            )

    st.markdown("---")
    with st.expander("🛠️ Editar costos y márgenes por SKU propio", expanded=fila_cadena.empty if propios is not None and not propios.empty else True):
        st.caption(
            "Un renglón por SKU propio. Los márgenes son **margen bruto sobre el precio de venta "
            "de esa etapa** (ej. margen_fabricante_pct = 30 significa que el margen del fabricante "
            "es 30% del precio de fábrica, no un markup de 30% sobre el costo)."
        )
        base_cadena = df_cadena.copy()
        # Asegura que todo SKU propio tenga al menos una fila para editar
        existentes = set(base_cadena["sku"])
        faltantes = [s for s in propios["sku"] if s not in existentes] if not propios.empty else []
        if faltantes:
            extra = pd.DataFrame({"sku": faltantes})
            base_cadena = pd.concat([base_cadena, extra], ignore_index=True)
            base_cadena = _normalize_cadena(base_cadena)

        edited_cadena = st.data_editor(
            base_cadena,
            num_rows="dynamic",
            use_container_width=True,
            height=320,
            column_config={
                "sku": st.column_config.TextColumn(),
                "costo_produccion": st.column_config.NumberColumn(format="%.2f"),
                "margen_fabricante_pct": st.column_config.NumberColumn(format="%.1f"),
                "margen_distribuidor_pct": st.column_config.NumberColumn(format="%.1f"),
                "margen_detallista_pct": st.column_config.NumberColumn(format="%.1f"),
            },
            key="editor_cadena",
        )
        if st.button("💾 Guardar costos y márgenes", type="primary"):
            save_cadena_data(edited_cadena)
            st.session_state.df_cadena = load_cadena_data()
            st.success("Guardado. Vuelve a seleccionar el SKU arriba para ver el puente actualizado.")
            st.rerun()

# ----------------------------------------------------------------------------
# TAB 5: Levantar precios (misma vista que usa la fuerza de ventas)
# ----------------------------------------------------------------------------
with tab_levant:
    st.subheader("Levantar precios de góndola")
    st.caption(
        "La misma pantalla que ve tu equipo comercial con la contraseña de ventas. "
        "Úsala para probar o para capturar precios tú mismo."
    )
    render_levantamiento_form(df, key_prefix="analista")

# ----------------------------------------------------------------------------
# TAB 6: Moda de precios — agrega los levantamientos y permite aplicarlos
# ----------------------------------------------------------------------------
with tab_moda:
    st.subheader("Moda de precios levantados en campo")
    st.caption(
        "Junta todos los precios que ha capturado el equipo por SKU, canal y cliente, "
        "y calcula la moda (el precio más repetido). Tú decides cuáles aplicar a la Grilla."
    )

    if df_levant.empty:
        st.info("Todavía no hay levantamientos capturados.")
    else:
        colf1, colf2, colf3 = st.columns(3)
        with colf1:
            agrupar_por = st.multiselect(
                "Agrupar por", ["canal", "cliente"], default=["canal"],
                help="Elige si quieres la moda por canal, por cliente puntual, o por ambos.",
            )
        with colf2:
            canales_lv = sorted([c for c in df_levant["canal"].dropna().unique() if c])
            f_canal_lv = st.multiselect("Filtrar canal", canales_lv, default=canales_lv)
        with colf3:
            clientes_lv = sorted([c for c in df_levant["cliente"].dropna().unique() if c])
            f_cliente_lv = st.multiselect("Filtrar cliente", clientes_lv, default=clientes_lv)

        lv = df_levant[
            df_levant["canal"].isin(f_canal_lv) & df_levant["cliente"].isin(f_cliente_lv)
        ].copy()

        if lv.empty or not agrupar_por:
            st.info("Ajusta los filtros o elige al menos un criterio de agrupación.")
        else:
            group_cols = ["sku", "nombre"] + agrupar_por
            resumen_rows = []
            for keys, sub in lv.groupby(group_cols):
                moda, n, exacta = compute_moda(sub["precio_gondola"])
                fila = dict(zip(group_cols, keys if isinstance(keys, tuple) else (keys,)))
                fila["moda_precio_gondola"] = round(moda, 2) if pd.notna(moda) else np.nan
                fila["n_muestras"] = n
                fila["min"] = sub["precio_gondola"].min()
                fila["max"] = sub["precio_gondola"].max()
                fila["moda_exacta"] = "Sí" if exacta else "Aprox. (mediana)"
                resumen_rows.append(fila)
            resumen = pd.DataFrame(resumen_rows).sort_values(group_cols)

            precio_actual = df.set_index("sku")["precio_gondola"]
            resumen["precio_actual_en_grilla"] = resumen["sku"].map(precio_actual)
            resumen["dif_vs_grilla_pct"] = (
                (resumen["moda_precio_gondola"] / resumen["precio_actual_en_grilla"] - 1) * 100
            ).round(1)

            resumen.insert(0, "aplicar", False)
            editado = st.data_editor(
                resumen,
                use_container_width=True,
                height=400,
                hide_index=True,
                disabled=[c for c in resumen.columns if c != "aplicar"],
                column_config={
                    "aplicar": st.column_config.CheckboxColumn("Aplicar a la Grilla"),
                    "moda_precio_gondola": st.column_config.NumberColumn(format="%.2f"),
                    "precio_actual_en_grilla": st.column_config.NumberColumn(format="%.2f"),
                    "dif_vs_grilla_pct": st.column_config.NumberColumn("Dif. vs. grilla %", format="%.1f"),
                },
                key="editor_moda",
            )
            st.caption(
                "'Moda exacta' = hubo al menos dos vendedores que reportaron el mismo precio. "
                "'Aprox. (mediana)' = todos los precios capturados fueron distintos, así que se usa "
                "la mediana como mejor estimación."
            )

            if st.button("📥 Aplicar seleccionados a la Grilla de Precios", type="primary"):
                a_aplicar = editado[editado["aplicar"] & editado["moda_precio_gondola"].notna()]
                if a_aplicar.empty:
                    st.warning("No marcaste ninguna fila para aplicar.")
                else:
                    df_actualizado = df.copy()
                    for _, r in a_aplicar.iterrows():
                        df_actualizado.loc[
                            df_actualizado["sku"] == r["sku"], "precio_gondola"
                        ] = r["moda_precio_gondola"]
                    save_data(df_actualizado)
                    st.session_state.df = load_data()
                    st.success(
                        f"Se actualizó precio_gondola de {len(a_aplicar)} SKU en la Grilla. "
                        "Revisa el PPK en Mantenimiento si depende del precio."
                    )
                    st.rerun()

    st.markdown("---")
    with st.expander("Ver todos los levantamientos crudos"):
        st.dataframe(df_levant.sort_values("fecha", ascending=False), use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Descargar levantamientos en CSV",
            data=df_levant.to_csv(index=False).encode("utf-8"),
            file_name="levantamientos.csv",
            mime="text/csv",
        )

# ----------------------------------------------------------------------------
# TAB 7: Mantenimiento (CRUD)
# ----------------------------------------------------------------------------
with tab_mant:
    st.subheader("Mantenimiento de artículos")
    st.caption(
        "Edita, agrega o elimina filas directamente en la tabla. "
        "Los cambios se guardan al presionar 'Guardar cambios' y quedan disponibles "
        "para todo el equipo de inmediato."
    )
    st.caption(
        "Columna **imagen_url**: pega el link de la foto del producto. Si es de Google Drive, "
        "puedes pegar el link tal cual te lo da el botón 'Compartir' (con que esté como "
        "'Cualquiera con el enlace', la app lo convierte sola al formato correcto). "
        "También funciona un link directo de SharePoint/intranet público."
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
