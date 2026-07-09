"""
Aplicativo de Forecast Colaborativo — Ventas y Unidades por Cliente-SKU
========================================================================
Corre localmente con:
    pip install -r requirements.txt
    streamlit run app_forecast_colaborativo.py

Flujo:
1. Sube tu histórico (Cliente, SKU, Mes, Unidades, Precio + N columnas macro que tú definas).
2. La app ajusta una Regresión Lineal por cada combinación Cliente-SKU:
   Unidades ~ tendencia + estacionalidad (sin/cos) + tus variables macro.
3. Editas en pantalla el escenario macro futuro, el % colaborativo por
   Cliente-SKU-Mes y el precio — todo con inputs interactivos.
4. Descargas el resultado en un Excel con varias hojas.
"""

import io
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from sklearn.linear_model import LinearRegression
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Forecast Colaborativo RGM", layout="wide")

# -------------------------------------------------------------------
# 0. Utilidades
# -------------------------------------------------------------------
REQUIRED_COLS = ["Cliente", "SKU", "Mes", "Unidades", "Precio"]


def datos_demo():
    """Genera un histórico ficticio de 24 meses para probar la app sin subir archivo."""
    np.random.seed(7)
    clientes = ["Supermercado Central", "Tienda Express"]
    skus = ["SKU-101 Bebida", "SKU-205 Snack"]
    meses = pd.date_range("2024-07-01", periods=24, freq="MS")
    filas = []
    for i, cli in enumerate(clientes):
        for j, sku in enumerate(skus):
            base, trend = 900 + i * 300, 4 + j * 2
            for k, mes in enumerate(meses):
                inflacion = 0.006 + 0.002 * np.sin(k / 6) + np.random.normal(0, 0.0008)
                tc = 0.004 + 0.003 * np.sin(k / 9) + np.random.normal(0, 0.001)
                season = 50 * np.sin(2 * np.pi * mes.month / 12)
                unidades = max(base + trend * k + season + np.random.normal(0, 15), 50)
                precio = round(25 + j * 10 + k * 0.15, 2)
                filas.append([cli, sku, mes.strftime("%Y-%m"), round(unidades), precio,
                              round(inflacion, 4), round(tc, 4)])
    return pd.DataFrame(filas, columns=REQUIRED_COLS + ["Inflacion_Mensual", "Var_TipoCambio"])


def meses_futuros(ultimo_mes: str, horizonte: int):
    inicio = pd.to_datetime(ultimo_mes) + pd.DateOffset(months=1)
    return pd.date_range(inicio, periods=horizonte, freq="MS").strftime("%Y-%m").tolist()


# -------------------------------------------------------------------
# 1. Carga de datos
# -------------------------------------------------------------------
st.title("📊 Modelo de Forecast Colaborativo — Cliente x SKU")
st.caption("Regresión lineal + variables macro editables + % colaborativo por iniciativa comercial")

with st.sidebar:
    st.header("1. Datos históricos")
    archivo = st.file_uploader("Sube tu histórico (CSV o Excel)", type=["csv", "xlsx"])
    usar_demo = st.checkbox("Usar datos de ejemplo (ficticios)", value=(archivo is None))

if archivo is not None and not usar_demo:
    hist = pd.read_csv(archivo) if archivo.name.endswith(".csv") else pd.read_excel(archivo)
else:
    hist = datos_demo()

faltantes = [c for c in REQUIRED_COLS if c not in hist.columns]
if faltantes:
    st.error(f"Faltan columnas obligatorias en tu archivo: {faltantes}. "
             f"Se requieren al menos: {REQUIRED_COLS}")
    st.stop()

posibles_macro = [c for c in hist.columns if c not in REQUIRED_COLS]

with st.sidebar:
    st.header("2. Variables macro / adicionales")
    macro_cols = st.multiselect(
        "Selecciona las columnas que usaremos como variables explicativas del modelo",
        options=posibles_macro, default=posibles_macro,
    )
    st.header("3. Horizonte de forecast")
    horizonte = st.slider("Meses a proyectar", 1, 24, 12)

st.subheader("Histórico cargado")
st.dataframe(hist.head(20), use_container_width=True)

# -------------------------------------------------------------------
# 2. Ajuste de Regresión Lineal por Cliente-SKU
# -------------------------------------------------------------------
hist = hist.sort_values(["Cliente", "SKU", "Mes"]).reset_index(drop=True)
combos = hist[["Cliente", "SKU"]].drop_duplicates().values.tolist()
ultimo_mes_global = hist["Mes"].max()
meses_fcst = meses_futuros(ultimo_mes_global, horizonte)


@st.cache_data(show_spinner=False)
def ajustar_modelos(hist_df: pd.DataFrame, macro_cols: list, meses_fcst: list, combos: list):
    resultados_bau, coeficientes, referencias = [], [], {}

    for col in macro_cols:
        referencias[col] = hist_df[col].tail(6).mean()

    for cli, sku in combos:
        serie = hist_df[(hist_df.Cliente == cli) & (hist_df.SKU == sku)].sort_values("Mes").reset_index(drop=True)
        n = len(serie)
        t = np.arange(n)
        sin_t, cos_t = np.sin(2 * np.pi * t / 12), np.cos(2 * np.pi * t / 12)

        X_cols = [t, sin_t, cos_t] + [serie[c].values for c in macro_cols]
        X = np.column_stack(X_cols)
        y = serie["Unidades"].values.astype(float)

        modelo = LinearRegression().fit(X, y)
        coefs = dict(zip(["trend", "sin", "cos"] + macro_cols, modelo.coef_))
        coefs["intercept"] = modelo.intercept_
        coeficientes.append({"Cliente": cli, "SKU": sku, **coefs})

        t_f = np.arange(n, n + len(meses_fcst))
        sin_f, cos_f = np.sin(2 * np.pi * t_f / 12), np.cos(2 * np.pi * t_f / 12)
        X_f_cols = [t_f, sin_f, cos_f] + [np.full(len(meses_fcst), referencias[c]) for c in macro_cols]
        X_f = np.column_stack(X_f_cols)
        bau = np.clip(modelo.predict(X_f), 0, None)

        for k, mes in enumerate(meses_fcst):
            resultados_bau.append({"Cliente": cli, "SKU": sku, "Mes": mes, "BAU_Unidades": round(bau[k])})

    return pd.DataFrame(resultados_bau), pd.DataFrame(coeficientes), referencias


bau_df, coef_df, referencias = ajustar_modelos(hist, macro_cols, meses_fcst, combos)

st.subheader("Modelo ajustado (Regresión Lineal por Cliente-SKU)")
st.dataframe(coef_df.round(4), use_container_width=True)
st.caption("Cada fila es un modelo independiente: Unidades = intercepto + tendencia + estacionalidad + Σ(coef_macro × variable_macro)")

# -------------------------------------------------------------------
# 3. Escenario macro editable
# -------------------------------------------------------------------
st.subheader("4. Escenario macro futuro (editable)")
if macro_cols:
    escenario_default = pd.DataFrame({"Mes": meses_fcst})
    for c in macro_cols:
        escenario_default[c] = referencias[c]
    escenario = st.data_editor(escenario_default, num_rows="fixed", use_container_width=True,
                                key="escenario_macro")
else:
    escenario = pd.DataFrame({"Mes": meses_fcst})
    st.info("No seleccionaste variables macro — el forecast usará solo tendencia y estacionalidad.")

# -------------------------------------------------------------------
# 4. % Colaborativo editable por Cliente-SKU-Mes
# -------------------------------------------------------------------
st.subheader("5. % Colaborativo por iniciativa comercial (editable)")
inic_default = bau_df[["Cliente", "SKU", "Mes"]].copy()
inic_default["%_Colaborativo"] = 0.0
inic_default["Iniciativa"] = ""
iniciativas = st.data_editor(inic_default, num_rows="fixed", use_container_width=True,
                              key="iniciativas", column_config={
                                  "%_Colaborativo": st.column_config.NumberColumn(format="%.1f%%", step=0.5)
                              })

# -------------------------------------------------------------------
# 5. Precio: híbrido = escalación automática por inflación + ajuste manual opcional
# -------------------------------------------------------------------
st.subheader("6. Precio proyectado (híbrido: automático + ajuste manual)")

col_a, col_b = st.columns([1, 1])
with col_a:
    st.markdown("**6a. Precio base por SKU (editable)**")
    precio_default = hist.groupby("SKU")["Precio"].last().reset_index().rename(columns={"Precio": "Precio_Base"})
    precios_base = st.data_editor(precio_default, num_rows="fixed", use_container_width=True, key="precios_base")

with col_b:
    st.markdown("**6b. Tasa de inflación mensual para escalar el precio**")
    tasa_sugerida = float(referencias.get("Inflacion_Mensual", 0.0) * 100) if "Inflacion_Mensual" in referencias else 0.5
    tasa_inflacion_precio = st.number_input(
        "Tasa mensual (%)", min_value=-5.0, max_value=10.0, value=round(tasa_sugerida, 2), step=0.05,
        help="El precio base se escala mes a mes con esta tasa compuesta. "
             "Por defecto toma el promedio histórico de tu variable de inflación, si la seleccionaste como macro.",
    )
    st.caption(f"Precio_Mes_n = Precio_Base × (1 + {tasa_inflacion_precio:.2f}%)ⁿ")

st.markdown("**6c. Ajuste manual de precio por Cliente-SKU-Mes (editable, 0% = usa solo el escalado automático)**")
st.caption("Úsalo solo para casos puntuales: aumento de lista negociado con un cliente específico, "
           "cambio de precio por contrato, etc. En el 95% de los casos lo dejas en 0%.")
ajuste_precio_default = bau_df[["Cliente", "SKU", "Mes"]].copy()
ajuste_precio_default["Ajuste_Manual_Precio_%"] = 0.0
ajuste_precio_default["Motivo"] = ""
ajuste_precio = st.data_editor(
    ajuste_precio_default, num_rows="fixed", use_container_width=True, key="ajuste_precio",
    column_config={"Ajuste_Manual_Precio_%": st.column_config.NumberColumn(format="%.1f%%", step=0.5)},
)

# -------------------------------------------------------------------
# 6. Cálculo del forecast final
# -------------------------------------------------------------------
# Renombramos los coeficientes del modelo para que no choquen con los nombres
# de las variables macro (el escenario usa el nombre original, ej. "Inflacion_Mensual";
# el coeficiente del modelo pasa a llamarse "coef_Inflacion_Mensual").
coef_df_ren = coef_df.rename(columns={c: f"coef_{c}" for c in macro_cols})

final = bau_df.merge(escenario, on="Mes", how="left") if macro_cols else bau_df.copy()
final = final.merge(coef_df_ren, on=["Cliente", "SKU"], how="left")
final = final.merge(iniciativas[["Cliente", "SKU", "Mes", "%_Colaborativo"]], on=["Cliente", "SKU", "Mes"], how="left")
final = final.merge(precios_base, on="SKU", how="left")
final = final.merge(ajuste_precio[["Cliente", "SKU", "Mes", "Ajuste_Manual_Precio_%"]], on=["Cliente", "SKU", "Mes"], how="left")

# Número de mes transcurrido (1, 2, 3...) para la escalación compuesta
mes_orden = {m: i + 1 for i, m in enumerate(meses_fcst)}
final["n_mes"] = final["Mes"].map(mes_orden)
final["Precio_Escalado"] = final["Precio_Base"] * (1 + tasa_inflacion_precio / 100.0) ** final["n_mes"]
final["Precio"] = (final["Precio_Escalado"] * (1 + final["Ajuste_Manual_Precio_%"].fillna(0) / 100.0)).round(2)

# Impacto macro en unidades = Σ coef_i × (valor_escenario_i − valor_referencia_i)
# (referencia = promedio de los últimos 6 meses reales, usado para entrenar el BAU)
final["Impacto_Macro_Unidades"] = 0.0
for c in macro_cols:
    final["Impacto_Macro_Unidades"] += (final[c] - referencias[c]) * final[f"coef_{c}"]

final["Unidades_Final"] = np.round(
    (final["BAU_Unidades"] + final["Impacto_Macro_Unidades"]).clip(lower=0)
    * (1 + final["%_Colaborativo"].fillna(0) / 100.0)
).astype(int)
final["Venta_Final"] = (final["Unidades_Final"] * final["Precio"]).round(0)

st.subheader("7. Forecast Final")
cols_mostrar = ["Cliente", "SKU", "Mes", "BAU_Unidades", "%_Colaborativo", "Unidades_Final",
                "Precio_Base", "Ajuste_Manual_Precio_%", "Precio", "Venta_Final"]
st.dataframe(final[cols_mostrar], use_container_width=True)
st.caption("Precio = Precio_Base escalado por la tasa de inflación mensual (6b) × (1 + Ajuste_Manual_Precio_% del mes).")

resumen = final.groupby("Cliente").agg(
    Unidades_BAU=("BAU_Unidades", "sum"),
    Unidades_Final=("Unidades_Final", "sum"),
    Venta_Final=("Venta_Final", "sum"),
).reset_index()
resumen["Impacto_%"] = (resumen["Unidades_Final"] - resumen["Unidades_BAU"]) / resumen["Unidades_BAU"]

st.subheader("8. Resumen Ejecutivo por Cliente")
st.dataframe(resumen, use_container_width=True)
st.bar_chart(resumen.set_index("Cliente")[["Unidades_BAU", "Unidades_Final"]])

# -------------------------------------------------------------------
# 7. Exportar a Excel
# -------------------------------------------------------------------
def exportar_excel(hist, coef_df, escenario, iniciativas, precios_base, ajuste_precio, tasa_inflacion_precio, final, resumen) -> bytes:
    wb = Workbook()
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", start_color="1F4E78")

    def volcar(ws, df, start_row=1):
        for c, col in enumerate(df.columns, 1):
            cell = ws.cell(row=start_row, column=c, value=col)
            cell.font = hdr_font
            cell.fill = hdr_fill
        for i, row in df.iterrows():
            for c, col in enumerate(df.columns, 1):
                ws.cell(row=start_row + 1 + i, column=c, value=row[col])
        for c in range(1, len(df.columns) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 20

    ws = wb.active
    ws.title = "Historico"
    volcar(ws, hist)

    ws = wb.create_sheet("Modelo_Coeficientes")
    volcar(ws, coef_df.round(4))

    ws = wb.create_sheet("Escenario_Macro")
    volcar(ws, escenario)

    ws = wb.create_sheet("Iniciativas_Colaborativo")
    volcar(ws, iniciativas)

    ws = wb.create_sheet("Precio_Base")
    ws.cell(row=1, column=1, value=f"Tasa de inflación mensual usada para escalar precio: {tasa_inflacion_precio:.2f}%").font = Font(bold=True)
    volcar(ws, precios_base, start_row=3)

    ws = wb.create_sheet("Ajuste_Manual_Precio")
    volcar(ws, ajuste_precio)

    ws = wb.create_sheet("Forecast_Final")
    volcar(ws, final[cols_mostrar])

    ws = wb.create_sheet("Resumen_Ejecutivo")
    volcar(ws, resumen.round(3))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


excel_bytes = exportar_excel(hist, coef_df, escenario, iniciativas, precios_base, ajuste_precio,
                              tasa_inflacion_precio, final, resumen)
st.download_button(
    "⬇️ Descargar Excel del modelo",
    data=excel_bytes,
    file_name=f"Forecast_Colaborativo_{datetime.now():%Y%m%d}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
