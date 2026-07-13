# Modelo de Proyección de Presupuesto Colaborativo — CBM CAM SUR

## Cómo correrlo

1. Instala Python 3.10+ (si no lo tienes) o usa Anaconda.
2. Abre una terminal en esta carpeta y corre:
   ```
   pip install -r requirements.txt
   streamlit run app_forecast_colaborativo.py
   ```
3. Se abrirá automáticamente en tu navegador (http://localhost:8501).

## Estructura del histórico a subir
Columnas obligatorias (nombres exactos): `Cliente, SKU, Mes, Unidades, Precio`
Mes en formato `YYYY-MM`. Agrega las columnas macro que quieras (cualquier nombre numérico).

**Opcionales (recomendadas) para habilitar los building blocks:**
- `Canal` (ej. Moderno, Tradicional) — atributo por Cliente, característica (no predictora)
- `Grupo_Cliente` (ej. Cadenas Nacionales, Mayoristas) — atributo por Cliente
- `Pais` (ej. Costa Rica, Panamá, Nicaragua) — atributo por Cliente, característica (no predictora)
- `Marca` (ej. Marca Azul, Marca Roja) — atributo por SKU
- `Categoria` (ej. Bebidas, Snacks) — atributo por SKU
- `SKU_ID` — identificador/código del SKU, separado del nombre descriptivo que va en la columna `SKU`
  (ej. `SKU_ID = "SKU-101"`, `SKU = "Bebida 600ml"`). Solo referencia, no se usa como predictor.

No afectan el modelo de proyección — Pais/Canal/Grupo_Cliente/Marca/Categoria solo habilitan el desglose
de la sección 10, y SKU_ID es puramente informativo (aparece en la tabla de Forecast Final y en el Excel
exportado). Puedes subir todas, algunas, o ninguna — la app solo muestra lo que sí tengas cargado.

### Variables macro sugeridas (según el contexto de CBM Cono Sur)
Si tienes estos datos por país y por mes, agrégalos como columnas numéricas y selecciónalos en el punto 2:
- `PIB_Pais` — crecimiento del PIB proyectado (CR 3.5–4.0%, Panamá 3.5–4.5%, Nicaragua 3–4%)
- `Inflacion_Pais` — inflación proyectada (CR dentro de la meta BCCR 3% ±1)
- `Tipo_Cambio_USD_Local` — relevante sobre todo para Costa Rica (USD/CRC); Panamá está dolarizado (0
  riesgo cambiario), así que ahí puede ir en 0 o quedar fuera
- `Costo_Commodities_Index` — normalización/volatilidad de costos de commodities y fletes
- `Remesas_Idx` — relevante para Nicaragua, donde el consumo está sostenido por remesas

Estas no vienen precargadas en los datos de ejemplo — son una sugerencia para cuando conectes tu
histórico real, tomadas del contexto macro de tu presentación de presupuesto 2027.

## Modo colaborativo (guardado compartido entre todo el equipo)

Por defecto cada persona trabaja aislada en su sesión (nada se guarda al refrescar). Si quieres que los
cambios (escenario macro, % colaborativo, precio) se guarden y todo el equipo vea lo último editado,
sigue la guía **SETUP_MODO_COLABORATIVO.md** — toma ~15 minutos, usa un Google Sheet gratis como base de
datos compartida. Sin esa configuración, la app funciona igual, solo sin guardado compartido.

## Apartados de la app
0. Logo (sidebar) — sube el logo de la empresa, se muestra junto al título.
1. Datos históricos (sidebar) — subes tu archivo o usas datos de ejemplo.
2. Variables macro / adicionales (sidebar) — eliges qué columnas numéricas usar como regresores.
3. Horizonte de forecast (sidebar) — cuántos meses proyectar.
4. Escenario macro futuro (editable).
5. Palancas de Ejecución Propia (editable) — Nuevos Productos, Profundización de Cuentas y Cuentas Nuevas,
   por Cliente-SKU-Mes. Mismo lenguaje que las palancas 3, 4 y 5 del puente de crecimiento del presupuesto.
   Afectan **volumen**, no precio.
6. Precio proyectado (híbrido):
   - 6a. Precio base por SKU (editable).
   - 6b. Tasa de inflación mensual para escalar el precio automáticamente mes a mes (compuesta).
   - 6c. Ajuste manual de precio por Cliente-SKU-Mes (editable, 0% por defecto).
7. Forecast Final — tabla detallada.
8. Resumen Ejecutivo por Cliente — BAU vs Final + gráfico de barras.
9. Tendencia de Volumen y Venta — línea Real (sólida) + Proyectado (punteada), todos los periodos.
10. Building Blocks por Canal, Grupo de Clientes, País, Marca y Categoría — desglose en barras (grilla 2x2),
    con selector de periodo: Acumulado (Real+Proyectado), Solo Real, Solo Proyectado, o un Mes específico;
    y selector de métrica (Unidades o Venta). Muestra solo las dimensiones que sí tengas cargadas.
11. Puente de Crecimiento — waterfall en el mismo lenguaje que la descomposición del presupuesto: Base (BAU)
    → Precio → Mercado Orgánico → Nuevos Productos → Profundización de Cuentas → Cuentas Nuevas → Total
    Proyectado. Muestra el $ y el % que aporta cada palanca sobre la Venta Base.

Botón de descarga al final exporta todo a Excel (una hoja por tabla).

## Qué hace el modelo (enfoque top-down)
En vez de ajustar un modelo independiente por cada Cliente-SKU (miles, la mayoría con poca historia y
mucho ruido), el modelo:

1. **Agrega todos los Cliente-SKU en una sola serie mensual total** y ajusta un único modelo de tendencia +
   estacionalidad (+ variables macro) sobre esa serie — mucho más estable, y con estacionalidad real y
   visible (no diluida entre miles de series individuales con poca o ninguna historia).

   `log(Unidades totales) = intercepto + coef_trend·mes_calendario + coef_sin·sin(mes_del_año) + coef_cos·cos(mes_del_año) + Σ coef_macro_i · variable_macro_i`

2. **Desagrega** ese total proyectado según la **participación histórica** de cada Cliente-SKU dentro del
   total (su volumen histórico ÷ el volumen histórico de todos). Los clientes de **compra única** (1 solo
   mes de historia) quedan **excluidos** — no se proyectan.

**3 topes de sensatez** para que el modelo no invente crecimiento o estacionalidad a partir de ruido:
- Tendencia: máximo 3%/mes si seleccionaste variables macro, o solo 1%/mes si no seleccionaste ninguna.
- Estacionalidad: la amplitud del ciclo anual se topa (con solo 2-3 años de historia, cada mes calendario
  tiene muy pocos datos para estimarse solo, y puede sobreajustarse a un pico/valle irreal).
- Nivel: el intercepto se recalibra siempre al promedio de los **últimos 6 meses** (no a los 2 años
  completos), para que el punto de partida sea "dónde está el negocio hoy", no un promedio histórico diluido.

El impacto de tu escenario macro editado se calcula como `coef_i × (valor_escenario − valor_referencia)`, en
**%** (Mercado_Organico_%), aplicado sobre el total ya desagregado. El % de las palancas de ejecución propia
se aplica multiplicativamente encima. El precio final = Precio_Base × (1 + tasa_inflación)^n_mes × (1 +
Ajuste_Manual_Precio_% del mes).

En el Forecast Final (sección 7) puedes ver `Categoria_Modelo` ("Excluido" o "Participación histórica") y
`Participacion_%` de cada Cliente-SKU, más un expander con el conteo de cada categoría y otro con el detalle
de participación completo.

## Descargas
Al final hay 2 botones: **Excel** (todas las hojas: histórico, coeficientes, escenario, palancas, precios,
forecast final, resumen, puente de crecimiento) y **CSV** (solo el Forecast Final, separado por `;` y
decimal `,` para que abra directo en Excel en español).

## Eficiencia del guardado compartido
Con miles de combinaciones Cliente-SKU-Mes, guardar la grilla completa en Google Sheets sería lentísimo.
Por eso el botón de guardado solo sube las filas que **sí editaste** (palancas ≠ 0%, o ajuste de precio ≠ 0%)
— el resto se completa solo con el valor por defecto (0%) al cargar. Esto no cambia nada de cómo usas la
app, solo la hace mucho más rápida con archivos grandes.
