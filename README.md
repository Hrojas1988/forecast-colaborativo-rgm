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
- `Canal` (ej. Moderno, Tradicional) — atributo por Cliente
- `Grupo_Cliente` (ej. Cadenas Nacionales, Mayoristas) — atributo por Cliente
- `Marca` (ej. Marca Azul, Marca Roja) — atributo por SKU
- `Categoria` (ej. Bebidas, Snacks) — atributo por SKU

No afectan el modelo de proyección, solo habilitan el desglose de la sección 10. Puedes subir todas, algunas,
o ninguna — la app solo muestra los gráficos de las dimensiones que sí tengas cargadas.

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
5. % Colaborativo por iniciativa comercial (editable) — afecta **volumen**, no precio.
6. Precio proyectado (híbrido):
   - 6a. Precio base por SKU (editable).
   - 6b. Tasa de inflación mensual para escalar el precio automáticamente mes a mes (compuesta).
   - 6c. Ajuste manual de precio por Cliente-SKU-Mes (editable, 0% por defecto).
7. Forecast Final — tabla detallada.
8. Resumen Ejecutivo por Cliente — BAU vs Final + gráfico de barras.
9. Tendencia de Volumen y Venta — línea Real (sólida) + Proyectado (punteada), todos los periodos.
10. Building Blocks por Canal, Grupo de Clientes, Marca y Categoría — desglose en barras (grilla 2x2),
    con selector de periodo: Acumulado (Real+Proyectado), Solo Real, Solo Proyectado, o un Mes específico;
    y selector de métrica (Unidades o Venta). Muestra solo las dimensiones que sí tengas cargadas.

Botón de descarga al final exporta todo a Excel (una hoja por tabla).

## Qué hace el modelo
Por cada combinación Cliente-SKU se entrena una Regresión Lineal:

Unidades = intercepto + coef_trend·t + coef_sin·sin(2πt/12) + coef_cos·cos(2πt/12) + Σ coef_macro_i · variable_macro_i

El BAU se proyecta usando el valor de referencia (promedio de los últimos 6 meses) de cada variable macro.
El impacto de tu escenario editado se calcula como coef_i × (valor_escenario − valor_referencia), evitando doble conteo.
El % colaborativo se aplica multiplicativamente sobre las unidades (BAU + impacto macro).
El precio final = Precio_Base × (1 + tasa_inflación)^n_mes × (1 + Ajuste_Manual_Precio_% del mes).
