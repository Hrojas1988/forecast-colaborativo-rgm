# Modelo de Forecast Colaborativo — Cliente x SKU

## Cómo correrlo

1. Instala Python 3.10+ (si no lo tienes).
2. Abre una terminal en esta carpeta y corre:
   ```
   pip install -r requirements.txt
   streamlit run app_forecast_colaborativo.py
   ```
3. Se abrirá automáticamente en tu navegador (http://localhost:8501).
4. En la barra lateral:
   - Sube tu histórico (CSV o Excel) con columnas mínimas: `Cliente, SKU, Mes, Unidades, Precio` + las variables macro que quieras (ej. `Inflacion_Mensual`, `Var_TipoCambio`, `PIB`, `Confianza_Consumidor`, etc. — cualquier nombre de columna numérica sirve).
   - O deja marcada la opción "Usar datos de ejemplo" para probar con datos ficticios.
   - Elige qué columnas usar como variables macro y el horizonte de meses a proyectar.
5. En el cuerpo de la página:
   - Revisa los coeficientes del modelo (una regresión lineal por Cliente-SKU).
   - Edita el escenario macro futuro directamente en la tabla.
   - Edita el % colaborativo por Cliente-SKU-Mes (tus iniciativas comerciales).
   - Edita el precio proyectado por SKU.
6. Descarga el resultado con el botón "⬇️ Descargar Excel del modelo".

## Qué hace el modelo
Por cada combinación Cliente-SKU se entrena una Regresión Lineal:

`Unidades = intercepto + coef_trend·t + coef_sin·sin(2πt/12) + coef_cos·cos(2πt/12) + Σ coef_macro_i · variable_macro_i`

El BAU se proyecta usando el valor de referencia (promedio de los últimos 6 meses) de cada variable macro.
El impacto de tu escenario editado se calcula como `coef_i × (valor_escenario − valor_referencia)`, evitando doble conteo.
El % colaborativo se aplica multiplicativamente sobre el resultado (BAU + impacto macro).
