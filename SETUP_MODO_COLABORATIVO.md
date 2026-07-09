# Cómo activar el Modo Colaborativo (guardado compartido)

Esto hace que cuando alguien edite el escenario macro, el % colaborativo, el precio base o el ajuste
manual de precio y le dé clic a "Guardar cambios para todo el equipo", la próxima persona que abra el
link vea esos mismos valores — usando un Google Sheet como base de datos compartida (gratis).

Si NO haces esta configuración, la app sigue funcionando normal, solo que cada quien trabaja aislado en
su propia sesión (como hasta ahora).

## Paso 1: Crear el Google Sheet

1. Ve a https://sheets.google.com y crea una hoja de cálculo nueva.
2. Ponle el nombre que quieras, ej. "Forecast_CBM_Datos_Compartidos".
3. Copia la URL completa de la hoja (la vas a necesitar en el Paso 4).
4. No hace falta crear pestañas ni columnas — la app las crea solas la primera vez que alguien guarda.

## Paso 2: Crear una cuenta de servicio de Google Cloud (las "credenciales")

1. Ve a https://console.cloud.google.com/
2. Crea un proyecto nuevo (o usa uno existente) — arriba a la izquierda, "Select a project" → "New Project".
3. En el buscador superior escribe **"Google Sheets API"** y haz clic en **"Enable"** (habilitar).
4. Busca también **"Google Drive API"** y habilítala igual.
5. Ve a **"APIs & Services" → "Credentials"** (Credenciales).
6. Clic en **"Create Credentials" → "Service Account"**.
7. Ponle un nombre (ej. "streamlit-forecast") y dale **"Create and Continue"**, luego **"Done"**.
8. En la lista de cuentas de servicio, haz clic en la que acabas de crear.
9. Ve a la pestaña **"Keys"** → **"Add Key" → "Create new key"** → tipo **JSON** → **Create**.
10. Se descarga un archivo `.json` a tu computadora — **guárdalo, es la credencial**.

## Paso 3: Compartir el Google Sheet con la cuenta de servicio

1. Abre el archivo `.json` que descargaste con el Bloc de Notas.
2. Busca el campo `"client_email"` — es un correo tipo `algo@tu-proyecto.iam.gserviceaccount.com`.
3. Vuelve a tu Google Sheet (Paso 1) → botón **"Compartir"** (arriba a la derecha).
4. Pega ese correo y dale permisos de **Editor**. Comparte.

## Paso 4: Agregar las credenciales a Streamlit Cloud (Secrets)

1. Ve a https://share.streamlit.io, entra a tu app.
2. Clic en los 3 puntos (⋮) → **"Settings" → "Secrets"**.
3. Pega esto, reemplazando los valores con los de tu archivo `.json` y la URL de tu Sheet:

```toml
[connections.gsheets]
spreadsheet = "PEGA_AQUI_LA_URL_COMPLETA_DE_TU_GOOGLE_SHEET"
type = "service_account"
project_id = "pega-aqui-el-project_id-del-json"
private_key_id = "pega-aqui-el-private_key_id-del-json"
private_key = "pega-aqui-el-private_key-del-json-COMPLETO-incluyendo-los---BEGIN/END---"
client_email = "pega-aqui-el-client_email-del-json"
client_id = "pega-aqui-el-client_id-del-json"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "pega-aqui-el-client_x509_cert_url-del-json"
```

4. Guarda. La app se reinicia sola en 1-2 minutos.

## Paso 5: Verificar que funcionó

1. Abre tu app — en la barra lateral debe aparecer marcada la casilla **"Guardar y compartir cambios con
   todo el equipo (Google Sheets)"** en vez del aviso de "Modo colaborativo no configurado".
2. Edita cualquier tabla (ej. el % colaborativo) y dale clic a **"💾 Guardar cambios para todo el equipo"**.
3. Revisa tu Google Sheet — deberían aparecer pestañas nuevas (`escenario_macro`, `iniciativas_colaborativo`,
   `precios_base`, `ajuste_manual_precio`) con los datos.
4. Pide a otra persona que abra el link de la app — debería ver los mismos valores que guardaste.

## Importante

- **No hay fusión de cambios en tiempo real**: si dos personas editan y guardan al mismo tiempo, gana la
  última en darle clic a "Guardar". No es colaboración simultánea tipo Google Docs, es "guardar y compartir".
- Nunca subas el archivo `.json` de credenciales a GitHub — solo va en "Secrets" de Streamlit Cloud, que es
  privado.
