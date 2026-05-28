# Sistema PRO de Seguimiento y Derivación Académica AVE

Aplicación en Streamlit para analizar estudiantes desde Canvas o desde reportes Excel/CSV, clasificar riesgo académico, generar mensajes, registrar historial y crear formatos de derivación para bienestar.

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Base de datos

La app permite cargar un Excel como base de datos desde la ventana **Base de datos**. El archivo puede contener estas hojas:

- `Estudiantes`
- `Asesores_Bienestar`
- `Historial_Estudiantes`
- `Derivaciones`
- `Mensajes_Enviados`
- `Consultas_Canvas`
- `Configuracion`

Si no se carga archivo, la app crea una estructura vacía en memoria y permite descargarla actualizada.

## Conexión con Canvas

La app solicita:

- URL institucional de Canvas
- Token personal
- ID del curso o selección manual del curso
- Semana de análisis

La disponibilidad de datos depende de los permisos del token. Si Canvas no permite leer alguna métrica, se puede usar la carga manual de reporte CSV/Excel.

## Google Sheets

La app incluye un conector opcional con Google Sheets mediante credenciales de cuenta de servicio. Para usarlo:

1. Crear una cuenta de servicio en Google Cloud.
2. Descargar el archivo JSON.
3. Compartir la hoja de cálculo con el correo de la cuenta de servicio.
4. Cargar el JSON en la app.


## Cambios versión PRO 3.0

- Selector de cursos desde Canvas con base en los permisos del token.
- Extracción más completa desde Canvas: matrículas, calificaciones, última actividad, actividades publicadas y entregas.
- Clasificación de riesgo corregida: los valores vacíos/None ya no se clasifican automáticamente como Bajo.
- Si Canvas no expone datos suficientes de avance, la app marca el caso como alerta para evitar falsos positivos de bajo riesgo.
- El análisis obtenido desde Canvas se registra automáticamente en el historial y en Consultas_Canvas.

## Conexión corregida con Google Sheets

La app puede usar Google Sheets como base de datos en línea. Para que funcione:

1. Crear una cuenta de servicio en Google Cloud.
2. Descargar el archivo JSON de credenciales.
3. Abrir el JSON y copiar el valor `client_email`.
4. Compartir el Google Sheets con ese correo, con permiso de Editor.
5. En la app, pegar la URL completa del Google Sheets y subir el JSON.
6. Usar primero `Probar conexión`.
7. Usar `Inicializar estructura` para crear las pestañas necesarias.
8. Usar `Leer Google Sheets` o `Guardar base actual en Google Sheets` según corresponda.

Si se despliega en Streamlit Cloud, también puede colocarse el contenido del JSON en `st.secrets` bajo `[gcp_service_account]` para no subirlo cada vez.
