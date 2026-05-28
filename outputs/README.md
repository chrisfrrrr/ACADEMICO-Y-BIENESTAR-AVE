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

