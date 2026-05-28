# AVE Seguimiento PRO v6 - Excel local o sincronizado

Esta versión elimina la conexión a Google Sheets con JSON y trabaja con una base de datos en Excel de forma más amigable.

## Modos de base de datos

1. **Cargar Excel manualmente**
   - El asesor sube el archivo `.xlsx`.
   - La app analiza, registra historial, mensajes y derivaciones.
   - Al finalizar, se descarga la base actualizada.

2. **Usar Excel desde ruta local / OneDrive / SharePoint sincronizado**
   - El archivo maestro vive en una carpeta sincronizada.
   - Ejemplo: `C:/Users/Usuario/Universidad del Valle de Guatemala/AVE/base_datos_seguimiento_ave.xlsx`
   - La app puede leer y guardar directamente en esa ruta.
   - Antes de guardar, crea un respaldo automático en `backups_ave/`.

3. **Crear base nueva**
   - Genera una base vacía con la estructura institucional.

## Recomendación importante

Si se usa SharePoint/OneDrive sincronizado, ejecuta la app en una computadora o servidor que tenga acceso a la carpeta sincronizada. En Streamlit Cloud, una ruta como `C:/Users/...` no existe porque la app corre en un servidor externo.

## Ejecutar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Hojas de la base

- Estudiantes
- Asesores_Bienestar
- Historial_Estudiantes
- Derivaciones
- Mensajes_Enviados
- Consultas_Canvas
- Configuracion

## Cambios v8
- Corrección de error de Canvas/Excel: `merge on str and int64 columns for key carne`.
- Normalización automática de carné, correo, login_id y canvas_user_id como texto.
- Limpieza de valores leídos desde Excel como `20261234.0` para convertirlos a `20261234`.
- Mejor compatibilidad al combinar estudiantes obtenidos desde Canvas con bases Excel locales, OneDrive o SharePoint sincronizado.
