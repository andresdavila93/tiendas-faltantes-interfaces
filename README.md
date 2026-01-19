# Faltantes en Interfaces (Streamlit)

App web en Streamlit para identificar qué Tiendas y CEDIS faltan en Interfaces, y generar un Excel con **Faltantes_Total** listo para correo.

## Usuario final
1. Entra a la URL de la app.
2. Sube 4 archivos:
   - Interfaces
   - Cecos (hoja: `JMC Cost Center Strucutre`)
   - MD mes anterior
   - Dash de tiendas
3. Descarga el Excel generado.

## Despliegue (Streamlit Cloud)
- Main file path: `streamlit run app.py`

## Nota (importante por el nombre del archivo)
Este repo usa un nombre de archivo con espacios:
`streamlit run app.py`

Si lo ejecutas local:
`streamlit run "streamlit run app.py"`
