# Faltantes en Interfaces (Streamlit)

App web en Streamlit para identificar qué **Tiendas** y **CEDIS** faltan en el consolidado de **Interfaces**, comparando contra:
1) Archivo **Cecos** (tiendas activas filtradas)
2) **MD mes anterior** (extrae CEDIS)
3) **Dash de tiendas** (enriquece el reporte final para correo)

La app genera un Excel descargable con la hoja **Faltantes_Total** lista para copiar/pegar en un correo.

---

## ✅ Usuario final (solo web)

1. Entra al enlace de la app.
2. Sube los 4 archivos en orden:
   - **Interfaces**
   - **Cecos** (hoja: `JMC Cost Center Strucutre`)
   - **MD mes anterior**
   - **Dash de tiendas**
3. La app mostrará la tabla **Faltantes_Total** y un botón para **Descargar Excel**.

---

## 🧾 Reglas de negocio implementadas

### Interfaces (Paso 1)
- Toma Cecos desde la columna **"Cecos"** (si existe).
- Si no existe, toma la **columna F**.

### Cecos (Paso 2)
- Hoja obligatoria: `JMC Cost Center Strucutre`
- Busca la celda con texto **Status**
- Usa esa fila como encabezado desde **A hasta AM**
- Toma datos desde la fila siguiente hasta la última fila con texto en columna **AM**
- Filtros:
  - **Status** = `ABIERTA`
  - **Concepto tienda** != `Franquicia`
- Tienda = columna C (se convierte a número)

### MD (Paso 3)
- Filtra **Ce. Coste** (col K) que inicia con `102`
- De **Centro de Coste** (col L) extrae 4 primeros caracteres si son numéricos => CEDIS

### Dash de tiendas (Paso 4)
Cruza por **TIENDA** (Value Tienda col C) para generar la tabla final para correo con:
- Ce.coste = col B (CECO)
- Centro de coste = col G (Nombre tienda)
- TIENDA = col C (Value Tienda)
- AM = col S
- DM = col R
- Región = col E

---

## 📦 Archivos del repo

- `app.py` => aplicación Streamlit
- `requirements.txt` => dependencias
- `README.md` => documentación
- `.gitignore` => evita subir archivos innecesarios

---

## 🚀 Despliegue en Streamlit Community Cloud (para el administrador)

1. Subir este repo a GitHub (público o privado según tu plan).
2. En Streamlit Community Cloud:
   - New app
   - Selecciona el repo
   - Main file path: `app.py`
3. Deploy.

El usuario final solo necesita la URL.
