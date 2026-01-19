import re
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Validación Cecos vs Interfaces", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def to_numeric_series(s: pd.Series) -> pd.Series:
    # Convierte textos tipo "0012", "12 ", "12.0" a número, dejando NaN si no puede.
    return pd.to_numeric(
        s.astype(str)
         .str.replace(r"\.0$", "", regex=True)
         .str.replace(r"[^\d]", "", regex=True)  # deja solo dígitos
         .replace("", pd.NA),
        errors="coerce"
    )

def find_cell_position(df_raw: pd.DataFrame, target: str):
    # Busca EXACTO (trim + case-insensitive) en todo el DataFrame
    target_norm = target.strip().lower()
    for r in range(df_raw.shape[0]):
        row = df_raw.iloc[r]
        for c in range(df_raw.shape[1]):
            val = normalize_str(row.iloc[c]).lower()
            if val == target_norm:
                return r, c
    return None

def last_nonempty_row_in_col(df_raw: pd.DataFrame, col_idx: int, start_row: int):
    # Retorna el índice de la última fila >= start_row donde en esa columna haya texto/valor
    last = None
    for r in range(start_row, df_raw.shape[0]):
        val = df_raw.iat[r, col_idx] if col_idx < df_raw.shape[1] else None
        if normalize_str(val) != "":
            last = r
    return last

def extract_cecos_from_interfaces(df_interfaces: pd.DataFrame) -> pd.Series:
    # Preferir columna llamada "Cecos" (case-insensitive)
    cols_map = {c.lower(): c for c in df_interfaces.columns}
    if "cecos" in cols_map:
        s = df_interfaces[cols_map["cecos"]]
    else:
        # Si no existe, usar columna F (index 5) como dijiste
        if df_interfaces.shape[1] < 6:
            raise ValueError("El consolidado de interfaces no tiene al menos 6 columnas para tomar la columna F.")
        s = df_interfaces.iloc[:, 5]
    return to_numeric_series(s).dropna().astype(int)

def load_excel_any_sheet(uploaded_file, sheet_name=None, header=0):
    # Lee Excel; si sheet_name es None, toma la primera hoja
    return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl", header=header)

# -----------------------------
# UI
# -----------------------------
st.title("Validación: Tiendas/CEDIS faltantes en Interfaces")

st.markdown(
    """
Este app:
- Carga **interfaces consolidadas** (saca **Cecos**).
- Carga archivo **Cecos** (extrae tabla desde celda **Status** en hoja **JMC Cost Center Strucutre**, filtra y toma **tiendas**).
- Carga **MD mes anterior** (filtra Ce. Coste inicia "102" y extrae **CEDIS**).
- Te entrega qué **tiendas y CEDIS faltan** en interfaces.
"""
)

col1, col2, col3 = st.columns(3)

# -----------------------------
# Paso 1: Interfaces
# -----------------------------
with col1:
    st.subheader("Paso 1) Interfaces consolidadas")
    f_interfaces = st.file_uploader("Sube el consolidado de interfaces (Excel)", type=["xlsx", "xls"], key="interfaces")

# -----------------------------
# Paso 2: Cecos
# -----------------------------
with col2:
    st.subheader("Paso 2) Archivo Cecos")
    f_cecos = st.file_uploader("Sube el archivo Cecos (Excel)", type=["xlsx", "xls"], key="cecos")

# -----------------------------
# Paso 3: MD mes anterior
# -----------------------------
with col3:
    st.subheader("Paso 3) MD mes anterior")
    f_md = st.file_uploader("Sube el MD mes anterior (Excel)", type=["xlsx", "xls"], key="md")

st.divider()

# -----------------------------
# Procesamiento
# -----------------------------
interfaces_cecos = None
tiendas_cecos = None
cedis_md = None

# --- Interfaces ---
if f_interfaces:
    try:
        df_int = load_excel_any_sheet(f_interfaces, sheet_name=None, header=0)
        # si viene como dict (varias hojas), tomar la primera
        if isinstance(df_int, dict):
            first_sheet = list(df_int.keys())[0]
            df_int = df_int[first_sheet]

        interfaces_cecos = extract_cecos_from_interfaces(df_int).unique()
        st.success(f"Interfaces cargadas. Cecos únicos detectados: {len(interfaces_cecos)}")
        with st.expander("Ver muestra de Cecos (interfaces)"):
            st.write(pd.DataFrame({"Cecos (interfaces)": pd.Series(interfaces_cecos).sort_values().head(50)}))
    except Exception as e:
        st.error(f"Error procesando interfaces: {e}")

# --- Cecos ---
if f_cecos:
    try:
        sheet_name = "JMC Cost Center Strucutre"
        df_raw = pd.read_excel(f_cecos, sheet_name=sheet_name, engine="openpyxl", header=None)

        pos = find_cell_position(df_raw, "Status")
        if not pos:
            raise ValueError('No encontré ninguna celda con el texto exacto "Status" en la hoja "JMC Cost Center Strucutre".')

        header_row_idx, status_col_idx = pos

        # Columnas A..AM => índices 0..38
        col_start = 0
        col_end = 38  # AM
        if df_raw.shape[1] <= col_end:
            raise ValueError("La hoja no tiene hasta la columna AM (índice 38).")

        # Última fila con texto en columna AM, desde la fila del header hacia abajo
        last_row = last_nonempty_row_in_col(df_raw, col_end, header_row_idx)
        if last_row is None or last_row <= header_row_idx:
            raise ValueError("No pude determinar el final de la tabla (columna AM sin datos).")

        # Encabezados desde A..AM en la fila header_row_idx
        headers = df_raw.iloc[header_row_idx, col_start:col_end + 1].tolist()
        headers = [normalize_str(h) if normalize_str(h) != "" else f"COL_{i+1}"
                   for i, h in enumerate(headers)]

        # Datos desde la siguiente fila hasta last_row
        data = df_raw.iloc[header_row_idx + 1:last_row + 1, col_start:col_end + 1].copy()
        data.columns = headers

        # Columna K (index 10) debe llamarse Status (pero por seguridad usamos posición si existe)
        # Columna AD (index 29) Concepto tienda
        # Columna C (index 2) tienda
        status_col = data.columns[10]
        concepto_col = data.columns[29]
        tienda_col = data.columns[2]

        # Filtros
        data_f = data.copy()
        data_f[status_col] = data_f[status_col].astype(str).str.strip()
        data_f[concepto_col] = data_f[concepto_col].astype(str).str.strip()

        data_f = data_f[data_f[status_col].str.upper() == "ABIERTA"]
        data_f = data_f[data_f[concepto_col].str.upper() != "FRANQUICIA"]

        # Tiendas a número
        tiendas_num = to_numeric_series(data_f[tienda_col]).dropna().astype(int).unique()
        tiendas_cecos = tiendas_num

        st.success(f"Cecos cargado y filtrado. Tiendas (únicas) detectadas: {len(tiendas_cecos)}")
        with st.expander("Ver muestra de tiendas (Cecos filtrado)"):
            st.write(pd.DataFrame({"Tienda (Cecos filtrado)": pd.Series(tiendas_cecos).sort_values().head(50)}))

    except Exception as e:
        st.error(f"Error procesando Cecos: {e}")

# --- MD ---
if f_md:
    try:
        df_md = load_excel_any_sheet(f_md, sheet_name=None, header=0)
        if isinstance(df_md, dict):
            first_sheet = list(df_md.keys())[0]
            df_md = df_md[first_sheet]

        # Columna K (index 10): Ce. Coste
        # Columna L (index 11): Centro de Coste
        if df_md.shape[1] < 12:
            raise ValueError("El archivo MD no tiene suficientes columnas para usar K y L.")

        ce_coste = df_md.iloc[:, 10].astype(str).str.strip()
        centro_coste = df_md.iloc[:, 11].astype(str).str.strip()

        # Filtrar Ce. Coste inicia con "102"
        mask_102 = ce_coste.str.startswith("102", na=False)
        df_md_f = df_md.loc[mask_102].copy()

        # Extraer 4 primeros caracteres y validar numéricos
        centro_vals = df_md_f.iloc[:, 11].astype(str).str.strip()
        first4 = centro_vals.str.slice(0, 4)

        # Solo los que son 4 dígitos
        first4_num = pd.to_numeric(first4.where(first4.str.fullmatch(r"\d{4}", na=False)), errors="coerce")
        cedis_md = first4_num.dropna().astype(int).unique()

        st.success(f"MD cargado y filtrado. CEDIS (únicos) detectados: {len(cedis_md)}")
        with st.expander("Ver muestra de CEDIS (MD filtrado)"):
            st.write(pd.DataFrame({"CEDIS (MD filtrado)": pd.Series(cedis_md).sort_values().head(50)}))

    except Exception as e:
        st.error(f"Error procesando MD: {e}")

st.divider()

# -----------------------------
# Comparaciones finales
# -----------------------------
st.subheader("Resultado: faltantes en Interfaces")

if interfaces_cecos is None:
    st.warning("Falta cargar el consolidado de interfaces (Paso 1).")
else:
    set_interfaces = set(map(int, interfaces_cecos))

    # Tiendas faltantes
    missing_tiendas = None
    if tiendas_cecos is not None:
        set_tiendas = set(map(int, tiendas_cecos))
        missing_tiendas = sorted(list(set_tiendas - set_interfaces))
        st.metric("Tiendas faltantes (Cecos filtrado vs Interfaces)", len(missing_tiendas))

    # CEDIS faltantes
    missing_cedis = None
    if cedis_md is not None:
        set_cedis = set(map(int, cedis_md))
        missing_cedis = sorted(list(set_cedis - set_interfaces))
        st.metric("CEDIS faltantes (MD vs Interfaces)", len(missing_cedis))

    # Unión faltantes
    if tiendas_cecos is not None or cedis_md is not None:
        union_source = set()
        if tiendas_cecos is not None:
            union_source |= set(map(int, tiendas_cecos))
        if cedis_md is not None:
            union_source |= set(map(int, cedis_md))

        missing_union = sorted(list(union_source - set_interfaces))
        st.metric("TOTAL faltantes en Interfaces (Tiendas ∪ CEDIS)", len(missing_union))

        # Mostrar tablas
        c1, c2, c3 = st.columns(3)

        with c1:
            st.write("**Tiendas faltantes**")
            st.dataframe(pd.DataFrame({"tienda": missing_tiendas or []}))

        with c2:
            st.write("**CEDIS faltantes**")
            st.dataframe(pd.DataFrame({"cedis": missing_cedis or []}))

        with c3:
            st.write("**Total faltantes (unión)**")
            st.dataframe(pd.DataFrame({"faltante": missing_union or []}))

        # Descargar a Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"Cecos_interfaces": sorted(list(set_interfaces))}).to_excel(writer, index=False, sheet_name="Interfaces_Cecos")

            if tiendas_cecos is not None:
                pd.DataFrame({"Tiendas_Cecos_filtrado": sorted(list(set(map(int, tiendas_cecos))))}).to_excel(
                    writer, index=False, sheet_name="Tiendas_Cecos"
                )
                pd.DataFrame({"Tiendas_faltantes": missing_tiendas}).to_excel(writer, index=False, sheet_name="Faltantes_Tiendas")

            if cedis_md is not None:
                pd.DataFrame({"CEDIS_MD": sorted(list(set(map(int, cedis_md))))}).to_excel(
                    writer, index=False, sheet_name="CEDIS_MD"
                )
                pd.DataFrame({"CEDIS_faltantes": missing_cedis}).to_excel(writer, index=False, sheet_name="Faltantes_CEDIS")

            pd.DataFrame({"Faltantes_union": missing_union}).to_excel(writer, index=False, sheet_name="Faltantes_Total")

        st.download_button(
            "Descargar resultado en Excel",
            data=output.getvalue(),
            file_name="faltantes_en_interfaces.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Cargaste interfaces, pero falta Cecos (Paso 2) y/o MD (Paso 3) para comparar.")
