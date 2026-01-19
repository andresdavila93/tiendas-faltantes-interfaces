import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Faltantes en Interfaces", layout="wide")
st.title("Validación: Tiendas y CEDIS faltantes en Interfaces")

# =========================================================
# Helpers generales
# =========================================================
def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def to_numeric_series(s: pd.Series) -> pd.Series:
    # Convierte textos con números a número (quita todo lo que no sea dígito)
    return pd.to_numeric(
        s.astype(str)
         .str.replace(r"\.0$", "", regex=True)
         .str.replace(r"[^\d]", "", regex=True)
         .replace("", pd.NA),
        errors="coerce"
    )

def load_excel_first_sheet(uploaded_file, header=0):
    x = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl", header=header)
    if isinstance(x, dict):
        return x[list(x.keys())[0]]
    return x

# =========================================================
# Paso 1 - Interfaces
# =========================================================
def extract_cecos_from_interfaces(df_interfaces: pd.DataFrame) -> pd.Series:
    # Preferir columna "Cecos" (por nombre). Si no existe, usar columna F (index 5)
    cols_map = {c.lower().strip(): c for c in df_interfaces.columns}
    if "cecos" in cols_map:
        s = df_interfaces[cols_map["cecos"]]
    else:
        if df_interfaces.shape[1] < 6:
            raise ValueError("El consolidado de interfaces no tiene al menos 6 columnas (para usar columna F).")
        s = df_interfaces.iloc[:, 5]  # F
    return to_numeric_series(s).dropna().astype(int)

# =========================================================
# Paso 2 - Cecos (robusto)
# =========================================================
def find_cell_position(df_raw: pd.DataFrame, target: str):
    target_norm = target.strip().lower()
    for r in range(df_raw.shape[0]):
        for c in range(df_raw.shape[1]):
            val = normalize_str(df_raw.iat[r, c]).lower()
            if val == target_norm:
                return r, c
    return None

def last_nonempty_row_in_col(df_raw: pd.DataFrame, col_idx: int, start_row: int):
    last = None
    for r in range(start_row, df_raw.shape[0]):
        if col_idx >= df_raw.shape[1]:
            break
        val = df_raw.iat[r, col_idx]
        if normalize_str(val) != "":
            last = r
    return last

def process_cecos_file(uploaded_file):
    sheet_name = "JMC Cost Center Strucutre"
    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl", header=None)

    pos = find_cell_position(df_raw, "Status")
    if not pos:
        raise ValueError('No encontré la palabra exacta "Status" en la hoja "JMC Cost Center Strucutre".')

    header_row, _ = pos

    # A..AM (0..38)
    col_start, col_end = 0, 38
    if df_raw.shape[1] <= col_end:
        raise ValueError("La hoja Cecos no llega hasta la columna AM.")

    # Cortar hasta última fila con texto en columna AM
    last_row = last_nonempty_row_in_col(df_raw, col_end, header_row)
    if last_row is None or last_row <= header_row:
        raise ValueError("No pude determinar el final de la tabla usando la columna AM.")

    # Encabezados desde A..AM en la fila donde está Status
    headers = df_raw.iloc[header_row, col_start:col_end + 1].tolist()
    headers = [normalize_str(h) if normalize_str(h) else f"COL_{i+1}" for i, h in enumerate(headers)]

    # Datos debajo del header
    data = df_raw.iloc[header_row + 1:last_row + 1, col_start:col_end + 1].copy()
    data.columns = headers

    # Detectar columnas por nombre (para filtrar sin depender de letras)
    cols_lower = {c.strip().lower(): c for c in data.columns}

    if "status" not in cols_lower:
        raise ValueError("Armé la tabla pero no existe una columna llamada 'Status' en los encabezados.")
    col_status = cols_lower["status"]

    # Buscar algo que contenga "concepto" y "tienda"
    col_concepto = None
    for k in cols_lower.keys():
        if "concepto" in k and "tienda" in k:
            col_concepto = cols_lower[k]
            break
    if col_concepto is None:
        raise ValueError("No encontré la columna 'Concepto tienda' (revisa el encabezado en el archivo Cecos).")

    # Tienda = columna C (index 2)
    col_tienda = data.columns[2]

    # Filtros
    df_f = data.copy()
    df_f[col_status] = df_f[col_status].astype(str).str.strip()
    df_f[col_concepto] = df_f[col_concepto].astype(str).str.strip()

    df_f = df_f[df_f[col_status].str.upper() == "ABIERTA"]
    df_f = df_f[df_f[col_concepto].str.upper() != "FRANQUICIA"]

    # Tienda a número
    tiendas = to_numeric_series(df_f[col_tienda]).dropna().astype(int).unique()
    return tiendas

# =========================================================
# Paso 3 - MD mes anterior
# =========================================================
def process_md_file(uploaded_file):
    df_md = load_excel_first_sheet(uploaded_file, header=0)

    if df_md.shape[1] < 12:
        raise ValueError("El archivo MD no tiene suficientes columnas para usar K y L.")

    # K index 10 = Ce. Coste, L index 11 = Centro de Coste
    ce_coste = df_md.iloc[:, 10].astype(str).str.strip()
    mask_102 = ce_coste.str.startswith("102", na=False)

    df_f = df_md.loc[mask_102].copy()
    centro_coste = df_f.iloc[:, 11].astype(str).str.strip()

    first4 = centro_coste.str.slice(0, 4)
    first4_num = pd.to_numeric(first4.where(first4.str.fullmatch(r"\d{4}", na=False)), errors="coerce")

    cedis = first4_num.dropna().astype(int).unique()
    return cedis

# =========================================================
# Paso 4 - Dash de tiendas
# =========================================================
def process_dash_tiendas(uploaded_file):
    df_dash = load_excel_first_sheet(uploaded_file, header=0)

    # Necesitamos hasta S (index 18)
    if df_dash.shape[1] < 19:
        raise ValueError("El Dash de tiendas no tiene suficientes columnas (necesito hasta la columna S).")

    # Por posición según tu regla:
    # B=CECO, C=Value Tienda, E=Región, G=Nombre tienda, R=DM, S=AM
    col_B = df_dash.columns[1]   # B
    col_C = df_dash.columns[2]   # C
    col_E = df_dash.columns[4]   # E
    col_G = df_dash.columns[6]   # G
    col_R = df_dash.columns[17]  # R
    col_S = df_dash.columns[18]  # S

    out = pd.DataFrame({
        "TIENDA": df_dash[col_C],
        "Ce.coste": df_dash[col_B],
        "Centro de coste": df_dash[col_G],
        "AM": df_dash[col_S],
        "DM": df_dash[col_R],
        "Región": df_dash[col_E],
    })

    out["TIENDA"] = to_numeric_series(out["TIENDA"]).astype("Int64")
    out = out.dropna(subset=["TIENDA"]).copy()
    out["TIENDA"] = out["TIENDA"].astype(int)
    out = out.drop_duplicates(subset=["TIENDA"], keep="first")

    # Orden final EXACTO para correo
    out = out[["Ce.coste", "Centro de coste", "TIENDA", "AM", "DM", "Región"]]
    return out

# =========================================================
# UI - Uploaders (Usuario final)
# =========================================================
st.info(
    "Sube los 4 archivos. La app calcula qué tiendas y CEDIS faltan en Interfaces y genera un Excel con la hoja "
    "**Faltantes_Total** lista para pegar en un correo."
)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.subheader("1) Interfaces")
    f_interfaces = st.file_uploader("Consolidado de Interfaces", type=["xlsx", "xls"], key="interfaces")

with c2:
    st.subheader("2) Cecos")
    f_cecos = st.file_uploader('Archivo Cecos (hoja "JMC Cost Center Strucutre")', type=["xlsx", "xls"], key="cecos")

with c3:
    st.subheader("3) MD mes anterior")
    f_md = st.file_uploader("MD mes anterior", type=["xlsx", "xls"], key="md")

with c4:
    st.subheader("4) Dash de tiendas")
    f_dash = st.file_uploader("Dash de tiendas", type=["xlsx", "xls"], key="dash")

st.divider()

# =========================================================
# Procesamiento (se ejecuta dentro de la app)
# =========================================================
interfaces_cecos = None
tiendas_cecos = None
cedis_md = None
dash_lookup = None

errors = []

if f_interfaces:
    try:
        df_int = load_excel_first_sheet(f_interfaces, header=0)
        interfaces_cecos = extract_cecos_from_interfaces(df_int).unique()
        st.success(f"Interfaces cargadas: {len(interfaces_cecos)} Cecos únicos.")
    except Exception as e:
        errors.append(f"Interfaces: {e}")

if f_cecos:
    try:
        tiendas_cecos = process_cecos_file(f_cecos)
        # Paso 2 interno: no mostramos tablas, solo confirmación mínima
        st.success(f"Cecos procesado: {len(tiendas_cecos)} tiendas activas (filtradas).")
    except Exception as e:
        errors.append(f"Cecos: {e}")

if f_md:
    try:
        cedis_md = process_md_file(f_md)
        st.success(f"MD procesado: {len(cedis_md)} CEDIS únicos.")
    except Exception as e:
        errors.append(f"MD: {e}")

if f_dash:
    try:
        dash_lookup = process_dash_tiendas(f_dash)
        st.success(f"Dash procesado: {len(dash_lookup)} tiendas disponibles para enriquecer el reporte.")
    except Exception as e:
        errors.append(f"Dash: {e}")

if errors:
    st.error("Se detectaron errores:\n- " + "\n- ".join(errors))

st.divider()
st.subheader("Resultado: faltantes en Interfaces")

if interfaces_cecos is None:
    st.warning("Primero carga el consolidado de Interfaces.")
else:
    set_interfaces = set(map(int, interfaces_cecos))

    union_source = set()
    if tiendas_cecos is not None:
        union_source |= set(map(int, tiendas_cecos))
    if cedis_md is not None:
        union_source |= set(map(int, cedis_md))

    if not union_source:
        st.info("Carga Cecos (Paso 2) y/o MD (Paso 3) para calcular faltantes.")
    else:
        missing_union = sorted(list(union_source - set_interfaces))

        st.metric("TOTAL faltantes (Tiendas ∪ CEDIS) vs Interfaces", len(missing_union))

        # Base de faltantes (columna TIENDA contiene tanto tiendas como cedis)
        df_total_base = pd.DataFrame({"TIENDA": missing_union})
        df_total_base["TIENDA"] = pd.to_numeric(df_total_base["TIENDA"], errors="coerce").astype("Int64")
        df_total_base = df_total_base.dropna(subset=["TIENDA"]).copy()
        df_total_base["TIENDA"] = df_total_base["TIENDA"].astype(int)

        # Enriquecer con Dash si existe (si no, deja columnas vacías)
        if dash_lookup is not None:
            df_faltantes_total = df_total_base.merge(dash_lookup, on="TIENDA", how="left")
        else:
            df_faltantes_total = df_total_base.copy()
            df_faltantes_total["Ce.coste"] = ""
            df_faltantes_total["Centro de coste"] = ""
            df_faltantes_total["AM"] = ""
            df_faltantes_total["DM"] = ""
            df_faltantes_total["Región"] = ""

        # Orden final EXACTO para correo
        df_faltantes_total = df_faltantes_total[["Ce.coste", "Centro de coste", "TIENDA", "AM", "DM", "Región"]]

        st.write("**Tabla para correo (Faltantes_Total):**")
        st.dataframe(df_faltantes_total, use_container_width=True)

        # Descargar Excel con hojas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"Cecos_interfaces": sorted(list(set_interfaces))}).to_excel(
                writer, index=False, sheet_name="Interfaces_Cecos"
            )

            if tiendas_cecos is not None:
                pd.DataFrame({"Tiendas_Cecos_filtrado": sorted(list(set(map(int, tiendas_cecos))))}).to_excel(
                    writer, index=False, sheet_name="Tiendas_Cecos"
                )

            if cedis_md is not None:
                pd.DataFrame({"CEDIS_MD": sorted(list(set(map(int, cedis_md))))}).to_excel(
                    writer, index=False, sheet_name="CEDIS_MD"
                )

            df_faltantes_total.to_excel(writer, index=False, sheet_name="Faltantes_Total")
            pd.DataFrame({"faltante_unio_
