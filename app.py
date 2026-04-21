import io
import hashlib
from io import BytesIO
from typing import Dict

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Comparador México vs Master",
    page_icon="📦",
    layout="wide",
)

# =========================
# HEADER
# =========================
col_title, col_logo = st.columns([6, 2], vertical_alignment="center")

with col_title:
    st.title("Comparador Planilla Externa vs Master")

with col_logo:
    st.markdown("""
    <div style="
        text-align: right;
        font-size: 32px;
        font-weight: 700;
        letter-spacing: 3px;
        color: #ec4899;
    ">
        MK
    </div>
    """, unsafe_allow_html=True)

# =========================
# ESTILOS
# =========================
st.markdown("""
<style>
.stApp {
    background-color: #0b1020;
}

/* Botones */
div.stButton > button,
div.stDownloadButton > button,
div[data-testid="stFormSubmitButton"] > button {
    background-color: #1f2937;
    color: white;
    border-radius: 12px;
    padding: 12px 16px;
    border: 1px solid #374151;
    font-weight: 600;
    width: 100%;
    min-height: 52px;
}

div.stButton > button:hover,
div.stDownloadButton > button:hover,
div[data-testid="stFormSubmitButton"] > button:hover {
    background-color: #374151;
    color: white;
    border: 1px solid #4b5563;
}

div.stButton > button:focus:not(:active),
div.stDownloadButton > button:focus:not(:active),
div[data-testid="stFormSubmitButton"] > button:focus:not(:active) {
    color: white !important;
    border: 1px solid #60a5fa !important;
    box-shadow: 0 0 0 0.2rem rgba(96, 165, 250, 0.25) !important;
}

/* Selectbox */
div[data-baseweb="select"] > div {
    background-color: #1f2937 !important;
    color: white !important;
    border-radius: 12px;
}

div[data-baseweb="select"] * {
    color: white !important;
}

/* Inputs numéricos */
div[data-testid="stNumberInput"] input {
    background-color: #111827 !important;
    color: white !important;
    border-radius: 10px !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background-color: rgba(255,255,255,0.02);
    border: 1px solid #374151;
    border-radius: 16px;
    padding: 10px;
}

/* Checkbox */
[data-testid="stCheckbox"] label {
    color: white !important;
}

/* Expander */
details {
    background-color: #111827 !important;
    border: 1px solid #374151 !important;
    border-radius: 14px !important;
    overflow: hidden !important;
    margin-bottom: 4px !important;
}

details summary {
    background-color: #1f2937 !important;
    color: white !important;
    padding: 14px 16px !important;
    font-weight: 700 !important;
    cursor: pointer !important;
}

details > div {
    padding: 16px !important;
}

/* Cards */
.metric-card {
    padding: 24px;
    border-radius: 24px;
    color: white;
    height: 220px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.25);
    position: relative;
    overflow: hidden;
    margin-bottom: 8px;
}

.metric-card::before {
    content: "";
    position: absolute;
    width: 180px;
    height: 180px;
    border-radius: 50%;
    top: -50px;
    right: -20px;
    background: rgba(255,255,255,0.10);
}

.metric-card::after {
    content: "";
    position: absolute;
    width: 120px;
    height: 120px;
    border-radius: 50%;
    bottom: -45px;
    left: -25px;
    background: rgba(255,255,255,0.08);
}

.metric-title {
    font-size: 20px;
    font-weight: 700;
    opacity: 0.95;
    position: relative;
    z-index: 1;
}

.metric-value {
    font-size: 68px;
    font-weight: 800;
    line-height: 1;
    margin: 10px 0;
    position: relative;
    z-index: 1;
}

.metric-sub {
    font-size: 16px;
    opacity: 0.92;
    position: relative;
    z-index: 1;
}

.card-gray {
    background: linear-gradient(135deg, #4b5563, #3b475d);
}

.card-green {
    background: linear-gradient(135deg, #15803d, #22c55e);
}

.card-orange {
    background: linear-gradient(135deg, #d97706, #f59e0b);
}

.card-red {
    background: linear-gradient(135deg, #dc2626, #ef4444);
}

[data-testid="stDataFrame"] {
    border-radius: 16px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)

# =========================
# CONFIG
# =========================
COLUMNAS_CANONICAS = ["imlitm", "imdsc1", "Width", "Depth", "Height", "Cubicate", "PESO_BRANCH"]
COLUMNAS_NUMERICAS = ["Width", "Depth", "Height", "Cubicate", "PESO_BRANCH"]

ALIASES = {
    "imlitm": ["imlitm", "codigo", "código"],
    "imdsc1": ["imdsc1", "descripcion", "descripción"],
    "Width": ["width", "widht", "ancho"],
    "Depth": ["depth", "profundidad"],
    "Height": ["height", "alto"],
    "Cubicate": ["cubicate", "cubicaje"],
    "PESO_BRANCH": ["peso_branch", "peso - kg", "peso_kg", "peso"],
}

if "resultado_comparacion" not in st.session_state:
    st.session_state.resultado_comparacion = None

if "filtro_resultado" not in st.session_state:
    st.session_state.filtro_resultado = "Todos"

# =========================
# FUNCIONES
# =========================
def limpiar_texto(texto):
    return str(texto).strip().replace("\n", " ").replace("\r", " ")


def normalizar(nombre):
    return limpiar_texto(nombre).lower().replace("_", " ").replace("-", " ")


def detectar_header(df_raw):
    max_filas = min(20, len(df_raw))
    for i in range(max_filas):
        fila = [normalizar(x) for x in df_raw.iloc[i].tolist()]
        if any(("codigo" in f) or ("código" in f) or ("imlitm" in f) for f in fila):
            return i
    return 0


def mapear_columnas(df):
    cols_norm = {col: normalizar(col) for col in df.columns}
    ren = {}
    ya_asignadas = set()

    for canon, alias_list in ALIASES.items():
        aliases_norm = [normalizar(a) for a in alias_list]

        for col, norm in cols_norm.items():
            if norm in aliases_norm and canon not in ya_asignadas:
                ren[col] = canon
                ya_asignadas.add(canon)
                break

    return df.rename(columns=ren)


def leer_archivo(file_obj):
    nombre = file_obj.name.lower()

    if nombre.endswith(".csv"):
        df = pd.read_csv(file_obj)
    else:
        raw = pd.read_excel(file_obj, header=None)
        header = detectar_header(raw)
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=header)

    df.columns = [limpiar_texto(c) for c in df.columns]
    df = mapear_columnas(df)
    return df


def normalizar_imlitm(codigo: str) -> str:
    codigo = str(codigo).strip()

    if codigo.lower() == "nan":
        return ""

    if codigo.endswith(".0"):
        codigo = codigo[:-2]

    codigo = "".join(ch for ch in codigo if ch.isdigit())

    if codigo == "":
        return ""

    if len(codigo) == 7 and codigo.startswith(("6", "8")):
        codigo = "0" + codigo

    return codigo


def preparar(df):
    df = df.copy()
    df.columns = [limpiar_texto(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()].copy()

    faltantes = [c for c in COLUMNAS_CANONICAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            f"Faltan columnas requeridas: {', '.join(faltantes)}. "
            f"Columnas detectadas: {', '.join(map(str, df.columns.tolist()))}"
        )

    df = df[COLUMNAS_CANONICAS].copy()

    df["imlitm"] = df["imlitm"].astype(str).apply(normalizar_imlitm)
    df["imdsc1"] = df["imdsc1"].astype(str).str.strip()

    for col in COLUMNAS_NUMERICAS:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .str.replace(",", ".", regex=False)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df[df["imlitm"] != ""]
    df = df[df["imlitm"].str.lower() != "nan"]

    return df


@st.cache_data(show_spinner=False)
def leer_y_preparar(file_bytes, name):
    file_obj = BytesIO(file_bytes)
    file_obj.name = name
    df = leer_archivo(file_obj)
    return preparar(df)


def comparar(row, tol: Dict[str, float]):
    if row["_merge"] == "left_only":
        return "No existe en master", ""

    dif = []

    desc_mex = str(row.get("imdsc1_mexico", "")).strip()
    desc_master = str(row.get("imdsc1_master", "")).strip()
    if desc_mex != desc_master:
        dif.append("Descripción")

    for col in COLUMNAS_NUMERICAS:
        a = row.get(f"{col}_mexico")
        b = row.get(f"{col}_master")

        if pd.isna(a) != pd.isna(b):
            dif.append(col)
            continue

        if pd.notna(a) and pd.notna(b) and abs(float(a) - float(b)) > tol[col]:
            dif.append(col)

    if not dif:
        return "Coincide", ""
    return "Con diferencias", ", ".join(dif)


@st.cache_data(show_spinner=False)
def ejecutar(df_mex, df_master, tol_items, hash_mex, hash_master):
    _ = hash_mex, hash_master
    tol = dict(tol_items)

    m1 = df_mex.rename(columns={c: f"{c}_mexico" for c in df_mex.columns if c != "imlitm"})
    m2 = df_master.rename(columns={c: f"{c}_master" for c in df_master.columns if c != "imlitm"})

    res = m1.merge(m2, how="left", on="imlitm", indicator=True)

    estados = res.apply(lambda r: comparar(r, tol), axis=1)
    res["resultado"] = [x[0] for x in estados]
    res["detalle"] = [x[1] for x in estados]

    columnas_ordenadas = [
        "imlitm",
        "resultado",
        "detalle",
        "imdsc1_mexico",
        "imdsc1_master",
        "Width_mexico",
        "Width_master",
        "Depth_mexico",
        "Depth_master",
        "Height_mexico",
        "Height_master",
        "Cubicate_mexico",
        "Cubicate_master",
        "PESO_BRANCH_mexico",
        "PESO_BRANCH_master",
    ]

    return res[[c for c in columnas_ordenadas if c in res.columns]]


def exportar_excel(data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre, df in data.items():
            df.to_excel(writer, sheet_name=nombre[:30], index=False)
    return output.getvalue()


def hash_archivo(b):
    return hashlib.md5(b).hexdigest()


# =========================
# FORM PRINCIPAL
# =========================
with st.form("form"):

    st.markdown("### Carga de archivos")
    mex = st.file_uploader("Archivo México", type=["xlsx", "xls", "csv"])
    master = st.file_uploader("Archivo Master", type=["xlsx", "xls", "csv"])

    ctrl1, ctrl2 = st.columns(2)

    with ctrl1:
        with st.expander("TOLERANCIAS", expanded=False):
            t1, t2 = st.columns(2)
            t3, t4, t5 = st.columns(3)

            with t1:
                tol_peso = st.number_input(
                    "Peso",
                    min_value=0.0,
                    max_value=100.0,
                    value=0.001,
                    step=0.001,
                    key="tol_peso"
                )

            with t2:
                tol_cub = st.number_input(
                    "Cubicate",
                    min_value=0.0,
                    max_value=10000.0,
                    value=0.5,
                    step=0.1,
                    key="tol_cub"
                )

            with t3:
                tol_width = st.number_input(
                    "Width",
                    min_value=0.0,
                    max_value=1000.0,
                    value=0.1,
                    step=0.1,
                    key="tol_width"
                )

            with t4:
                tol_depth = st.number_input(
                    "Depth",
                    min_value=0.0,
                    max_value=1000.0,
                    value=0.1,
                    step=0.1,
                    key="tol_depth"
                )

            with t5:
                tol_height = st.number_input(
                    "Height",
                    min_value=0.0,
                    max_value=1000.0,
                    value=0.1,
                    step=0.1,
                    key="tol_height"
                )

    with ctrl2:
        with st.expander("COLUMNAS A VISUALIZAR", expanded=False):
            v1, v2 = st.columns(2)
            v3, v4 = st.columns(2)
            v5, v6 = st.columns(2)

            with v1:
                mostrar_descripcion = st.checkbox("Descripción", value=False, key="ver_desc")
            with v2:
                mostrar_width = st.checkbox("Width", value=False, key="ver_width")
            with v3:
                mostrar_depth = st.checkbox("Depth", value=False, key="ver_depth")
            with v4:
                mostrar_height = st.checkbox("Height", value=False, key="ver_height")
            with v5:
                mostrar_cubicate = st.checkbox("Cubicate", value=False, key="ver_cub")
            with v6:
                mostrar_peso = st.checkbox("Peso", value=False, key="ver_peso")

    tol = {
        "PESO_BRANCH": tol_peso,
        "Cubicate": tol_cub,
        "Width": tol_width,
        "Depth": tol_depth,
        "Height": tol_height,
    }

    columnas_mostrar = {
        "Descripción": mostrar_descripcion,
        "Width": mostrar_width,
        "Depth": mostrar_depth,
        "Height": mostrar_height,
        "Cubicate": mostrar_cubicate,
        "PESO_BRANCH": mostrar_peso,
    }

    run = st.form_submit_button("Comparar")

# =========================
# PROCESO
# =========================
if run:
    if mex and master:
        try:
            mex_b = mex.getvalue()
            master_b = master.getvalue()

            df_mex = leer_y_preparar(mex_b, mex.name)
            df_master = leer_y_preparar(master_b, master.name)

            res = ejecutar(
                df_mex,
                df_master,
                tuple(tol.items()),
                hash_archivo(mex_b),
                hash_archivo(master_b),
            )

            st.session_state.resultado_comparacion = res
            st.session_state.filtro_resultado = "Todos"
            st.session_state.columnas_mostrar = columnas_mostrar.copy()

        except Exception as e:
            st.error(f"Error al procesar archivos: {e}")
    else:
        st.error("Tenés que cargar ambos archivos.")

# Mantener columnas visibles después del submit
columnas_mostrar_actual = st.session_state.get("columnas_mostrar", {
    "Descripción": False,
    "Width": False,
    "Depth": False,
    "Height": False,
    "Cubicate": False,
    "PESO_BRANCH": False,
})

# =========================
# RESULTADOS
# =========================
if st.session_state.resultado_comparacion is not None:
    res = st.session_state.resultado_comparacion.copy()

    total = len(res)
    coinciden = len(res[res["resultado"] == "Coincide"])
    dif = len(res[res["resultado"] == "Con diferencias"])
    no_master = len(res[res["resultado"] == "No existe en master"])

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class='metric-card card-gray'>
            <div class='metric-title'>Total analizados</div>
            <div class='metric-value'>{total}</div>
            <div class='metric-sub'>Registros comparados desde México</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class='metric-card card-green'>
            <div class='metric-title'>Coinciden</div>
            <div class='metric-value'>{coinciden}</div>
            <div class='metric-sub'>Sin diferencias detectadas</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class='metric-card card-orange'>
            <div class='metric-title'>Con diferencias</div>
            <div class='metric-value'>{dif}</div>
            <div class='metric-sub'>Cambios en descripción o valores</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class='metric-card card-red'>
            <div class='metric-title'>No existe en master</div>
            <div class='metric-value'>{no_master}</div>
            <div class='metric-sub'>Código presente en México, ausente en Master</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("###")

    b1, b2 = st.columns(2)

    with b1:
        if st.button("🧹 Limpiar filtro"):
            st.session_state.filtro_resultado = "Todos"
            st.rerun()

    with b2:
        if st.button("🔄 Nueva comparación"):
            st.session_state.resultado_comparacion = None
            st.session_state.filtro_resultado = "Todos"
            st.session_state.columnas_mostrar = {
                "Descripción": False,
                "Width": False,
                "Depth": False,
                "Height": False,
                "Cubicate": False,
                "PESO_BRANCH": False,
            }
            st.cache_data.clear()
            st.rerun()

    filtro = st.selectbox(
        "Filtrar",
        ["Todos", "Coincide", "Con diferencias", "No existe en master"],
        key="filtro_resultado"
    )

    if filtro != "Todos":
        vista = res[res["resultado"] == filtro]
    else:
        vista = res

    columnas_visibles = ["imlitm", "resultado", "detalle"]

    if columnas_mostrar_actual.get("Descripción", False):
        columnas_visibles += ["imdsc1_mexico", "imdsc1_master"]

    if columnas_mostrar_actual.get("Width", False):
        columnas_visibles += ["Width_mexico", "Width_master"]

    if columnas_mostrar_actual.get("Depth", False):
        columnas_visibles += ["Depth_mexico", "Depth_master"]

    if columnas_mostrar_actual.get("Height", False):
        columnas_visibles += ["Height_mexico", "Height_master"]

    if columnas_mostrar_actual.get("Cubicate", False):
        columnas_visibles += ["Cubicate_mexico", "Cubicate_master"]

    if columnas_mostrar_actual.get("PESO_BRANCH", False):
        columnas_visibles += ["PESO_BRANCH_mexico", "PESO_BRANCH_master"]

    vista_mostrar = vista[[c for c in columnas_visibles if c in vista.columns]]

    st.dataframe(vista_mostrar, use_container_width=True)

    st.download_button(
        "📥 Descargar Excel",
        exportar_excel({
            "resultado": res,
            "coinciden": res[res["resultado"] == "Coincide"],
            "diferencias": res[res["resultado"] == "Con diferencias"],
            "no_existe_master": res[res["resultado"] == "No existe en master"],
        }),
        "resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# =========================
# FOOTER
# =========================
st.markdown("""
<hr style="margin-top:40px; border: 0.5px solid #374151;">

<div style='text-align:center; color:#9ca3af; font-size:14px;'>
    Desarrollado por <b>Krana&#123;dev&#125;</b> © 2026<br>
    Todos los derechos reservados
</div>
""", unsafe_allow_html=True)