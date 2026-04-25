import io
import hashlib
from io import BytesIO
from typing import Dict, Tuple, List

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Comparador Planilla Externa vs Master",
    page_icon="📦",
    layout="wide",
)

# =========================================================
# CONFIG BASE
# =========================================================
APP_TITLE = "Comparador Planilla Externa vs Master"
APP_SUBTITLE = (
    "Compara una planilla externa contra una base master, detecta diferencias "
    "en descripción, dimensiones, cubicaje y peso, y exporta los resultados."
)

COLUMNAS_CANONICAS = [
    "imlitm",
    "imdsc1",
    "Width",
    "Depth",
    "Height",
    "Cubicate",
    "PESO_BRANCH",
]

COLUMNAS_NUMERICAS = ["Width", "Depth", "Height", "Cubicate", "PESO_BRANCH"]

ALIASES = {
    "imlitm": ["imlitm", "codigo", "código", "code", "item", "sku"],
    "imdsc1": ["imdsc1", "descripcion", "descripción", "description"],
    "Width": ["width", "widht", "ancho"],
    "Depth": ["depth", "profundidad"],
    "Height": ["height", "alto"],
    "Cubicate": ["cubicate", "cubicaje", "cubication"],
    "PESO_BRANCH": ["peso_branch", "peso - kg", "peso_kg", "peso", "weight"],
}

if "resultado_comparacion" not in st.session_state:
    st.session_state.resultado_comparacion = None

if "filtro_resultado" not in st.session_state:
    st.session_state.filtro_resultado = "Todos"

if "columnas_mostrar" not in st.session_state:
    st.session_state.columnas_mostrar = {
        "Descripción": False,
        "Width": False,
        "Depth": False,
        "Height": False,
        "Cubicate": False,
        "PESO_BRANCH": False,
    }

if "ultima_ejecucion_ok" not in st.session_state:
    st.session_state.ultima_ejecucion_ok = False


# =========================================================
# ESTILOS
# =========================================================
st.markdown(
    """
<style>

:root {
    --bg: #071122;
    --panel: #0d1729;
    --panel-2: #111d33;
    --border: rgba(148, 163, 184, 0.18);
    --text: #f8fafc;
    --muted: #94a3b8;
    --primary: #60a5fa;
    --primary-2: #2563eb;
    --green: #22c55e;
    --orange: #f59e0b;
    --red: #ef4444;
    --gray: #64748b;
    --pink: #ec4899;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(59,130,246,0.08), transparent 25%),
        radial-gradient(circle at top right, rgba(236,72,153,0.08), transparent 25%),
        linear-gradient(180deg, #050b16 0%, #08111f 100%);
    color: var(--text);
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}

/* Header */
.app-hero {
    background: linear-gradient(135deg, rgba(15,23,42,0.95), rgba(8,17,31,0.92));
    border: 1px solid var(--border);
    border-radius: 28px;
    padding: 28px 30px;
    box-shadow: 0 20px 50px rgba(0, 0, 0, 0.22);
    animation: fadeUp 0.5s ease;
}

.app-kicker {
    display: inline-block;
    font-size: 12px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    padding: 8px 12px;
    border-radius: 999px;
    border: 1px solid rgba(96,165,250,0.25);
    background: rgba(96,165,250,0.08);
    color: #bfdbfe;
    margin-bottom: 14px;
    font-weight: 700;
}

.app-title-row {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 20px;
}

.app-title {
    font-size: 48px;
    font-weight: 800;
    line-height: 1.05;
    margin: 0;
}

.app-subtitle {
    margin-top: 12px;
    color: var(--muted);
    font-size: 16px;
    max-width: 820px;
}

.app-brand {
    font-size: 42px;
    font-weight: 800;
    letter-spacing: 3px;
    color: var(--pink);
    line-height: 1;
    opacity: 0.95;
}
.st-emotion-cache-gquqoo {
    display:none;
}

/* Form panel */


.section-title {
    font-size: 22px;
    font-weight: 800;
    color: white;
    margin-bottom: 16px;
}

.helper-text {
    color: var(--muted);
    font-size: 14px;
    margin-top: -6px;
    margin-bottom: 14px;
}

/* Buttons */

div.stButton > button,
div.stDownloadButton > button,
div[data-testid="stFormSubmitButton"] > button {
    background: linear-gradient(135deg, var(--primary-2), var(--primary));
    color: white;
    border-radius: 14px;
    padding: 12px 18px;
    border: none;
    font-weight: 700;
    width: 100%;
    min-height: 52px;
    box-shadow: 0 12px 28px rgba(37, 99, 235, 0.24);
    transition: transform 0.18s ease, box-shadow 0.18s ease, opacity 0.18s ease;
}

div.stButton > button:hover,
div.stDownloadButton > button:hover,
div[data-testid="stFormSubmitButton"] > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 16px 30px rgba(37, 99, 235, 0.28);
    color: white;
}

div.stButton > button:disabled,
div.stDownloadButton > button:disabled,
div[data-testid="stFormSubmitButton"] > button:disabled {
    opacity: 0.55;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
}

/* Secondary buttons */
.secondary-button div.stButton > button {
    background: #172338 !important;
    border: 1px solid var(--border) !important;
    box-shadow: none !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: var(--panel-2);
    border: 1px dashed var(--border);
    border-radius: 18px;
    padding: 14px;
    transition: all 0.2s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: rgba(96,165,250,0.6);
    box-shadow: 0 0 0 1px rgba(96,165,250,0.15) inset;
}

[data-testid="stFileUploader"] label {
    display: block;
    margin-bottom: 10px;
}

[data-testid="stFileUploader"] section {
    min-height: 78px;
    display: flex;
    align-items: center;
    padding: 0 14px;
}

[data-testid="stFileUploader"] section button {
    margin: 0;
    align-self: center;
}

[data-testid="stFileUploader"] small {
    align-self: center;
}

/* Inputs */
div[data-testid="stNumberInput"] input {
    background-color: #0f172a !important;
    color: white !important;
    border-radius: 10px !important;
}

div[data-baseweb="select"] > div {
    background-color: #111827 !important;
    color: white !important;
    border-radius: 12px;
    border: 1px solid rgba(148,163,184,0.14);
}

div[data-baseweb="select"] * {
    color: white !important;
}

/* Expander */
details {
    background: rgba(17,24,39,0.72) !important;
    border: 1px solid var(--border) !important;
    border-radius: 16px !important;
    overflow: hidden !important;
    margin-bottom: 8px !important;
}

details summary {
    background: linear-gradient(180deg, rgba(30,41,59,0.9), rgba(30,41,59,0.75)) !important;
    color: white !important;
    padding: 14px 16px !important;
    font-weight: 700 !important;
    cursor: pointer !important;
}

details > div {
    padding: 16px !important;
}

/* Metric cards */
.metric-card {
    padding: 28px 24px;
    border-radius: 26px;
    color: white;

    height: 260px;              /* 🔥 más alto */
    
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;

    text-align: center;

    position: relative;
    overflow: hidden;

    gap: 12px;                  /* 🔥 más aire entre elementos */
}

.metric-title {
    font-size: 20px;
    font-weight: 700;
    opacity: 0.95;
}

.metric-value {
    font-size: 72px;           /* 🔥 más protagonista */
    font-weight: 800;
    line-height: 1;
}

.metric-sub {
    font-size: 15px;
    opacity: 0.9;
    max-width: 75%;            /* 🔥 evita líneas largas */
    line-height: 1.4;
}

.metric-pct {
    margin-top: 10px;
    padding: 8px 14px;
    border-radius: 999px;
    background: rgba(255,255,255,0.18);
    font-size: 13px;
    font-weight: 700;
}

.metric-card::before {
    content: "";
    position: absolute;
    width: 180px;
    height: 180px;
    background: rgba(255,255,255,0.08);
    border-radius: 50%;
    top: -40px;
    right: -40px;
}

.metric-card::after {
    content: "";
    position: absolute;
    width: 125px;
    height: 125px;
    border-radius: 50%;
    bottom: -50px;
    left: -24px;
    background: rgba(255,255,255,0.08);
}

.metric-title {
    font-size: 19px;
    font-weight: 700;
    opacity: 0.96;
    position: relative;
    z-index: 1;
}

.metric-value {
    font-size: 66px;
    font-weight: 800;
    line-height: 1;
    margin: 10px 0;
    position: relative;
    z-index: 1;
}

.metric-sub {
    font-size: 15px;
    opacity: 0.92;
    position: relative;
    z-index: 1;
}

.metric-pct {
    position: relative;
    z-index: 1;
    display: inline-block;
    padding: 7px 10px;
    border-radius: 999px;
    background: rgba(255,255,255,0.14);
    font-size: 13px;
    font-weight: 700;
    width: fit-content;
}

.card-gray {
    background: linear-gradient(135deg, #4b5563, #344256);
}
.card-green {
    background: linear-gradient(135deg, #16a34a, #22c55e);
}
.card-orange {
    background: linear-gradient(135deg, #d97706, #f59e0b);
}
.card-red {
    background: linear-gradient(135deg, #dc2626, #ef4444);
}

/* Summary block */
.summary-box {
    background: linear-gradient(180deg, rgba(13,23,41,0.96), rgba(13,23,41,0.84));
    border: 1px solid var(--border);
    border-radius: 22px;
    padding: 22px;
    margin-bottom: 18px;
    animation: fadeUp 0.6s ease;
}

.summary-title {
    font-size: 22px;
    font-weight: 800;
    margin-bottom: 12px;
}

.summary-text {
    color: #dbeafe;
    font-size: 15px;
    line-height: 1.65;
}

.progress-wrap {
    margin-top: 18px;
}

.progress-labels {
    display: flex;
    justify-content: space-between;
    color: var(--muted);
    font-size: 13px;
    margin-bottom: 8px;
}

.progress-bar-main {
    display: flex;
    width: 100%;
    height: 16px;
    overflow: hidden;
    border-radius: 999px;
    background: #0f172a;
    border: 1px solid rgba(148,163,184,0.12);
}

.seg-green { background: linear-gradient(90deg, #22c55e, #4ade80); }
.seg-orange { background: linear-gradient(90deg, #f59e0b, #fbbf24); }
.seg-red { background: linear-gradient(90deg, #ef4444, #f87171); }

/* Mini badges */
.badge-row {
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    margin-top: 16px;
}

.badge {
    border-radius: 999px;
    padding: 8px 12px;
    font-size: 13px;
    font-weight: 700;
    border: 1px solid transparent;
}

.badge-blue {
    background: rgba(96,165,250,0.1);
    color: #bfdbfe;
    border-color: rgba(96,165,250,0.2);
}

.badge-green {
    background: rgba(34,197,94,0.1);
    color: #bbf7d0;
    border-color: rgba(34,197,94,0.2);
}

.badge-orange {
    background: rgba(245,158,11,0.12);
    color: #fde68a;
    border-color: rgba(245,158,11,0.22);
}

.badge-red {
    background: rgba(239,68,68,0.12);
    color: #fecaca;
    border-color: rgba(239,68,68,0.22);
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 18px;
    overflow: hidden;
    border: 1px solid rgba(148,163,184,0.12);
}

/* Footer */
.footer-note {
    text-align: center;
    color: #94a3b8;
    font-size: 14px;
    margin-top: 28px;
    padding-top: 16px;
    border-top: 1px solid rgba(148,163,184,0.12);
}

/* Animations */
@keyframes fadeUp {
    from {
        opacity: 0;
        transform: translateY(12px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.toolbar-actions {
    margin-bottom: 12px;
}

/* 📱 MOBILE */
@media (max-width: 768px) {

    .app-title{
        text-align:center;
    }
    
    .section-title{
        text-align:center;
    }
    
    .helper-text{
        padding: 10px 0 0 10px;
    }
    .app-kicker {
        text-align: center;
    }
    
    .st-emotion-cache-epvm6 {
        margin-top:0.75rem;
        text-align: center;
    }
    
    .st-emotion-cache-1ivd0y6 .e3v525e3 {
        align-self: center;
    }
    
    .st-emotion-cache-wfksaw {
        align-items: center;
    }
    
    /* Cards */
    .metric-card {
        height: auto;
        padding: 20px;
        border-radius: 22px;
        width: 85%;
        margin: 0 auto 12px auto;
    }

    .metric-value {
        font-size: 48px;
    }

    .metric-title {
        font-size: 16px;
    }

    .metric-sub {
        font-size: 13px;
        max-width: 100%;
    }

    .metric-pct {
        font-size: 12px;
        padding: 6px 10px;
    }

    .toolbar-actions div[data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 10px !important;
        align-items: stretch !important;
    }

    .toolbar-actions div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        flex: 1 1 0 !important;
        min-width: 0 !important;
        width: 50% !important;
        max-width: 50% !important;
    }

    .toolbar-actions .stButton {
        width: 100% !important;
        margin-bottom: 0 !important;
    }

    .toolbar-actions .stButton > button {
        width: 100% !important;
        min-width: 0 !important;
        white-space: nowrap !important;
        font-size: 14px !important;
        min-height: 46px !important;
        padding: 10px 12px !important;
    }
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# HELPERS
# =========================================================
def limpiar_texto(texto) -> str:
    return str(texto).strip().replace("\n", " ").replace("\r", " ")


def normalizar(nombre) -> str:
    return limpiar_texto(nombre).lower().replace("_", " ").replace("-", " ")


def detectar_header(df_raw: pd.DataFrame) -> int:
    max_filas = min(20, len(df_raw))
    for i in range(max_filas):
        fila = [normalizar(x) for x in df_raw.iloc[i].tolist()]
        if any(("codigo" in f) or ("código" in f) or ("imlitm" in f) for f in fila):
            return i
    return 0


def mapear_columnas(df: pd.DataFrame) -> pd.DataFrame:
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


def leer_archivo(file_obj) -> pd.DataFrame:
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


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [limpiar_texto(c) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()].copy()

    faltantes = [c for c in COLUMNAS_CANONICAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas requeridas: "
            + ", ".join(faltantes)
            + ". Columnas detectadas: "
            + ", ".join(map(str, df.columns.tolist()))
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


def validar_archivo(file_obj, etiqueta: str) -> List[str]:
    mensajes = []

    if file_obj is None:
        mensajes.append(f"No se cargó el archivo {etiqueta}.")
        return mensajes

    nombre = file_obj.name.lower()
    extensiones_validas = (".xlsx", ".xls", ".csv")

    if not nombre.endswith(extensiones_validas):
        mensajes.append(
            f"El archivo {etiqueta} debe ser XLSX, XLS o CSV. Archivo recibido: {file_obj.name}"
        )

    if getattr(file_obj, "size", 0) == 0:
        mensajes.append(f"El archivo {etiqueta} está vacío.")

    return mensajes


def comparar(row, tol: Dict[str, float]) -> Tuple[str, str]:
    if row["_merge"] == "left_only":
        return "No existe en master", ""

    dif = []

    desc_ext = str(row.get("imdsc1_externa", "")).strip()
    desc_master = str(row.get("imdsc1_master", "")).strip()
    if desc_ext != desc_master:
        dif.append("Descripción")

    for col in COLUMNAS_NUMERICAS:
        a = row.get(f"{col}_externa")
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
def leer_y_preparar(file_bytes: bytes, name: str) -> pd.DataFrame:
    file_obj = BytesIO(file_bytes)
    file_obj.name = name
    df = leer_archivo(file_obj)
    return preparar(df)


@st.cache_data(show_spinner=False)
def ejecutar(
    df_ext: pd.DataFrame,
    df_master: pd.DataFrame,
    tol_items,
    hash_ext: str,
    hash_master: str,
) -> pd.DataFrame:
    _ = hash_ext, hash_master
    tol = dict(tol_items)

    m1 = df_ext.rename(columns={c: f"{c}_externa" for c in df_ext.columns if c != "imlitm"})
    m2 = df_master.rename(columns={c: f"{c}_master" for c in df_master.columns if c != "imlitm"})

    res = m1.merge(m2, how="left", on="imlitm", indicator=True)

    estados = res.apply(lambda r: comparar(r, tol), axis=1)
    res["resultado"] = [x[0] for x in estados]
    res["detalle"] = [x[1] for x in estados]

    columnas_ordenadas = [
        "imlitm",
        "resultado",
        "detalle",
        "imdsc1_externa",
        "imdsc1_master",
        "Width_externa",
        "Width_master",
        "Depth_externa",
        "Depth_master",
        "Height_externa",
        "Height_master",
        "Cubicate_externa",
        "Cubicate_master",
        "PESO_BRANCH_externa",
        "PESO_BRANCH_master",
    ]

    return res[[c for c in columnas_ordenadas if c in res.columns]]


def exportar_excel(data: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for nombre, df in data.items():
            df.to_excel(writer, sheet_name=nombre[:31], index=False)
    return output.getvalue()


def hash_archivo(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def porcentaje(parte: int, total: int) -> float:
    if total == 0:
        return 0.0
    return (parte / total) * 100


def construir_resumen(total: int, coinciden: int, dif: int, no_master: int) -> str:
    pct_ok = porcentaje(coinciden, total)
    pct_dif = porcentaje(dif, total)
    pct_missing = porcentaje(no_master, total)

    if total == 0:
        return "No hay registros para analizar."

    if coinciden == total:
        return (
            f"Resultado excelente: los {total} registros analizados coinciden con la base master "
            f"({pct_ok:.1f}% de coincidencia)."
        )

    if dif > 0 and no_master == 0:
        return (
            f"Se analizaron {total} registros. {dif} presentan diferencias "
            f"({pct_dif:.1f}%) y {coinciden} coinciden completamente."
        )

    if no_master > 0:
        return (
            f"Se analizaron {total} registros. {dif} presentan diferencias "
            f"({pct_dif:.1f}%) y {no_master} no existen en la base master "
            f"({pct_missing:.1f}%)."
        )

    return f"Se analizaron {total} registros. {coinciden} coinciden y {dif} presentan diferencias."


def render_metric_card(title: str, value: int, subtitle: str, pct_text: str, css_class: str):
    st.markdown(
        f"""
        <div class='metric-card {css_class}'>
            <div class='metric-title'>{title}</div>
            <div class='metric-value'>{value}</div>
            <div class='metric-sub'>{subtitle}</div>
            <div class='metric-pct'>{pct_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def limpiar_estado_comparacion():
    st.session_state.resultado_comparacion = None
    st.session_state.filtro_resultado = "Todos"
    st.session_state.ultima_ejecucion_ok = False


# =========================================================
# HEADER
# =========================================================
st.markdown(
    f"""
    <div class="app-hero">
        <div class="app-kicker">Data comparison · logistics workflow · streamlit app</div>
        <div class="app-title-row">
            <div>
                <h1 class="app-title">{APP_TITLE}</h1>
                <div class="app-subtitle">{APP_SUBTITLE}</div>
            </div>
             <!-- <div class="app-brand">MK</div> -->
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# PANEL PRINCIPAL
# =========================================================
st.markdown("<div class='panel'>", unsafe_allow_html=True)
st.markdown("<div class='section-title'>Carga de archivos</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='helper-text'>Subí la planilla externa y el archivo master. Luego configurá tolerancias y columnas visibles.</div>",
    unsafe_allow_html=True,
)

with st.form("form_comparacion"):
    col_up1, col_up2 = st.columns(2)

    with col_up1:
        externa = st.file_uploader(
            "Archivo externo",
            type=["xlsx", "xls", "csv"],
            help="Arrastrá o seleccioná la planilla externa a comparar.",
        )

    with col_up2:
        master = st.file_uploader(
            "Archivo Master",
            type=["xlsx", "xls", "csv"],
            help="Arrastrá o seleccioná el archivo master.",
        )

    ctrl1, ctrl2 = st.columns(2)

    with ctrl1:
        with st.expander("TOLERANCIAS", expanded=False):
            st.caption("Definí la tolerancia máxima permitida para considerar que dos valores coinciden.")
            t1, t2 = st.columns(2)
            t3, t4, t5 = st.columns(3)

            with t1:
                tol_peso = st.number_input(
                    "Peso",
                    min_value=0.0,
                    max_value=100.0,
                    value=0.001,
                    step=0.001,
                    key="tol_peso",
                )

            with t2:
                tol_cub = st.number_input(
                    "Cubicate",
                    min_value=0.0,
                    max_value=10000.0,
                    value=0.5,
                    step=0.1,
                    key="tol_cub",
                )

            with t3:
                tol_width = st.number_input(
                    "Width",
                    min_value=0.0,
                    max_value=1000.0,
                    value=0.1,
                    step=0.1,
                    key="tol_width",
                )

            with t4:
                tol_depth = st.number_input(
                    "Depth",
                    min_value=0.0,
                    max_value=1000.0,
                    value=0.1,
                    step=0.1,
                    key="tol_depth",
                )

            with t5:
                tol_height = st.number_input(
                    "Height",
                    min_value=0.0,
                    max_value=1000.0,
                    value=0.1,
                    step=0.1,
                    key="tol_height",
                )

    with ctrl2:
        with st.expander("COLUMNAS A VISUALIZAR", expanded=False):
            st.caption("Elegí qué columnas querés mostrar en la tabla final.")
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

    ready_to_compare = externa is not None and master is not None
    run = st.form_submit_button("Comparar")

st.markdown("</div>", unsafe_allow_html=True)

if not ready_to_compare:
    st.info("Cargá ambos archivos para habilitar la comparación.")

# =========================================================
# PROCESO DE COMPARACIÓN
# =========================================================
if run:
    errores = []
    errores.extend(validar_archivo(externa, "externo"))
    errores.extend(validar_archivo(master, "master"))

    if errores:
        for err in errores:
            st.error(err)
        limpiar_estado_comparacion()
    else:
        try:
            progress_placeholder = st.empty()
            status_placeholder = st.empty()

            with status_placeholder.container():
                st.info("Procesando archivos. Esto puede tardar unos segundos...")

            progress_bar = progress_placeholder.progress(0, text="Inicializando comparación...")

            externa_b = externa.getvalue()
            master_b = master.getvalue()

            progress_bar.progress(15, text="Leyendo planilla externa...")
            df_ext = leer_y_preparar(externa_b, externa.name)

            progress_bar.progress(45, text="Leyendo archivo master...")
            df_master = leer_y_preparar(master_b, master.name)

            if df_ext.empty:
                raise ValueError("La planilla externa no contiene registros válidos luego de la limpieza.")
            if df_master.empty:
                raise ValueError("El archivo master no contiene registros válidos luego de la limpieza.")

            progress_bar.progress(70, text="Comparando registros...")
            res = ejecutar(
                df_ext,
                df_master,
                tuple(tol.items()),
                hash_archivo(externa_b),
                hash_archivo(master_b),
            )

            progress_bar.progress(100, text="Comparación finalizada.")

            st.session_state.resultado_comparacion = res
            st.session_state.filtro_resultado = "Todos"
            st.session_state.columnas_mostrar = columnas_mostrar.copy()
            st.session_state.ultima_ejecucion_ok = True

            status_placeholder.empty()
            progress_placeholder.empty()
            st.success("Comparación finalizada correctamente.")

        except Exception as e:
            limpiar_estado_comparacion()
            st.error(f"Error al procesar archivos: {e}")


# =========================================================
# RESULTADOS
# =========================================================
if st.session_state.resultado_comparacion is not None:
    res = st.session_state.resultado_comparacion.copy()
    columnas_mostrar_actual = st.session_state.get("columnas_mostrar", {})

    total = len(res)
    coinciden = len(res[res["resultado"] == "Coincide"])
    dif = len(res[res["resultado"] == "Con diferencias"])
    no_master = len(res[res["resultado"] == "No existe en master"])

    pct_ok = porcentaje(coinciden, total)
    pct_dif = porcentaje(dif, total)
    pct_missing = porcentaje(no_master, total)

    # Resumen ejecutivo
    st.markdown(f"""
        <div class="summary-box">
        <div class="summary-title">Resumen ejecutivo</div>
        <div class="summary-text">{construir_resumen(total, coinciden, dif, no_master)}</div>

        <div class="badge-row">
        <div class="badge badge-blue">Total: {total}</div>
        <div class="badge badge-green">Coinciden: {coinciden} · {pct_ok:.1f}%</div>
        <div class="badge badge-orange">Con diferencias: {dif} · {pct_dif:.1f}%</div>
        <div class="badge badge-red">No existe en master: {no_master} · {pct_missing:.1f}%</div>
        </div>

        <div class="progress-wrap">
        <div class="progress-labels">
        <span>Distribución del resultado</span>
        <span>100%</span>
        </div>
        <div class="progress-bar-main">
        <div class="seg-green" style="width:{pct_ok:.2f}%"></div>
        <div class="seg-orange" style="width:{pct_dif:.2f}%"></div>
        <div class="seg-red" style="width:{pct_missing:.2f}%"></div>
        </div>
        </div>
        </div>
        """, unsafe_allow_html=True)


    c1, c2, c3, c4 = st.columns(4)
    with c1:
        render_metric_card(
            "Total analizados",
            total,
            "Registros comparados desde la planilla externa",
            "Base de análisis",
            "card-gray",
        )
    with c2:
        render_metric_card(
            "Coinciden",
            coinciden,
            "Sin diferencias detectadas",
            f"{pct_ok:.1f}% del total",
            "card-green",
        )
    with c3:
        render_metric_card(
            "Con diferencias",
            dif,
            "Cambios en descripción o valores",
            f"{pct_dif:.1f}% del total",
            "card-orange",
        )
    with c4:
        render_metric_card(
            "No existe en master",
            no_master,
            "Código presente en externa, ausente en master",
            f"{pct_missing:.1f}% del total",
            "card-red",
        )

    st.markdown('<div class="toolbar-actions">', unsafe_allow_html=True)

    left_btn, right_btn = st.columns([1, 1], gap="small")

    with left_btn:
        if st.button("🧹 Limpiar", key="btn_limpiar", use_container_width=True):
            st.session_state.filtro_resultado = "Todos"
            st.rerun()

    with right_btn:
        if st.button("🔄 Reiniciar", key="btn_reiniciar", use_container_width=True):
            limpiar_estado_comparacion()
            st.cache_data.clear()
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    filtro = st.selectbox(
        "Filtrar resultado",
        ["Todos", "Coincide", "Con diferencias", "No existe en master"],
        key="filtro_resultado",
    )

    if filtro != "Todos":
        vista = res[res["resultado"] == filtro].copy()
    else:
        vista = res.copy()

    columnas_visibles = ["imlitm", "resultado", "detalle"]

    if columnas_mostrar_actual.get("Descripción", False):
        columnas_visibles += ["imdsc1_externa", "imdsc1_master"]

    if columnas_mostrar_actual.get("Width", False):
        columnas_visibles += ["Width_externa", "Width_master"]

    if columnas_mostrar_actual.get("Depth", False):
        columnas_visibles += ["Depth_externa", "Depth_master"]

    if columnas_mostrar_actual.get("Height", False):
        columnas_visibles += ["Height_externa", "Height_master"]

    if columnas_mostrar_actual.get("Cubicate", False):
        columnas_visibles += ["Cubicate_externa", "Cubicate_master"]

    if columnas_mostrar_actual.get("PESO_BRANCH", False):
        columnas_visibles += ["PESO_BRANCH_externa", "PESO_BRANCH_master"]

    vista_mostrar = vista[[c for c in columnas_visibles if c in vista.columns]].copy()

    st.dataframe(
        vista_mostrar,
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "📥 Descargar Excel con resultados",
        exportar_excel(
            {
                "resultado": res,
                "coinciden": res[res["resultado"] == "Coincide"],
                "diferencias": res[res["resultado"] == "Con diferencias"],
                "no_existe_master": res[res["resultado"] == "No existe en master"],
            }
        ),
        "resultado_comparacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# =========================================================
# FOOTER
# =========================================================
st.markdown(
    """
    <div class="footer-note">
        Desarrollado por <b>Krana{dev}</b> © 2026 · Todos los derechos reservados
    </div>
    """,
    unsafe_allow_html=True,
)