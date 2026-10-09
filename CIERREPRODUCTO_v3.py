# CIERREPRODUCTO.py
# Dashboard profesional de Cierre de Productos
#
# Instalar dependencias:
# pip install streamlit pandas openpyxl plotly streamlit-plotly-events
#
# Ejecutar:
# streamlit run CIERREPRODUCTO.py

from __future__ import annotations

import io
import base64
import html
import os
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st




# ============================================================
# FAVICON CORPORATIVO DICASA S.A. — EMBEBIDO
# ============================================================

DICASA_FAVICON_BYTES = base64.b64decode(
    "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"
)
DICASA_FAVICON = io.BytesIO(DICASA_FAVICON_BYTES)


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Dashboard de Cierre de Productos",
    page_icon=DICASA_FAVICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "Dashboard de Cierre de Productos"
DEFAULT_FILE_NAME = "CIERRE PRODUCTOS.xlsx"

REQUIRED_COLUMNS = [
    "NUMERO DE FAMILIA",
    "FAMILIA",
    "#COD.",
    "DESCRIPCION",
    "EMPAQ",
    "#COMPRAS MES ACTUAL",
    "FECHA ULT COMPRA",
    "COSTO DE BODEGA",
    "COSTO GENERACION",
    "COSTO ACTUAL",
    "INVENTARIO TOTAL",
    "#VENTAS AÑO ACTUAL",
    "#VENTAS MES ANTERIOR",
    "#VENTAS MES ACTUAL",
    "#VENTAS PERDIDAS",
    "FECHA ULT VENTAS",
    "$VENTAS MES ACTUAL",
    "PRECIO PROMEDIO",
    "% MAGERN S/VENTA",
    "$UTILIDAD MES ACTUAL",
    "$GENERA S/VENTAS",
]

NUMERIC_COLUMNS = [
    "NUMERO DE FAMILIA",
    "#COMPRAS MES ACTUAL",
    "COSTO DE BODEGA",
    "COSTO GENERACION",
    "COSTO ACTUAL",
    "INVENTARIO TOTAL",
    "#VENTAS AÑO ACTUAL",
    "#VENTAS MES ANTERIOR",
    "#VENTAS MES ACTUAL",
    "#VENTAS PERDIDAS",
    "$VENTAS MES ACTUAL",
    "PRECIO PROMEDIO",
    "% MAGERN S/VENTA",
    "$UTILIDAD MES ACTUAL",
    "$GENERA S/VENTAS",
]

DATE_COLUMNS = [
    "FECHA ULT COMPRA",
    "FECHA ULT VENTAS",
]

TEXT_COLUMNS = [
    "FAMILIA",
    "#COD.",
    "DESCRIPCION",
    "EMPAQ",
]

METRIC_OPTIONS = [
    "$VENTAS MES ACTUAL",
    "$UTILIDAD MES ACTUAL",
    "INVENTARIO TOTAL",
    "#VENTAS MES ACTUAL",
    "#VENTAS PERDIDAS",
]


# ============================================================
# RENDER UI/UX EJECUTIVO PREMIUM — BEHANCE STYLE
# ============================================================

st.markdown(
    """
    <style>
    :root {
        --bg-0:#050B14;
        --bg-1:#07111D;
        --panel:#0E1B2B;
        --panel-2:#122136;
        --panel-3:#17283C;
        --border:#223B54;
        --text:#F6F8FB;
        --muted:#93A5B8;
        --gold:#D7AE58;
        --gold-soft:#F0D487;
        --emerald:#27C99B;
        --cyan:#43B9E6;
        --red:#E96573;
        --purple:#A984D8;
    }

    html, body, [class*="css"] {
        font-family:"Inter","SF Pro Display","Segoe UI Variable","Segoe UI",Arial,sans-serif !important;
    }

    .stApp {
        background:
            radial-gradient(circle at 78% -12%, rgba(39,201,155,.08), transparent 28%),
            radial-gradient(circle at 12% -8%, rgba(215,174,88,.055), transparent 23%),
            linear-gradient(180deg,#07111D 0%,#050A12 100%) !important;
        color:var(--text) !important;
    }

    .block-container {
        padding-top:.8rem !important;
        padding-bottom:3rem !important;
        max-width:1880px !important;
    }

    header[data-testid="stHeader"] { background:transparent !important; }
    #MainMenu, footer { visibility:hidden; }

    h1,h2,h3 {
        color:#FFFFFF !important;
        font-family:"Inter","Segoe UI Variable","Segoe UI",Arial,sans-serif !important;
        font-weight:780 !important;
        letter-spacing:-.025em !important;
    }

    h1 { font-size:1.95rem !important; }
    h2 { font-size:1.22rem !important; margin-top:.5rem !important; }
    h3 { font-size:.98rem !important; }
    p,label,.stCaption { color:#A5B4C4 !important; }

    .dash-header {
        display:flex;
        justify-content:space-between;
        align-items:center;
        gap:24px;
        margin:0 0 14px 0;
        padding:18px 20px;
        border:1px solid #1E344C;
        border-radius:16px;
        background:linear-gradient(120deg,rgba(16,32,50,.97),rgba(8,19,32,.98));
        box-shadow:0 14px 36px rgba(0,0,0,.20),inset 0 1px 0 rgba(255,255,255,.025);
    }

    .dash-header-left {
        display:flex;
        align-items:center;
        gap:13px;
        min-width:0;
    }

    .dash-logo {
        width:42px;
        height:42px;
        border-radius:11px;
        display:flex;
        align-items:center;
        justify-content:center;
        color:#07101C;
        font-size:1.25rem;
        font-weight:900;
        background:linear-gradient(145deg,#F1D17C,#B88632);
        box-shadow:0 7px 18px rgba(216,174,83,.16);
    }

    .dash-title {
        color:#FFFFFF;
        font-size:1.58rem;
        font-weight:820;
        letter-spacing:-.035em;
        line-height:1.05;
    }

    .dash-subtitle {
        color:#8094AA;
        font-size:.78rem;
        margin-top:5px;
        line-height:1.35;
    }

    .dash-badge {
        color:#E7CA7A;
        border:1px solid rgba(215,174,88,.32);
        background:rgba(215,174,88,.07);
        border-radius:999px;
        padding:7px 12px;
        font-size:.67rem;
        font-weight:780;
        letter-spacing:.05em;
        white-space:nowrap;
    }

    section[data-testid="stSidebar"] {
        background:linear-gradient(180deg,#0A1625 0%,#07101B 100%) !important;
        border-right:1px solid #1B3046 !important;
        box-shadow:14px 0 40px rgba(0,0,0,.14);
    }

    .sidebar-brand {
        padding:14px 13px;
        margin-bottom:12px;
        border:1px solid #20384F;
        border-radius:13px;
        background:linear-gradient(135deg,#112139,#0A1727);
    }

    .sidebar-brand-main {
        color:#F3D884;
        font-size:.92rem;
        font-weight:820;
        letter-spacing:.045em;
    }

    .sidebar-brand-sub {
        color:#6F8398;
        font-size:.62rem;
        margin-top:3px;
        letter-spacing:.08em;
        text-transform:uppercase;
    }

    div[data-baseweb="select"] > div,
    div[data-testid="stDateInput"] input,
    div[data-testid="stTextInput"] input,
    div[data-testid="stFileUploader"] section {
        background:#0F1E30 !important;
        border:1px solid #284159 !important;
        color:#EDF3F8 !important;
        border-radius:10px !important;
        box-shadow:none !important;
    }

    .exec-kpi,.leader-card {
        position:relative;
        overflow:hidden;
        background:linear-gradient(145deg,#16273C 0%,#101E30 100%) !important;
        border:1px solid #2A425B !important;
        border-radius:13px !important;
        padding:18px 18px 15px !important;
        height:158px !important;
        width:100% !important;
        box-sizing:border-box;
        box-shadow:0 10px 26px rgba(0,0,0,.17),inset 0 1px 0 rgba(255,255,255,.025);
        margin:0 !important;
        display:flex;
        flex-direction:column;
    }

    .exec-kpi::before,.leader-card::before {
        content:"";
        position:absolute;
        left:12px;
        right:12px;
        top:0;
        height:2px;
        border-radius:0 0 4px 4px;
        background:var(--accent);
    }

    .exec-kpi-label,.leader-label {
        color:#D5E1EC !important;
        font-size:16px !important;
        line-height:1.18 !important;
        font-weight:780 !important;
        text-transform:uppercase;
        letter-spacing:.045em !important;
        margin-bottom:8px !important;
    }

    .exec-kpi-value {
        color:var(--accent) !important;
        font-size:20px !important;
        line-height:1.06 !important;
        font-weight:860 !important;
        letter-spacing:-.025em !important;
        margin-bottom:7px !important;
        overflow-wrap:anywhere;
    }

    .exec-kpi-note {
        margin-top:auto;
        border-top:1px solid rgba(116, 170, 224, .45) !important;
        padding-top:8px !important;
        color:#EAF4FF !important;
        font-size:15px !important;
        font-weight:780 !important;
        line-height:1.20 !important;
        text-shadow:0 0 1px rgba(255,255,255,.08);
    }

    .leader-name {
        color:var(--accent) !important;
        font-size:18px !important;
        font-weight:840 !important;
        line-height:1.14 !important;
        margin-bottom:4px !important;
        overflow:hidden;
        display:-webkit-box;
        -webkit-line-clamp:2;
        -webkit-box-orient:vertical;
    }

    .leader-code { color:#F7FBFF !important; font-size:15px !important; font-weight:780 !important; line-height:1.22 !important; text-shadow:0 0 1px rgba(255,255,255,.08); }
    .leader-value { color:#FFFFFF !important; font-size:20px !important; font-weight:860 !important; line-height:1.18 !important; }

    div[data-baseweb="tab-list"] {
        gap:6px !important;
        padding:5px !important;
        border:1px solid #21384F !important;
        border-radius:12px !important;
        background:#091522 !important;
        margin:7px 0 14px !important;
    }

    button[data-baseweb="tab"] {
        min-height:37px !important;
        padding:6px 12px !important;
        border-radius:8px !important;
        color:#8599AE !important;
        background:transparent !important;
        border:1px solid transparent !important;
        font-size:.70rem !important;
        font-weight:720 !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color:#F7F9FB !important;
        background:linear-gradient(135deg,#17283D,#112035) !important;
        border-color:#35516D !important;
        box-shadow:inset 0 -2px 0 #D7AE58 !important;
    }

    div[data-testid="stPlotlyChart"] {
        background:linear-gradient(145deg,#0F1D2E 0%,#0B1726 100%) !important;
        border:1px solid #223B53 !important;
        border-radius:14px !important;
        padding:8px 9px 2px !important;
        box-shadow:0 10px 28px rgba(0,0,0,.15),inset 0 1px 0 rgba(255,255,255,.018);
        overflow:hidden !important;
    }

    div[data-testid="stButton"] > button,
    div[data-testid="stButton"] > button:hover,
    div[data-testid="stButton"] > button:focus,
    div[data-testid="stButton"] > button:active {
        background:linear-gradient(135deg,#17273A,#102033) !important;
        color:#DCE6EF !important;
        border:1px solid #314A64 !important;
        border-radius:9px !important;
        box-shadow:none !important;
        font-size:.72rem !important;
        font-weight:680 !important;
    }

    .stDownloadButton button,
    .stDownloadButton button:hover {
        min-height:44px !important;
        background:linear-gradient(135deg,#15263A,#102033) !important;
        color:#E5EDF5 !important;
        border:1px solid #314A64 !important;
        border-left:3px solid #D7AE58 !important;
        border-radius:9px !important;
        font-weight:700 !important;
    }

    div[data-testid="stDataFrame"] {
        background:#0C1928 !important;
        border:1px solid #233B53 !important;
        border-radius:12px !important;
        box-shadow:0 8px 24px rgba(0,0,0,.14);
        overflow:hidden;
    }

    details[data-testid="stExpander"] {
        background:#0C1928 !important;
        border:1px solid #233B53 !important;
        border-radius:11px !important;
    }

    div[data-testid="stAlert"] {
        background:#0D1A2A !important;
        border:1px solid #2A425A !important;
        border-radius:10px !important;
    }

    hr { border-color:#1B3046 !important; opacity:.75 !important; }

    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        min-width:0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# PALETA Y ESTILO PROFESIONAL PARA GRÁFICAS PLOTLY
# ============================================================

EXECUTIVE_COLORS = [
    "#27C99B", "#D7AE58", "#43B9E6", "#6F91B2", "#A984D8",
    "#E57D6A", "#59C6B7", "#C79356", "#7FA7C7", "#8BC9AA",
    "#BC7F87", "#6BC2C8", "#8E9FD0", "#E1C267", "#62B590",
    "#91A8BF", "#D09276", "#7CBBC6", "#BAA56B", "#789283",
]


def apply_executive_bar_style(fig):
    """Estilo premium para gráficas de barras."""
    bar_traces = [
        trace for trace in fig.data
        if getattr(trace, "type", "") == "bar"
    ]

    if len(bar_traces) == 1:
        trace = bar_traces[0]
        values = trace.x if trace.orientation == "h" else trace.y
        count = len(values) if values is not None else 0
        trace.marker.color = [
            EXECUTIVE_COLORS[i % len(EXECUTIVE_COLORS)]
            for i in range(count)
        ]
        trace.marker.line.color = "rgba(255,255,255,.10)"
        trace.marker.line.width = .7
        trace.opacity = .95
    else:
        for i, trace in enumerate(bar_traces):
            trace.marker.color = EXECUTIVE_COLORS[i % len(EXECUTIVE_COLORS)]
            trace.marker.line.color = "rgba(255,255,255,.08)"
            trace.marker.line.width = .7
            trace.opacity = .95

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#0E1B2B",
        plot_bgcolor="#0E1B2B",
        font=dict(
            family="Inter, Segoe UI, Arial, sans-serif",
            color="#C5D0DB",
            size=10,
        ),
        title_font=dict(
            family="Inter, Segoe UI, Arial, sans-serif",
            color="#F6F8FB",
            size=16,
        ),
        bargap=.38,
        bargroupgap=.14,
        margin=dict(l=42, r=24, t=58, b=50),
        hoverlabel=dict(
            bgcolor="#07121E",
            font_color="#FFFFFF",
            bordercolor="#35506B",
            font_size=11,
        ),
        legend=dict(
            bgcolor="rgba(7,18,30,.50)",
            bordercolor="#274057",
            borderwidth=1,
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color="#B3C2D0", size=9),
        ),
    )

    fig.update_xaxes(
        showgrid=False,
        zeroline=False,
        linecolor="#30465E",
        tickfont=dict(color="#8297AC", size=9),
        title_font=dict(color="#AEBECC", size=10),
        automargin=True,
    )

    fig.update_yaxes(
        showgrid=True,
        gridcolor="rgba(65,88,111,.22)",
        griddash="dot",
        zeroline=False,
        linecolor="#30465E",
        tickfont=dict(color="#8297AC", size=9),
        title_font=dict(color="#AEBECC", size=10),
        automargin=True,
    )

    return fig


def apply_executive_pie_style(fig):
    """Estilo premium para pastel y dona."""
    fig.update_traces(
        marker=dict(
            colors=EXECUTIVE_COLORS,
            line=dict(color="#07121E", width=2),
        ),
        textfont=dict(
            color="#FFFFFF",
            size=10,
            family="Inter, Segoe UI, Arial, sans-serif",
        ),
        hoverlabel=dict(
            bgcolor="#07121E",
            font_color="#FFFFFF",
            bordercolor="#35506B",
        ),
    )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#0E1B2B",
        plot_bgcolor="#0E1B2B",
        font=dict(
            family="Inter, Segoe UI, Arial, sans-serif",
            color="#C5D0DB",
            size=10,
        ),
        title_font=dict(color="#F6F8FB", size=16),
        legend=dict(
            font=dict(color="#B3C2D0", size=9),
            bgcolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=18, r=18, t=58, b=22),
    )

    return fig


def render_exec_kpi(label: str, value: str, accent: str, note: str = "") -> None:
    """Renderiza una tarjeta KPI con valor de color llamativo."""
    note_html = f'<div class="exec-kpi-note">{note}</div>' if note else ""
    st.markdown(
        f'''<div class="exec-kpi" style="--accent:{accent};">
<div class="exec-kpi-label">{label}</div>
<div class="exec-kpi-value">{value}</div>
{note_html}
</div>''',
        unsafe_allow_html=True,
    )


# ============================================================
# NORMALIZACIÓN DE ENCABEZADOS
# ============================================================

def normalize_header_name(value: object) -> str:
    """
    Normaliza encabezados para tolerar:
    - espacios adicionales
    - espacios al inicio o final
    - acentos
    - puntos
    - caracteres #, $, %
    - pequeñas variaciones de formato
    """

    if value is None:
        return ""

    text = str(value).strip().upper()

    # Mantener significado de caracteres especiales.
    text = text.replace("#", " NUM ")
    text = text.replace("$", " USD ")
    text = text.replace("%", " PCT ")
    text = text.replace("/", " ")

    # Eliminar acentos.
    text = unicodedata.normalize("NFKD", text)

    text = "".join(
        character
        for character in text
        if not unicodedata.combining(character)
    )

    # La columna original proporcionada utiliza MAGERN.
    # Internamente permitimos también MARGEN.
    text = text.replace("MAGERN", "MARGEN")

    # Quitar puntuación.
    text = re.sub(
        r"[^A-Z0-9]+",
        " ",
        text,
    )

    # Eliminar espacios duplicados.
    text = re.sub(
        r"\s+",
        " ",
        text,
    ).strip()

    return text


# ============================================================
# ALIAS DE COLUMNAS
# ============================================================

COLUMN_ALIASES = {

    "NUMERO DE FAMILIA": [
    "NUMERO DE FAMILIA",
    "NUMERO FAMILIA",
    "NUMERO",
    "Numero",
    "NRO DE FAMILIA",
    "NRO FAMILIA",
    "NO DE FAMILIA",
    "NO FAMILIA",
],

    "FAMILIA": [
        "FAMILIA",
    ],

    "#COD.": [
        "#COD.",
        "#COD",
        "COD",
        "COD.",
        "CODIGO",
        "# COD",
        "# COD.",
    ],

    "DESCRIPCION": [
        "DESCRIPCION",
        "DESCRIPCIÓN",
    ],

    "EMPAQ": [
        "EMPAQ",
        "EMPAQUE",
    ],

    "#COMPRAS MES ACTUAL": [
        "#COMPRAS MES ACTUAL",
        "# COMPRAS MES ACTUAL",
        "COMPRAS MES ACTUAL",
    ],

    "FECHA ULT COMPRA": [
        "FECHA ULT COMPRA",
        "FECHA DE ULT COMPRA",
        "FECHA ULTIMA COMPRA",
        "FECHA DE ULTIMA COMPRA",
    ],

    "COSTO DE BODEGA": [
        "COSTO DE BODEGA",
        "COSTO BODEGA",
    ],

    "COSTO GENERACION": [
        "COSTO GENERACION",
        "COSTO DE GENERACION",
    ],

    "COSTO ACTUAL": [
        "COSTO ACTUAL",
    ],

    "INVENTARIO TOTAL": [
        "INVENTARIO TOTAL",
    ],

    "#VENTAS AÑO ACTUAL": [
        "#VENTAS AÑO ACTUAL",
        "# VENTAS AÑO ACTUAL",
        "#VENTAS ANO ACTUAL",
        "VENTAS AÑO ACTUAL",
        "VENTAS ANO ACTUAL",
    ],

    "#VENTAS MES ANTERIOR": [
        "#VENTAS MES ANTERIOR",
        "# VENTAS MES ANTERIOR",
        "VENTAS MES ANTERIOR",
    ],

    "#VENTAS MES ACTUAL": [
        "#VENTAS MES ACTUAL",
        "# VENTAS MES ACTUAL",
        "VENTAS MES ACTUAL",
    ],

    "#VENTAS PERDIDAS": [
        "#VENTAS PERDIDAS",
        "# VENTAS PERDIDAS",
        "VENTAS PERDIDAS",
    ],

    "FECHA ULT VENTAS": [
        "FECHA ULT VENTAS",
        "FECHA ULT VENTA",
        "FECHA DE ULT VENTA",
        "FECHA DE ULT VENTAS",
        "FECHA ULTIMA VENTA",
        "FECHA ULTIMAS VENTAS",
    ],

    "$VENTAS MES ACTUAL": [
        "$VENTAS MES ACTUAL",
        "$ VENTAS MES ACTUAL",
        "VENTAS $ MES ACTUAL",
        "VENTAS MES ACTUAL $",
    ],

    "PRECIO PROMEDIO": [
        "PRECIO PROMEDIO",
    ],

    "% MAGERN S/VENTA": [
        "% MAGERN S/VENTA",
        "% MARGEN S/VENTA",
        "%MARGEN S/VENTA",
        "% MARGEN S/VENTAS",
        "%MARGEN S/VENTAS",
        "MARGEN S/VENTA",
        "MARGEN S/VENTAS",
    ],

    "$UTILIDAD MES ACTUAL": [
        "$UTILIDAD MES ACTUAL",
        "$ UTILIDAD MES ACTUAL",
        "UTILIDAD MES ACTUAL",
    ],

    "$GENERA S/VENTAS": [
        "$GENERA S/VENTAS",
        "$ GENERA S/VENTAS",
        "GENERA S/VENTAS",
        "$GENERA S/VENTA",
        "GENERA S/VENTA",
    ],
}


def build_alias_lookup() -> Dict[str, str]:
    """
    Crea el diccionario:
    nombre normalizado -> nombre canónico.
    """

    lookup: Dict[str, str] = {}

    for canonical, aliases in COLUMN_ALIASES.items():

        lookup[
            normalize_header_name(canonical)
        ] = canonical

        for alias in aliases:

            lookup[
                normalize_header_name(alias)
            ] = canonical

    return lookup


ALIAS_LOOKUP = build_alias_lookup()


# ============================================================
# FUNCIONES PARA ENCABEZADOS
# ============================================================

def clean_header_cell(value: object) -> str:
    """
    Limpia una celda utilizada como encabezado.
    """

    if value is None or pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() == "nan":
        return ""

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text


def map_columns_to_canonical(
    columns: List[object],
) -> Tuple[List[str], Dict[str, str]]:

    mapped: List[str] = []

    report: Dict[str, str] = {}

    repetitions: Dict[str, int] = {}

    for original in columns:

        original_clean = clean_header_cell(
            original
        )

        normalized = normalize_header_name(
            original_clean
        )

        canonical = ALIAS_LOOKUP.get(
            normalized,
            original_clean or "COLUMNA_SIN_NOMBRE",
        )

        # Evitar nombres duplicados.
        if canonical in repetitions:

            repetitions[canonical] += 1

            final_name = (
                f"{canonical}__DUP"
                f"{repetitions[canonical]}"
            )

        else:

            repetitions[canonical] = 0

            final_name = canonical

        mapped.append(
            final_name
        )

        if original_clean:

            report[
                original_clean
            ] = final_name

    return mapped, report


def combine_header_rows(
    row1: pd.Series,
    row2: pd.Series,
) -> List[str]:

    """
    Combina encabezados divididos en dos filas.

    Ejemplo:

    #VENTAS
    MES ACTUAL

    se convierte en:

    #VENTAS MES ACTUAL
    """

    headers: List[str] = []

    max_len = max(
        len(row1),
        len(row2),
    )

    for i in range(max_len):

        first = clean_header_cell(
            row1.iloc[i]
            if i < len(row1)
            else ""
        )

        second = clean_header_cell(
            row2.iloc[i]
            if i < len(row2)
            else ""
        )

        if first and second:

            if (
                normalize_header_name(first)
                ==
                normalize_header_name(second)
            ):

                combined = first

            else:

                combined = (
                    f"{first} {second}"
                )

        else:

            combined = (
                first or second
            )

        combined = re.sub(
            r"\s+",
            " ",
            combined,
        ).strip()

        headers.append(
            combined
        )

    return headers


def count_recognized_columns(
    headers: List[str],
) -> int:

    mapped, _ = map_columns_to_canonical(
        headers
    )

    return sum(
        1
        for column in mapped
        if column in REQUIRED_COLUMNS
    )


def score_header_row(
    row: pd.Series,
) -> int:

    """
    Calcula qué tan probable es que una fila sea
    parte del encabezado del reporte.
    """

    joined = " | ".join(
        normalize_header_name(value)
        for value in row.tolist()
    )

    keywords = [

        "FAMILIA",
        "COD",
        "DESCRIPCION",
        "EMPAQ",
        "COMPRAS",
        "VENTAS",
        "COSTO",
        "INVENTARIO",
        "PRECIO",
        "UTILIDAD",
        "MARGEN",

    ]

    return sum(
        1
        for word in keywords
        if word in joined
    )


# ============================================================
# DETECTAR ENCABEZADOS Y LEER EXCEL
# ============================================================

def detect_header_and_read_excel(
    source,
) -> Tuple[pd.DataFrame, Dict[str, object]]:

    """
    Detecta automáticamente:

    - encabezado de una fila
    - encabezado dividido en dos filas
    """

    xls = pd.ExcelFile(
        source,
        engine="openpyxl",
    )

    if not xls.sheet_names:

        raise ValueError(
            "El archivo Excel no contiene hojas."
        )

    selected_sheet = xls.sheet_names[0]

    preview = pd.read_excel(
        xls,
        sheet_name=selected_sheet,
        header=None,
        nrows=40,
        engine="openpyxl",
    )

    if preview.empty:

        raise ValueError(
            "La primera hoja del archivo está vacía."
        )

    scores = {

        idx: score_header_row(
            preview.iloc[idx]
        )

        for idx
        in range(len(preview))

    }

    best_row = max(
        scores,
        key=scores.get,
    )

    if scores[best_row] < 3:

        raise ValueError(
            "No fue posible identificar automáticamente "
            "los encabezados."
        )

    candidates: List[
        Dict[str, object]
    ] = []


    # --------------------------------------------------------
    # Posibilidad 1:
    # Encabezado de una sola fila
    # --------------------------------------------------------

    one_row_headers = [

        clean_header_cell(value)

        for value
        in preview.iloc[
            best_row
        ].tolist()

    ]

    candidates.append({

        "header_rows": [
            best_row
        ],

        "headers":
            one_row_headers,

        "matches":
            count_recognized_columns(
                one_row_headers
            ),

        "data_start":
            best_row + 1,

    })


    # --------------------------------------------------------
    # Posibilidad 2:
    # fila anterior + fila encontrada
    # --------------------------------------------------------

    if best_row > 0:

        previous_headers = combine_header_rows(

            preview.iloc[
                best_row - 1
            ],

            preview.iloc[
                best_row
            ],

        )

        candidates.append({

            "header_rows": [
                best_row - 1,
                best_row,
            ],

            "headers":
                previous_headers,

            "matches":
                count_recognized_columns(
                    previous_headers
                ),

            "data_start":
                best_row + 1,

        })


    # --------------------------------------------------------
    # Posibilidad 3:
    # fila encontrada + fila siguiente
    # --------------------------------------------------------

    if (
        best_row + 1
        <
        len(preview)
    ):

        next_headers = combine_header_rows(

            preview.iloc[
                best_row
            ],

            preview.iloc[
                best_row + 1
            ],

        )

        candidates.append({

            "header_rows": [
                best_row,
                best_row + 1,
            ],

            "headers":
                next_headers,

            "matches":
                count_recognized_columns(
                    next_headers
                ),

            "data_start":
                best_row + 2,

        })


    best = max(

        candidates,

        key=lambda item:
        int(
            item["matches"]
        ),

    )


    # --------------------------------------------------------
    # Leer hoja completa
    # --------------------------------------------------------

    full_raw = pd.read_excel(

        xls,

        sheet_name=
        selected_sheet,

        header=None,

        engine="openpyxl",

    )


    data_start = int(
        best[
            "data_start"
        ]
    )

    headers = list(
        best[
            "headers"
        ]
    )


    df = full_raw.iloc[
        data_start:
    ].copy()


    # --------------------------------------------------------
    # Asegurar que haya un nombre por columna
    # --------------------------------------------------------

    if (
        len(headers)
        <
        df.shape[1]
    ):

        headers.extend(

            [

                f"COLUMNA_SIN_NOMBRE_{i}"

                for i
                in range(
                    len(headers) + 1,
                    df.shape[1] + 1,
                )

            ]

        )

    elif (
        len(headers)
        >
        df.shape[1]
    ):

        headers = headers[
            :df.shape[1]
        ]


    df.columns = headers


    # --------------------------------------------------------
    # Eliminar filas y columnas totalmente vacías
    # --------------------------------------------------------

    df = df.dropna(
        axis=1,
        how="all",
    )

    df = df.dropna(
        axis=0,
        how="all",
    )

    df = df.reset_index(
        drop=True
    )


    # --------------------------------------------------------
    # Convertir nombres a los nombres canónicos
    # --------------------------------------------------------

    mapped_columns, mapping_report = (
        map_columns_to_canonical(
            df.columns.tolist()
        )
    )

    df.columns = mapped_columns


    info = {

        "sheet_name":
            selected_sheet,

        "header_rows": [

            int(row) + 1

            for row
            in list(
                best[
                    "header_rows"
                ]
            )

        ],

        "recognized_columns":
            int(
                best[
                    "matches"
                ]
            ),

        "mapping_report":
            mapping_report,

    }


    return df, info


# ============================================================
# VALORES VACÍOS / MARCADORES DE NEGOCIO
# ============================================================

BUSINESS_EMPTY_VALUES = {
    "",
    "NAN",
    "NONE",
    "NULL",
    "N/A",
    "NA",
    "N/D",
    "ND",
    "S/D",
    "-",
    "--",
    "---",
    "SIN DATO",
    "SIN DATOS",
    "SIN VENTA",
    "SIN VENTAS",
    "SIN COMPRA",
    "SIN COMPRAS",
    "SIN MOVIMIENTO",
    "NO APLICA",
    "NO APLICA.",
}


def is_business_empty(value: object) -> bool:
    """
    Devuelve True cuando una celda representa realmente un valor vacío
    o un marcador de negocio que no debe contarse como error de limpieza.
    """

    if value is None:
        return True

    try:
        if pd.isna(value):
            return True
    except Exception:
        pass

    text = str(value).strip().upper()

    return text in BUSINESS_EMPTY_VALUES


# ============================================================
# CONVERSIÓN DE NÚMEROS
# ============================================================

def parse_number(
    value: object,
) -> float:

    """
    Convierte correctamente valores como:

    $1,234.56
    1,234.56
    1.234,56
    25%
    (1,250.00)

    Si no puede convertir devuelve NaN.
    """

    if is_business_empty(value):

        return float("nan")


    if isinstance(
        value,
        (int, float),
    ):

        return float(value)


    text = str(
        value
    ).strip()


    if is_business_empty(text):

        return float("nan")


    negative_parentheses = (

        text.startswith("(")
        and
        text.endswith(")")

    )


    if negative_parentheses:

        text = text[
            1:-1
        ]


    text = (

        text

        .replace(
            "$",
            "",
        )

        .replace(
            "%",
            "",
        )

        .replace(
            "B/.",
            "",
        )

        .replace(
            "B/",
            "",
        )

        .replace(
            "\u00A0",
            "",
        )

        .replace(
            " ",
            "",
        )

    )


    text = re.sub(

        r"[^0-9,\.\-+]",

        "",

        text,

    )


    if not text:

        return float("nan")


    # --------------------------------------------------------
    # Detectar formato decimal
    # --------------------------------------------------------

    if (
        ","
        in text
        and
        "."
        in text
    ):

        if (
            text.rfind(",")
            >
            text.rfind(".")
        ):

            # 1.234,56
            text = text.replace(
                ".",
                "",
            )

            text = text.replace(
                ",",
                ".",
            )

        else:

            # 1,234.56
            text = text.replace(
                ",",
                "",
            )


    elif "," in text:

        parts = text.split(
            ","
        )

        if len(parts) > 2:

            text = "".join(
                parts
            )

        elif len(parts) == 2:

            left, right = parts

            if (
                len(right) == 3
                and
                left
                not in {
                    "0",
                    "+0",
                    "-0",
                }
            ):

                text = (
                    left
                    +
                    right
                )

            else:

                text = (
                    left
                    +
                    "."
                    +
                    right
                )


    elif (
        text.count(".")
        >
        1
    ):

        parts = text.split(
            "."
        )

        text = (

            "".join(
                parts[:-1]
            )

            +
            "."

            +
            parts[-1]

        )


    try:

        number = float(
            text
        )

        if negative_parentheses:

            number = -abs(
                number
            )

        return number


    except (
        TypeError,
        ValueError,
    ):

        return float(
            "nan"
        )


# ============================================================
# CONVERSIÓN DE FECHAS
# ============================================================

SPANISH_MONTH_REPLACEMENTS = {

    # Abreviaturas estándar
    "ENE": "JAN",
    "FEB": "FEB",
    "MAR": "MAR",
    "MZO": "MAR",
    "ABR": "APR",
    "MAY": "MAY",
    "JUN": "JUN",
    "JUL": "JUL",
    "AGO": "AUG",
    "SEP": "SEP",
    "SEPT": "SEP",
    "OCT": "OCT",
    "NOV": "NOV",
    "DIC": "DEC",

    # Abreviaturas compactas del archivo, por ejemplo 27Jl2026
    "EN": "JAN",
    "FB": "FEB",
    "MZ": "MAR",
    "AB": "APR",
    "MY": "MAY",
    "JN": "JUN",
    "JL": "JUL",
    "AG": "AUG",
    "SP": "SEP",
    "OC": "OCT",
    "NV": "NOV",
    "DC": "DEC",

}


def parse_single_date(
    value: object,
):
    """
    Convierte fechas válidas de Excel o texto.

    Soporta:
    - datetime / Timestamp
    - seriales numéricos de Excel
    - seriales de Excel guardados como texto, por ejemplo "45568" o "45568.0"
    - fechas dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd
    - fechas con meses abreviados en español
    - marcadores como "-", "N/A", "SIN VENTA" o "SIN COMPRA"
    """

    if is_business_empty(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    # --------------------------------------------------------
    # datetime / date
    # --------------------------------------------------------
    if (
        hasattr(value, "year")
        and hasattr(value, "month")
        and hasattr(value, "day")
    ):
        try:
            return pd.Timestamp(value)
        except Exception:
            pass

    # --------------------------------------------------------
    # Fecha serial numérica de Excel
    # --------------------------------------------------------
    if isinstance(value, (int, float)):
        numeric = float(value)

        if 20000 <= numeric <= 80000:
            return (
                pd.Timestamp("1899-12-30")
                + pd.to_timedelta(numeric, unit="D")
            )

        # Ceros y números fuera del rango de fecha de Excel:
        # tratarlos como vacíos, no como errores.
        return pd.NaT

    text = str(value).strip()

    if is_business_empty(text):
        return pd.NaT

    upper_text = text.upper().strip()

    # Marcadores adicionales frecuentes en columnas de fecha.
    if upper_text in {
        "0",
        "0.0",
        "00/00/0000",
        "00-00-0000",
        "0000-00-00",
    }:
        return pd.NaT

    # --------------------------------------------------------
    # Serial de Excel guardado como texto
    # --------------------------------------------------------
    numeric_text = (
        upper_text
        .replace(",", ".")
        .replace(" ", "")
    )

    if re.fullmatch(r"\d+(?:\.\d+)?", numeric_text):
        try:
            numeric = float(numeric_text)

            if 20000 <= numeric <= 80000:
                return (
                    pd.Timestamp("1899-12-30")
                    + pd.to_timedelta(numeric, unit="D")
                )
        except (TypeError, ValueError):
            pass

    # --------------------------------------------------------
    # Formato compacto del sistema: 27Jl2026, 05En2026, etc.
    # --------------------------------------------------------

    compact_match = re.fullmatch(
        r"(\d{1,2})\s*([A-ZÁÉÍÓÚÑ]{2,4})\s*(\d{4})",
        upper_text,
    )

    if compact_match:

        compact_day = compact_match.group(1)
        compact_month = compact_match.group(2)
        compact_year = compact_match.group(3)

        compact_month_en = (
            SPANISH_MONTH_REPLACEMENTS.get(
                compact_month
            )
        )

        if compact_month_en:

            compact_cleaned = (
                f"{compact_day}"
                f"{compact_month_en}"
                f"{compact_year}"
            )

            compact_date = pd.to_datetime(
                compact_cleaned,
                format="%d%b%Y",
                errors="coerce",
            )

            if not pd.isna(compact_date):
                return compact_date

    # --------------------------------------------------------
    # Intento directo: conserva formatos estándar de fecha/hora
    # --------------------------------------------------------
    direct = pd.to_datetime(
        text,
        errors="coerce",
        dayfirst=True,
    )

    if not pd.isna(direct):
        return direct

    # --------------------------------------------------------
    # Normalizar meses abreviados en español
    # --------------------------------------------------------
    cleaned = (
        upper_text
        .replace(".", "")
    )

    for es_month, en_month in sorted(
        SPANISH_MONTH_REPLACEMENTS.items(),
        key=lambda x: len(x[0]),
        reverse=True,
    ):
        cleaned = re.sub(
            rf"\b{es_month}\b",
            en_month,
            cleaned,
        )

        cleaned = re.sub(
            rf"(?<=\d){es_month}(?=\d)",
            en_month,
            cleaned,
        )

    cleaned = re.sub(
        r"\s+",
        " ",
        cleaned,
    ).strip()

    return pd.to_datetime(
        cleaned,
        errors="coerce",
        dayfirst=True,
    )


# ============================================================
# LIMPIEZA DE FILAS
# ============================================================

def remove_non_product_rows(
    df: pd.DataFrame,
) -> pd.DataFrame:

    result = df.copy()


    meaningful = [

        column

        for column
        in [

            "#COD.",
            "DESCRIPCION",
            "FAMILIA",
            "$VENTAS MES ACTUAL",
            "INVENTARIO TOTAL",

        ]

        if column
        in result.columns

    ]


    if meaningful:

        result = result.dropna(

            subset=
            meaningful,

            how="all",

        )


    if (
        "DESCRIPCION"
        in result.columns
    ):

        descriptions = (

            result[
                "DESCRIPCION"
            ]

            .astype(
                "string"
            )

            .fillna(
                ""
            )

            .str.strip()

            .str.upper()

        )


        result = result[

            descriptions.ne(
                "DESCRIPCION"
            )

        ]


    return result.reset_index(
        drop=True
    )


# ============================================================
# LIMPIEZA COMPLETA
# ============================================================

def clean_dataframe(
    df: pd.DataFrame,
) -> Tuple[
    pd.DataFrame,
    List[str],
]:

    result = df.copy()

    warnings: List[str] = []


    # --------------------------------------------------------
    # Texto
    # --------------------------------------------------------

    for column in TEXT_COLUMNS:

        if column not in result.columns:

            continue


        result[column] = (

            result[column]

            .astype(
                "string"
            )

            .str.strip()

        )


        result[column] = (
            result[column]
            .replace(
                {
                    "":
                        pd.NA,
                    "<NA>":
                        pd.NA,
                    "nan":
                        pd.NA,
                    "None":
                        pd.NA,
                }
            )
        )


    # --------------------------------------------------------
    # REGLA GLOBAL AJUSTE / AJUSTES
    # --------------------------------------------------------
    # Toda descripción que COMIENCE con "AJUSTE" queda marcada.
    # startswith("AJUSTE") incluye tanto AJUSTE como AJUSTES.
    # Esta marca se utiliza en todo el dashboard para:
    # - excluir esos productos de selectores/listas de productos;
    # - excluirlos de cálculos de inventario;
    # - excluirlos del análisis de EMPAQ.
    if "DESCRIPCION" in result.columns:
        result["AJUSTE_PRODUCTO"] = (
            result["DESCRIPCION"]
            .astype("string")
            .fillna("")
            .str.strip()
            .str.upper()
            .str.startswith(
                "AJUSTE",
                na=False,
            )
        )
    else:
        result["AJUSTE_PRODUCTO"] = False

    # --------------------------------------------------------
    # EMPAQ PARA ANÁLISIS
    # --------------------------------------------------------
    if "EMPAQ" in result.columns:
        result["EMPAQ_ANALISIS"] = (
            result["EMPAQ"]
            .mask(
                result["AJUSTE_PRODUCTO"],
                pd.NA,
            )
        )


    # --------------------------------------------------------
    # Código
    # --------------------------------------------------------

    if (
        "#COD."
        in result.columns
    ):

        result[
            "#COD."
        ] = (

            result[
                "#COD."
            ]

            .astype(
                "string"
            )

            .str.replace(
                r"\.0$",
                "",
                regex=True,
            )

            .str.strip()

        )


    # --------------------------------------------------------
    # Números
    # --------------------------------------------------------

    for column in NUMERIC_COLUMNS:

        if column not in result.columns:

            continue


        source_values = result[column]

        expected_numeric_mask = (
            ~source_values.apply(is_business_empty)
        )

        converted = (
            source_values
            .apply(parse_number)
        )

        failed = int(
            (
                expected_numeric_mask
                & converted.isna()
            )
            .sum()
        )


        if failed:

            warnings.append(

                f"{failed:,} valores de "
                f"'{column}' no pudieron "
                "convertirse a número. Revise esos valores; "
                "los marcadores vacíos válidos no se cuentan aquí."

            )


        result[
            column
        ] = pd.to_numeric(

            converted,

            errors="coerce",

        )


    # --------------------------------------------------------
    # INVENTARIO PARA ANÁLISIS — EXCLUSIÓN GLOBAL DE AJUSTES
    # --------------------------------------------------------
    # Se conserva el valor fuente en una columna interna y el
    # INVENTARIO TOTAL analítico se lleva a 0 para AJUSTE/AJUSTES.
    # De esta forma ningún KPI, gráfica o tabla de inventario
    # suma esos productos en ninguna pestaña del dashboard.
    if "INVENTARIO TOTAL" in result.columns:
        result["INVENTARIO_TOTAL_ORIGINAL"] = (
            result["INVENTARIO TOTAL"].copy()
        )

        result.loc[
            result["AJUSTE_PRODUCTO"],
            "INVENTARIO TOTAL",
        ] = 0.0


    # --------------------------------------------------------
    # Fechas
    # --------------------------------------------------------

    for column in DATE_COLUMNS:

        if column not in result.columns:

            continue


        source_values = result[column]

        def _date_should_convert(value: object) -> bool:
            if is_business_empty(value):
                return False

            text = str(value).strip().upper()

            if text in {
                "0",
                "0.0",
                "00/00/0000",
                "00-00-0000",
                "0000-00-00",
            }:
                return False

            return True

        expected_date_mask = (
            source_values.apply(_date_should_convert)
        )

        converted = (
            source_values
            .apply(parse_single_date)
        )

        converted = pd.to_datetime(
            converted,
            errors="coerce",
        )

        failed = int(
            (
                expected_date_mask
                & converted.isna()
            )
            .sum()
        )


        if failed:

            warnings.append(

                f"{failed:,} valores de "
                f"'{column}' no pudieron "
                "interpretarse como fecha. Revise esos valores; "
                "los marcadores vacíos válidos no se cuentan aquí."

            )


        result[
            column
        ] = converted


    result = remove_non_product_rows(
        result
    )


    return (
        result,
        warnings,
    )


# ============================================================
# CARGA CACHEADA
# ============================================================

@st.cache_data(
    show_spinner=False
)
def load_uploaded_excel(
    file_bytes: bytes,
):

    buffer = io.BytesIO(
        file_bytes
    )

    raw_df, info = (
        detect_header_and_read_excel(
            buffer
        )
    )

    clean_df, warnings = (
        clean_dataframe(
            raw_df
        )
    )

    return (
        clean_df,
        info,
        warnings,
    )


@st.cache_data(
    show_spinner=False
)
def load_local_excel(
    file_path: str,
    modified_time: float,
):

    # Se utiliza modified_time para invalidar
    # automáticamente el caché si cambia el Excel.
    _ = modified_time

    raw_df, info = (
        detect_header_and_read_excel(
            file_path
        )
    )

    clean_df, warnings = (
        clean_dataframe(
            raw_df
        )
    )

    return (
        clean_df,
        info,
        warnings,
    )


# ============================================================
# VALIDACIÓN
# ============================================================

def validate_required_columns(
    df: pd.DataFrame,
) -> List[str]:

    return [

        column

        for column
        in REQUIRED_COLUMNS

        if column
        not in df.columns

    ]


# ============================================================
# FORMATOS
# ============================================================

def money(
    value: object,
) -> str:

    try:

        if (
            value is None
            or pd.isna(value)
        ):

            return "$0.00"

        return (
            f"${float(value):,.2f}"
        )

    except (
        TypeError,
        ValueError,
    ):

        return "$0.00"


def quantity(
    value: object,
    decimals: int = 0,
) -> str:

    """
    Función corregida.
    El error de SyntaxError estaba anteriormente
    en el f-string de esta función.
    """

    try:

        if (
            value is None
            or pd.isna(value)
        ):

            value = 0


        if decimals <= 0:

            return (
                f"{float(value):,.0f}"
            )


        # CORRECCIÓN DEL F-STRING
        return (
            f"{float(value):,.{decimals}f}"
        )


    except (
        TypeError,
        ValueError,
    ):

        return "0"


def percent(
    value: object,
) -> str:

    try:

        if (
            value is None
            or pd.isna(value)
        ):

            return "0.00%"

        return (
            f"{float(value):,.2f}%"
        )

    except (
        TypeError,
        ValueError,
    ):

        return "0.00%"


def metric_axis_title(
    metric: str,
) -> str:

    labels = {

        "$VENTAS MES ACTUAL":
            "Ventas ($)",

        "$UTILIDAD MES ACTUAL":
            "Utilidad ($)",

        "INVENTARIO TOTAL":
            "Inventario",

        "#VENTAS MES ACTUAL":
            "Unidades vendidas",

        "#VENTAS PERDIDAS":
            "Ventas perdidas",

    }

    return labels.get(
        metric,
        metric,
    )


def safe_sum(
    series: pd.Series,
) -> float:

    return float(

        pd.to_numeric(

            series,

            errors="coerce",

        )

        .fillna(0)

        .sum()

    )


def safe_mean(
    series: pd.Series,
) -> float:

    values = pd.to_numeric(

        series,

        errors="coerce",

    ).dropna()


    if values.empty:

        return 0.0


    return float(
        values.mean()
    )


# ============================================================
# AGREGACIÓN POR FAMILIA
# ============================================================

def family_summary(
    df: pd.DataFrame,
) -> pd.DataFrame:

    if df.empty:

        return pd.DataFrame()


    summary = (

        df

        .groupby(

            "FAMILIA",

            dropna=False,

            as_index=False,

        )

        .agg(

            {

                "$VENTAS MES ACTUAL":
                    "sum",

                "$UTILIDAD MES ACTUAL":
                    "sum",

                "INVENTARIO TOTAL":
                    "sum",

                "#VENTAS PERDIDAS":
                    "sum",

                "% MAGERN S/VENTA":
                    "mean",

                "#VENTAS MES ANTERIOR":
                    "sum",

                "#VENTAS MES ACTUAL":
                    "sum",

            }

        )

        .rename(

            columns={

                "% MAGERN S/VENTA":
                    "MARGEN PROMEDIO %",

            }

        )

        .sort_values(

            "$VENTAS MES ACTUAL",

            ascending=False,

        )

        .reset_index(
            drop=True
        )

    )


    return summary


# ============================================================
# AGREGACIÓN POR PRODUCTO
# ============================================================

def product_summary(
    df: pd.DataFrame,
) -> pd.DataFrame:

    """
    Agrupa productos correctamente por:

    FAMILIA
    #COD.
    DESCRIPCION

    Si un producto aparece varias veces:
    - suma ventas
    - suma inventario
    - suma utilidad
    - suma ventas perdidas
    - promedia costo
    - promedia precio
    - promedia margen
    """

    if df.empty:

        return pd.DataFrame()


    result = (

        df

        .groupby(

            [
                "FAMILIA",
                "#COD.",
                "DESCRIPCION",
            ],

            dropna=False,

            as_index=False,

        )

        .agg(

            {

                "$VENTAS MES ACTUAL":
                    "sum",

                "#VENTAS MES ANTERIOR":
                    "sum",

                "#VENTAS MES ACTUAL":
                    "sum",

                "#VENTAS PERDIDAS":
                    "sum",

                "INVENTARIO TOTAL":
                    "sum",

                "COSTO ACTUAL":
                    "mean",

                "PRECIO PROMEDIO":
                    "mean",

                "% MAGERN S/VENTA":
                    "mean",

                "$UTILIDAD MES ACTUAL":
                    "sum",

            }

        )

        .sort_values(

            "$VENTAS MES ACTUAL",

            ascending=False,

        )

        .reset_index(
            drop=True
        )

    )


    return result


def aggregate_family_metric(
    df: pd.DataFrame,
    metric: str,
) -> pd.DataFrame:

    result = (

        df

        .groupby(

            "FAMILIA",

            dropna=False,

            as_index=False,

        )[metric]

        .sum()

        .sort_values(

            metric,

            ascending=False,

        )

        .reset_index(
            drop=True
        )

    )


    return result


def aggregate_product_metric(
    df: pd.DataFrame,
    metric: str,
) -> pd.DataFrame:

    result = (

        df

        .groupby(

            [
                "#COD.",
                "DESCRIPCION",
            ],

            dropna=False,

            as_index=False,

        )[metric]

        .sum()

        .sort_values(

            metric,

            ascending=False,

        )

        .reset_index(
            drop=True
        )

    )


    return result


# ============================================================
# EXPORTACIÓN CSV
# ============================================================

def dataframe_to_csv_bytes(
    df: pd.DataFrame,
) -> bytes:

    return (

        df

        .to_csv(
            index=False
        )

        .encode(
            "utf-8-sig"
        )

    )


# ============================================================
# EXPORTACIÓN EXCEL
# ============================================================

def dataframe_to_excel_bytes(
    df: pd.DataFrame,
    sheet_name: str,
) -> bytes:

    output = io.BytesIO()


    with pd.ExcelWriter(

        output,

        engine="openpyxl",

    ) as writer:


        clean_sheet_name = (
            sheet_name[:31]
        )


        df.to_excel(

            writer,

            index=False,

            sheet_name=
            clean_sheet_name,

        )


        worksheet = (
            writer.book[
                clean_sheet_name
            ]
        )


        worksheet.freeze_panes = (
            "A2"
        )


        if (

            worksheet.max_row
            >=
            1

            and

            worksheet.max_column
            >=
            1

        ):

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )


        # ----------------------------------------------------
        # Ajustar ancho de columnas
        # ----------------------------------------------------

        for column_cells in worksheet.columns:


            column_letter = (
                column_cells[
                    0
                ].column_letter
            )


            max_length = 0


            for cell in column_cells:


                value = (

                    ""

                    if cell.value is None

                    else

                    str(
                        cell.value
                    )

                )


                max_length = max(

                    max_length,

                    len(
                        value
                    ),

                )


            worksheet.column_dimensions[

                column_letter

            ].width = min(

                max(
                    max_length + 2,
                    10,
                ),

                45,

            )


    output.seek(0)


    return output.getvalue()


# ============================================================
# BUSCAR ARCHIVO EXCEL LOCAL
# ============================================================

def find_default_excel() -> Optional[Path]:

    directories: List[Path] = []


    try:

        directories.append(

            Path(
                __file__
            )

            .resolve()

            .parent

        )

    except NameError:

        pass


    directories.append(
        Path.cwd()
    )


    possible_names = [

        "CIERRE PRODUCTOS.xlsx",

        "CIERRE DE PRODUCTOS.xlsx",

    ]


    seen = set()


    for directory in directories:


        try:

            directory_key = str(

                directory.resolve()

            )

        except Exception:

            directory_key = str(
                directory
            )


        if directory_key in seen:

            continue


        seen.add(
            directory_key
        )


        for filename in possible_names:


            candidate = (

                directory
                /
                filename

            )


            if (

                candidate.exists()

                and

                candidate.is_file()

            ):

                return candidate


    return None


# ============================================================
# LOGO CORPORATIVO DICASA S.A. EMBEBIDO
# ============================================================

DICASA_LOGO_URI = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAQDAwMDAgQDAwMEBAQFBgoGBgUFBgwICQcKDgwPDg4MDQ0PERYTDxAVEQ0NExoTFRcYGRkZDxIbHRsYHRYYGRj/2wBDAQQEBAYFBgsGBgsYEA0QGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBgYGBj/wAARCADcANwDASIAAhEBAxEB/8QAHQAAAgEFAQEAAAAAAAAAAAAAAAECAwQFBggHCf/EAEIQAAEDAgMEBQkFBwQDAQAAAAEAAgMEEQUGIQcSMUETUWFxkQgVIjI0U3Kx0RRSc4HBIyQzNWKCoRZCQ5IlY6Lh/8QAGgEBAAIDAQAAAAAAAAAAAAAAAAECAwQFBv/EADURAAIBAgQDBAgFBQAAAAAAAAABAgMRBAUhMRITQVGBkbEiIzJCYXGhwQYUJDPRUpLh8PH/2gAMAwEAAhEDEQA/AO/kIRyQAhLeb94eKN5v3h4oBoSuOsIuOsIBoRcdaEAIQhACEIQAhFx1ouOtACErjrCN4dYQDQlvDrCN5v3h4oBoS32feHilvs++3xQEkKPSM++3xR0kfvG+KAkhQ6WL3rP+wR00PvWf9ggJoUOmh99H/wBgm2SN5syRrj1A3QEkIQgBJ3qHuTSd6ju5AW40CYtxUbaKVlUEgpKI1TClAkOCaQ4pqQCEIQAhCEAuaRKfBIoBJaIv1Jd6ARSdZF1E8VFwI9aiepMlQNkuBOKgdFIqJKglEHKJ0FkzxUShJFXWH+1O+D9QrXjeyusPH7274P1CBmSQhCsVBJ3qO7k0neo7uQFs3hqpKF1IWuqgkFIa81EKV1YDTukndAF0XRwWDxHN+XcMmfBU4nE6dmjoYbyvaeoht9387K0YuTtFXKykoq8mZy6V+1aPJtQwIOtDT1Drc3uY3/Fyos2o4I54D4JW/wB4/Wy2FgcQ1fgfga7x2HTtxrxN6ulfRa1SZ7y7VWvVOhv7wA/5aSs9S1dLWx9JSVMU7euNwNu/qWCdOcPaVjPCrCesHcrc1EnVMixskSQVQuRSKZUTzVQRPBRUiolARdw0UCpnRUzxQkiVE8FIqJ1QkirrDzeqd8H6hWp7FdYf7U74f1CEMySEIViASd6ju5NJ/wDDd3IC1Ck1UwpgqoJBSHFRB1TurAldDpI4ojJK9rGDi5xsAkDzXPe1baZK/ae3LmG1JbRYJZ9SGnSarc3eDT1iNpabfef/AErYwuFnianLgYq1eNGPHPY3vabjNdDB9ngxIxUjm2MNOC0yde+8a2/pFu2/BeJyVb3S2vZgOjQLNHcBoFWjzXLiMjYa2V0sLjqL636x2rzTaJiuYavGabAMgVEVTDVx7zq+mfd1iSLAj1BcEaek7lYan1eHjSy2narv29p5/EQqZhP1Wy+huWN58ynlaAvx7HqSjkAv0Bdvyn+xtz42WCw7bJl/G6Z1Xl/Bsw4rTh2709NRt3CezeeD/hed4xsAxLKmH4bj2YcPmxcV1V0UzWhz3xEi4c2PgTe/HeOh4rIw5fw7B6EGlwjMFPA3Xo6eKdo4andjb1DjbktGvn7X7a8bnZyz8IxxMXOrPRdjSd/k39TaMw7VcZwzA34wMpOoKCGRolmxSa0pYT6RbGzhYAm5d+S3zIG3vJ8czJ6XH4o942Da9roiey7vR8CvFMJy9T592kYNlXD6evZDOenqzWOlu6MHUNbJbSwILrH1gAvbMZ8k7LM0BnwyM0sxF7xEst4fqteWbSnHhxEbp/QyYvJcPhaqjhZNOK1bad33adhvlVt/OE4kyaso46/C5dbwvAkZ8DvVcOw2+Jeq5WznlzOWE/b8vYnFVMbbpYvVlhJ5PYdW9/A8iVwrmHYjnzJYkfg1RJPT8XM4b3eLbp/MA9q1jAM747lXMcdS6oq8CxOmPo1cJLA0X4Pab+ieYO8w9av+VweMj+nfDPsexpc3E4Z+tXFE+lt+SXFeNbItu+F58njy3jvQ4fmUNvGGG0GIAC5dDfg+2pjJJtq0uHD2LeXDr0Z0ZuFRWZ0adSNSPFB6EioHTmi44padawmQiVEqTlDnxQlETfiomykSoki6Ekb6FXWHe1u+D9QrXRXWH+1u+D9QhBk0IQrEAov/AITu4qSjJ/Cd3FAWYJJUhZQHBSB1VQTTUUb2tlYDfK2KJ8r/AFWNLj3AXXz/AMSxl1fmPEMQfIXS1dVNUvceJc95cfmB+S73rwXYTVNA1MLx/wDJXzTdXOixOWAu1Y9zD+RsvUfhqK4qkuun3OHnTdoLpqZrNOM10uEU+X8LL/t2LSfZ2lhs5sf+8g8ibhoP9d+S6n2LbJqPKeVKZ1VGJapzQ5z3C/pW4jqHIdgC5/2N4HT5r240jqlu+ygp+lF9dbu/W3gF7rtT28YhsqzFJlqLJjZjJRtloMRlqHGKVxab3jYy/okG7d4EgcgbrWz6pKriOXHZI28moeqVt5d3mbntUxvJ2AZap8NzFWPZVVzw2ipaaISzucCP2gbwDW3F3OsBw1JsuUA50WJ9OD0bYw9obG1oAd0kDG68bBr3gDhr3LU8Vzv/AKkzSczZzjxvEq17mST1cEJY5zWHea1kbfRDRybfhfUnU4ybaXhUuHzQwUGMVNdK8kBlA+G4cWkt3nus3VjTvWcRyBWlgqsKV+La8Xs9Um7nYzHLZw4Ippyals1ZN2trfxNtgxXG8GzgMewSRlFXMia2Ooppuhka17WPc2wBaRvHgQQvR6Tyj9ptJThlXHgNWQ2wM9MAXEcyWPb/AIAXi0NPhP8Ap3zzjOJl9VNA6qexkUjGtPqmMelpuuAiAte4Gh4rC1M9G0maPL+KPdbgXPC0Wmn7Oj1Wj2PT06WFxcFzJJyjZN3gtUu3Rs7kyvtjyJmbIOG4jmfEqHBsUqYd6poJGyfsnXI5tNgQA6xNxexWt5kyFs12ow1lNlrGKCtraZgkkFI/9pAHkhrtRcAkEdRsuK5ah8vpRZZxBpvYkyvsAuifI4q5X5yzQx9HJS/udOx0UhLnE9I8glx15cO9WS3kk1b5/wAHKx2V08PSc41FL4Xi/KT8jyPO+z/NWzPEm9C2okpWSB8MsJLXROBu10ZGrHAi4A/tsdF1T5Ou36HaNhrcqZnq4hmmmh34p9GtxOFvF4HASt/3tHH1hzA9OzZk/DMy4RNSVtNHIJGkHebe64T2pbOMxbIc8w5uy9PPA2nqRVQzxXDontN98W5/eHMXP3r9GliI4yHIr+10Z5idF0Jcylt1R9GwbouvPdju06h2rbK6HM9O1kVYP2FfTN/4Z2j0gP6T6w7D2Lfr31XJqQlTk4S3RuwkppSWxIqJ4pFx5JHgqFg5qDuKdyokoSI8Fd4b7W74P1CsidVeYaf3t3wH5hCDKoQhWIBRk/gv7ipKMv8ABf8ACUBYBSVNptxUgVUFS6LqF9EXQDfbo3A8CNe5fMzaDQnK+13MWBzHdfTV0gbfm1x3mn8w5fTJ3AhcN+WbkabDsz0G0KgicIapraKtLR6sjb9G497bt72jrXbyTFcms4vaRzsyocylddDI+SvUU8m06ruWmSWiO7/aQfkV0ptByBg+dsLMOJUrJXBtmuI1auCfJxz5Fljbbh8ldNuwSu6NzidA13ou/wAEH8l9KwAW2vcW4pncbYjmLqvInLJep4H0Pn/tA2XZs2fVE8+HVNdUYdclsbHk9GOwHl3W7ua82iqjPvzf6mmY8GzonXY5p7QQvpfjWX6LGKN8NVA17Xda4h8oTZniWR810uOYHTwS4fXnoHwzAbrJhcgt7XAHwXI0n018z2OV5nUc40K0rx2V3a3fZ6GhVWIsbhcYnxWT0cEkk+1tN3A/bbA8Py4K2weF2I07KmszliFNSSy9DEAXdJO7mI22Gg5kkAeK1+bF89vrBJHg2DhzKU0TZejcHNiPFtt/d4m993jrxVeSrrJdomXaOqs18dHCAxgG6HOje42t/UStqjTWIlaSsox8iMRjKuX0p8p+lOpo73sn1tbXb/BVxyiw7D6ZtdR50xKsgfI1oLXvN2vaXNdvAkEENPAkdvEDpLyNaeF9bmbEoquapDzTwiSYkucAHnW/Vey5Wpoc7w4DT04o4Rhu/wDs2GGMND7En0rbx9Z2hNgXOtxK9w8nfNueMr7T8JytHh8ENLj1ZHNVvmhDh9njYd4xel6Bs4cRzB5rVgrt2VtOz5G7mMqssDw1KqlJPW0vn0svh17jvmwLbLTs/ZLw7N2V6nDqyFrt9uhtctPIjtW3RvvGCk/0hY6qibTujydjh3YfimJbF/KqnyFijjDgeOkwMa4+g2UH9mRftNu5wXcofyOi5R8rPJksOBUWe8IjLKzCqhk+8wa2ab/K/gF0dlDHY8y5FwjHonbza2kinv2lov8A5uuhjHzacK/XZ/NGtQXLnKn03RsOnFRJAKjfRK65xtBcpHVCSAFd4Z7a74D8wrLrV5hZ/fHfAfmEIZl0IQrEAoS/wH/CVNQl/gP+E/JAYxp5FSBVMaJhyqSVN7Wyd7lU7ovqhBUvdaRtNybh+d8hYjgGJQ9JT1URY63Fp4hzeoggEdoW6XVOZgewgi6lNxd0LJ6M+Qea8AxvZztFqMIxGMtrKGYOY+1mzM5OH9Lhfu1HJfRzydtqdNtM2SUxkqQ/FsMYynqmOPpPZa0clu0AtP8AU0rRfKU2IwZ7wA4phsUcWM0jS6nmIsHjiY3H7p6+R161yHsnz9mTY3tTZiDGSQ/Z5TT1tFOS1pbcB8cg5A2BvyIa4dvoFJZhQ4ffXn/DOW4/lKnF7rPqpe4XPnlf4eyo2BMqHU8sgixal3jELlrXb7Sbc+NvzXseUM34HnnKNLmPL1UJqScWcxxG/A8etG8Dg4f50IuCFqm3TJ1VnzYjiuX6CqNNV78NVA/d3rvikDw3svYi/LRcNR4J2lodnDVFGpCpfRNP6/DXwPmw6hwcuDRhmKONtLRcf/pVI6ihosRopThuMBtGd6GaKBrns9IkxuY51nsN3cHNLbniDYery7B8/wAOW6atnrpzVbhE7IpHAE304C/5rz3EcIxjBWuGLHFoJG+tuvLwPzJCtGm4Pig/D/h6iWPw2Yx5NVu99Lt928rXKZrMLzDLFgNNgeMUNMZIYqdromBkDGl93cS65MriS4n1b35D2byaJ6au8o/FKGsMXTxxOrcODNQInMYxzR3MbHcci0rxPCMWwmlqZKmqx+sp3BhbGK0vDJC70XatDtQ0kgW421C2PZ9nfDsr7fsAzrFmPC6g0sjqWoh6OWlNTC5zmMDRuOaLsc3VxGtr8yopShaSvq1p3dy32Ofm9GrBxppO0Xd3v73e18fM+n1NC58bQGuNuoKVQwwu3TdcMbYtudZJ5R+L5UxrFa/C8BpIGwYdJTSvjZDOACZX7hBdc6X5clcQ+U7jOz/Z1gOAUlQ3MeLw1clXiVdiMplFVDJKXthjffiY3C8vBpAABubYnTkrOWz1+By6GHniJ8ukry7DqDazhMON7IccopWBwNM42I6v/wAutb8m6tkqfJ3y/FI/edTMkpSfw5HN/RX+L5/y/mbyb67PeD1JfhlVROIMo3XxOuGvjeOT2m4I7NNCFjPJxopcP2BYCJWFr6iN9YQeXTSOkH+HBbjdsK0/6vsaLj67u+57ETYJX1UdTwQVpGckTqodqSSAe8r3CvbX/AfmFYcVfYT7c8f0H5hAzMoQhWKgoTezSfCfkpqnP7NJ8J+SAxAdZMu0VO5CYKqWKm8i6hfRO+iEEw4pX0UL80yTZCS0r6SOqpnRyNDg4W1XKG3vYGzH+kxzAWspcVY3SW3oTNH+yTs6ncR3aLrci/FWNfQxVUDo3tBBFtVkpVZUpccHqUnGM1wy2Pmxsv2vZx2DZ7koa+kkFI4iOrw6oJDJGDhqL8Nd17b27RcL6C5Gz9lbablFuNZarRMwtHT0shAmpnHk9o5Hk4Xa7kV5JtY2FYLm2ie+aiD3tuWPYLPYf6T+i5WGVtquxTNbcbydXVcsUDibRHcfu82kcCDzFiDzBXUlOjjVr6M/P/fE01GeH21j5H0cfhdO5hYYhunsWh5w2VZfzFh9RFPQxnpWFrgBa4IXmWy/yw8rZkkiwXPVM7BMXFmOlDC1jz2sOo/tJ7guisNxfB8epRPg2JUtfGRfeppA+3eOI/MBaFTD1KL1RtQqxmtD5pZ5yTmnJOcqnLdVXkQR3kpaiSMESQ3sLuI1cOBB1581qtTFXtiEE+M0kbZAW2cWN3hz7V9Ks5bPsMzO6OSpga6Rh00C572seT/O6bCsRy8yCN0b5GVQliD95pA3bDsIKtFRn1s/m/5O7h84nBRhUimtr8Mb2/tZ5FTDD875bmo83OZNX01FAKLEJ2ODa9rY7O/bC/7VvAO5gW9KxC83oYm4dhxpcXrXRwYZHJLJGwRunYx0g6NhDgbH0g4jlvW0Oi9QzHsBxrDcj43meF8sNRSUxmDox0Qdum7mjXgQOF7di8uwzCDi+R62ljpaenM0bpdGdH6jd5mu8SSXOGhtwTg4EouXot692/Xrc1qNVKVTEQtdaK6XW9tLWsrfA3XIdTnjFsWq9l2H4tKcDxWrZ5xiaQ6NscZ1laAf2b3Ns3T1ri+oBH0Ryrh8eG5fpaSGMRxxRNYxjeDQBYAdwsuafJMwTCcxbMqXNENNGMQjeaKteG+k50YG6Se1pabd66up4uiiDRyCw1Zv9p+7fxNfE1lXqOta17dn2SX0Lm+iLpJXWIwDJSJ1S7ikUAFyv8IN61/wH5hY7tWQwf29/wCGfmEDM2hCFYqCpz+yy/AfkqipVHscvwH5IDBh3WmDqqd9EwVUsVL6o3ioXRvDkgJ30sgFRvqjeQEnFQOqC5RLkBRmhZIwhzQbrUMw5Jw7FonCSBm8edluhcOpQIBQHLOePJ5wfGI3iXCqec8n7lnDuK8yi2PZ0yjViXK2b8Ww4x+pDNeVjewXuR+RC7slpmSaFt1janBKSYnfhafyWxDGVYKyeniYpUKctWjnTIGctt2DYtEzH8TwzGKRpALZnSsc7s1cbLK5xzLtix7F3TYNU4DhVGdOidUue5ndutBPivZX5Uw9z7inZ4Jx5VoI3X6Bqo8RJy4rK5dUko8N2c80+zPH8zTtmzjm3EMSbzo6Jr4Y3dhke5z7dxCs4dhUMElRTUtCaeJkTY4OjLrWa7eaCCbHW+vE31K6jp8Kp4WgMjaPyVf7FFe+6Fjq1J1NJMvTSp+yeNbBtnkWzurxKGkp5qOnqwC6nZM8xOf98tc4+lYWuOV17s2wCs4qWON+81oCugVW7bu9wyR1F0roUVJA76qJJQSlfRRckLrI4Mf39/4Z+YWNusjgv8wf+GfmEW4exnUIQrlAVKo9jl+A/JVVSqfYpvgd8kBr1z1oJuFDnqnfRULE76IvooA6qRd1aqQSujtUL2SDrlAT1uok6oJCSgAi6XcjS6kEgouF0+9GgGqATWgOF1F9i7RS0ulYXUWJIWCOakAOaWiEgmOCiTYI3rhLgZKLqNwhLgZKiSLpEpXUAZNlksEP/kHj/wBZ+YWLusngf8wf+GfmFK3Iexn0IQrlAVKp9im+B3yVVUqr2Gb8N3yQGtFF7CyV9NVHe7FQvYndO/YoX0QHHrQWG46cUgepBNwoXUCxU3kXUEwgsSv1pbyR4ospBK6d9FDW6d9LKQHNLeN0EpE68FUkLovqlyUedkBI8UkibJaoCQsj81FF0A7a3SKV7ovogDrCyeBfzB/4Z+YWLusngRviL/wz8wpW5D2NhQhCuUBUar2Gf8N3yVZUav2Cf8N3yQGrlyV+pJKyxmQki/UkgoCQOtkAgg2Ubo6kBLVPkoc072QDJQkTc3ugFAO+qWvFHJGvJAO6ROqXJJAO6LpJFAGp4oSJSQEkieSSDogH2JJckckAG6ymA/zF/wCGfmFillcB/mUn4Z+YUrch7GwoQhXKAqNV7BP+G75Ksqc7DJTSRttdzS0X7QgNTJ14JLJnBKo/8kPifol5kq/eQ+J+ipZl7oxw67IvqdFkhglWP+SHxP0R5lq7/wASHxP0SzF0Y3RF9FkvMlV7yHxP0R5kquckPifolmLoxhvfimO1ZHzJV+8h8T9E/MlX7yHxP0SzF0Y3RCyXmSr95D4n6I8y1fvIfE/RLMXRjdbcUgbarJ+ZKr3kPifojzJV+8h8T9EsxdGM4ngglZPzJV+8h8T9EvMdX7yHxP0SzF0Yy6RWU8x1fvIfE/RI4HV+8h8T9EsxdGLQsp5iq/eQeJ+iPMVXb+JD4n6JZi6MWlzWU8xVnvYPE/RHmKs95B/2P0SzF0Yvnojksp5iq/eQeJ+iXmGs95B4n6JZi6MYsngP8yf+GfmEeYaz3sHifor3DMMnoqt0sr4yCzd9Em/EdnYiQbMqhCFcoCEIQAhCEAIQhACEIQAhCEAIQhACEIQAhCEAIQhACEIQAhCEAIQhACEIQAhCEB//2Q=="



# ============================================================
# LOGO DE BIENVENIDA DICASA S.A. — EMBEBIDO
# ============================================================

DICASA_WELCOME_LOGO_URI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABwYAAAcGCAYAAAAMUMRFAAAACXBIWXMAAC4jAAAuIwF4pT92AAAgAElEQVR4nOzdTZAt53kf9rfP3E98EBckwQ9TJCASpPk9IINopEooQi4ypiw5pJ24LFtJSFdFJTuyJLiUZIOKxSymskgcgk72BndeUsusDCyx4806VQEWqVQqFRkQaX0Q93Snuvvt7rf79DlzZubMzJl5fz9q7vSc0/12nw8Ao/O/z/MUVVUFAAAAAAAA4GZbeH0BAAAAAADg5hMMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAIAAAAAAAAGRAMAgAAAAAAQAYEgwAAAAAAAJABwSAAAAAAAABkQDAI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EAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAFZAYBAAAAAAAgApIDAIAAAAAAEAFJAYBAAAAAACgAhKDAAAAAAAAUAGJQQAAAAAAAKiAxCAAAAAAAABUQGIQAAAAAAAAKiAxCAAAAAAAABWQGAQAAAAAAIAKSAwCAAAAAABABSQGAQAAAAAAoAISgwAAAAAAAFABiUEAAAAAAACogMQgAAAAAAAAVEBiEAAAAAAAACogMQgAAAAAAAAVkBgEAAAAAACACkgMAgAAAAAAQAUkBgEAAAAAAKACEoMAAAAAAABQAYlBAAAAAAAAqIDEIAAAAAAAAFRAYhAAAAAAAAAqIDEIAAAAAAAAS7darf4/JLMznwQhVPIAAAAASUVORK5CYII="


# ============================================================
# AUTENTICACIÓN / PANTALLA DE BIENVENIDA
# ============================================================

DEFAULT_DASH_USER = os.getenv(
    "DICASA_DASH_USER",
    "admin",
)

DEFAULT_DASH_PASSWORD = os.getenv(
    "DICASA_DASH_PASSWORD",
    "dicasa2026",
)


def render_welcome_login_screen() -> None:
    st.markdown(
        """
        <style>
        /* ==========================================
           LOGIN DICASA — COMPACTO Y CENTRADO
           ========================================== */

        .block-container {
            /* Centrado vertical real para Streamlit.
               El conjunto visible del login mide aproximadamente 36.5rem.
               El espacio libre restante se reparte por igual arriba y abajo. */
            min-height: 100vh !important;
            padding-top:
                max(
                    2rem,
                    calc((100vh - 36.5rem) / 2)
                ) !important;
            padding-bottom:
                max(
                    2rem,
                    calc((100vh - 36.5rem) / 2)
                ) !important;
            box-sizing: border-box !important;
        }

        .welcome-shell {
            width: 100%;
            margin: 0 auto 0.72rem auto;
        }

        .welcome-hero {
            background:
                radial-gradient(
                    circle at top left,
                    rgba(25,245,255,0.15),
                    transparent 30%
                ),
                radial-gradient(
                    circle at top right,
                    rgba(38,208,124,0.11),
                    transparent 28%
                ),
                linear-gradient(
                    135deg,
                    rgba(21,33,58,0.97) 0%,
                    rgba(13,24,47,0.99) 100%
                );
            border: 1px solid rgba(88,200,255,0.24);
            border-top: 3px solid #20F0FF;
            border-radius: 18px;
            padding: 1.35rem 1.20rem 1.25rem 1.20rem;
            box-shadow: 0 18px 52px rgba(0,0,0,0.32);
            text-align: center;
        }

        .welcome-logo-wrap {
            display: flex;
            justify-content: center;
            margin-bottom: 0.62rem;
        }

        .welcome-logo-frame {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            background: rgba(255,255,255,0.98);
            padding: 0.34rem 0.50rem;
            border-radius: 10px;
            box-shadow:
                0 10px 26px rgba(32,240,255,0.08),
                0 0 0 1px rgba(255,255,255,0.04);
        }

        .welcome-logo-frame img {
            max-width: 118px;
            max-height: 126px;
            width: auto;
            height: auto;
            object-fit: contain;
            display: block;
        }

        .welcome-title {
            margin: 0.10rem 0 0.20rem 0;
            color: #F7FBFF;
            font-size: 1.34rem;
            line-height: 1.12;
            font-weight: 850;
            letter-spacing: 0.018em;
            text-transform: uppercase;
        }

        .welcome-subtitle {
            margin: 0.18rem 0 0.22rem 0;
            color: #20F0FF;
            font-size: 0.82rem;
            font-weight: 800;
            letter-spacing: 0.085em;
            text-transform: uppercase;
        }

        .welcome-caption {
            margin: 0.20rem 0 0 0;
            color: #DDE8F7;
            font-size: 0.66rem;
            font-weight: 500;
        }

        div[data-testid="stForm"] {
            border: 1px solid rgba(62,90,130,0.28) !important;
            background:
                linear-gradient(
                    180deg,
                    rgba(6,15,29,0.92) 0%,
                    rgba(4,10,20,0.96) 100%
                ) !important;
            border-radius: 15px !important;
            padding: 1.05rem 0.92rem 1.00rem 0.92rem !important;
            box-shadow: 0 10px 34px rgba(0,0,0,0.18) !important;
        }

        div[data-testid="stTextInput"] {
            margin-bottom: 0.10rem !important;
        }

        div[data-testid="stTextInput"] label,
        div[data-testid="stTextInput"] p {
            color: #FFFFFF !important;
            font-weight: 750 !important;
            font-size: 0.73rem !important;
        }

        div[data-testid="stTextInput"] input {
            background: rgba(255,255,255,0.98) !important;
            color: #1B3152 !important;
            border-radius: 10px !important;
            border: 1px solid rgba(87,115,150,0.34) !important;
            padding-top: 0.70rem !important;
            padding-bottom: 0.70rem !important;
            font-size: 0.86rem !important;
            min-height: 3.05rem !important;
        }

        div[data-testid="stFormSubmitButton"] button {
            background:
                linear-gradient(
                    135deg,
                    #1B2C49 0%,
                    #223556 100%
                ) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(32,240,255,0.28) !important;
            border-radius: 10px !important;
            font-weight: 800 !important;
            font-size: 0.72rem !important;
            padding: 0.48rem 0.80rem !important;
            min-height: 2.85rem !important;
            box-shadow: 0 8px 22px rgba(0,0,0,0.16) !important;
        }

        div[data-testid="stFormSubmitButton"] button:hover {
            border-color: rgba(32,240,255,0.55) !important;
            color: #FFFFFF !important;
        }

        .welcome-help {
            color: #AFC2D8;
            font-family:
                Inter,
                "Segoe UI",
                "Helvetica Neue",
                Arial,
                sans-serif;
            font-size: 0.88rem;
            font-weight: 600;
            line-height: 1.55;
            letter-spacing: 0.018em;
            margin-top: 0.78rem;
            margin-bottom: 0.28rem;
            text-align: center;
        }

        .welcome-corporate-footer {
            margin: 0.58rem auto 0 auto;
            padding: 0.34rem 0.40rem 0.18rem 0.40rem;
            text-align: center;
            color: #A9BDD3;
            font-family:
                Inter,
                "Segoe UI",
                "Helvetica Neue",
                Arial,
                sans-serif;
            font-size: 0.82rem;
            font-weight: 500;
            line-height: 1.65;
            letter-spacing: 0.018em;
        }

        .welcome-corporate-footer strong {
            color: #F3F7FC;
            font-size: 0.86rem;
            font-weight: 700;
            letter-spacing: 0.012em;
        }

        @media (max-height: 760px) {
            .block-container {
                min-height: 100vh !important;
                padding-top: 1.25rem !important;
                padding-bottom: 1.25rem !important;
            }
        }

        @media (max-width: 900px) {
            .welcome-title {
                font-size: 1.12rem;
            }

            .welcome-subtitle {
                font-size: 0.70rem;
            }

            .welcome-logo-frame img {
                max-width: 105px;
                max-height: 112px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Columna central REAL de Streamlit:
    # controla también el ancho de inputs y botón.
    _, login_center, _ = st.columns(
        [1.0, 1.22, 1.0],
        gap="small",
    )

    with login_center:
        welcome_hero_html = (
            f'<div class="welcome-shell">'
            f'<div class="welcome-hero">'
            f'<div class="welcome-logo-wrap">'
            f'<div class="welcome-logo-frame">'
            f'<img src="{DICASA_WELCOME_LOGO_URI}" '
            f'alt="Logo DICASA S.A." />'
            f'</div>'
            f'</div>'
            f'<div class="welcome-title">'
            f'Dashboard Cierre de Productos DICASA S.A'
            f'</div>'
            f'<div class="welcome-subtitle">'
            f'Executive Analytics &amp; Intelligence Portal'
            f'</div>'
            f'<div class="welcome-caption">'
            f'Módulo de Autenticación Comercial Restringida'
            f'</div>'
            f'</div>'
            f'</div>'
        )

        st.markdown(
            welcome_hero_html,
            unsafe_allow_html=True,
        )

        with st.form(
            "login_form_dicasa",
            clear_on_submit=False,
        ):
            user_value = st.text_input(
                "Usuario",
                placeholder="Ingrese su usuario autorizado...",
                key="login_user",
            )

            password_value = st.text_input(
                "Password",
                type="password",
                placeholder="Ingrese su credencial comercial...",
                key="login_password",
            )

            submitted = st.form_submit_button(
                "🔐 Ingresar al Dashboard",
                use_container_width=True,
            )

        st.markdown(
            '<div class="welcome-help">'
            'Acceso exclusivo para personal comercial autorizado.'
            '</div>',
            unsafe_allow_html=True,
        )

        welcome_footer_html = (
            '<div class="welcome-corporate-footer">'
            'Created by '
            '<strong>'
            'IT Department | Distribuidora Centroamérica S.A.'
            '</strong>'
            '<br>'
            'Derechos Reservados © 2026'
            '</div>'
        )

        st.markdown(
            welcome_footer_html,
            unsafe_allow_html=True,
        )

    if submitted:
        if (
            user_value.strip() == DEFAULT_DASH_USER
            and password_value == DEFAULT_DASH_PASSWORD
        ):
            st.session_state["dashboard_authenticated"] = True
            st.session_state["dashboard_user"] = user_value.strip()
            st.success(
                "Acceso autorizado. Cargando dashboard..."
            )
            st.rerun()
        else:
            st.error(
                "Usuario o password incorrectos. "
                "Verifique sus credenciales."
            )

    st.stop()


if "dashboard_authenticated" not in st.session_state:
    st.session_state["dashboard_authenticated"] = False


if not st.session_state["dashboard_authenticated"]:
    render_welcome_login_screen()

# ============================================================
# TÍTULO
# ============================================================

dashboard_header_html = (
    f'<div class="dash-header dicasa-corporate-header">'
    f'<div class="dicasa-header-logo-zone">'
    f'<div class="dicasa-logo-frame">'
    f'<img class="dicasa-logo-image" '
    f'src="{DICASA_LOGO_URI}" '
    f'alt="Logo DICASA S.A." />'
    f'</div>'
    f'</div>'
    f'<div class="dicasa-header-copy">'
    f'<div class="dicasa-title-line">'
    f'<span class="dicasa-dashboard-title">'
    f'{APP_TITLE}'
    f'</span>'
    f'<span class="dicasa-title-separator">—</span>'
    f'<span class="dicasa-company-name">'
    f'DICASA S.A.'
    f'</span>'
    f'</div>'
    f'<div class="dash-subtitle dicasa-subtitle">'
    f'Ventas, utilidad, inventario y análisis ejecutivo '
    f'por familia y producto'
    f'</div>'
    f'</div>'
    f'<div class="dicasa-header-badge-zone">'
    f'<div class="dash-badge">EXECUTIVE ANALYTICS</div>'
    f'</div>'
    f'</div>'
)

st.markdown(
    dashboard_header_html,
    unsafe_allow_html=True,
)


# ============================================================
# CARGAR ARCHIVO
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="sidebar-brand-main">▥ CIERRE ANALYTICS</div>
            <div class="sidebar-brand-sub">Executive Product Intelligence</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.header(
        "📁 Datos"
    )


    uploaded_file = st.file_uploader(

        "Cargar archivo Excel",

        type=[
            "xlsx"
        ],

        help=(

            "Seleccione el archivo "
            "CIERRE PRODUCTOS.xlsx. "
            "Si no carga un archivo, "
            "la aplicación intentará "
            "leerlo desde la carpeta "
            "del programa."

        ),

    )


    st.caption(

        "Si no selecciona un archivo, "
        "se intentará leer automáticamente "
        "`CIERRE PRODUCTOS.xlsx`."

    )


df: Optional[
    pd.DataFrame
] = None


load_info: Optional[
    Dict[str, object]
] = None


cleaning_warnings: List[
    str
] = []


source_description = ""


try:


    # --------------------------------------------------------
    # Archivo seleccionado manualmente
    # --------------------------------------------------------

    if uploaded_file is not None:


        with st.spinner(
            "Procesando archivo cargado..."
        ):


            (

                df,

                load_info,

                cleaning_warnings,

            ) = load_uploaded_excel(

                uploaded_file.getvalue()

            )


            source_description = (
                uploaded_file.name
            )


    # --------------------------------------------------------
    # Archivo automático
    # --------------------------------------------------------

    else:


        default_file = (
            find_default_excel()
        )


        if default_file is None:


            st.info(

                "📂 Cargue el archivo "
                "**CIERRE PRODUCTOS.xlsx** "
                "desde la barra lateral o "
                "colóquelo en la misma carpeta "
                "del archivo Python."

            )


            st.stop()


        with st.spinner(
            "Cargando archivo local..."
        ):


            modified_time = (
                os.path.getmtime(
                    default_file
                )
            )


            (

                df,

                load_info,

                cleaning_warnings,

            ) = load_local_excel(

                str(
                    default_file
                ),

                modified_time,

            )


            source_description = (
                default_file.name
            )


except Exception as exc:


    st.error(

        "No fue posible procesar "
        "el archivo Excel."

    )


    st.exception(
        exc
    )


    st.stop()


if df is None:


    st.error(

        "No se pudo cargar información "
        "del archivo."

    )


    st.stop()


# ============================================================
# VALIDAR COLUMNAS
# ============================================================

missing_columns = (
    validate_required_columns(
        df
    )
)


if missing_columns:


    st.error(

        "❌ El archivo no contiene "
        "todas las columnas necesarias."

    )


    st.markdown(

        "### Columnas que no "
        "fueron identificadas"

    )


    for column in missing_columns:


        st.write(
            f"- `{column}`"
        )


    with st.expander(
        "Ver columnas encontradas"
    ):


        st.write(
            list(
                df.columns
            )
        )


    if load_info:


        with st.expander(

            "Información de detección "
            "de encabezados"

        ):


            st.write(

                f"**Hoja:** "
                f"{load_info.get('sheet_name', '')}"

            )


            st.write(

                f"**Fila(s) del encabezado:** "
                f"{load_info.get('header_rows', [])}"

            )


            st.write(

                "**Columnas requeridas reconocidas:** "
                f"{load_info.get('recognized_columns', 0)} "
                f"de {len(REQUIRED_COLUMNS)}"

            )


            st.json(

                load_info.get(
                    "mapping_report",
                    {},
                )

            )


    st.stop()


# ============================================================
# INFORMACIÓN DEL ARCHIVO
# ============================================================

with st.sidebar:


    st.success(

        f"✅ Archivo cargado: "
        f"{source_description}"

    )


    if load_info:


        st.caption(

            f"Hoja: "
            f"{load_info.get('sheet_name', '')}"

        )


        st.caption(

            "Encabezado detectado "
            "en fila(s): "

            +

            ", ".join(

                str(value)

                for value
                in load_info.get(
                    "header_rows",
                    [],
                )

            )

        )


# ============================================================
# FILTROS
# ============================================================

filtered_df = df.copy()


def _format_filter_code(value: object) -> str:
    """Normaliza códigos para mostrarlos sin .0 innecesario."""
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.endswith(".0"):
        integer_part = text[:-2]
        if integer_part.replace("-", "", 1).isdigit():
            text = integer_part

    return text


def _clear_product_dependent_filters() -> None:
    """
    Sincroniza el filtro lateral de familia con Análisis de productos
    y limpia cualquier selección de producto que ya no sea válida.
    """
    st.session_state["sidebar_product_filter"] = []
    st.session_state["selected_product_from_treemap"] = None
    st.session_state["selected_products_from_treemap"] = []
    st.session_state["selected_product_family"] = None

    sidebar_families = st.session_state.get(
        "sidebar_family_filter",
        [],
    )

    # Si el usuario escogió una sola familia lateralmente,
    # el selector interno de Análisis de productos usa esa misma.
    if len(sidebar_families) == 1:
        st.session_state[
            "family_product_chart"
        ] = sidebar_families[0]

    elif len(sidebar_families) > 1:
        current_internal_family = st.session_state.get(
            "family_product_chart"
        )

        if current_internal_family not in sidebar_families:
            st.session_state[
                "family_product_chart"
            ] = sidebar_families[0]


with st.sidebar:

    st.divider()
    st.header("🔎 Filtros")

    # --------------------------------------------------------
    # Familia
    # --------------------------------------------------------

    family_values = sorted(
        str(value)
        for value in df["FAMILIA"].dropna().unique()
        if str(value).strip()
    )

    # Sanitizar estado antes de crear el widget.
    current_family_state = st.session_state.get(
        "sidebar_family_filter",
        [],
    )

    valid_family_state = [
        value
        for value in current_family_state
        if value in family_values
    ]

    if current_family_state != valid_family_state:
        st.session_state["sidebar_family_filter"] = (
            valid_family_state
        )

    selected_families = st.multiselect(
        "FAMILIA",
        options=family_values,
        placeholder="Todas las familias",
        help=(
            "Si no selecciona nada, se incluyen todas. "
            "Al cambiar la familia se limpia el filtro de producto."
        ),
        key="sidebar_family_filter",
        on_change=_clear_product_dependent_filters,
    )

    if selected_families:
        filtered_df = filtered_df[
            filtered_df["FAMILIA"]
            .astype(str)
            .isin(selected_families)
        ].copy()

    # --------------------------------------------------------
    # Producto — Código + Descripción
    # --------------------------------------------------------

    product_filter_source = (
        filtered_df[
            ~filtered_df[
                "AJUSTE_PRODUCTO"
            ].fillna(False)
        ]
        .copy()
    )

    # Construcción vectorizada del texto de búsqueda.
    # Evita ejecutar una función Python fila por fila en cada rerun.
    product_code_text = (
        product_filter_source["#COD."]
        .astype("string")
        .fillna("")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True,
        )
    )

    product_filter_source["_PRODUCTO_FILTRO"] = (
        product_code_text
        + " - "
        + product_filter_source["DESCRIPCION"]
        .astype("string")
        .fillna("")
        .str.strip()
    )

    product_filter_values = sorted(
        value
        for value in product_filter_source[
            "_PRODUCTO_FILTRO"
        ]
        .dropna()
        .astype(str)
        .unique()
        if value.strip(" -")
    )

    # Sanitizar estado si la familia cambió y alguno de los
    # productos seleccionados ya no pertenece a las opciones.
    current_product_state = st.session_state.get(
        "sidebar_product_filter",
        [],
    )

    valid_product_state = [
        value
        for value in current_product_state
        if value in product_filter_values
    ]

    if current_product_state != valid_product_state:
        st.session_state["sidebar_product_filter"] = (
            valid_product_state
        )

    selected_products_filter = st.multiselect(
        "Producto / CÓDIGO + DESCRIPCIÓN",
        options=product_filter_values,
        placeholder="Todos los productos",
        help=(
            "Busque escribiendo el código o cualquier parte "
            "de la descripción."
        ),
        key="sidebar_product_filter",
    )

    if selected_products_filter:
        selected_product_rows = (
            product_filter_source[
                product_filter_source[
                    "_PRODUCTO_FILTRO"
                ]
                .astype(str)
                .isin(selected_products_filter)
            ]
            .index
        )

        filtered_df = filtered_df.loc[
            filtered_df.index.isin(
                selected_product_rows
            )
        ].copy()

    # --------------------------------------------------------
    # Métrica fija
    # --------------------------------------------------------

    selected_metric = "$VENTAS MES ACTUAL"

    # --------------------------------------------------------
    # TOP global de visualización
    # --------------------------------------------------------
    # Se usa selectbox para máxima compatibilidad y respuesta inmediata.

    valid_top_values = ["TODOS", 10, 20, 50]

    if "global_top_n" not in st.session_state:
        st.session_state["global_top_n"] = "TODOS"

    if st.session_state["global_top_n"] not in valid_top_values:
        st.session_state["global_top_n"] = "TODOS"

    top_choice = st.selectbox(
        "Cantidad de productos / familias a mostrar",
        options=valid_top_values,
        key="global_top_n",
        help=(
            "TODOS conserva todos los datos filtrados. "
            "Las gráficas pesadas limitan automáticamente "
            "la cantidad visible para mantener el dashboard rápido."
        ),
    )

    top_label = (
        "TODOS"
        if top_choice == "TODOS"
        else str(top_choice)
    )

    # Límite visual global para evitar que las pestañas ocultas
    # intenten dibujar cientos o miles de categorías en cada rerun.
    top_n = (
        50
        if top_choice == "TODOS"
        else int(top_choice)
    )

    # Límite de datos: TODOS conserva todos los registros en tablas.
    top_data_n = (
        max(len(filtered_df), 1)
        if top_choice == "TODOS"
        else int(top_choice)
    )

    st.caption(
        (
            "TODOS aplicado al dashboard"
            if top_choice == "TODOS"
            else f"TOP {top_label} aplicado al dashboard"
        )
        + f" · Registros filtrados: {len(filtered_df):,}"
    )

# ============================================================
# VALIDAR FILTROS
# ============================================================

if filtered_df.empty:


    st.warning(

        "⚠️ Los filtros seleccionados "
        "no devolvieron registros. "
        "Modifique los filtros para continuar."

    )


    st.stop()


# ============================================================
# CREAR RESÚMENES
# ============================================================

family_df = (
    family_summary(
        filtered_df
    )
)


product_df = (
    product_summary(
        filtered_df
    )
)


# ============================================================
# KPI
# ============================================================

total_sales = (
    safe_sum(
        df[
            "$VENTAS MES ACTUAL"
        ]
    )
)


total_profit = (
    safe_sum(
        df[
            "$UTILIDAD MES ACTUAL"
        ]
    )
)


total_inventory = (
    safe_sum(
        (
            df[
                "INVENTARIO TOTAL"
            ]
            .fillna(0)
            *
            df[
                "COSTO ACTUAL"
            ]
            .fillna(0)
        )
    )
)


lost_sales = (
    safe_sum(
        df[
            "#VENTAS PERDIDAS"
        ]
    )
)


average_margin = (
    safe_mean(
        df[
            "% MAGERN S/VENTA"
        ]
    )
)


product_count = (

    df[
        [
            "#COD.",
            "DESCRIPCION",
        ]
    ]

    .drop_duplicates()

    .shape[0]

)


family_count = (

    df[
        "FAMILIA"
    ]

    .dropna()

    .nunique()

)


# ============================================================
# FAMILIA LÍDER
# ============================================================

family_sales = (

    df

    .groupby(

        "FAMILIA",

        dropna=False,

        as_index=False,

    )[
        "$VENTAS MES ACTUAL"
    ]

    .sum()

    .sort_values(

        "$VENTAS MES ACTUAL",

        ascending=False,

    )

)


if not family_sales.empty:


    top_family_name = str(

        family_sales.iloc[
            0
        ][
            "FAMILIA"
        ]

    )


    top_family_sales = float(

        family_sales.iloc[
            0
        ][
            "$VENTAS MES ACTUAL"
        ]

    )


else:


    top_family_name = (
        "Sin datos"
    )


    top_family_sales = 0.0


# ============================================================
# PRODUCTO LÍDER
# ============================================================

product_sales = (

    df

    .groupby(

        [
            "#COD.",
            "DESCRIPCION",
        ],

        dropna=False,

        as_index=False,

    )[
        "$VENTAS MES ACTUAL"
    ]

    .sum()

    .sort_values(

        "$VENTAS MES ACTUAL",

        ascending=False,

    )

)


if not product_sales.empty:


    top_product_code = str(

        product_sales.iloc[
            0
        ][
            "#COD."
        ]

    )


    top_product_description = str(

        product_sales.iloc[
            0
        ][
            "DESCRIPCION"
        ]

    )


    top_product_sales = float(

        product_sales.iloc[
            0
        ][
            "$VENTAS MES ACTUAL"
        ]

    )


else:


    top_product_code = ""


    top_product_description = (
        "Sin datos"
    )


    top_product_sales = 0.0





def render_executive_html_table(
    dataframe: pd.DataFrame,
    column_labels: Dict[str, str],
    formats: Optional[Dict[str, str]] = None,
    max_height: int = 520,
    compact: bool = False,
):
    """
    Renderiza una tabla HTML oscura y ejecutiva como el ejemplo visual:
    - encabezado azul/cian
    - fondo azul oscuro
    - líneas finas
    - filas uniformes
    - números alineados a la derecha
    """

    if dataframe.empty:
        st.info("No se encontraron registros.")
        return

    formats = formats or {}

    numeric_columns = set(
        dataframe.select_dtypes(
            include="number"
        ).columns.tolist()
    )

    def format_value(column: str, value: object) -> str:
        if value is None or pd.isna(value):
            return ""

        fmt = formats.get(column)

        try:
            if fmt == "money":
                return f"${float(value):,.2f}"
            if fmt == "percent":
                return f"{float(value):,.2f}%"
            if fmt == "integer":
                return f"{float(value):,.0f}"
            if fmt == "number2":
                return f"{float(value):,.2f}"
        except (TypeError, ValueError):
            pass

        return str(value)

    # Encabezados
    header_cells = "".join(
        (
            f'<th class="exec-th">'
            f'{html.escape(column_labels.get(column, column))}'
            f'</th>'
        )
        for column in dataframe.columns
    )

    # Filas
    body_rows = []

    for _, row in dataframe.iterrows():
        cells = []

        for column in dataframe.columns:
            value = format_value(
                column,
                row[column],
            )

            alignment_class = (
                "exec-num"
                if column in numeric_columns
                else "exec-text"
            )

            cells.append(
                f'<td class="exec-td {alignment_class}">'
                f'{html.escape(value)}'
                f'</td>'
            )

        body_rows.append(
            "<tr class='exec-tr'>"
            + "".join(cells)
            + "</tr>"
        )

    compact_class = " exec-table-compact" if compact else ""

    table_html = f"""
<div class="exec-table-shell{compact_class}" style="max-height:{max_height}px;">
<table class="exec-table{compact_class}">
<thead>
<tr>{header_cells}</tr>
</thead>
<tbody>
{''.join(body_rows)}
</tbody>
</table>
</div>
"""

    st.markdown(
        table_html,
        unsafe_allow_html=True,
    )

# ============================================================
# ANÁLISIS EN PESTAÑAS
# ============================================================

st.markdown("---")
st.markdown("## Exploración ejecutiva")

(
    tab_resumen_ejecutivo,
    tab_familia,
    tab_productos,
    tab_ventas_utilidad_familia,
    tab_pareto,
    tab_insights,
    tab_alertas,
    tab_descargas,
) = st.tabs(
    [
        "📊 Resumen ejecutivo",
        "🗂️ Análisis por familia",
        "📦 Análisis de productos",
        "💰 Ventas y utilidad por familia",
        "📊 Pareto",
        "💡 Insights",
        "🚨 Alertas de productos",
        "⬇️ Descargar información",
    ]
)


# ============================================================
# PESTAÑA 1 — RESUMEN EJECUTIVO
# ============================================================

with tab_resumen_ejecutivo:

    st.markdown("## Resumen ejecutivo")

    # ------------------------------------------------------------
    # FILA 1 — KPI FINANCIEROS Y OPERATIVOS
    # ------------------------------------------------------------
    k1, k2, k3, k4, k5 = st.columns(5, gap="small")

    with k1:
        render_exec_kpi("💵 Ventas mes", money(total_sales), "#43B9E6")

    with k2:
        render_exec_kpi("📈 Utilidad mes", money(total_profit), "#27C99B")

    with k3:
        render_exec_kpi(
            "📦 Costo total del inventario",
            money(total_inventory),
            "#A984D8",
        )

    with k4:
        render_exec_kpi("⚠️ Ventas perdidas", quantity(lost_sales, 2), "#E96573")

    with k5:
        render_exec_kpi("◉ Margen promedio", percent(average_margin), "#D7AE58")

    st.markdown(
        '<div style="height:28px;"></div>',
        unsafe_allow_html=True,
    )

    # ------------------------------------------------------------
    # FILA 2 — DIMENSIÓN Y LIDERAZGO
    # ------------------------------------------------------------
    k6, k7, k8, k9 = st.columns(4, gap="small")

    with k6:
        render_exec_kpi("▣ Productos", f"{product_count:,}", "#6F91B2")

    with k7:
        render_exec_kpi("◆ Familias", f"{family_count:,}", "#C79356")

    with k8:
        st.markdown(
            f"""<div class="leader-card" style="--accent:#27C99B;">
    <div class="leader-label">Familia mayor venta</div>
    <div class="leader-name">{html.escape(top_family_name)}</div>
    <div class="leader-value">{money(top_family_sales)}</div>
    </div>""",
            unsafe_allow_html=True,
        )

    with k9:
        st.markdown(
            f"""<div class="leader-card" style="--accent:#D7AE58;">
    <div class="leader-label">Producto mayor venta</div>
    <div class="leader-name">{html.escape(top_product_description)}</div>
    <div class="leader-code">Código: {html.escape(top_product_code)}</div>
    <div class="leader-value">{money(top_product_sales)}</div>
    </div>""",
            unsafe_allow_html=True,
        )


# ============================================================
# PESTAÑA 2 — ANÁLISIS POR FAMILIA
# ============================================================

with tab_familia:

    # ========================================================
    # CONFIGURACIÓN VISUAL DE FAMILIAS
    # ========================================================
    family_total_count = int(
        filtered_df["FAMILIA"]
        .astype("string")
        .fillna("SIN FAMILIA")
        .nunique()
    )

    family_visual_limit = (
        min(30, family_total_count)
        if top_choice == "TODOS"
        else min(int(top_choice), family_total_count)
    )

    family_main_title = (
        f"TOP {family_visual_limit} DE {family_total_count:,} FAMILIAS"
        if top_choice == "TODOS"
        else f"TOP {family_visual_limit} FAMILIAS"
    )

    # ========================================================
    # ESTADO DEL FILTRO INTERACTIVO DE FAMILIA
    # ========================================================

    if "family_selection_reset_token" not in st.session_state:
        st.session_state["family_selection_reset_token"] = 0

    def clear_family_selection() -> None:
        """
        Restablece por completo los filtros de familia/producto:
        - filtros laterales de FAMILIA y PRODUCTO
        - selección de familia en gráficas
        - selección de productos del treemap
        - estados internos dependientes
        """
        # Filtros laterales
        st.session_state["sidebar_family_filter"] = []
        st.session_state["sidebar_product_filter"] = []

        # Selecciones de familia
        st.session_state["selected_family_from_chart"] = None
        st.session_state["family_legend_radio"] = None
        st.session_state["family_selection_reset_token"] += 1

        # Selecciones de producto dependientes
        st.session_state["selected_product_from_treemap"] = None
        st.session_state["selected_products_from_treemap"] = []
        st.session_state["selected_product_family"] = None

        # Forzar recreación limpia del treemap si existe
        st.session_state["treemap_click_reset_token"] = (
            st.session_state.get(
                "treemap_click_reset_token",
                0,
            )
            + 1
        )

    # ========================================================
    # SELECTOR INTERACTIVO DE FAMILIA
    # ========================================================

    family_sales_selector = (
        filtered_df
        .groupby(
            "FAMILIA",
            dropna=False,
            as_index=False,
        )["$VENTAS MES ACTUAL"]
        .sum()
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .head(family_visual_limit)
        .reset_index(drop=True)
    )

    family_sales_selector["FAMILIA_COMPLETA"] = (
        family_sales_selector["FAMILIA"]
        .astype("string")
        .fillna("SIN FAMILIA")
        .astype(str)
    )

    def _short_family_name(value: object, max_chars: int = 19) -> str:
        """
        Abrevia solo la etiqueta visual.
        Conserva el final del nombre para diferenciar familias
        como PAPEL HIGIENICO HD y PAPEL HIGIENICO SD.
        """
        text = str(value).strip()

        if len(text) <= max_chars:
            return text

        parts = text.split()

        if len(parts) >= 2:
            suffix = parts[-1]

            # Conservar sufijos cortos y significativos: HD, SD, CR, etc.
            if len(suffix) <= 4:
                available = max_chars - len(suffix) - 2

                if available >= 5:
                    return (
                        text[:available].rstrip()
                        + "… "
                        + suffix
                    )

        return text[: max_chars - 1].rstrip() + "…"

    family_sales_selector["FAMILIA_CORTA"] = (
        family_sales_selector["FAMILIA_COMPLETA"]
        .apply(_short_family_name)
    )

    # IMPORTANTE:
    # FAMILIA_COMPLETA se usa como categoría real del eje.
    # FAMILIA_CORTA se usa únicamente como texto visual.
    # Esto evita fusionar familias diferentes con abreviaciones iguales.

    family_sales_selector["ETIQUETA_VALOR"] = ""

    # Mostrar valores solo en las barras principales para evitar ruido visual.
    top_labels_count = min(
        12,
        len(family_sales_selector),
    )

    if top_labels_count:
        family_sales_selector.loc[
            : top_labels_count - 1,
            "ETIQUETA_VALOR",
        ] = (
            family_sales_selector.loc[
                : top_labels_count - 1,
                "$VENTAS MES ACTUAL",
            ]
            .apply(
                lambda value:
                f"${float(value):,.0f}"
                if pd.notna(value)
                else ""
            )
        )

    available_family_names = set(
        family_sales_selector["FAMILIA_COMPLETA"].tolist()
    )

    if (
        "selected_family_from_chart"
        not in st.session_state
    ):
        st.session_state[
            "selected_family_from_chart"
        ] = None

    selected_family = st.session_state.get(
        "selected_family_from_chart"
    )

    # Si los filtros laterales excluyen la familia seleccionada,
    # limpiar automáticamente la selección.
    if (
        selected_family
        and selected_family
        not in available_family_names
    ):
        st.session_state[
            "selected_family_from_chart"
        ] = None
        selected_family = None


    # Base de datos que alimentará las demás visualizaciones.
    if selected_family:
        family_focus_df = (
            filtered_df[
                filtered_df["FAMILIA"]
                .astype("string")
                .fillna("SIN FAMILIA")
                .astype(str)
                .eq(
                    selected_family
                )
            ]
            .copy()
        )
    else:
        family_focus_df = (
            filtered_df.copy()
        )



    # ========================================================
    # KPI INTERACTIVOS — ANÁLISIS POR FAMILIA
    # ========================================================
    # Los KPI usan family_focus_df:
    # - si hay una familia seleccionada, muestran solo esa familia;
    # - si no hay selección, muestran el conjunto filtrado actual.

    family_kpi_label = (
        str(selected_family)
        if selected_family
        else "TODAS LAS FAMILIAS FILTRADAS"
    )

    family_kpi_sales = safe_sum(
        family_focus_df[
            "$VENTAS MES ACTUAL"
        ]
    )

    family_kpi_inventory = safe_sum(
        family_focus_df[
            "INVENTARIO TOTAL"
        ]
    )

    family_kpi_profit = safe_sum(
        family_focus_df[
            "$UTILIDAD MES ACTUAL"
        ]
    )

    family_kpi_lost_sales = safe_sum(
        family_focus_df[
            "#VENTAS PERDIDAS"
        ]
    )

    family_kpi_inventory_value = safe_sum(
        (
            family_focus_df[
                "INVENTARIO TOTAL"
            ]
            .fillna(0)
            *
            family_focus_df[
                "COSTO ACTUAL"
            ]
            .fillna(0)
        )
    )

    family_kpi_pack_values = (
        family_focus_df[
            "EMPAQ_ANALISIS"
        ]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    family_kpi_pack_values = [
        value
        for value in family_kpi_pack_values.unique().tolist()
        if value
        and value.upper() not in {"NAN", "NONE"}
    ]

    if len(family_kpi_pack_values) == 1:
        family_kpi_pack = family_kpi_pack_values[0]
    elif len(family_kpi_pack_values) > 1:
        family_kpi_pack = "Varios empaques"
    else:
        family_kpi_pack = "Sin empaque"

    family_kpi_top_product = (
        family_focus_df
        .groupby(
            [
                "#COD.",
                "DESCRIPCION",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            {
                "$VENTAS MES ACTUAL": "sum",
                "INVENTARIO TOTAL": "sum",
                "EMPAQ_ANALISIS": (
                    lambda series:
                    next(
                        (
                            str(value).strip()
                            for value in series
                            if pd.notna(value)
                            and str(value).strip()
                            and str(value).strip().upper()
                            not in {"NAN", "NONE"}
                        ),
                        "",
                    )
                ),
            }
        )
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    if not family_kpi_top_product.empty:
        family_kpi_top_row = (
            family_kpi_top_product.iloc[0]
        )

        family_kpi_top_name = str(
            family_kpi_top_row[
                "DESCRIPCION"
            ]
        ).strip()

        family_kpi_top_code = str(
            family_kpi_top_row[
                "#COD."
            ]
        ).strip()

        family_kpi_top_sales = float(
            family_kpi_top_row[
                "$VENTAS MES ACTUAL"
            ]
        )

        family_kpi_top_inventory = float(
            family_kpi_top_row[
                "INVENTARIO TOTAL"
            ]
        )

        family_kpi_top_pack = str(
            family_kpi_top_row[
                "EMPAQ_ANALISIS"
            ]
        ).strip() or "Sin empaque"
    else:
        family_kpi_top_name = "SIN PRODUCTO"
        family_kpi_top_code = "-"
        family_kpi_top_sales = 0.0
        family_kpi_top_inventory = 0.0
        family_kpi_top_pack = "Sin empaque"

    # --------------------------------------------------------
    # Estilo de los KPI, equivalente al usado en productos.
    # --------------------------------------------------------
    st.markdown(
        """
        <style>
        .fa-kpi-card {
            min-height: 132px;
            height: 132px;
            position: relative;
            overflow: hidden;
            border: 1px solid #29445F;
            border-radius: 14px;
            background:
                linear-gradient(145deg,#10243A 0%,#0D1C2D 100%);
            padding: 7px 9px 6px 9px;
            display: flex;
            flex-direction: column;
        }

        .fa-kpi-card::before {
            content: "";
            position: absolute;
            left: 12px;
            right: 12px;
            top: 0;
            height: 3px;
            border-radius: 0 0 5px 5px;
            background: var(--fa-accent);
        }

        .fa-kpi-head {
            display: flex;
            align-items: center;
            gap: 7px;
            min-height: 29px;
        }

        .fa-kpi-icon {
            width: 25px;
            height: 25px;
            min-width: 25px;
            border-radius: 8px;
            border: 1px solid var(--fa-accent);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: .78rem;
        }

        .fa-kpi-title {
            color: #4DD5FF !important;
            font-size: .70rem !important;
            line-height: 1.15 !important;
            font-weight: 900 !important;
            letter-spacing: .02em !important;
        }

        .fa-kpi-divider {
            height: 1px;
            background: #24415D;
            margin: 2px 0 4px 0;
        }

        .fa-kpi-main {
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 43px;
        }

        .fa-kpi-value {
            color: var(--fa-value) !important;
            font-size: clamp(1.26rem,1.52vw,1.62rem) !important;
            line-height: 1 !important;
            font-weight: 950 !important;
            letter-spacing: -.025em !important;
        }

        .fa-kpi-footer {
            min-height: 25px;
            border: 1px solid #274B6B;
            border-radius: 9px;
            padding: 4px 7px;
            display: flex;
            align-items: center;
            gap: 7px;
            color: #D8E5F1 !important;
            font-size: .64rem !important;
            line-height: 1.15 !important;
            font-weight: 760 !important;
        }

        .fa-kpi-card.sales {
            --fa-accent:#22D3EE;
            --fa-value:#53DDF8;
        }

        .fa-kpi-card.inventory {
            --fa-accent:#A56CF5;
            --fa-value:#B174FF;
        }

        .fa-kpi-card.leader {
            --fa-accent:#F3C64E;
            --fa-value:#F3C64E;
        }

        .fa-kpi-card.profit {
            --fa-accent:#28D7A1;
            --fa-value:#35E4B0;
        }

        .fa-kpi-card.stockvalue {
            --fa-accent:#3DA8FF;
            --fa-value:#55B8FF;
        }

        .fa-kpi-card.lostsales {
            --fa-accent:#FF596B;
            --fa-value:#FF6C7A;
        }

        .fa-leader-main {
            flex: 1;
            min-height: 47px;
        }

        .fa-leader-name {
            color: #F4F7FB !important;
            font-size: .72rem !important;
            line-height: 1.15 !important;
            font-weight: 900 !important;
            margin-bottom: 3px !important;
        }

        .fa-leader-code {
            color: #AFC6DC !important;
            font-size: .62rem !important;
            font-weight: 740 !important;
            margin-bottom: 3px !important;
        }

        .fa-leader-value {
            color: #F3C64E !important;
            font-size: 1.18rem !important;
            font-weight: 950 !important;
            line-height: 1 !important;
        }

        @media (max-width: 1350px) {
            .fa-kpi-title {
                font-size: .64rem !important;
            }

            .fa-kpi-footer {
                font-size: .59rem !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        "<div style='height:12px'></div>",
        unsafe_allow_html=True,
    )

    fa1, fa2, fa3 = st.columns(
        3,
        gap="small",
    )

    with fa1:
        st.markdown(
            (
                '<div class="fa-kpi-card sales" '
                'style="--fa-accent:#22D3EE;--fa-value:#53DDF8;">'
                '<div class="fa-kpi-head">'
                '<div class="fa-kpi-icon">💵</div>'
                '<div class="fa-kpi-title">VENTA TOTAL POR FAMILIA</div>'
                '</div>'
                '<div class="fa-kpi-divider"></div>'
                '<div class="fa-kpi-main">'
                '<div class="fa-kpi-value">'
                f'{money(family_kpi_sales)}'
                '</div>'
                '</div>'
                '<div class="fa-kpi-footer">'
                f'🏷️ {html.escape(family_kpi_label)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with fa2:
        st.markdown(
            (
                '<div class="fa-kpi-card inventory" '
                'style="--fa-accent:#A56CF5;--fa-value:#B174FF;">'
                '<div class="fa-kpi-head">'
                '<div class="fa-kpi-icon">📦</div>'
                '<div class="fa-kpi-title">INVENTARIO TOTAL POR FAMILIA</div>'
                '</div>'
                '<div class="fa-kpi-divider"></div>'
                '<div class="fa-kpi-main">'
                '<div class="fa-kpi-value">'
                f'{quantity(family_kpi_inventory, 0)}'
                '</div>'
                '</div>'
                '<div class="fa-kpi-footer">'
                f'◈ EMPAQ: {html.escape(family_kpi_pack)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with fa3:
        st.markdown(
            (
                '<div class="fa-kpi-card leader" '
                'style="--fa-accent:#F3C64E;--fa-value:#F3C64E;">'
                '<div class="fa-kpi-head">'
                '<div class="fa-kpi-icon">🏆</div>'
                '<div class="fa-kpi-title">PRODUCTO CON MAYOR VENTA</div>'
                '</div>'
                '<div class="fa-kpi-divider"></div>'
                '<div class="fa-leader-main">'
                '<div class="fa-leader-name">'
                f'{html.escape(family_kpi_top_name)}'
                '</div>'
                '<div class="fa-leader-code">Código: '
                f'{html.escape(family_kpi_top_code)}</div>'
                '<div class="fa-leader-value">'
                f'{money(family_kpi_top_sales)}'
                '</div>'
                '</div>'
                '<div class="fa-kpi-footer">'
                f'▧ Inventario: {quantity(family_kpi_top_inventory, 0)}'
                f' | EMPAQ: {html.escape(family_kpi_top_pack)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    st.markdown(
        "<div style='height:10px'></div>",
        unsafe_allow_html=True,
    )

    fa4, fa5, fa6 = st.columns(
        3,
        gap="small",
    )

    with fa4:
        st.markdown(
            (
                '<div class="fa-kpi-card profit" '
                'style="--fa-accent:#28D7A1;--fa-value:#35E4B0;">'
                '<div class="fa-kpi-head">'
                '<div class="fa-kpi-icon">📈</div>'
                '<div class="fa-kpi-title">'
                'UTILIDAD BRUTA DE LA FAMILIA'
                '</div>'
                '</div>'
                '<div class="fa-kpi-divider"></div>'
                '<div class="fa-kpi-main">'
                '<div class="fa-kpi-value">'
                f'{money(family_kpi_profit)}'
                '</div>'
                '</div>'
                '<div class="fa-kpi-footer">'
                f'🏷️ {html.escape(family_kpi_label)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with fa5:
        st.markdown(
            (
                '<div class="fa-kpi-card stockvalue" '
                'style="--fa-accent:#3DA8FF;--fa-value:#55B8FF;">'
                '<div class="fa-kpi-head">'
                '<div class="fa-kpi-icon">💰</div>'
                '<div class="fa-kpi-title">'
                'VALOR DEL INVENTARIO DE LA FAMILIA'
                '</div>'
                '</div>'
                '<div class="fa-kpi-divider"></div>'
                '<div class="fa-kpi-main">'
                '<div class="fa-kpi-value">'
                f'{money(family_kpi_inventory_value)}'
                '</div>'
                '</div>'
                '<div class="fa-kpi-footer">'
                '◈ INVENTARIO × COSTO ACTUAL'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with fa6:
        st.markdown(
            (
                '<div class="fa-kpi-card lostsales" '
                'style="--fa-accent:#FF596B;--fa-value:#FF6C7A;">'
                '<div class="fa-kpi-head">'
                '<div class="fa-kpi-icon">⚠️</div>'
                '<div class="fa-kpi-title">'
                'VENTAS PERDIDAS DE LA FAMILIA'
                '</div>'
                '</div>'
                '<div class="fa-kpi-divider"></div>'
                '<div class="fa-kpi-main">'
                '<div class="fa-kpi-value">'
                f'{quantity(family_kpi_lost_sales, 0)}'
                '</div>'
                '</div>'
                '<div class="fa-kpi-footer">'
                '▧ UNIDADES PERDIDAS'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    st.markdown(
        "<div style='height:20px'></div>",
        unsafe_allow_html=True,
    )


    st.markdown("### 💰 Ventas mes actual por familia")
    if top_choice == "TODOS":
        st.caption(
            f"Mostrando las {family_visual_limit} familias con mayores ventas "
            f"de {family_total_count:,} familias disponibles. "
            "TODOS conserva el conjunto completo para análisis y tablas. "
            "Haz clic sobre una barra para analizar esa familia."
        )
    else:
        st.caption(
            f"Se muestran las {family_visual_limit} familias con mayores ventas "
            f"de {family_total_count:,} familias disponibles. "
            "Haz clic sobre una barra para analizar esa familia en las demás "
            "gráficas y tablas de esta pestaña."
        )


    selector_fig = px.bar(
        family_sales_selector,
        x="FAMILIA_COMPLETA",
        y="$VENTAS MES ACTUAL",
        custom_data=[
            "FAMILIA_COMPLETA",
        ],
        text="ETIQUETA_VALOR",
        labels={
            "FAMILIA_COMPLETA": "Familia",
            "$VENTAS MES ACTUAL": "Ventas ($)",
        },
    )

    apply_executive_bar_style(
        selector_fig
    )

    selector_fig.update_traces(
        textposition="outside",
        cliponaxis=False,
        textfont=dict(
            size=9,
            color="#DDE7F2",
        ),
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "Ventas: $%{y:,.2f}"
            "<extra></extra>"
        ),
    )

    selector_title_html = (
        "<b><span style='font-size:28px;color:#00E5FF;'>"
        "VENTAS MES ACTUAL"
        "</span></b>"
        "<br>"
        "<span style='font-size:22px;color:#FFFFFF;'>"
        + family_main_title
        + "</span>"
    )

    selector_fig.update_layout(
        height=570,
        bargap=0.30,
        title=dict(
            text=selector_title_html,
            x=0.5,
            xanchor="center",
            y=0.93,
            yanchor="top",
            font=dict(
                family="Segoe UI, Arial, sans-serif",
            ),
        ),
        margin=dict(
            l=25,
            r=20,
            t=130,
            b=125,
        ),
        xaxis=dict(
            title="Familia",
            tickangle=-30,
            tickfont=dict(
                size=10,
                color="#C4D3E3",
                family="Segoe UI, Arial, sans-serif",
            ),
            automargin=True,
            showgrid=False,
            categoryorder="array",
            categoryarray=(
                family_sales_selector[
                    "FAMILIA_COMPLETA"
                ].tolist()
            ),
            tickmode="array",
            tickvals=(
                family_sales_selector[
                    "FAMILIA_COMPLETA"
                ].tolist()
            ),
            ticktext=(
                family_sales_selector[
                    "FAMILIA_CORTA"
                ].tolist()
            ),
        ),
        yaxis=dict(
            title="Ventas ($)",
            tickfont=dict(
                size=10,
                color="#AFC4DA",
            ),
            gridcolor="#31445D",
            zeroline=False,
        ),
        hovermode="closest",
    )

    # Resaltar visualmente la familia ya seleccionada.
    if selected_family:
        selected_indices = (
            family_sales_selector.index[
                family_sales_selector[
                    "FAMILIA_COMPLETA"
                ].eq(
                    selected_family
                )
            ]
            .tolist()
        )

        if selected_indices:
            selector_fig.update_traces(
                selectedpoints=selected_indices,
                selected=dict(
                    marker=dict(
                        opacity=1.0,
                    )
                ),
                unselected=dict(
                    marker=dict(
                        opacity=0.32,
                    )
                ),
            )

    selector_event = st.plotly_chart(
        selector_fig,
        use_container_width=True,
        key=(
            "family_click_selector_chart_"
            f"{st.session_state['family_selection_reset_token']}"
        ),
        on_select="rerun",
        selection_mode="points",
    )

    # Leer la familia pulsada.
    selected_points = []

    try:
        selected_points = (
            selector_event
            .selection
            .points
        )
    except Exception:
        try:
            selected_points = (
                selector_event
                .get(
                    "selection",
                    {},
                )
                .get(
                    "points",
                    [],
                )
            )
        except Exception:
            selected_points = []

    if selected_points:
        # Plotly puede conservar la selección anterior y devolver
        # más de un punto. El último elemento corresponde al clic
        # más reciente y debe reemplazar la familia anterior.
        point = selected_points[-1]

        try:
            custom_data = point.get(
                "customdata"
            )
        except Exception:
            custom_data = None

        clicked_family = None

        if isinstance(
            custom_data,
            (list, tuple),
        ):
            if custom_data:
                clicked_family = str(
                    custom_data[0]
                )
        elif custom_data is not None:
            clicked_family = str(
                custom_data
            )

        if (
            clicked_family
            and clicked_family
            in available_family_names
        ):
            previous_family = st.session_state.get(
                "selected_family_from_chart"
            )

            st.session_state[
                "selected_family_from_chart"
            ] = clicked_family

            selected_family = clicked_family

            # Si el usuario pulsa otra barra, recreamos solamente
            # este componente con un key nuevo. Esto borra la
            # selección visual anterior y evita tener que hacer
            # un segundo clic.
            if clicked_family != previous_family:
                st.session_state[
                    "family_selection_reset_token"
                ] += 1
                st.rerun()

    # Controles de selección
    selection_col1, selection_col2 = (
        st.columns(
            [4, 1],
            gap="large",
        )
    )

    with selection_col1:

        if selected_family:
            st.markdown(
                f"**Familia seleccionada:** "
                f"🟦 {selected_family}"
            )
        else:
            st.caption(
                "Vista general: todas las familias."
            )

    with selection_col2:

        if selected_family:
            st.button(
                "Ver todas",
                key="clear_family_chart_selection",
                use_container_width=True,
                on_click=clear_family_selection,
            )

    # ========================================================
    # VENTAS Y UTILIDAD POR FAMILIA
    # ========================================================
    # Estas visualizaciones fueron trasladadas a la pestaña:
    # "Ventas y utilidad por familia".
    # ========================================================


    # ========================================================
    # MES ANTERIOR VS MES ACTUAL
    # ========================================================

    st.markdown("---")

    if selected_family:
        st.markdown(
            f"### 📈 Mes anterior vs. mes actual "
            f"— {selected_family}"
        )

        comparison_products = (
            family_focus_df
            .groupby(
                [
                    "#COD.",
                    "DESCRIPCION",
                ],
                dropna=False,
                as_index=False,
            )
            .agg(
                {
                    "#VENTAS MES ANTERIOR":
                        "sum",
                    "#VENTAS MES ACTUAL":
                        "sum",
                }
            )
        )

        comparison_products[
            "TOTAL_COMPARACION"
        ] = (
            comparison_products[
                "#VENTAS MES ANTERIOR"
            ]
            .fillna(0)
            +
            comparison_products[
                "#VENTAS MES ACTUAL"
            ]
            .fillna(0)
        )

        comparison_products = (
            comparison_products
            .sort_values(
                "TOTAL_COMPARACION",
                ascending=False,
            )
            .head(top_n)
        )

        comparison_products[
            "PRODUCTO"
        ] = (
            comparison_products[
                "DESCRIPCION"
            ]
            .astype(str)
            .apply(
                lambda value:
                (
                    value
                    if len(value) <= 22
                    else value[:21] + "…"
                )
            )
        )

        comparison_long = (
            comparison_products
            .melt(
                id_vars="PRODUCTO",
                value_vars=[
                    "#VENTAS MES ANTERIOR",
                    "#VENTAS MES ACTUAL",
                ],
                var_name="PERIODO",
                value_name="VENTAS",
            )
        )

        fig = px.bar(
            comparison_long,
            x="PRODUCTO",
            y="VENTAS",
            color="PERIODO",
            barmode="group",
            title=(
                "Comparativo por producto"
            ),
            labels={
                "PRODUCTO": "Producto",
                "VENTAS": "Cantidad",
                "PERIODO": "Periodo",
            },
            color_discrete_sequence=
                EXECUTIVE_COLORS,
        )

        apply_executive_bar_style(
            fig
        )

        fig.update_layout(
            height=540,
            xaxis_tickangle=-35,
            xaxis_tickfont=dict(
                size=9,
            ),
            margin=dict(
                l=20,
                r=20,
                t=60,
                b=115,
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    else:

        st.markdown(
            "### 📈 Mes anterior vs. "
            "mes actual por familia"
        )

        comparison = (
            family_focus_df
            .groupby(
                "FAMILIA",
                dropna=False,
                as_index=False,
            )
            .agg(
                {
                    "#VENTAS MES ANTERIOR":
                        "sum",
                    "#VENTAS MES ACTUAL":
                        "sum",
                }
            )
            .sort_values(
                "#VENTAS MES ACTUAL",
                ascending=False,
            )
            .head(family_visual_limit)
        )

        comparison[
            "FAMILIA_CORTA"
        ] = (
            comparison["FAMILIA"]
            .astype(str)
            .apply(
                _short_family_name
            )
        )

        comparison_long = (
            comparison
            .melt(
                id_vars=[
                    "FAMILIA",
                    "FAMILIA_CORTA",
                ],
                value_vars=[
                    "#VENTAS MES ANTERIOR",
                    "#VENTAS MES ACTUAL",
                ],
                var_name="PERIODO",
                value_name="VENTAS",
            )
        )

        fig = px.bar(
            comparison_long,
            x="FAMILIA_CORTA",
            y="VENTAS",
            color="PERIODO",
            custom_data=[
                "FAMILIA",
            ],
            barmode="group",
            title=(
                (
                    f"Comparativo de ventas — Top {family_visual_limit} "
                    f"de {family_total_count:,} familias"
                )
                if top_choice == "TODOS"
                else (
                    f"Comparativo de ventas — "
                    f"Top {family_visual_limit} familias"
                )
            ),
            labels={
                "FAMILIA_CORTA":
                    "Familia",
                "VENTAS":
                    "Cantidad vendida",
                "PERIODO":
                    "Periodo",
            },
            color_discrete_sequence=
                EXECUTIVE_COLORS,
        )

        apply_executive_bar_style(
            fig
        )

        fig.update_traces(
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Cantidad: %{y:,.0f}"
                "<extra></extra>"
            ),
        )

        fig.update_layout(
            height=540,
            xaxis_tickangle=-35,
            xaxis_tickfont=dict(
                size=9,
            ),
            margin=dict(
                l=20,
                r=20,
                t=60,
                b=110,
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
        )


    # ========================================================
    # VENTAS TOTALES
    # ========================================================

    st.markdown("---")

    if selected_family:

        st.markdown(
            f"### 📊 Ventas por producto "
            f"— {selected_family}"
        )

        total_product_sales = (
            family_focus_df
            .groupby(
                [
                    "#COD.",
                    "DESCRIPCION",
                ],
                dropna=False,
                as_index=False,
            )["$VENTAS MES ACTUAL"]
            .sum()
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=False,
            )
            .head(top_n)
        )

        total_product_sales[
            "PRODUCTO"
        ] = (
            total_product_sales[
                "DESCRIPCION"
            ]
            .astype(str)
        )

        total_product_sales[
            "PRODUCTO_CORTO"
        ] = (
            total_product_sales[
                "PRODUCTO"
            ]
            .apply(
                lambda value:
                (
                    value
                    if len(value) <= 22
                    else value[:21] + "…"
                )
            )
        )

        fig = px.bar(
            total_product_sales,
            x="PRODUCTO_CORTO",
            y="$VENTAS MES ACTUAL",
            custom_data=[
                "PRODUCTO",
            ],
            title=(
                "Productos con mayores ventas "
                "de la familia"
            ),
            labels={
                "PRODUCTO_CORTO":
                    "Producto",
                "$VENTAS MES ACTUAL":
                    "Ventas ($)",
            },
        )

        apply_executive_bar_style(
            fig
        )

        fig.update_traces(
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Ventas: $%{y:,.2f}"
                "<extra></extra>"
            ),
        )

        fig.update_layout(
            height=540,
            xaxis_tickangle=-35,
            xaxis_tickfont=dict(
                size=9,
            ),
            margin=dict(
                l=20,
                r=20,
                t=60,
                b=115,
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    else:

        st.markdown(
            "### 📊 Ventas totales por familia"
        )

        total_by_family_all = (
            family_focus_df
            .groupby(
                "FAMILIA",
                dropna=False,
                as_index=False,
            )["$VENTAS MES ACTUAL"]
            .sum()
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=False,
            )
            .reset_index(drop=True)
        )

        if top_choice == "TODOS":
            pie_top_limit = min(
                10,
                len(total_by_family_all),
            )

            total_by_family = (
                total_by_family_all
                .head(pie_top_limit)
                .copy()
            )

            other_sales = float(
                total_by_family_all
                .iloc[pie_top_limit:][
                    "$VENTAS MES ACTUAL"
                ]
                .sum()
            )

            if other_sales != 0:
                total_by_family = pd.concat(
                    [
                        total_by_family,
                        pd.DataFrame(
                            {
                                "FAMILIA": [
                                    "OTRAS FAMILIAS"
                                ],
                                "$VENTAS MES ACTUAL": [
                                    other_sales
                                ],
                            }
                        ),
                    ],
                    ignore_index=True,
                )

            pie_title = (
                f"Participación de ventas — "
                f"Top {pie_top_limit} + OTRAS FAMILIAS"
            )
        else:
            total_by_family = (
                total_by_family_all
                .head(family_visual_limit)
                .copy()
            )

            pie_title = (
                f"Participación de ventas — "
                f"Top {family_visual_limit} familias"
            )

        fig = px.pie(
            total_by_family,
            names="FAMILIA",
            values="$VENTAS MES ACTUAL",
            title=pie_title,
            hole=0.45,
            color_discrete_sequence=
                EXECUTIVE_COLORS,
        )

        fig.update_traces(
            textposition="auto",
            textinfo="none",
            texttemplate=(
                "<b>%{percent:.1%}</b><br>"
                "$%{value:,.2f}"
            ),
            insidetextorientation="horizontal",
            textfont=dict(
                color="#FFFFFF",
                size=10,
                family=(
                    "Inter, Segoe UI, Arial, sans-serif"
                ),
            ),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Ventas: $%{value:,.2f}<br>"
                "Participación: %{percent:.2%}"
                "<extra></extra>"
            ),
        )

        apply_executive_pie_style(
            fig
        )

        fig.update_layout(
            height=585,
            margin=dict(
                l=35,
                r=35,
                t=72,
                b=28,
            ),
            legend=dict(
                font=dict(
                    size=9,
                ),
            ),
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
        )


    # ========================================================
    # RESUMEN POR FAMILIA
    # ========================================================

    st.markdown("---")

    if selected_family:
        summary_family_focus = (
            family_summary(
                family_focus_df
            )
        )

        table_title = (
            f"📋 Resumen de familia — "
            f"{selected_family}"
        )
    else:
        if top_choice == "TODOS":
            summary_family_focus = (
                family_df
                .copy()
            )

            table_title = (
                f"📋 Resumen por familia — "
                f"TODAS ({family_total_count:,})"
            )
        else:
            summary_family_focus = (
                family_df
                .copy()
                .head(family_visual_limit)
            )

            table_title = (
                f"📋 Resumen por familia — "
                f"TOP {family_visual_limit}"
            )

    with st.expander(
        f"{table_title}  ·  Mostrar / ocultar",
        expanded=False,
    ):
        st.markdown(
            '<div class="executive-table-title" '
            'style="border-left-color:#00D8F0;'
            'margin-top:2px;">'
            f'{table_title}'
            '</div>',
            unsafe_allow_html=True,
        )

        render_executive_html_table(
            summary_family_focus,
            column_labels={
                "FAMILIA":
                    "Familia",
                "$VENTAS MES ACTUAL":
                    "Ventas Mes Actual ($)",
                "$UTILIDAD MES ACTUAL":
                    "Utilidad Mes Actual ($)",
                "INVENTARIO TOTAL":
                    "Inventario Total",
                "#VENTAS PERDIDAS":
                    "Ventas Perdidas",
                "MARGEN PROMEDIO %":
                    "Margen Promedio (%)",
                "#VENTAS MES ANTERIOR":
                    "Ventas Mes Anterior",
                "#VENTAS MES ACTUAL":
                    "Ventas Mes Actual",
            },
            formats={
                "$VENTAS MES ACTUAL":
                    "money",
                "$UTILIDAD MES ACTUAL":
                    "money",
                "INVENTARIO TOTAL":
                    "number2",
                "#VENTAS PERDIDAS":
                    "number2",
                "MARGEN PROMEDIO %":
                    "percent",
                "#VENTAS MES ANTERIOR":
                    "number2",
                "#VENTAS MES ACTUAL":
                    "number2",
            },
            max_height=560,
        )


# ============================================================
# PESTAÑA 2 — ANÁLISIS DE PRODUCTOS
# ============================================================

with tab_productos:

    st.markdown("### 🎯 Análisis por métrica")
    st.caption(
        f"Métrica seleccionada: **{selected_metric}**"
    )

    # La gráfica dinámica de "Análisis por métrica" fue eliminada
    # de esta sección por solicitud del usuario.


    # ========================================================
    # VENTAS POR PRODUCTO DENTRO DE UNA FAMILIA
    # ========================================================

    st.markdown("---")

    families_for_chart = sorted(
        str(value)
        for value in filtered_df["FAMILIA"]
        .dropna()
        .unique()
        if str(value).strip()
    )

    if families_for_chart:

        if top_family_name in families_for_chart:
            default_index = families_for_chart.index(
                top_family_name
            )
        else:
            default_index = 0

        default_chart_family = (
            families_for_chart[
                default_index
            ]
        )

        # Familia actualmente visible en el selector.
        current_chart_family = st.session_state.get(
            "family_product_chart",
            default_chart_family,
        )

        if current_chart_family not in families_for_chart:
            current_chart_family = default_chart_family
            st.session_state[
                "family_product_chart"
            ] = default_chart_family

        st.markdown(
            "### 🧩 Ventas por producto dentro de una familia"
        )

        if "treemap_click_reset_token" not in st.session_state:
            st.session_state["treemap_click_reset_token"] = 0

        def clear_family_product_filter() -> None:
            """
            Mantiene la familia elegida y limpia únicamente los
            productos seleccionados. TOP vuelve a TODOS, que en esta
            sección significa TODOS LOS PRODUCTOS DE LA FAMILIA ACTUAL.
            """
            current_family = st.session_state.get(
                "family_product_chart",
                default_chart_family,
            )

            st.session_state[
                "selected_product_from_treemap"
            ] = None

            st.session_state[
                "selected_products_from_treemap"
            ] = []

            st.session_state[
                "selected_product_family"
            ] = current_family

            st.session_state[
                "sidebar_product_filter"
            ] = []

            st.session_state[
                "global_top_n"
            ] = "TODOS"

            st.session_state[
                "treemap_click_reset_token"
            ] = (
                st.session_state.get(
                    "treemap_click_reset_token",
                    0,
                )
                + 1
            )

        family_selector_col, family_clear_col = st.columns(
            [5.2, 0.9],
            gap="small",
        )

        with family_selector_col:
            chart_family = st.selectbox(
                "Seleccione una familia",
                options=families_for_chart,
                index=default_index,
                key="family_product_chart",
            )

        with family_clear_col:
            st.markdown(
                "<div style='height:25px'></div>",
                unsafe_allow_html=True,
            )

            st.button(
                "🧹 Borrar filtro",
                key="clear_family_product_treemap",
                use_container_width=True,
                on_click=clear_family_product_filter,
            )

        # Leer nuevamente el estado actual para garantizar que
        # todos los componentes de esta sección usen exactamente
        # la familia que muestra el selectbox.
        chart_family = st.session_state.get(
            "family_product_chart",
            default_chart_family,
        )

        # Si cambia la familia, una selección de producto
        # perteneciente a la familia anterior deja de ser válida.
        previous_product_family = st.session_state.get(
            "selected_product_family"
        )

        if (
            previous_product_family is not None
            and str(previous_product_family) != str(chart_family)
        ):
            st.session_state[
                "selected_product_from_treemap"
            ] = None
            st.session_state[
                "selected_products_from_treemap"
            ] = []
            st.session_state[
                "selected_product_family"
            ] = chart_family
            st.session_state[
                "treemap_click_reset_token"
            ] = (
                st.session_state.get(
                    "treemap_click_reset_token",
                    0,
                )
                + 1
            )

        # ====================================================
        # KPI DINÁMICOS — ANÁLISIS DE PRODUCTOS POR FAMILIA
        # ====================================================

        selected_family_kpi_df = (
            filtered_df[
                filtered_df["FAMILIA"]
                .astype(str)
                .eq(str(chart_family))
            ]
            .copy()
        )

        kpi_family_sales = safe_sum(
            selected_family_kpi_df[
                "$VENTAS MES ACTUAL"
            ]
        )

        kpi_family_inventory = safe_sum(
            selected_family_kpi_df[
                "INVENTARIO TOTAL"
            ]
        )

        # Nuevos KPI interactivos por familia.
        kpi_family_gross_profit = safe_sum(
            selected_family_kpi_df[
                "$UTILIDAD MES ACTUAL"
            ]
        )

        kpi_family_inventory_value = float(
            (
                pd.to_numeric(
                    selected_family_kpi_df[
                        "INVENTARIO TOTAL"
                    ],
                    errors="coerce",
                ).fillna(0.0)
                * pd.to_numeric(
                    selected_family_kpi_df[
                        "COSTO ACTUAL"
                    ],
                    errors="coerce",
                ).fillna(0.0)
            ).sum()
        )

        kpi_family_lost_sales = safe_sum(
            selected_family_kpi_df[
                "#VENTAS PERDIDAS"
            ]
        )

        family_pack_values = sorted(
            {
                str(value).strip()
                for value in selected_family_kpi_df[
                    "EMPAQ_ANALISIS"
                ]
                .dropna()
                .tolist()
                if str(value).strip()
                and str(value).strip().lower()
                not in {"nan", "<na>", "none"}
            }
        )

        if len(family_pack_values) == 1:
            kpi_family_pack = family_pack_values[0]
        elif len(family_pack_values) > 1:
            kpi_family_pack = "Varios empaques"
        else:
            kpi_family_pack = "Sin empaque"

        kpi_product_summary = (
            selected_family_kpi_df
            .groupby(
                [
                    "#COD.",
                    "DESCRIPCION",
                ],
                dropna=False,
                as_index=False,
            )
            .agg(
                {
                    "$VENTAS MES ACTUAL": "sum",
                    "INVENTARIO TOTAL": "sum",
                    "EMPAQ_ANALISIS": (
                        lambda values:
                        next(
                            (
                                str(value).strip()
                                for value in values
                                if pd.notna(value)
                                and str(value).strip()
                            ),
                            "",
                        )
                    ),
                }
            )
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=False,
            )
            .reset_index(drop=True)
        )

        if not kpi_product_summary.empty:
            kpi_top_product_code = str(
                kpi_product_summary.iloc[0][
                    "#COD."
                ]
            )

            kpi_top_product_name = str(
                kpi_product_summary.iloc[0][
                    "DESCRIPCION"
                ]
            )

            kpi_top_product_sales = float(
                kpi_product_summary.iloc[0][
                    "$VENTAS MES ACTUAL"
                ]
            )

            kpi_top_product_inventory = float(
                kpi_product_summary.iloc[0][
                    "INVENTARIO TOTAL"
                ]
            )

            kpi_top_product_pack = str(
                kpi_product_summary.iloc[0][
                    "EMPAQ_ANALISIS"
                ]
            ).strip()

            if (
                not kpi_top_product_pack
                or kpi_top_product_pack.lower()
                in {"nan", "<na>", "none"}
            ):
                kpi_top_product_pack = "Sin empaque"
        else:
            kpi_top_product_code = ""
            kpi_top_product_name = "Sin datos"
            kpi_top_product_sales = 0.0
            kpi_top_product_inventory = 0.0
            kpi_top_product_pack = "Sin empaque"

        st.markdown(
            "<div style='height:10px'></div>",
            unsafe_allow_html=True,
        )

        # ====================================================
        # KPI EJECUTIVOS — ANÁLISIS DE PRODUCTOS
        # Diseño exclusivo para evitar conflictos con CSS global
        # ====================================================

        st.markdown(
            """
            <style>
            /* ==================================================
               KPI ANÁLISIS DE PRODUCTOS — DISEÑO EJECUTIVO
               ================================================== */

            .pa-kpi-card {
                position: relative;
                min-height: 142px;
                height: 142px;
                border-radius: 14px;
                padding: 7px 9px 6px 9px;
                background:
                    linear-gradient(
                        145deg,
                        rgba(18, 39, 64, .98) 0%,
                        rgba(9, 25, 45, .99) 100%
                    );
                border: 1px solid rgba(100, 135, 170, .32);
                box-shadow:
                    0 10px 26px rgba(0, 0, 0, .20),
                    inset 0 1px 0 rgba(255, 255, 255, .025);
                overflow: hidden;
                display: flex;
                flex-direction: column;
            }

            .pa-kpi-card::before {
                content: "";
                position: absolute;
                top: 0;
                left: 16px;
                right: 16px;
                height: 3px;
                border-radius: 0 0 6px 6px;
                background: var(--pa-accent);
                box-shadow: 0 0 14px var(--pa-glow);
            }

            .pa-kpi-card.sales {
                --pa-accent: #26C9F5;
                --pa-glow: rgba(38, 201, 245, .55);
                --pa-value: #7DE0FF;
            }

            .pa-kpi-card.inventory {
                --pa-accent: #A47AE8;
                --pa-glow: rgba(164, 122, 232, .48);
                --pa-value: #B88AF7;
            }

            .pa-kpi-card.leader {
                --pa-accent: #E2B84E;
                --pa-glow: rgba(226, 184, 78, .48);
                --pa-value: #F2CF69;
            }

            .pa-kpi-card.profit {
                --pa-accent: #35D0A1;
                --pa-glow: rgba(53, 208, 161, .46);
                --pa-value: #68E3BB;
            }

            .pa-kpi-card.stockvalue {
                --pa-accent: #4FA8FF;
                --pa-glow: rgba(79, 168, 255, .46);
                --pa-value: #79BEFF;
            }

            .pa-kpi-card.lostsales {
                --pa-accent: #FF6F7D;
                --pa-glow: rgba(255, 111, 125, .43);
                --pa-value: #FF8D98;
            }

            .pa-kpi-head {
                display: flex;
                align-items: center;
                gap: 8px;
                min-height: 27px;
            }

            .pa-kpi-icon {
                width: 26px;
                height: 26px;
                min-width: 26px;
                border-radius: 10px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: .80rem;
                background:
                    linear-gradient(
                        145deg,
                        rgba(15, 42, 69, .95),
                        rgba(7, 24, 44, .98)
                    );
                border: 1px solid var(--pa-accent);
                box-shadow:
                    0 0 14px var(--pa-glow),
                    inset 0 0 14px rgba(255, 255, 255, .025);
            }

            .pa-kpi-title {
                color: #48C9F5 !important;
                font-size: .70rem !important;
                line-height: 1.14 !important;
                font-weight: 900 !important;
                letter-spacing: .028em !important;
                text-transform: uppercase !important;
                text-align: left !important;
                margin: 0 !important;
                white-space: normal !important;
                overflow: visible !important;
            }

            .pa-kpi-divider {
                height: 1px;
                margin: 4px 0 4px 0;
                background:
                    linear-gradient(
                        90deg,
                        rgba(85, 139, 190, .62),
                        rgba(85, 139, 190, .10)
                    );
            }

            .pa-kpi-main {
                flex: 1;
                display: flex;
                align-items: center;
                justify-content: center;
                min-height: 34px;
                padding: 0 2px 1px 2px;
            }

            .pa-kpi-value {
                color: var(--pa-value) !important;
                font-size: clamp(1.28rem, 1.42vw, 1.72rem) !important;
                line-height: .98 !important;
                font-weight: 950 !important;
                letter-spacing: -.035em !important;
                text-align: center !important;
                white-space: nowrap !important;
                text-shadow: 0 0 12px var(--pa-glow);
                margin: 0 !important;
            }

            .pa-kpi-footer {
                min-height: 27px;
                border-radius: 10px;
                border: 1px solid rgba(91, 137, 184, .30);
                background: rgba(12, 35, 59, .78);
                padding: 4px 7px;
                display: flex;
                align-items: center;
                gap: 9px;
                color: #CFE2F5 !important;
                font-size: .66rem !important;
                line-height: 1.18 !important;
                font-weight: 780 !important;
                text-align: left !important;
            }

            .pa-kpi-footer-icon {
                color: #79C9FF;
                font-size: .75rem;
                line-height: 1;
            }

            /* Producto líder */
            .pa-leader-main {
                flex: 1;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: flex-start;
                min-height: 46px;
                padding: 0 2px 2px 2px;
            }

            .pa-leader-name {
                width: 100%;
                color: #E8EDF5 !important;
                font-size: .70rem !important;
                line-height: 1.22 !important;
                font-weight: 880 !important;
                margin-bottom: 5px !important;
                white-space: normal !important;
                overflow-wrap: anywhere !important;
                word-break: normal !important;
                display: block !important;
            }

            .pa-leader-code {
                color: #AFC6DC !important;
                font-size: .64rem !important;
                line-height: 1.18 !important;
                font-weight: 740 !important;
                margin-bottom: 5px !important;
            }

            .pa-leader-value {
                color: #F2CF69 !important;
                font-size: clamp(1.15rem, 1.28vw, 1.50rem) !important;
                line-height: .98 !important;
                font-weight: 950 !important;
                letter-spacing: -.025em !important;
                text-shadow: 0 0 11px rgba(226, 184, 78, .36);
                margin: 0 !important;
            }

            .pa-leader-footer {
                min-height: 27px;
                border-radius: 10px;
                border: 1px solid rgba(226, 184, 78, .26);
                background: rgba(26, 38, 53, .86);
                padding: 4px 7px;
                display: flex;
                align-items: center;
                gap: 9px;
                color: #D7E1EC !important;
                font-size: .66rem !important;
                line-height: 1.18 !important;
                font-weight: 760 !important;
                white-space: normal !important;
            }

            .pa-leader-footer strong {
                color: #F2CF69 !important;
                font-weight: 900 !important;
            }

            @media (max-width: 1350px) {
                .pa-kpi-card {
                    min-height: 132px;
                    height: 132px;
                    padding: 6px 8px 5px 8px;
                }

                .pa-kpi-title {
                    font-size: .70rem !important;
                }

                .pa-kpi-icon {
                    width: 24px;
                    height: 24px;
                    min-width: 24px;
                    font-size: .75rem;
                }

                .pa-kpi-footer,
                .pa-leader-footer {
                    font-size: .62rem !important;
                }

                .pa-leader-name {
                    font-size: .70rem !important;
                }
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        spacer_left, pk1, pk2, pk3, spacer_right = st.columns(
            [0.42, 1.0, 1.0, 1.15, 0.42],
            gap="small",
        )

        with pk1:
            st.markdown(
                (
                    '<div class="pa-kpi-card sales">'
                    '<div class="pa-kpi-head">'
                    '<div class="pa-kpi-icon">💵</div>'
                    '<div class="pa-kpi-title">'
                    'VENTA TOTAL POR FAMILIA'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-divider"></div>'
                    '<div class="pa-kpi-main">'
                    '<div class="pa-kpi-value">'
                    f'{money(kpi_family_sales)}'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-footer">'
                    '<span class="pa-kpi-footer-icon">🏷️</span>'
                    f'<span>{html.escape(str(chart_family))}</span>'
                    '</div>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

        with pk2:
            st.markdown(
                (
                    '<div class="pa-kpi-card inventory">'
                    '<div class="pa-kpi-head">'
                    '<div class="pa-kpi-icon">📦</div>'
                    '<div class="pa-kpi-title">'
                    'INVENTARIO TOTAL POR FAMILIA'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-divider"></div>'
                    '<div class="pa-kpi-main">'
                    '<div class="pa-kpi-value">'
                    f'{quantity(kpi_family_inventory, 0)}'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-footer">'
                    '<span class="pa-kpi-footer-icon">◈</span>'
                    '<span>EMPAQ: '
                    f'{html.escape(kpi_family_pack)}'
                    '</span>'
                    '</div>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

        with pk3:
            st.markdown(
                (
                    '<div class="pa-kpi-card leader">'
                    '<div class="pa-kpi-head">'
                    '<div class="pa-kpi-icon">🏆</div>'
                    '<div class="pa-kpi-title">'
                    'PRODUCTO CON MAYOR VENTA'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-divider"></div>'
                    '<div class="pa-leader-main">'
                    '<div class="pa-leader-name">'
                    f'{html.escape(kpi_top_product_name)}'
                    '</div>'
                    '<div class="pa-leader-code">'
                    'Código: '
                    f'<strong>{html.escape(kpi_top_product_code)}</strong>'
                    '</div>'
                    '<div class="pa-leader-value">'
                    f'{money(kpi_top_product_sales)}'
                    '</div>'
                    '</div>'
                    '<div class="pa-leader-footer">'
                    '<span class="pa-kpi-footer-icon">▧</span>'
                    '<span>Inventario: '
                    f'<strong>{quantity(kpi_top_product_inventory, 0)}</strong>'
                    '&nbsp;&nbsp;|&nbsp;&nbsp; EMPAQ: '
                    f'<strong>{html.escape(kpi_top_product_pack)}</strong>'
                    '</span>'
                    '</div>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

        # ----------------------------------------------------
        # SEGUNDA FILA KPI — RENTABILIDAD, VALOR Y PÉRDIDAS
        # ----------------------------------------------------
        st.markdown(
            "<div style='height:12px'></div>",
            unsafe_allow_html=True,
        )

        spacer2_left, pk4, pk5, pk6, spacer2_right = st.columns(
            [0.42, 1.0, 1.0, 1.15, 0.42],
            gap="small",
        )

        with pk4:
            st.markdown(
                (
                    '<div class="pa-kpi-card profit">'
                    '<div class="pa-kpi-head">'
                    '<div class="pa-kpi-icon">📈</div>'
                    '<div class="pa-kpi-title">'
                    'UTILIDAD BRUTA DE LA FAMILIA'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-divider"></div>'
                    '<div class="pa-kpi-main">'
                    '<div class="pa-kpi-value">'
                    f'{money(kpi_family_gross_profit)}'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-footer">'
                    '<span class="pa-kpi-footer-icon">🏷️</span>'
                    f'<span>{html.escape(str(chart_family))}</span>'
                    '</div>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

        with pk5:
            st.markdown(
                (
                    '<div class="pa-kpi-card stockvalue">'
                    '<div class="pa-kpi-head">'
                    '<div class="pa-kpi-icon">💰</div>'
                    '<div class="pa-kpi-title">'
                    'VALOR DEL INVENTARIO DE LA FAMILIA'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-divider"></div>'
                    '<div class="pa-kpi-main">'
                    '<div class="pa-kpi-value">'
                    f'{money(kpi_family_inventory_value)}'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-footer">'
                    '<span class="pa-kpi-footer-icon">◈</span>'
                    '<span>INVENTARIO × COSTO ACTUAL</span>'
                    '</div>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

        with pk6:
            st.markdown(
                (
                    '<div class="pa-kpi-card lostsales">'
                    '<div class="pa-kpi-head">'
                    '<div class="pa-kpi-icon">⚠️</div>'
                    '<div class="pa-kpi-title">'
                    'VENTAS PERDIDAS DE LA FAMILIA'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-divider"></div>'
                    '<div class="pa-kpi-main">'
                    '<div class="pa-kpi-value">'
                    f'{quantity(kpi_family_lost_sales, 0)}'
                    '</div>'
                    '</div>'
                    '<div class="pa-kpi-footer">'
                    '<span class="pa-kpi-footer-icon">▧</span>'
                    '<span>UNIDADES PERDIDAS</span>'
                    '</div>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

        # Separación visual adicional entre los KPI y el treemap.
        st.markdown(
            "<div style='height:32px'></div>",
            unsafe_allow_html=True,
        )


    # ========================================================
    # SCOPE DE PRODUCTOS DE LA FAMILIA ACTUAL
    # ========================================================
    # IMPORTANTE: en Análisis de productos, TODOS significa
    # todos los productos de chart_family, nunca todo filtered_df.
    family_product_scope_df = (
        filtered_df[
            (
                filtered_df["FAMILIA"]
                .astype(str)
                .eq(str(chart_family))
            )
            &
            (
                ~filtered_df["DESCRIPCION"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
                .str.startswith(
                    "AJUSTE",
                    na=False,
                )
            )
        ]
        .groupby(
            ["#COD.", "DESCRIPCION"],
            dropna=False,
            as_index=False,
        )["$VENTAS MES ACTUAL"]
        .sum()
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    family_product_count = len(
        family_product_scope_df
    )

    family_top_n = (
        min(
            50,
            max(family_product_count, 1),
        )
        if top_choice == "TODOS"
        else min(
            int(top_choice),
            max(family_product_count, 1),
        )
    )

    family_top_label = (
        (
            f"{family_top_n} DE "
            f"{family_product_count}"
        )
        if top_choice == "TODOS"
        and family_product_count > family_top_n
        else (
            "TODOS"
            if top_choice == "TODOS"
            else str(top_choice)
        )
    )

    # ========================================================
    # TOP PRODUCTOS
    # ========================================================

    st.markdown("---")
    st.markdown(
        f"### 🏅 Top {family_top_label} productos "
        "por ventas del mes actual"
    )

    top_products = (
        family_product_scope_df
        .head(family_top_n)
        .copy()
    )

    top_products["PRODUCTO"] = (
        top_products["#COD."]
        .astype(str)
        + " - "
        + top_products["DESCRIPCION"]
        .astype(str)
    )

    top_products["VENTA_LABEL"] = (
        top_products["$VENTAS MES ACTUAL"]
        .apply(money)
    )

    product_top_colors = [
        "#2ED3B7",
        "#D7AE58",
        "#4FC3F7",
        "#8FB1D0",
        "#B68AE6",
        "#FF8A72",
        "#57D4C3",
        "#E6A85C",
        "#8EB6DB",
        "#93C9A5",
        "#FF6F91",
        "#7FD3FF",
    ]

    fig = px.bar(
        top_products.sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        ),
        x="PRODUCTO",
        y="$VENTAS MES ACTUAL",
        labels={
            "$VENTAS MES ACTUAL": "Ventas ($)",
            "PRODUCTO": "Producto",
        },
        text="VENTA_LABEL",
        color="PRODUCTO",
        color_discrete_sequence=product_top_colors,
    )

    fig.update_traces(
        textposition="outside",
        textfont=dict(
            size=12,
            color="#EAF3FA",
            family=(
                "Inter, Segoe UI, Arial, sans-serif"
            ),
        ),
        cliponaxis=False,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Ventas: %{text}<extra></extra>"
        ),
        marker_line_color="rgba(255,255,255,.16)",
        marker_line_width=1.0,
    )

    fig.update_layout(
        title=dict(
            text=(
                "$ VENTAS MES ACTUAL"
                "<br>"
                "<span style='font-size:18px;"
                "font-weight:700;"
                "letter-spacing:0.06em;"
                "color:#D9B85C;'>"
                f"TOP {family_top_label} PRODUCTOS"
                "</span>"
            ),
            x=0.5,
            xanchor="center",
            y=0.96,
            yanchor="top",
            font=dict(
                size=28,
                color="#FFFFFF",
                family=(
                    "Aptos Display, Inter, "
                    "Segoe UI, Arial, sans-serif"
                ),
            ),
        ),
        height=600,
        showlegend=False,
        xaxis_tickangle=-42,
        margin=dict(
            l=28,
            r=24,
            t=110,
            b=165,
        ),
    )

    apply_executive_bar_style(fig)

    st.plotly_chart(
        fig,
        use_container_width=True,
    )



    if families_for_chart:

        family_products = (
            selected_family_kpi_df[
                (
                    ~selected_family_kpi_df["DESCRIPCION"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    .str.startswith(
                        "AJUSTE",
                        na=False,
                    )
                )
            ]
            .groupby(
                [
                    "#COD.",
                    "DESCRIPCION",
                ],
                dropna=False,
                as_index=False,
            )
            .agg(
                {
                    "$VENTAS MES ACTUAL": "sum",
                    "INVENTARIO TOTAL": "sum",
                    "FECHA ULT COMPRA": "max",
                    "FECHA ULT VENTAS": "max",
                    "EMPAQ_ANALISIS": (
                        lambda values:
                        next(
                            (
                                str(value).strip()
                                for value in values
                                if pd.notna(value)
                                and str(value).strip()
                            ),
                            "",
                        )
                    ),
                }
            )
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=False,
            )
            .head(family_top_n)
            .reset_index(drop=True)
        )

        family_products["PRODUCTO"] = (
            family_products["#COD."]
            .astype(str)
            + " - "
            + family_products["DESCRIPCION"]
            .astype(str)
        )

        treemap_palette = [
            "#2D6CDF",
            "#13A89E",
            "#D6A936",
            "#7C5CE5",
            "#3B9FE8",
            "#2FA36B",
            "#D96872",
            "#C77C32",
            "#536CB8",
            "#2696A6",
            "#A06CD5",
            "#4AAE8A",
        ]

        treemap_color_map = {
            product: treemap_palette[idx % len(treemap_palette)]
            for idx, product in enumerate(
                family_products["PRODUCTO"].astype(str).tolist()
            )
        }

        family_products["EMPAQ_TEXTO"] = (
            family_products[
                "EMPAQ_ANALISIS"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        family_products["INVENTARIO_TEXTO"] = (
            family_products.apply(
                lambda row:
                (
                    f"{float(row['INVENTARIO TOTAL']):,.0f}"
                    + (
                        f" {row['EMPAQ_TEXTO']}"
                        if row["EMPAQ_TEXTO"]
                        else ""
                    )
                ),
                axis=1,
            )
        )

        family_products["ULT_COMPRA_TEXTO"] = (
            pd.to_datetime(
                family_products[
                    "FECHA ULT COMPRA"
                ],
                errors="coerce",
            )
            .dt.strftime(
                "%d/%m/%Y"
            )
            .fillna(
                "Sin fecha"
            )
        )

        family_products["ULT_VENTA_TEXTO"] = (
            pd.to_datetime(
                family_products[
                    "FECHA ULT VENTAS"
                ],
                errors="coerce",
            )
            .dt.strftime(
                "%d/%m/%Y"
            )
            .fillna(
                "Sin fecha"
            )
        )

        family_products["VENTA_TEXTO"] = (
            family_products["$VENTAS MES ACTUAL"]
            .apply(lambda value: f"${float(value):,.2f}")
        )

        # Plotly Treemap no dibuja categorías con venta 0 porque
        # no tienen área. La regla visual será:
        # - Si VENTAS > 0: el tamaño usa la venta real.
        # - Si VENTAS <= 0 pero INVENTARIO > 0: el producto permanece
        #   visible usando un tamaño visual mínimo basado en inventario.
        # - Los valores reales de venta e inventario NO se alteran.
        positive_sales_for_size = (
            family_products["$VENTAS MES ACTUAL"]
            .astype(float)
            .clip(lower=0)
        )

        inventory_for_size = (
            family_products["INVENTARIO TOTAL"]
            .astype(float)
            .clip(lower=0)
        )

        positive_total_for_size = float(
            positive_sales_for_size.sum()
        )

        visual_floor = (
            max(
                positive_total_for_size * 0.0025,
                0.01,
            )
            if positive_total_for_size > 0
            else 1.0
        )

        max_inventory_for_size = float(
            inventory_for_size.max()
        ) if len(inventory_for_size) else 0.0

        if max_inventory_for_size > 0:
            inventory_visual_factor = (
                inventory_for_size
                / max_inventory_for_size
            )
        else:
            inventory_visual_factor = (
                inventory_for_size * 0
            )

        # Para ventas 0 con inventario positivo, asignamos entre
        # 1x y 3x el piso visual según su inventario relativo.
        inventory_visual_size = (
            visual_floor
            * (
                1.0
                + 2.0 * inventory_visual_factor
            )
        )

        family_products["TREEMAP_SIZE"] = (
            positive_sales_for_size.copy()
        )

        zero_sales_with_inventory = (
            (positive_sales_for_size <= 0)
            & (inventory_for_size > 0)
        )

        family_products.loc[
            zero_sales_with_inventory,
            "TREEMAP_SIZE",
        ] = inventory_visual_size.loc[
            zero_sales_with_inventory
        ]

        # Si tanto ventas como inventario son 0, se deja un área mínima
        # para mantener consistencia con el TOP visible seleccionado.
        zero_sales_zero_inventory = (
            (positive_sales_for_size <= 0)
            & (inventory_for_size <= 0)
        )

        family_products.loc[
            zero_sales_zero_inventory,
            "TREEMAP_SIZE",
        ] = visual_floor * 0.35

        positive_sales_total = float(
            family_products[
                "$VENTAS MES ACTUAL"
            ]
            .clip(lower=0)
            .sum()
        )

        def _wrap_treemap_words(
            text: str,
            width: int,
            max_lines: int,
        ) -> list[str]:
            words = str(text).split()
            lines = []
            current = ""

            for word in words:
                candidate = word if not current else f"{current} {word}"

                if len(candidate) <= width:
                    current = candidate
                else:
                    if current:
                        lines.append(current)
                    current = word

                if len(lines) >= max_lines:
                    break

            if current and len(lines) < max_lines:
                lines.append(current)

            return lines[:max_lines]

        def build_treemap_text(row) -> str:
            code_text = str(row["#COD."]).strip()
            description = str(row["DESCRIPCION"]).strip()
            sales_value = float(row["$VENTAS MES ACTUAL"])
            inventory_text = str(
                row["INVENTARIO_TEXTO"]
            ).strip()

            if positive_sales_total > 0 and sales_value > 0:
                share = sales_value / positive_sales_total
            else:
                share = 0.0

            # Todos los bloques conservan el nombre del producto.
            # Los bloques grandes muestran código, ventas e inventario.
            # Los bloques pequeños muestran al menos nombre + inventario.

            if share >= 0.10:
                lines = _wrap_treemap_words(
                    description,
                    18,
                    3,
                )
                body = "<br>".join(lines)
                return (
                    f"<b>{code_text} - {body}</b><br>"
                    f"Ventas: {row['VENTA_TEXTO']}<br>"
                    f"Inventario: {inventory_text}"
                )

            if share >= 0.045:
                lines = _wrap_treemap_words(
                    description,
                    17,
                    2,
                )
                body = "<br>".join(lines)
                return (
                    f"<b>{code_text} - {body}</b><br>"
                    f"Ventas: {row['VENTA_TEXTO']}<br>"
                    f"Inventario: {inventory_text}"
                )

            if share >= 0.018:
                lines = _wrap_treemap_words(
                    description,
                    16,
                    2,
                )
                body = "<br>".join(lines)
                return (
                    f"<b>{body}</b><br>"
                    f"Ventas: {row['VENTA_TEXTO']}<br>"
                    f"Inventario: {inventory_text}"
                )

            if share > 0:
                lines = _wrap_treemap_words(
                    description,
                    14,
                    1,
                )
                body = (
                    lines[0]
                    if lines
                    else description
                )
                return (
                    f"<b>{body}</b><br>"
                    f"Inv.: {inventory_text}"
                )

            lines = _wrap_treemap_words(
                description,
                14,
                1,
            )
            body = (
                lines[0]
                if lines
                else description
            )
            return (
                f"<b>{body}</b><br>"
                f"Inv.: {inventory_text}"
            )

        family_products["TREEMAP_TEXT"] = family_products.apply(
            build_treemap_text,
            axis=1,
        )

        fig = px.treemap(
            family_products,
            path=["PRODUCTO"],
            values="TREEMAP_SIZE",
            title=(
                f"Distribución de ventas por producto — "
                f"TOP {top_label} — {chart_family}"
            ),
            color="PRODUCTO",
            color_discrete_map=treemap_color_map,
            custom_data=[
                "PRODUCTO",
                "INVENTARIO_TEXTO",
                "ULT_COMPRA_TEXTO",
                "ULT_VENTA_TEXTO",
                "VENTA_TEXTO",
                "TREEMAP_TEXT",
            ],
        )

        fig.update_traces(
            texttemplate="%{customdata[5]}",
            textposition="middle center",
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Ventas mes actual: %{customdata[4]}<br>"
                "Inventario: %{customdata[1]}<br>"
                "Última compra: %{customdata[2]}<br>"
                "Última venta: %{customdata[3]}"
                "<extra></extra>"
            ),
            textfont=dict(
                size=10,
                color="#FFFFFF",
                family=(
                    "Inter, Segoe UI, Arial, sans-serif"
                ),
            ),
            marker=dict(
                line=dict(
                    color="rgba(235,242,248,.70)",
                    width=1.0,
                ),
                pad=dict(
                    t=6,
                    l=6,
                    r=6,
                    b=6,
                ),
            ),
            tiling=dict(
                packing="squarify",
                pad=3,
            ),
            pathbar=dict(
                visible=False,
            ),
        )

        fig.update_layout(
            height=500,
            autosize=True,
            template="plotly_dark",
            paper_bgcolor="#091625",
            plot_bgcolor="#091625",
            font=dict(
                color="#E7EFF6",
                family=(
                    "Inter, Segoe UI, Arial, sans-serif"
                ),
            ),
            title=dict(
                text=(
                    "Distribución de ventas por producto"
                    "<br>"
                    f"<span style='font-size:14px;"
                    "color:#D7AE58;'>"
                    f"{chart_family} · "
                    f"{len(family_products)} productos"
                    "</span>"
                ),
                x=0.5,
                xanchor="center",
                y=0.975,
                yanchor="top",
                font=dict(
                    color="#FFFFFF",
                    size=19,
                    family=(
                        "Inter, Segoe UI, Arial, sans-serif"
                    ),
                ),
            ),
            uniformtext=dict(
                minsize=8,
                mode="hide",
            ),
            showlegend=False,
            hoverlabel=dict(
                bgcolor="#06131F",
                bordercolor="#6886A0",
                font_color="#FFFFFF",
                font_size=11,
            ),
            transition=dict(duration=0),
            uirevision=str(chart_family),
            margin=dict(
                l=18,
                r=18,
                t=82,
                b=18,
            ),
        )

        # ----------------------------------------------------
        # TREEMAP — INTERACCIÓN NATIVA DE STREAMLIT
        # ----------------------------------------------------
        # Se elimina streamlit-plotly-events. El componente nativo
        # evita bloquear el rerun completo del dashboard.
        fig.update_layout(
            clickmode="event+select",
        )

        treemap_event = st.plotly_chart(
            fig,
            use_container_width=True,
            key=(
                "family_product_treemap_native_"
                f"{chart_family}_"
                f"{st.session_state.get('treemap_click_reset_token', 0)}"
            ),
            config={
                "displaylogo": False,
                "responsive": True,
            },
            on_select="rerun",
            selection_mode="points",
        )

        selected_points = []

        try:
            selected_points = list(
                treemap_event.selection.points
            )
        except Exception:
            try:
                selected_points = list(
                    treemap_event
                    .get("selection", {})
                    .get("points", [])
                )
            except Exception:
                selected_points = []

        clicked_products = []

        for selected_point in selected_points:
            product_name = None

            if isinstance(selected_point, dict):
                customdata = selected_point.get(
                    "customdata"
                )

                if isinstance(
                    customdata,
                    (list, tuple),
                ) and customdata:
                    product_name = str(
                        customdata[0]
                    ).strip()

                if not product_name:
                    raw_label = selected_point.get(
                        "label"
                    )

                    if raw_label is not None:
                        product_name = str(
                            raw_label
                        ).strip()

                if not product_name:
                    point_index = selected_point.get(
                        "pointIndex",
                        selected_point.get(
                            "pointNumber"
                        ),
                    )

                    try:
                        point_index = int(
                            point_index
                        )
                    except (TypeError, ValueError):
                        point_index = None

                    if point_index is not None:
                        if (
                            0 <= point_index
                            < len(family_products)
                        ):
                            product_name = str(
                                family_products.iloc[
                                    point_index
                                ]["PRODUCTO"]
                            ).strip()

            if (
                product_name
                and product_name
                in set(
                    family_products[
                        "PRODUCTO"
                    ]
                    .astype(str)
                    .tolist()
                )
                and product_name
                not in clicked_products
            ):
                clicked_products.append(
                    product_name
                )

        if clicked_products:
            existing_products = list(
                st.session_state.get(
                    "selected_products_from_treemap",
                    [],
                )
            )

            # Cada nuevo clic se agrega a la selección existente.
            # Así el usuario puede escoger más de un producto
            # mediante clics sucesivos en el treemap.
            for clicked_product in clicked_products:
                if clicked_product not in existing_products:
                    existing_products.append(
                        clicked_product
                    )

            # Solo conservar productos que pertenecen actualmente
            # al treemap/familia visible.
            valid_treemap_products = set(
                family_products["PRODUCTO"]
                .astype(str)
                .tolist()
            )

            existing_products = [
                product
                for product in existing_products
                if product in valid_treemap_products
            ]

            st.session_state[
                "selected_products_from_treemap"
            ] = existing_products

            st.session_state[
                "selected_product_from_treemap"
            ] = (
                existing_products[-1]
                if existing_products
                else None
            )

            st.session_state[
                "selected_product_family"
            ] = chart_family

        selected_treemap_products_display = list(
            st.session_state.get(
                "selected_products_from_treemap",
                [],
            )
        )

        # El cambio de familia invalida selecciones anteriores.
        if (
            st.session_state.get(
                "selected_product_family"
            )
            not in (None, chart_family)
        ):
            selected_treemap_products_display = []
            st.session_state[
                "selected_products_from_treemap"
            ] = []
            st.session_state[
                "selected_product_from_treemap"
            ] = None
            st.session_state[
                "selected_product_family"
            ] = chart_family

        if selected_treemap_products_display:
            st.caption(
                "🎯 Productos seleccionados "
                f"({len(selected_treemap_products_display)}): "
                + " | ".join(
                    selected_treemap_products_display
                )
                + " · Haz clic en otros recuadros para agregarlos."
            )

    # ========================================================
    # PRODUCTOS DE LA FAMILIA SELECCIONADA
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 📦 Ventas por producto de la familia seleccionada"
    )

    selected_family_products_chart = (
        filtered_df[
            (
                filtered_df["FAMILIA"]
                .astype(str)
                .eq(
                    chart_family
                )
            )
            &
            (
                ~filtered_df["DESCRIPCION"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
                .str.startswith(
                    "AJUSTE",
                    na=False,
                )
            )
        ]
        .groupby(
            [
                "#COD.",
                "DESCRIPCION",
            ],
            dropna=False,
            as_index=False,
        )["$VENTAS MES ACTUAL"]
        .sum()
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .head(family_top_n)
        .reset_index(drop=True)
    )

    selected_family_products_chart[
        "PRODUCTO"
    ] = (
        selected_family_products_chart[
            "#COD."
        ]
        .astype(str)
        + " - "
        + selected_family_products_chart[
            "DESCRIPCION"
        ]
        .astype(str)
    )

    selected_family_products_chart[
        "PRODUCTO_CORTO"
    ] = (
        selected_family_products_chart[
            "PRODUCTO"
        ]
        .astype(str)
    )

    selected_treemap_products = list(
        st.session_state.get(
            "selected_products_from_treemap",
            [],
        )
    )

    if selected_treemap_products:
        selected_family_products_chart = (
            selected_family_products_chart[
                selected_family_products_chart["PRODUCTO"]
                .astype(str)
                .isin(selected_treemap_products)
            ]
            .copy()
            .reset_index(drop=True)
        )

    # El filtrado de productos se controla desde el sidebar.
    # Se elimina el estado oculto del antiguo clic del treemap.

    selected_family_products_chart[
        "VENTA_LABEL"
    ] = (
        selected_family_products_chart[
            "$VENTAS MES ACTUAL"
        ]
        .apply(
            money
        )
    )

    selected_product_colors = [
        "#28C7A4",
        "#D8AE53",
        "#43B9E6",
        "#8DA9C4",
        "#A57DCE",
        "#E47C68",
        "#55BEB4",
        "#D99A50",
        "#739CC1",
        "#7FC5A1",
        "#D96C92",
        "#55C7D5",
        "#8E7BC7",
        "#E0BD58",
        "#56AE91",
    ]

    import plotly.graph_objects as go

    chart_df = (
        selected_family_products_chart
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .head(len(selected_family_products_chart) if selected_treemap_products else family_top_n)
        .copy()
        .reset_index(drop=True)
    )

    total_family_sales_chart = safe_sum(
        filtered_df.loc[
            filtered_df["FAMILIA"]
            .astype(str)
            .eq(chart_family),
            "$VENTAS MES ACTUAL",
        ]
    )

    chart_colors = [
        "#22D3EE",  # cian brillante
        "#F6C453",  # dorado elegante
        "#A855F7",  # violeta intenso
        "#FF6B6B",  # coral vivo
        "#4ADE80",  # verde esmeralda
        "#60A5FA",  # azul eléctrico
        "#F97316",  # naranja moderno
        "#2DD4BF",  # turquesa
        "#FB7185",  # rosa coral
        "#818CF8",  # índigo
    ]

    fig = go.Figure()

    # --------------------------------------------------------
    # Barras principales
    # --------------------------------------------------------

    for idx, row in chart_df.iterrows():

        product_name = str(
            row["PRODUCTO_CORTO"]
        )

        full_product = str(
            row["PRODUCTO"]
        )

        sales_value = float(
            row["$VENTAS MES ACTUAL"]
        )

        sales_label = money(
            sales_value
        )

        fig.add_trace(
            go.Bar(
                x=[sales_value],
                y=[product_name],
                orientation="h",
                name=product_name,
                marker=dict(
                    color=chart_colors[
                        idx % len(chart_colors)
                    ],
                    opacity=0.98,
                    line=dict(
                        color="rgba(255,255,255,.42)",
                        width=1.25,
                    ),
                ),
                text=[sales_label],
                textposition="outside",
                cliponaxis=False,
                customdata=[
                    [full_product]
                ],
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Ventas: %{text}"
                    "<extra></extra>"
                ),
                showlegend=False,
            )
        )

    fig.update_layout(
        template="plotly_dark",
        barmode="overlay",
        height=470,
        paper_bgcolor="#121A28",
        plot_bgcolor="#121A28",
        font=dict(
            color="#E6EDF5",
            family=(
                "Aptos, Inter, Segoe UI, Arial, sans-serif"
            ),
        ),
        title=dict(
            text=(
                "<b>VENTAS POR PRODUCTO</b>"
                "<br>"
                "<span style='font-size:12px;"
                "color:#D8AE53;"
                "letter-spacing:0.05em;'>"
                f"{html.escape(str(chart_family))}"
                "</span>"
            ),
            x=0.5,
            xanchor="center",
            y=0.97,
            yanchor="top",
            font=dict(
                size=19,
                color="#FFFFFF",
                family=(
                    "Aptos Display, Inter, "
                    "Segoe UI, Arial, sans-serif"
                ),
            ),
        ),
        margin=dict(
            l=28,
            r=92,
            t=78,
            b=34,
        ),
        bargap=0.14,
        xaxis=dict(
            title="",
            tickprefix="$",
            tickformat="~s",
            showgrid=True,
            rangemode="tozero",
            gridcolor="rgba(120,150,180,.11)",
            griddash="dot",
            zeroline=False,
            linecolor="#314257",
            tickfont=dict(
                size=9,
                color="#AFC2D6",
            ),
        ),
        yaxis=dict(
            title="",
            categoryorder="array",
            categoryarray=list(
                reversed(
                    chart_df[
                        "PRODUCTO_CORTO"
                    ].tolist()
                )
            ),
            tickfont=dict(
                size=8,
                color="#F2F6FA",
            ),
            ticklabelposition="outside",
            ticklabelstandoff=6,
            ticks="outside",
            ticklen=4,
            tickcolor="rgba(160,185,210,.36)",
            showline=True,
            linecolor="rgba(74,101,130,.34)",
            showgrid=False,
            automargin=True,
        ),
        hoverlabel=dict(
            bgcolor="#0A1220",
            bordercolor="#40546A",
            font_color="#FFFFFF",
            font_size=11,
        ),
        shapes=[
            dict(
                type="rect",
                xref="paper",
                yref="paper",
                x0=-0.025,
                y0=-0.05,
                x1=1.025,
                y1=1.04,
                line=dict(
                    color="rgba(88,126,160,.36)",
                    width=1.2,
                ),
                fillcolor="rgba(0,0,0,0)",
                layer="below",
            ),
        ],
    )

    fig.update_traces(
        textfont=dict(
            size=10,
            color="#F8FBFF",
            family=(
                "Aptos, Inter, Segoe UI, Arial, sans-serif"
            ),
        ),
    )

    # ========================================================
    # COMPARATIVO DE UNIDADES VENDIDAS:
    # MES ANTERIOR VS MES ACTUAL
    # ========================================================

    comparison_products_df = (
        filtered_df[
            (
                filtered_df["FAMILIA"]
                .astype(str)
                .eq(str(chart_family))
            )
            &
            (
                ~filtered_df["DESCRIPCION"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
                .str.startswith(
                    "AJUSTE",
                    na=False,
                )
            )
        ]
        .groupby(
            [
                "#COD.",
                "DESCRIPCION",
            ],
            dropna=False,
            as_index=False,
        )
        .agg(
            {
                "#VENTAS MES ANTERIOR": "sum",
                "#VENTAS MES ACTUAL": "sum",
            }
        )
    )

    comparison_products_df[
        "PRODUCTO"
    ] = (
        comparison_products_df[
            "#COD."
        ]
        .astype(str)
        + " - "
        + comparison_products_df[
            "DESCRIPCION"
        ]
        .astype(str)
    )

    if selected_treemap_products:
        comparison_products_df = (
            comparison_products_df[
                comparison_products_df["PRODUCTO"]
                .astype(str)
                .isin(selected_treemap_products)
            ]
            .copy()
            .reset_index(drop=True)
        )

    # Comparación producto por producto.
    # Si hay un producto del treemap, se conserva solo ese producto.
    comparison_products_df = (
        comparison_products_df
        .sort_values(
            "#VENTAS MES ACTUAL",
            ascending=False,
        )
        .head(len(comparison_products_df) if selected_treemap_products else family_top_n)
        .copy()
        .reset_index(drop=True)
    )

    # Orden invertido para que el producto con mayor venta
    # aparezca en la parte superior de la gráfica horizontal.
    comparison_order = list(
        reversed(
            comparison_products_df[
                "PRODUCTO"
            ].tolist()
        )
    )

    comparison_chart_height = max(
        495,
        150 + (
            len(comparison_products_df)
            * 34
        ),
    )

    fig_compare = go.Figure()

    fig_compare.add_trace(
        go.Bar(
            x=comparison_products_df[
                "#VENTAS MES ANTERIOR"
            ],
            y=comparison_products_df[
                "PRODUCTO"
            ],
            orientation="h",
            name="Ventas mes anterior",
            marker=dict(
                color="#F6B73C",
                opacity=0.98,
                line=dict(
                    color="#FFD978",
                    width=1.25,
                ),
            ),
            text=comparison_products_df[
                "#VENTAS MES ANTERIOR"
            ].apply(
                lambda value:
                f"{value:,.0f}"
            ),
            textposition="outside",
            cliponaxis=False,
            customdata=comparison_products_df[
                ["PRODUCTO"]
            ],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Mes anterior: %{x:,.0f}"
                "<extra></extra>"
            ),
        )
    )

    fig_compare.add_trace(
        go.Bar(
            x=comparison_products_df[
                "#VENTAS MES ACTUAL"
            ],
            y=comparison_products_df[
                "PRODUCTO"
            ],
            orientation="h",
            name="Ventas mes actual",
            marker=dict(
                color="#16D7D0",
                opacity=0.98,
                line=dict(
                    color="#71FFF8",
                    width=1.25,
                ),
            ),
            text=comparison_products_df[
                "#VENTAS MES ACTUAL"
            ].apply(
                lambda value:
                f"{value:,.0f}"
            ),
            textposition="outside",
            cliponaxis=False,
            customdata=comparison_products_df[
                ["PRODUCTO"]
            ],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Mes actual: %{x:,.0f}"
                "<extra></extra>"
            ),
        )
    )

    compare_title_text = (
        (
            "<b>COMPARACIÓN DE VENTAS DE PRODUCTOS SELECCIONADOS</b>"
            if len(selected_treemap_products) > 1
            else "<b>COMPARACIÓN DE VENTAS DEL PRODUCTO SELECCIONADO</b>"
        )
        if selected_treemap_products
        else (
            f"<b>COMPARACIÓN DE VENTAS POR PRODUCTO — "
            f"TOP {top_label}</b>"
        )
    )

    fig_compare.update_layout(
        template="plotly_dark",
        barmode="group",
        height=comparison_chart_height,
        paper_bgcolor="#121A28",
        plot_bgcolor="#121A28",
        font=dict(
            color="#E6EDF5",
            family=(
                "Aptos, Inter, Segoe UI, Arial, sans-serif"
            ),
        ),
        title=dict(
            text=(
                compare_title_text
                + "<br>"
                "<span style='font-size:13px;"
                "color:#D8AE53;"
                "letter-spacing:0.05em;'>"
                f"{html.escape(str(chart_family))}"
                "</span>"
            ),
            x=0.5,
            xanchor="center",
            y=0.97,
            yanchor="top",
            font=dict(
                size=18,
                color="#FFFFFF",
                family=(
                    "Aptos Display, Inter, "
                    "Segoe UI, Arial, sans-serif"
                ),
            ),
        ),
        margin=dict(
            l=230,
            r=72,
            t=92,
            b=88,
        ),
        bargap=0.20,
        bargroupgap=0.08,
        xaxis=dict(
            title="# Ventas",
            showgrid=True,
            gridcolor="rgba(120,150,180,.11)",
            griddash="dot",
            zeroline=False,
            linecolor="#314257",
            tickfont=dict(
                size=9,
                color="#AFC2D6",
            ),
        ),
        yaxis=dict(
            title="",
            categoryorder="array",
            categoryarray=comparison_order,
            tickfont=dict(
                size=8,
                color="#EEF5FB",
            ),
            ticklabelposition="outside",
            ticks="outside",
            ticklen=4,
            tickcolor="rgba(160,185,210,.42)",
            showgrid=False,
            automargin=True,
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.13,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(8,22,38,.45)",
            bordercolor="rgba(110,145,180,.22)",
            borderwidth=1,
            font=dict(
                size=10,
                color="#F3F7FB",
            ),
        ),
        hoverlabel=dict(
            bgcolor="#081625",
            bordercolor="#5B7898",
            font_color="#FFFFFF",
            font_size=11,
        ),
    )

    fig_compare.update_traces(
        textfont=dict(
            size=9,
            color="#FFFFFF",
            family=(
                "Aptos, Inter, Segoe UI, Arial, sans-serif"
            ),
        ),
    )

    # ========================================================
    # MOSTRAR AMBAS GRÁFICAS LADO A LADO
    # ========================================================

    chart_left_col, chart_right_col = st.columns(
        [1.18, 0.82],
        gap="medium",
    )

    with chart_left_col:
        st.plotly_chart(
            fig,
            use_container_width=True,
            key="products_selected_family_chart",
        )

    with chart_right_col:
        st.plotly_chart(
            fig_compare,
            use_container_width=True,
            key="products_month_comparison_chart",
        )


    # ========================================================
    # DETALLE POR PRODUCTO
    # ========================================================

    st.markdown("---")

    selected_detail_family = (
        st.session_state.get(
            "family_product_chart"
        )
    )

    if (
        selected_detail_family
        and "FAMILIA" in product_df.columns
    ):
        product_detail_filtered = (
            product_df[
                product_df["FAMILIA"]
                .astype(str)
                .eq(
                    str(
                        selected_detail_family
                    )
                )
            ]
            .copy()
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=False,
            )
            .reset_index(drop=True)
        )
    else:
        product_detail_filtered = (
            product_df.copy()
        )

    product_detail_filtered = (
        product_detail_filtered
        .head(top_data_n)
        .copy()
        .reset_index(drop=True)
    )

    detail_family_label = (
        str(selected_detail_family)
        if selected_detail_family
        else "Todas las familias"
    )

    with st.expander(
        f"🧾 Detalle por producto · TOP {top_label} · {detail_family_label}",
        expanded=False,
    ):
        st.markdown(
            (
                '<div class="executive-table-title" '
                'style="border-left-color:#00D8F0;'
                'margin-top:2px;">'
                f'🧾 Detalle por producto — TOP {top_label}'
                '<span style="margin-left:10px;'
                'color:#D7AE58;'
                'font-size:.72rem;'
                'font-weight:800;">'
                f'{html.escape(detail_family_label)}'
                '</span>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

        render_executive_html_table(
            product_detail_filtered,
            column_labels={
                "FAMILIA": "Familia",
                "#COD.": "Código",
                "DESCRIPCION": "Descripción",
                "$VENTAS MES ACTUAL": "Ventas Mes Actual ($)",
                "#VENTAS MES ANTERIOR": "Ventas Mes Anterior",
                "#VENTAS MES ACTUAL": "Ventas Mes Actual",
                "#VENTAS PERDIDAS": "Ventas Perdidas",
                "INVENTARIO TOTAL": "Inventario Total",
                "COSTO ACTUAL": "Costo Actual ($)",
                "PRECIO PROMEDIO": "Precio Promedio ($)",
                "% MAGERN S/VENTA": "Margen (%)",
                "$UTILIDAD MES ACTUAL": "Utilidad Mes Actual ($)",
            },
            formats={
                "$VENTAS MES ACTUAL": "money",
                "#VENTAS MES ANTERIOR": "number2",
                "#VENTAS MES ACTUAL": "number2",
                "#VENTAS PERDIDAS": "number2",
                "INVENTARIO TOTAL": "number2",
                "COSTO ACTUAL": "money",
                "PRECIO PROMEDIO": "money",
                "% MAGERN S/VENTA": "percent",
                "$UTILIDAD MES ACTUAL": "money",
            },
            max_height=620,
        )


# ============================================================
# PESTAÑA 3 — VENTAS Y UTILIDAD POR FAMILIA
# ============================================================

with tab_ventas_utilidad_familia:

    FAMILY_EXEC_COLORS = [
        "#22C7A9",  # esmeralda vivo
        "#D8AE53",  # dorado
        "#3AAFE8",  # azul brillante
        "#7E8CE0",  # índigo
        "#A86DD6",  # violeta
        "#E66F61",  # coral
        "#36B8C7",  # turquesa
        "#E39A4A",  # ámbar
        "#6FA8DC",  # azul acero
        "#58C98D",  # verde
        "#D96C92",  # rosa elegante
        "#49C2D0",  # cian
        "#8D79C6",  # púrpura
        "#E5C25A",  # amarillo dorado
        "#4FB18F",  # jade
        "#8CA6BF",  # azul gris
        "#D8896F",  # salmón
        "#5AB6A9",  # verde agua
        "#B89B5F",  # bronce
        "#708A79",  # verde pizarra
    ]

    st.markdown(
        "### 💰 Ventas y utilidad por familia"
    )

    st.caption(
        "Análisis ejecutivo de participación, utilidad y concentración "
        f"de ventas por familia. Se muestran las {top_n} familias principales."
    )

    # --------------------------------------------------------
    # DATOS BASE TOP DINÁMICO
    # --------------------------------------------------------

    family_sales_top20 = (
        filtered_df
        .groupby(
            "FAMILIA",
            dropna=False,
            as_index=False,
        )["$VENTAS MES ACTUAL"]
        .sum()
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .head(top_n)
        .reset_index(drop=True)
    )

    family_profit_top20 = (
        filtered_df
        .groupby(
            "FAMILIA",
            dropna=False,
            as_index=False,
        )["$UTILIDAD MES ACTUAL"]
        .sum()
        .sort_values(
            "$UTILIDAD MES ACTUAL",
            ascending=False,
        )
        .head(top_n)
        .reset_index(drop=True)
    )

    selected_family_summary = (
        st.session_state.get(
            "selected_family_from_chart"
        )
    )

    if selected_family_summary:
        family_summary_df = (
            filtered_df[
                filtered_df["FAMILIA"]
                .astype(str)
                .eq(
                    str(selected_family_summary)
                )
            ]
            .copy()
        )

        family_summary_label = str(
            selected_family_summary
        )
    else:
        family_summary_df = (
            filtered_df.copy()
        )

        family_summary_label = (
            "TODAS LAS FAMILIAS"
        )

    total_family_sales_kpi = safe_sum(
        family_summary_df[
            "$VENTAS MES ACTUAL"
        ]
    )

    total_family_profit_kpi = safe_sum(
        family_summary_df[
            "$UTILIDAD MES ACTUAL"
        ]
    )

    total_family_inventory_kpi = safe_sum(
        family_summary_df[
            "INVENTARIO TOTAL"
        ]
    )

    family_summary_empaq_values = (
        family_summary_df["EMPAQ_ANALISIS"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    family_summary_empaq_values = [
        value
        for value in family_summary_empaq_values.unique().tolist()
        if value
        and value.upper() not in {"NAN", "NONE"}
    ]

    if len(family_summary_empaq_values) == 1:
        family_summary_inventory_empaq = (
            family_summary_empaq_values[0]
        )
    elif len(family_summary_empaq_values) > 1:
        family_summary_inventory_empaq = (
            "VARIOS EMPAQUES"
        )
    else:
        family_summary_inventory_empaq = (
            "SIN EMPAQUE"
        )

    k_sales, k_profit, k_inventory = st.columns(
        3,
        gap="large",
    )

    with k_sales:
        st.markdown(
            (
                '<div class="family-compact-kpi" '
                'style="--family-kpi-accent:#43B9E6;">'
                '<div class="family-compact-kpi-label">'
                f'💵 VENTAS · {html.escape(family_summary_label)}'
                '</div>'
                '<div class="family-compact-kpi-value">'
                f'{money(total_family_sales_kpi)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with k_profit:
        st.markdown(
            (
                '<div class="family-compact-kpi" '
                'style="--family-kpi-accent:#27C99B;">'
                '<div class="family-compact-kpi-label">'
                f'📈 UTILIDAD · {html.escape(family_summary_label)}'
                '</div>'
                '<div class="family-compact-kpi-value">'
                f'{money(total_family_profit_kpi)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with k_inventory:
        st.markdown(
            (
                '<div class="family-compact-kpi" '
                'style="--family-kpi-accent:#A984D8;">'
                '<div class="family-compact-kpi-label">'
                f'📦 INVENTARIO · {html.escape(family_summary_label)}'
                '</div>'
                '<div class="family-compact-kpi-value">'
                f'{quantity(total_family_inventory_kpi, 2)}'
                '</div>'
                '<div class="family-compact-kpi-note">'
                f'EMPAQ: {html.escape(family_summary_inventory_empaq)}'
                '</div>'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    st.markdown(
        "<div style='height:8px'></div>",
        unsafe_allow_html=True,
    )

    st.markdown("---")

    st.markdown(
        "<div style='height:10px'></div>",
        unsafe_allow_html=True,
    )

    # ========================================================
    # FILA 1 — PARTICIPACIÓN + UTILIDAD
    # ========================================================

    st.markdown(
        """
        <div style="
            text-align:center;
            font-family:Inter,'Segoe UI','Helvetica Neue',Arial,sans-serif;
            font-size:1.05rem;
            font-weight:750;
            color:#DCE8F5;
            letter-spacing:0.025em;
            margin:0.20rem 0 0.90rem 0;
        ">
            HAGA CLIC EN CUALQUIERA DE LAS BARRAS PARA MOSTRAR
            LA INFORMACIÓN DE LA FAMILIA.
        </div>
        """,
        unsafe_allow_html=True,
    )

    chart_sales_col, chart_clear_col, chart_profit_col = st.columns(
        [1.08, 0.16, 1.08],
        gap="small",
    )

    with chart_clear_col:
        st.markdown(
            "<div style='height:245px'></div>",
            unsafe_allow_html=True,
        )

        st.button(
            "Borrar filtro",
            key="clear_family_filter_between_charts",
            use_container_width=True,
            on_click=clear_family_selection,
        )

    # --------------------------------------------------------
    # PARTICIPACIÓN DE VENTAS — TOP DINÁMICO
    # RANKING EJECUTIVO INTERACTIVO
    # --------------------------------------------------------

    with chart_sales_col:

        participation_base_df = (
            family_sales_top20
            .copy()
            .reset_index(drop=True)
        )

        participation_total = safe_sum(
            participation_base_df[
                "$VENTAS MES ACTUAL"
            ]
        )

        selected_family_exec = (
            st.session_state.get(
                "selected_family_from_chart"
            )
        )

        if (
            selected_family_exec
            and selected_family_exec
            in participation_base_df[
                "FAMILIA"
            ].astype(str).tolist()
        ):
            participation_df = (
                participation_base_df[
                    participation_base_df[
                        "FAMILIA"
                    ]
                    .astype(str)
                    .eq(
                        selected_family_exec
                    )
                ]
                .copy()
                .reset_index(drop=True)
            )
        else:
            participation_df = (
                participation_base_df
                .copy()
            )

        if participation_total > 0:
            participation_df[
                "PARTICIPACION_PCT"
            ] = (
                participation_df[
                    "$VENTAS MES ACTUAL"
                ]
                / participation_total
                * 100
            )
        else:
            participation_df[
                "PARTICIPACION_PCT"
            ] = 0.0

        participation_df[
            "FAMILIA_DISPLAY"
        ] = (
            participation_df["FAMILIA"]
            .astype(str)
            .apply(
                lambda value:
                value
                if len(value) <= 30
                else value[:29].rstrip()
                + "…"
            )
        )

        participation_df = (
            participation_df
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=True,
            )
            .reset_index(drop=True)
        )

        # Colores ejecutivos:
        # seleccionado en dorado, resto en esmeralda/azul suave.
        bar_colors = []

        for color_index, (_, row) in enumerate(
            participation_df.iterrows()
        ):

            family_name = str(
                row["FAMILIA"]
            )

            base_color = FAMILY_EXEC_COLORS[
                color_index
                % len(FAMILY_EXEC_COLORS)
            ]

            if (
                selected_family_exec
                and family_name
                == selected_family_exec
            ):
                bar_colors.append(
                    "#F0D487"
                )
            else:
                bar_colors.append(
                    base_color
                )

        import plotly.graph_objects as go

        participation_fig = go.Figure()

        participation_fig.add_trace(
            go.Bar(
                x=participation_df[
                    "PARTICIPACION_PCT"
                ],
                y=participation_df[
                    "FAMILIA_DISPLAY"
                ],
                orientation="h",
                customdata=[
                    [
                        str(row["FAMILIA"]),
                        float(
                            row[
                                "$VENTAS MES ACTUAL"
                            ]
                        ),
                    ]
                    for _, row
                    in participation_df.iterrows()
                ],
                marker=dict(
                    color=bar_colors,
                    line=dict(
                        color="rgba(255,255,255,.18)",
                        width=.9,
                    ),
                ),
                text=[
                    (
                        f"{pct:.1f}%  ·  "
                        f"${sales:,.0f}"
                    )
                    for pct, sales
                    in zip(
                        participation_df[
                            "PARTICIPACION_PCT"
                        ],
                        participation_df[
                            "$VENTAS MES ACTUAL"
                        ],
                    )
                ],
                textposition="outside",
                textfont=dict(
                    size=10,
                    color="#C8D6E4",
                    family=(
                        "Inter, Segoe UI, Arial, "
                        "sans-serif"
                    ),
                ),
                cliponaxis=False,
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Ventas: $%{customdata[1]:,.2f}<br>"
                    "Participación: %{x:.2f}%"
                    "<extra></extra>"
                ),
                showlegend=False,
            )
        )

        participation_fig.update_layout(
            template="plotly_dark",
            height=(
                320
                if selected_family_exec
                and selected_family_exec
                in participation_base_df[
                    "FAMILIA"
                ].astype(str).tolist()
                else max(
                    640,
                    180 + len(participation_df) * 17,
                )
            ),
            paper_bgcolor="#0E1B2B",
            plot_bgcolor="#0E1B2B",
            title=dict(
                text=(
                    "<b>PARTICIPACIÓN DE VENTAS</b>"
                    "<br>"
                    "<span style='font-size:14px;"
                    "color:#D7AE58;'>"
                    + (
                        html.escape(
                            selected_family_exec
                        )
                        if selected_family_exec
                        and selected_family_exec
                        in participation_base_df[
                            "FAMILIA"
                        ].astype(str).tolist()
                        else f"TOP {top_label} FAMILIAS"
                    )
                    + "</span>"
                ),
                x=0.5,
                xanchor="center",
                y=0.93,
                yanchor="top",
                pad=dict(
                    t=8,
                    b=6,
                ),
                font=dict(
                    size=18,
                    color="#FFFFFF",
                    family=(
                        "Inter, Segoe UI, Arial, "
                        "sans-serif"
                    ),
                ),
            ),
            margin=dict(
                l=205,
                r=115,
                t=112,
                b=58,
            ),
            xaxis=dict(
                title="Participación en ventas (%)",
                ticksuffix="%",
                showgrid=True,
                gridcolor="rgba(65,88,111,.20)",
                griddash="dot",
                zeroline=False,
                linecolor="#30465E",
                tickfont=dict(
                    size=9,
                    color="#8297AC",
                ),
                title_font=dict(
                    size=10,
                    color="#AEBECC",
                ),
            ),
            yaxis=dict(
                title="",
                tickfont=dict(
                    size=10,
                    color="#D3DFEA",
                ),
                automargin=True,
                showgrid=False,
            ),
            hoverlabel=dict(
                bgcolor="#07121E",
                font_color="#FFFFFF",
                bordercolor="#35506B",
                font_size=11,
            ),
        )

        participation_event = st.plotly_chart(
            participation_fig,
            use_container_width=True,
            key=(
                "family_participation_rank_"
                f"{st.session_state.get('family_selection_reset_token', 0)}"
            ),
            on_select="rerun",
            selection_mode="points",
        )

        participation_points = []

        try:
            participation_points = (
                participation_event
                .selection
                .points
            )
        except Exception:
            try:
                participation_points = (
                    participation_event
                    .get(
                        "selection",
                        {},
                    )
                    .get(
                        "points",
                        [],
                    )
                )
            except Exception:
                participation_points = []

        if participation_points:

            selected_point = (
                participation_points[0]
            )

            selected_custom = (
                selected_point.get(
                    "customdata"
                )
            )

            clicked_family = None

            if isinstance(
                selected_custom,
                (list, tuple),
            ):
                if selected_custom:
                    clicked_family = str(
                        selected_custom[0]
                    )

            elif selected_custom is not None:
                clicked_family = str(
                    selected_custom
                )

            if (
                clicked_family
                and clicked_family
                != st.session_state.get(
                    "selected_family_from_chart"
                )
            ):
                st.session_state[
                    "selected_family_from_chart"
                ] = clicked_family

                st.rerun()

        # ----------------------------------------------------
        # FICHA EJECUTIVA DE LA FAMILIA SELECCIONADA
        # ----------------------------------------------------

        selected_family_exec = (
            st.session_state.get(
                "selected_family_from_chart"
            )
        )

        if selected_family_exec:

            selected_df = (
                filtered_df[
                    filtered_df["FAMILIA"]
                    .astype(str)
                    .eq(
                        selected_family_exec
                    )
                ]
                .copy()
            )

            if not selected_df.empty:

                selected_sales = safe_sum(
                    selected_df[
                        "$VENTAS MES ACTUAL"
                    ]
                )

                selected_profit = safe_sum(
                    selected_df[
                        "$UTILIDAD MES ACTUAL"
                    ]
                )

                selected_inventory = safe_sum(
                    selected_df[
                        "INVENTARIO TOTAL"
                    ]
                )

                selected_units = safe_sum(
                    selected_df[
                        "#VENTAS MES ACTUAL"
                    ]
                )

                selected_margin = safe_mean(
                    selected_df[
                        "% MAGERN S/VENTA"
                    ]
                )

                selected_product_count = int(
                    selected_df[
                        "#COD."
                    ]
                    .nunique(
                        dropna=True
                    )
                )

                selected_share = (
                    (
                        selected_sales
                        / participation_total
                        * 100
                    )
                    if participation_total > 0
                    else 0
                )

                family_card_html = (
                    f'<div class="family-meeting-card">'
                    f'<div class="family-meeting-top">'
                    f'<div>'
                    f'<div class="family-meeting-eyebrow">'
                    f'FAMILIA SELECCIONADA'
                    f'</div>'
                    f'<div class="family-meeting-name">'
                    f'{html.escape(selected_family_exec)}'
                    f'</div>'
                    f'</div>'
                    f'<div class="family-meeting-share">'
                    f'{selected_share:.1f}%'
                    f'<span>participación</span>'
                    f'</div>'
                    f'</div>'
                    f'<div class="family-meeting-grid">'
                    f'<div><span>Ventas</span>'
                    f'<b>{money(selected_sales)}</b></div>'
                    f'<div><span>Utilidad</span>'
                    f'<b>{money(selected_profit)}</b></div>'
                    f'<div><span>Margen</span>'
                    f'<b>{percent(selected_margin)}</b></div>'
                    f'<div><span>Inventario</span>'
                    f'<b>{quantity(selected_inventory, 2)}</b></div>'
                    f'<div><span>Productos</span>'
                    f'<b>{selected_product_count:,}</b></div>'
                    f'<div><span>Unidades</span>'
                    f'<b>{quantity(selected_units, 0)}</b></div>'
                    f'</div>'
                    f'</div>'
                )

                st.markdown(
                    family_card_html,
                    unsafe_allow_html=True,
                )

                st.button(
                    "Borrar selección",
                    key="family_meeting_clear",
                    use_container_width=True,
                    on_click=clear_family_selection,
                )

    # --------------------------------------------------------
    # UTILIDAD POR FAMILIA — TOP DINÁMICO
    # --------------------------------------------------------

    with chart_profit_col:

        selected_family_for_utility = (
            st.session_state.get(
                "selected_family_from_chart"
            )
        )

        if selected_family_for_utility:
            utility_chart = (
                filtered_df[
                    filtered_df["FAMILIA"]
                    .astype(str)
                    .eq(
                        selected_family_for_utility
                    )
                ]
                .groupby(
                    "FAMILIA",
                    dropna=False,
                    as_index=False,
                )["$UTILIDAD MES ACTUAL"]
                .sum()
                .reset_index(drop=True)
            )
        else:
            utility_chart = (
                family_profit_top20
                .copy()
            )

        utility_chart[
            "FAMILIA_CORTA"
        ] = (
            utility_chart["FAMILIA"]
            .astype(str)
            .apply(
                lambda value:
                value
                if len(value) <= 30
                else value[:29].rstrip()
                + "…"
            )
        )

        fig = px.bar(
            utility_chart.sort_values(
                "$UTILIDAD MES ACTUAL",
                ascending=True,
            ),
            x="$UTILIDAD MES ACTUAL",
            y="FAMILIA_CORTA",
            orientation="h",
            custom_data=[
                "FAMILIA",
            ],
            title=None,
            labels={
                "$UTILIDAD MES ACTUAL":
                    "Utilidad ($)",
                "FAMILIA_CORTA":
                    "Familia",
            },
        )

        apply_executive_bar_style(
            fig
        )

        utility_sorted = (
            utility_chart.sort_values(
                "$UTILIDAD MES ACTUAL",
                ascending=True,
            )
            .reset_index(drop=True)
        )

        utility_colors = [
            (
                "#F0D487"
                if selected_family_for_utility
                else FAMILY_EXEC_COLORS[
                    i % len(FAMILY_EXEC_COLORS)
                ]
            )
            for i in range(
                len(utility_sorted)
            )
        ]

        if fig.data:
            fig.data[0].marker.color = (
                utility_colors
            )
            fig.data[0].marker.line.color = (
                "rgba(255,255,255,.12)"
            )
            fig.data[0].marker.line.width = .8

        fig.update_traces(
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Utilidad: $%{x:,.2f}"
                "<extra></extra>"
            ),
        )

        fig.update_layout(
            height=(
                320
                if selected_family_for_utility
                else max(
                    640,
                    180 + len(utility_chart) * 17,
                )
            ),
            title=dict(
                text=(
                    "<b>UTILIDAD POR FAMILIA</b>"
                    "<br>"
                    "<span style='font-size:14px;"
                    "color:#D7AE58;'>"
                    f"TOP {top_label} FAMILIAS"
                    "</span>"
                ),
                x=0.5,
                xanchor="center",
                y=0.93,
                yanchor="top",
                pad=dict(
                    t=8,
                    b=6,
                ),
                font=dict(
                    size=18,
                    color="#FFFFFF",
                    family=(
                        "Inter, Segoe UI, Arial, "
                        "sans-serif"
                    ),
                ),
            ),
            margin=dict(
                l=205,
                r=35,
                t=112,
                b=58,
            ),
            yaxis=dict(
                tickfont=dict(
                    size=10,
                    color="#D3DFEA",
                ),
                automargin=True,
            ),
        )

        profit_event = st.plotly_chart(
            fig,
            use_container_width=True,
            key=(
                "family_profit_top20_chart_"
                f"{st.session_state.get('family_selection_reset_token', 0)}"
            ),
            on_select="rerun",
            selection_mode="points",
        )

        profit_points = []

        try:
            profit_points = (
                profit_event
                .selection
                .points
            )
        except Exception:
            try:
                profit_points = (
                    profit_event
                    .get(
                        "selection",
                        {},
                    )
                    .get(
                        "points",
                        [],
                    )
                )
            except Exception:
                profit_points = []

        if profit_points:

            profit_custom = (
                profit_points[0]
                .get(
                    "customdata"
                )
            )

            profit_family = None

            if isinstance(
                profit_custom,
                (list, tuple),
            ):
                if profit_custom:
                    profit_family = str(
                        profit_custom[0]
                    )

            elif profit_custom is not None:
                profit_family = str(
                    profit_custom
                )

            if (
                profit_family
                and profit_family
                != st.session_state.get(
                    "selected_family_from_chart"
                )
            ):
                st.session_state[
                    "selected_family_from_chart"
                ] = profit_family

                st.rerun()

    # ========================================================
    # TABLA AMPLIA — PRODUCTOS DE LA FAMILIA SELECCIONADA
    # ========================================================

    selected_family_table = (
        st.session_state.get(
            "selected_family_from_chart"
        )
    )

    if selected_family_table:

        family_products_full_df = (
            filtered_df[
                filtered_df["FAMILIA"]
                .astype(str)
                .eq(
                    selected_family_table
                )
            ]
            .copy()
        )

        if not family_products_full_df.empty:

            family_products_full_table = (
                family_products_full_df[
                    [
                        "#COD.",
                        "DESCRIPCION",
                        "$VENTAS MES ACTUAL",
                        "#VENTAS MES ACTUAL",
                        "$UTILIDAD MES ACTUAL",
                        "% MAGERN S/VENTA",
                        "INVENTARIO TOTAL",
                    ]
                ]
                .copy()
                .sort_values(
                    "$VENTAS MES ACTUAL",
                    ascending=False,
                )
                .reset_index(drop=True)
            )

            st.markdown("---")

            st.markdown(
                (
                    '<div class="family-products-wide-title">'
                    '<div>'
                    '<span class="family-products-wide-eyebrow">'
                    'DETALLE DE PRODUCTOS'
                    '</span>'
                    '<strong>'
                    f'{html.escape(str(selected_family_table))}'
                    '</strong>'
                    '</div>'
                    '<span class="family-products-wide-count">'
                    f'{len(family_products_full_table):,} productos'
                    '</span>'
                    '</div>'
                ),
                unsafe_allow_html=True,
            )

            # --------------------------------------------
            # Construcción de tabla HTML a ancho completo
            # --------------------------------------------

            family_table_rows = []

            for _, product_row in (
                family_products_full_table
                .iterrows()
            ):

                code_value = (
                    ""
                    if pd.isna(
                        product_row["#COD."]
                    )
                    else html.escape(
                        str(
                            product_row[
                                "#COD."
                            ]
                        )
                    )
                )

                product_value = (
                    ""
                    if pd.isna(
                        product_row[
                            "DESCRIPCION"
                        ]
                    )
                    else html.escape(
                        str(
                            product_row[
                                "DESCRIPCION"
                            ]
                        )
                    )
                )

                sales_value = (
                    money(
                        product_row[
                            "$VENTAS MES ACTUAL"
                        ]
                    )
                )

                units_value = (
                    quantity(
                        product_row[
                            "#VENTAS MES ACTUAL"
                        ],
                        2,
                    )
                )

                profit_value = (
                    money(
                        product_row[
                            "$UTILIDAD MES ACTUAL"
                        ]
                    )
                )

                margin_value = (
                    percent(
                        product_row[
                            "% MAGERN S/VENTA"
                        ]
                    )
                )

                inventory_value = (
                    quantity(
                        product_row[
                            "INVENTARIO TOTAL"
                        ],
                        2,
                    )
                )

                family_table_rows.append(
                    (
                        '<tr>'
                        f'<td class="code-cell">{code_value}</td>'
                        f'<td class="product-cell">{product_value}</td>'
                        f'<td class="num-cell">{sales_value}</td>'
                        f'<td class="num-cell">{units_value}</td>'
                        f'<td class="num-cell">{profit_value}</td>'
                        f'<td class="num-cell margin-cell">{margin_value}</td>'
                        f'<td class="num-cell">{inventory_value}</td>'
                        '</tr>'
                    )
                )

            family_table_html = (
                '<div class="family-products-wide-shell">'
                '<table class="family-products-wide-table">'
                '<colgroup>'
                '<col style="width:8%">'
                '<col style="width:32%">'
                '<col style="width:14%">'
                '<col style="width:11%">'
                '<col style="width:14%">'
                '<col style="width:9%">'
                '<col style="width:12%">'
                '</colgroup>'
                '<thead>'
                '<tr>'
                '<th>Código</th>'
                '<th>Producto</th>'
                '<th>Ventas Mes Actual ($)</th>'
                '<th>Unidades Vendidas</th>'
                '<th>Utilidad Mes Actual ($)</th>'
                '<th>Margen (%)</th>'
                '<th>Inventario</th>'
                '</tr>'
                '</thead>'
                '<tbody>'
                + "".join(
                    family_table_rows
                )
                + '</tbody>'
                '</table>'
                '</div>'
            )

            st.markdown(
                family_table_html,
                unsafe_allow_html=True,
            )


    # ========================================================
    # PARETO
    # ========================================================
    # La gráfica Pareto fue trasladada a la pestaña "Pareto".
    # ========================================================


# ============================================================
# PESTAÑA 4 — PARETO
# ============================================================

with tab_pareto:

    st.markdown(
        "### 📊 Pareto 80/20 por familia"
    )

    st.caption(
        "Análisis estático de concentración de ventas por familia. "
        "Las familias se ordenan de mayor a menor venta y la línea "
        "muestra el porcentaje acumulado. La referencia del 80% "
        "identifica el grupo de familias que concentra la mayor parte "
        "de las ventas."
    )

    # ========================================================
    # PARETO 80/20 — TODAS LAS FAMILIAS FILTRADAS
    # ========================================================
    # Para aplicar correctamente Pareto, el porcentaje acumulado
    # debe calcularse contra el total de TODAS las familias, no
    # únicamente contra TOP 10/20/50.
    pareto_df = (
        filtered_df
        .groupby(
            "FAMILIA",
            dropna=False,
            as_index=False,
        )["$VENTAS MES ACTUAL"]
        .sum()
        .sort_values(
            "$VENTAS MES ACTUAL",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    # Excluir familias sin nombre únicamente de la visualización.
    pareto_df["FAMILIA"] = (
        pareto_df["FAMILIA"]
        .astype("string")
        .fillna("SIN FAMILIA")
        .astype(str)
    )

    pareto_total = safe_sum(
        pareto_df[
            "$VENTAS MES ACTUAL"
        ]
    )

    if pareto_total > 0:
        pareto_df[
            "PORCENTAJE"
        ] = (
            pareto_df[
                "$VENTAS MES ACTUAL"
            ]
            / pareto_total
            * 100
        )

        pareto_df[
            "PORCENTAJE_ACUMULADO"
        ] = (
            pareto_df[
                "PORCENTAJE"
            ]
            .cumsum()
        )
    else:
        pareto_df[
            "PORCENTAJE"
        ] = 0.0

        pareto_df[
            "PORCENTAJE_ACUMULADO"
        ] = 0.0

    total_families_pareto = int(
        len(pareto_df)
    )

    # Primera posición que alcanza o supera el 80%.
    if (
        pareto_total > 0
        and total_families_pareto > 0
    ):
        reaches_80 = (
            pareto_df[
                "PORCENTAJE_ACUMULADO"
            ]
            >= 80.0
        )

        if reaches_80.any():
            pareto_cutoff_index = int(
                reaches_80.idxmax()
            )
        else:
            pareto_cutoff_index = (
                total_families_pareto - 1
            )

        pareto_critical_count = (
            pareto_cutoff_index + 1
        )
    else:
        pareto_cutoff_index = 0
        pareto_critical_count = 0

    pareto_critical_pct = (
        (
            pareto_critical_count
            / total_families_pareto
            * 100
        )
        if total_families_pareto > 0
        else 0.0
    )

    pareto_df[
        "GRUPO_PARETO"
    ] = [
        (
            "Familias que concentran el 80%"
            if index
            < pareto_critical_count
            else "Resto de familias"
        )
        for index in range(
            total_families_pareto
        )
    ]

    # Para el Pareto mostramos el nombre COMPLETO de cada familia.
    # No se recorta con puntos suspensivos.
    pareto_df[
        "FAMILIA_CORTA"
    ] = (
        pareto_df["FAMILIA"]
        .astype(str)
        .str.strip()
    )

    # La lógica Pareto se calcula con TODAS las familias,
    # pero la gráfica muestra únicamente las familias necesarias
    # para alcanzar o superar el 80% de las ventas.
    pareto_display_df = (
        pareto_df
        .head(
            max(
                pareto_critical_count,
                1,
            )
        )
        .copy()
        .reset_index(drop=True)
    )

    # --------------------------------------------------------
    # Resumen ejecutivo Pareto 80/20
    # --------------------------------------------------------
    pareto_explanation_html = (
        '<div style="'
        'background:#102236;'
        'border:1px solid #29445F;'
        'border-left:4px solid #FFD166;'
        'border-radius:10px;'
        'padding:12px 16px;'
        'margin:0.35rem 0 0.85rem 0;'
        'font-family:Inter,Segoe UI,Helvetica Neue,Arial,sans-serif;'
        'color:#DCE8F5;'
        'line-height:1.55;'
        '">'
        '<div style="font-size:.95rem;font-weight:800;'
        'color:#FFD166;margin-bottom:5px;">'
        '¿Qué significa la regla Pareto 80/20?'
        '</div>'
        '<div style="font-size:.84rem;">'
        'La regla Pareto indica que una parte pequeña de las familias '
        'puede concentrar la mayor parte de las ventas. '
        'En esta gráfica, las barras muestran las ventas de cada familia '
        'ordenadas de mayor a menor, mientras que la línea amarilla muestra '
        'el porcentaje acumulado de ventas. '
        'La línea roja horizontal marca el 80%. '
        'Las familias ubicadas antes del punto de corte representan el grupo '
        'que, en conjunto, concentra aproximadamente el 80% de las ventas. '
        'Esto permite identificar rápidamente cuáles familias tienen mayor '
        'impacto comercial y requieren mayor atención en inventario, compras '
        'y estrategia de ventas.'
        '</div>'
        '</div>'
    )

    st.markdown(
        pareto_explanation_html,
        unsafe_allow_html=True,
    )

    pareto_summary_html = (
        '<div style="display:flex;flex-wrap:wrap;gap:10px;'
        'margin:0.35rem 0 0.85rem 0;">'

        '<div style="flex:1 1 240px;background:#102236;'
        'border:1px solid #29445F;border-radius:10px;'
        'padding:10px 14px;">'
        '<div style="font-size:.68rem;color:#8FA8C1;'
        'font-weight:700;letter-spacing:.04em;">'
        'FAMILIAS QUE ALCANZAN EL 80%'
        '</div>'
        '<div style="font-size:1.20rem;color:#FFD166;'
        'font-weight:850;margin-top:3px;">'
        f'{pareto_critical_count:,} de {total_families_pareto:,}'
        '</div>'
        '</div>'

        '<div style="flex:1 1 240px;background:#102236;'
        'border:1px solid #29445F;border-radius:10px;'
        'padding:10px 14px;">'
        '<div style="font-size:.68rem;color:#8FA8C1;'
        'font-weight:700;letter-spacing:.04em;">'
        '% DE FAMILIAS NECESARIAS'
        '</div>'
        '<div style="font-size:1.20rem;color:#43D7C0;'
        'font-weight:850;margin-top:3px;">'
        f'{pareto_critical_pct:.1f}%'
        '</div>'
        '</div>'

        '<div style="flex:1 1 240px;background:#102236;'
        'border:1px solid #29445F;border-radius:10px;'
        'padding:10px 14px;">'
        '<div style="font-size:.68rem;color:#8FA8C1;'
        'font-weight:700;letter-spacing:.04em;">'
        'VENTAS TOTALES ANALIZADAS'
        '</div>'
        '<div style="font-size:1.20rem;color:#54C8FF;'
        'font-weight:850;margin-top:3px;">'
        f'{money(pareto_total)}'
        '</div>'
        '</div>'

        '</div>'
    )

    st.markdown(
        pareto_summary_html,
        unsafe_allow_html=True,
    )

    import plotly.graph_objects as go

    pareto_fig = go.Figure()

    # Barras:
    # - grupo que alcanza el 80%: destacado
    # - resto: tono tenue
    pareto_executive_palette = [
        "#2FC3A4",  # verde turquesa
        "#4DA3FF",  # azul ejecutivo
        "#F2B84B",  # dorado
        "#8B6FE8",  # violeta
        "#E86E7A",  # coral
        "#39B7C8",  # cyan
        "#6CBF84",  # verde
        "#D8894B",  # cobre
        "#6F8FC9",  # azul acero
        "#B879D6",  # púrpura
        "#4FB6A8",  # teal
        "#D6A84F",  # mostaza
        "#5FA8D3",  # azul claro
        "#C66C84",  # rosa ejecutivo
        "#7DAA55",  # oliva vivo
        "#A978D1",  # lavanda
        "#E08E45",  # naranja sobrio
        "#3E9E8F",  # verde petróleo
        "#7089C4",  # índigo suave
        "#C98B55",  # bronce
    ]

    pareto_bar_colors = [
        pareto_executive_palette[
            index
            % len(
                pareto_executive_palette
            )
        ]
        for index in range(
            len(pareto_display_df)
        )
    ]

    pareto_fig.add_trace(
        go.Bar(
            x=pareto_display_df[
                "FAMILIA_CORTA"
            ],
            y=pareto_display_df[
                "$VENTAS MES ACTUAL"
            ],
            name="Ventas",
            marker=dict(
                color=pareto_bar_colors,
                line=dict(
                    color=(
                        "rgba(255,255,255,.12)"
                    ),
                    width=.7,
                ),
            ),
            hoverinfo="skip",
        )
    )

    pareto_fig.add_trace(
        go.Scatter(
            x=pareto_display_df[
                "FAMILIA_CORTA"
            ],
            y=pareto_display_df[
                "PORCENTAJE_ACUMULADO"
            ],
            name="% acumulado",
            mode="lines+markers",
            yaxis="y2",
            line=dict(
                color="#F4BE45",
                width=3.0,
            ),
            marker=dict(
                size=5.5,
                color="#FFD166",
                line=dict(
                    color="#07121E",
                    width=.8,
                ),
            ),
            hoverinfo="skip",
        )
    )

    # Línea horizontal del 80%.
    pareto_fig.add_hline(
        y=80,
        line_dash="dash",
        line_width=2,
        line_color="#EF6B73",
        annotation_text="80% DE VENTAS",
        annotation_position="top right",
        yref="y2",
        annotation_font=dict(
            color="#FF8A92",
            size=11,
        ),
    )

    # Línea vertical que marca la última familia necesaria para
    # alcanzar / superar el 80%.
    if (
        pareto_critical_count > 0
        and total_families_pareto > 0
    ):
        cutoff_family_short = str(
            pareto_display_df.iloc[
                -1
            ]["FAMILIA_CORTA"]
        )

        # Plotly add_vline() puede fallar con ejes categóricos
        # en algunas versiones al intentar promediar strings.
        # Usamos add_shape + add_annotation para mantener el
        # mismo marcador vertical de forma compatible.
        pareto_fig.add_shape(
            type="line",
            x0=cutoff_family_short,
            x1=cutoff_family_short,
            y0=0,
            y1=1,
            xref="x",
            yref="paper",
            line=dict(
                color="#FFD166",
                width=1.5,
                dash="dot",
            ),
        )

        pareto_fig.add_annotation(
            x=cutoff_family_short,
            y=1.02,
            xref="x",
            yref="paper",
            text=(
                f"{pareto_critical_count} familias "
                f"({pareto_critical_pct:.1f}%)"
            ),
            showarrow=False,
            xanchor="center",
            yanchor="bottom",
            font=dict(
                color="#FFD166",
                size=10,
            ),
            bgcolor="rgba(14,27,43,.85)",
            bordercolor="rgba(255,209,102,.30)",
            borderwidth=1,
            borderpad=3,
        )

    pareto_fig.update_layout(
        template="plotly_dark",
        height=650,
        paper_bgcolor="#0E1B2B",
        plot_bgcolor="#0E1B2B",
        title=dict(
            text=(
                "<b>PARETO 80/20 DE VENTAS POR FAMILIA</b>"
                "<br>"
                "<span style='font-size:13px;"
                "color:#AFC1D3;'>"
                f"Mostrando {pareto_critical_count} de "
                f"{total_families_pareto} familias: "
                "las necesarias para alcanzar al menos "
                "el 80% de las ventas"
                "</span>"
            ),
            x=0.5,
            xanchor="center",
            font=dict(
                size=18,
                color="#FFFFFF",
            ),
        ),
        margin=dict(
            l=60,
            r=70,
            t=95,
            b=210,
        ),
        bargap=.28,
        hovermode=False,
        dragmode=False,
        xaxis=dict(
            title="Familia",
            tickangle=-38,
            tickfont=dict(
                size=11,
                color="#C3D1DF",
            ),
            showgrid=False,
            automargin=True,
            fixedrange=True,
        ),
        yaxis=dict(
            title="Ventas ($)",
            tickfont=dict(
                size=9,
                color="#A5B7C9",
            ),
            gridcolor=(
                "rgba(65,88,111,.22)"
            ),
            griddash="dot",
            zeroline=False,
            fixedrange=True,
        ),
        yaxis2=dict(
            title="% acumulado",
            overlaying="y",
            side="right",
            range=[0, 105],
            ticksuffix="%",
            tickfont=dict(
                size=9,
                color="#D7AE58",
            ),
            title_font=dict(
                color="#D7AE58",
            ),
            showgrid=False,
            zeroline=False,
            fixedrange=True,
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.015,
            xanchor="right",
            x=1,
            bgcolor="rgba(0,0,0,0)",
            font=dict(
                size=10,
                color="#B8C6D4",
            ),
        ),
    )

    # Gráfica deliberadamente NO interactiva.
    st.plotly_chart(
        pareto_fig,
        use_container_width=True,
        key="family_pareto_chart_static_80_20",
        config={
            "staticPlot": True,
            "displayModeBar": False,
            "responsive": True,
        },
    )


# PESTAÑA 5 — INSIGHTS
# ============================================================

with tab_insights:

    st.markdown("### 💡 Insights")


    i1, i2 = (
        st.columns(2)
    )


    # ============================================================
    # FAMILIA LÍDER
    # ============================================================

    with i1:


        st.subheader(
            "🏆 Familia líder en ventas"
        )


        if not family_sales.empty:


            st.success(

                f"**{top_family_name}** "
                "lidera las ventas con "
                f"**{money(top_family_sales)}**."

            )


        else:


            st.info(
                "No se encontraron registros."
            )


    # ============================================================
    # PRODUCTO LÍDER
    # ============================================================

    with i2:


        st.subheader(
            "🥇 Producto líder"
        )


        if not product_sales.empty:


            st.success(

                f"**{top_product_description}** "
                f"(Código: **{top_product_code}**) "
                "lidera las ventas con "
                f"**{money(top_product_sales)}**."

            )


        else:


            st.info(
                "No se encontraron registros."
            )


    # ============================================================
    # FAMILIA CON MAYOR UTILIDAD
    # ============================================================

    st.subheader(
        "💵 Familia con más utilidad"
    )


    family_profit_insight = (

        filtered_df

        .groupby(

            "FAMILIA",

            dropna=False,

            as_index=False,

        )[
            "$UTILIDAD MES ACTUAL"
        ]

        .sum()

        .sort_values(

            "$UTILIDAD MES ACTUAL",

            ascending=False,

        )

    )


    if not family_profit_insight.empty:


        best_profit_family = str(

            family_profit_insight.iloc[
                0
            ][
                "FAMILIA"
            ]

        )


        best_profit_amount = float(

            family_profit_insight.iloc[
                0
            ][
                "$UTILIDAD MES ACTUAL"
            ]

        )


        st.success(

            f"**{best_profit_family}** "
            "presenta la mayor utilidad con "
            f"**{money(best_profit_amount)}**."

        )


    else:


        st.info(
            "No se encontraron registros."
        )







def executive_alert_table_style(
    dataframe: pd.DataFrame,
    alert_type: str,
):
    """
    Estilo profesional para las tablas de alertas.
    No modifica los datos; únicamente su presentación.
    """

    palettes = {
        "lost": {
            "accent": "#E05D68",
            "accent_soft": "rgba(224,93,104,0.18)",
            "header": "#22344A",
            "row_a": "#162437",
            "row_b": "#1B2B40",
            "text": "#EAF1F8",
            "muted": "#CAD6E2",
        },
        "inventory": {
            "accent": "#E2A84A",
            "accent_soft": "rgba(226,168,74,0.18)",
            "header": "#22344A",
            "row_a": "#162437",
            "row_b": "#1B2B40",
            "text": "#EAF1F8",
            "muted": "#CAD6E2",
        },
        "margin": {
            "accent": "#9D7CE2",
            "accent_soft": "rgba(157,124,226,0.18)",
            "header": "#22344A",
            "row_a": "#162437",
            "row_b": "#1B2B40",
            "text": "#EAF1F8",
            "muted": "#CAD6E2",
        },
    }

    palette = palettes[alert_type]

    styler = (
        dataframe.style
        .set_table_styles(
            [
                {
                    "selector": "thead th",
                    "props": [
                        ("background-color", palette["header"]),
                        ("color", "#FFFFFF"),
                        ("font-weight", "800"),
                        ("font-size", "12px"),
                        ("text-align", "center"),
                        ("border-bottom", f"3px solid {palette['accent']}"),
                    ],
                },
                {
                    "selector": "tbody td",
                    "props": [
                        ("color", palette["text"]),
                        ("font-size", "12px"),
                        ("border-bottom", "1px solid #2C4058"),
                        ("padding", "8px 10px"),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(odd) td",
                    "props": [
                        ("background-color", palette["row_a"]),
                    ],
                },
                {
                    "selector": "tbody tr:nth-child(even) td",
                    "props": [
                        ("background-color", palette["row_b"]),
                    ],
                },
                {
                    "selector": "tbody tr:hover td",
                    "props": [
                        ("background-color", "#263B54"),
                    ],
                },
            ]
        )
        .set_properties(
            **{
                "font-family": "Arial, sans-serif",
            }
        )
    )

    if alert_type == "lost":
        critical_columns = [
            column for column in [
                "#VENTAS PERDIDAS",
                "INVENTARIO TOTAL",
            ]
            if column in dataframe.columns
        ]

    elif alert_type == "inventory":
        critical_columns = [
            column for column in [
                "INVENTARIO TOTAL",
                "COSTO ACTUAL",
            ]
            if column in dataframe.columns
        ]

    else:
        critical_columns = [
            column for column in [
                "% MAGERN S/VENTA",
                "$UTILIDAD MES ACTUAL",
            ]
            if column in dataframe.columns
        ]

    if critical_columns:
        styler = styler.set_properties(
            subset=critical_columns,
            **{
                "background-color": palette["accent_soft"],
                "color": palette["accent"],
                "font-weight": "800",
            }
        )

    if "FAMILIA" in dataframe.columns:
        styler = styler.set_properties(
            subset=["FAMILIA"],
            **{
                "color": "#79C8FF",
                "font-weight": "700",
            }
        )

    if "#COD." in dataframe.columns:
        styler = styler.set_properties(
            subset=["#COD."],
            **{
                "color": "#91E6D0",
                "font-weight": "700",
            }
        )

    return styler


# ============================================================
# PESTAÑA 6 — ALERTAS DE PRODUCTOS
# ============================================================

with tab_alertas:

    st.markdown("### 🚨 Alertas de productos")
    st.caption(
        "Productos que requieren atención por ventas perdidas, "
        "inventario sin movimiento o margen negativo."
    )

    # --------------------------------------------------------
    # Estilo visual de tarjetas ejecutivas
    # --------------------------------------------------------
    st.markdown(
        """
        <style>
        .alert-summary-card {
            background: linear-gradient(145deg, #1D2A3D 0%, #172437 100%);
            border: 1px solid #354A64;
            border-radius: 14px;
            padding: 14px 16px;
            min-height: 112px;
            box-shadow: 0 7px 20px rgba(0,0,0,.20);
            margin-bottom: 12px;
        }

        .alert-summary-title {
            color: #9FB3C8;
            font-size: .82rem;
            font-weight: 700;
            margin-bottom: 8px;
            letter-spacing: .01em;
        }

        .alert-summary-value {
            font-size: .80rem;
            font-weight: 900;
            line-height: 1.05;
            margin-bottom: 5px;
        }

        .alert-summary-note {
            color: #B9C8D8;
            font-size: .78rem;
            line-height: 1.25;
        }

        .alert-red {
            border-top: 4px solid #FF5A6A;
        }

        .alert-red .alert-summary-value {
            color: #FF6D7A;
        }

        .alert-amber {
            border-top: 4px solid #FFB84D;
        }

        .alert-amber .alert-summary-value {
            color: #FFC766;
        }

        .alert-purple {
            border-top: 4px solid #B88CFF;
        }

        .alert-purple .alert-summary-value {
            color: #C7A6FF;
        }

        .executive-table-title {
            background: linear-gradient(90deg, #20344C 0%, #192B40 100%);
            border-left: 5px solid #3FA7D6;
            border-radius: 10px;
            padding: 11px 14px;
            margin-top: 16px;
            margin-bottom: 10px;
            color: #F3F7FB;
            font-weight: 800;
            font-size: 1rem;
        }

        /* Tablas ejecutivas estilo dashboard */
        .exec-table-shell {
            width: 100%;
            overflow: auto;
            border: 1px solid #31455F;
            border-radius: 7px;
            background: #1E2B3E;
            box-shadow: 0 8px 22px rgba(0,0,0,.18);
            margin-bottom: 22px;
        }

        .exec-table {
            width: 100%;
            border-collapse: collapse;
            border-spacing: 0;
            font-family: "Segoe UI", "Inter", "Helvetica Neue", Arial, sans-serif;
            font-size: 0.82rem;
            letter-spacing: 0.01em;
            color: #FFFFFF;
            min-width: 980px;
        }

        .exec-th {
            position: sticky;
            top: 0;
            z-index: 2;
            background: #0E192A;
            color: #00E5FF;
            text-align: left;
            font-weight: 700;
            font-size: 0.78rem;
            letter-spacing: 0.025em;
            padding: 12px 12px;
            border-right: 1px solid #00D8F0;
            border-bottom: 2px solid #354B66;
            white-space: nowrap;
        }

        .exec-th:last-child {
            border-right: none;
        }

        .exec-td {
            background: #202D40;
            color: #FFFFFF;
            padding: 9px 12px;
            border-right: 1px solid #AFC0D2;
            border-bottom: 1px solid #35465D;
            vertical-align: middle;
            font-weight: 400;
            line-height: 1.25;
        }

        .exec-td:last-child {
            border-right: none;
        }

        .exec-tr:nth-child(even) .exec-td {
            background: #1E2A3C;
        }

        .exec-tr:hover .exec-td {
            background: #293A51;
        }

        .exec-text {
            text-align: left;
        }

        .exec-num {
            text-align: right;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }

        /* Variante compacta: solo para "Ver datos originales después de filtros" */
        .exec-table-shell.exec-table-compact {
            overflow-x: auto;
        }

        .exec-table.exec-table-compact {
            min-width: 0;
            width: 100%;
            font-size: 0.68rem;
            letter-spacing: 0;
            table-layout: auto;
        }

        .exec-table.exec-table-compact .exec-th {
            padding: 7px 5px;
            font-size: 0.66rem;
            line-height: 1.08;
            white-space: normal;
            word-break: normal;
            overflow-wrap: anywhere;
            min-width: 78px;
            max-width: 150px;
        }

        .exec-table.exec-table-compact .exec-td {
            padding: 5px 6px;
            font-size: 0.67rem;
            line-height: 1.12;
            min-width: 72px;
            max-width: 165px;
        }

        .exec-table.exec-table-compact .exec-text {
            white-space: normal;
            overflow-wrap: anywhere;
        }

        .exec-table.exec-table-compact .exec-num {
            white-space: nowrap;
            min-width: 78px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ========================================================
    # PREPARAR DATOS
    # ========================================================

    lost_products = product_df[
        product_df["#VENTAS PERDIDAS"]
        .fillna(0)
        .gt(0)
    ].copy()

    if not lost_products.empty:
        lost_products = (
            lost_products
            .sort_values(
                "#VENTAS PERDIDAS",
                ascending=False,
            )
        )

    no_sales_inventory = product_df[
        product_df["#VENTAS MES ACTUAL"]
        .fillna(0)
        .eq(0)
        &
        product_df["INVENTARIO TOTAL"]
        .fillna(0)
        .gt(0)
    ].copy()

    if not no_sales_inventory.empty:
        no_sales_inventory = (
            no_sales_inventory
            .sort_values(
                "INVENTARIO TOTAL",
                ascending=False,
            )
        )

    negative_margin = product_df[
        product_df["% MAGERN S/VENTA"]
        .notna()
        &
        product_df["% MAGERN S/VENTA"]
        .lt(0)
    ].copy()

    if not negative_margin.empty:
        negative_margin = (
            negative_margin
            .sort_values(
                "% MAGERN S/VENTA",
                ascending=True,
            )
        )

    total_lost = (
        safe_sum(lost_products["#VENTAS PERDIDAS"])
        if not lost_products.empty
        else 0
    )

    inventory_without_sales = (
        safe_sum(no_sales_inventory["INVENTARIO TOTAL"])
        if not no_sales_inventory.empty
        else 0
    )

    # ========================================================
    # TARJETAS RESUMEN
    # ========================================================

    a1, a2, a3 = st.columns(3, gap="large")

    with a1:
        st.markdown(
            f"""<div class="alert-summary-card alert-red">
<div class="alert-summary-title">⚠️ PRODUCTOS CON VENTAS PERDIDAS</div>
<div class="alert-summary-value">{len(lost_products):,}</div>
<div class="alert-summary-note">Ventas perdidas acumuladas: {quantity(total_lost, 2)}</div>
</div>""",
            unsafe_allow_html=True,
        )

    with a2:
        st.markdown(
            f"""<div class="alert-summary-card alert-amber">
<div class="alert-summary-title">📦 SIN VENTAS CON INVENTARIO</div>
<div class="alert-summary-value">{len(no_sales_inventory):,}</div>
<div class="alert-summary-note">Inventario acumulado: {quantity(inventory_without_sales, 2)}</div>
</div>""",
            unsafe_allow_html=True,
        )

    with a3:
        st.markdown(
            f"""<div class="alert-summary-card alert-purple">
<div class="alert-summary-title">📉 MARGEN NEGATIVO</div>
<div class="alert-summary-value">{len(negative_margin):,}</div>
<div class="alert-summary-note">Productos que requieren revisión de rentabilidad</div>
</div>""",
            unsafe_allow_html=True,
        )

    # ========================================================
    # TABLA 1 — PRODUCTOS CON VENTAS PERDIDAS
    # ========================================================

    st.markdown(
        '<div class="executive-table-title" style="border-left-color:#E05D68;">'
        f'⚠️ TOP {top_label} productos con ventas perdidas'
        '</div>',
        unsafe_allow_html=True,
    )

    if lost_products.empty:
        st.info("No se encontraron registros.")
    else:
        lost_table = lost_products[
            [
                "FAMILIA",
                "#COD.",
                "DESCRIPCION",
                "#VENTAS PERDIDAS",
                "INVENTARIO TOTAL",
                "$VENTAS MES ACTUAL",
            ]
        ].copy().head(top_data_n)

        render_executive_html_table(
            lost_table,
            column_labels={
                "FAMILIA": "Familia",
                "#COD.": "Código",
                "DESCRIPCION": "Descripción",
                "#VENTAS PERDIDAS": "Ventas Perdidas",
                "INVENTARIO TOTAL": "Inventario Total",
                "$VENTAS MES ACTUAL": "Ventas Mes Actual ($)",
            },
            formats={
                "#VENTAS PERDIDAS": "number2",
                "INVENTARIO TOTAL": "number2",
                "$VENTAS MES ACTUAL": "money",
            },
            max_height=520,
        )

    # ========================================================
    # TABLA 2 — SIN VENTAS PERO CON INVENTARIO
    # ========================================================

    st.markdown(
        '<div class="executive-table-title" style="border-left-color:#E2A84A;">'
        f'📦 TOP {top_label} productos sin ventas en el mes actual pero con inventario'
        '</div>',
        unsafe_allow_html=True,
    )

    if no_sales_inventory.empty:
        st.info("No se encontraron registros.")
    else:
        no_sales_table = no_sales_inventory[
            [
                "FAMILIA",
                "#COD.",
                "DESCRIPCION",
                "#VENTAS MES ACTUAL",
                "$VENTAS MES ACTUAL",
                "INVENTARIO TOTAL",
                "COSTO ACTUAL",
            ]
        ].copy().head(top_data_n)

        render_executive_html_table(
            no_sales_table,
            column_labels={
                "FAMILIA": "Familia",
                "#COD.": "Código",
                "DESCRIPCION": "Descripción",
                "#VENTAS MES ACTUAL": "Unidades Vendidas",
                "$VENTAS MES ACTUAL": "Ventas Mes Actual ($)",
                "INVENTARIO TOTAL": "Inventario Total",
                "COSTO ACTUAL": "Costo Actual ($)",
            },
            formats={
                "#VENTAS MES ACTUAL": "integer",
                "$VENTAS MES ACTUAL": "money",
                "INVENTARIO TOTAL": "number2",
                "COSTO ACTUAL": "money",
            },
            max_height=520,
        )

    # ========================================================
    # TABLA 3 — MARGEN NEGATIVO
    # ========================================================

    st.markdown(
        '<div class="executive-table-title" style="border-left-color:#9D7CE2;">'
        f'📉 TOP {top_label} productos con margen negativo'
        '</div>',
        unsafe_allow_html=True,
    )

    if negative_margin.empty:
        st.info("No se encontraron registros.")
    else:
        negative_margin_table = negative_margin[
            [
                "FAMILIA",
                "#COD.",
                "DESCRIPCION",
                "% MAGERN S/VENTA",
                "$VENTAS MES ACTUAL",
                "$UTILIDAD MES ACTUAL",
                "COSTO ACTUAL",
                "PRECIO PROMEDIO",
            ]
        ].copy().head(top_data_n)

        render_executive_html_table(
            negative_margin_table,
            column_labels={
                "FAMILIA": "Familia",
                "#COD.": "Código",
                "DESCRIPCION": "Descripción",
                "% MAGERN S/VENTA": "Margen (%)",
                "$VENTAS MES ACTUAL": "Ventas Mes Actual ($)",
                "$UTILIDAD MES ACTUAL": "Utilidad Mes Actual ($)",
                "COSTO ACTUAL": "Costo Actual ($)",
                "PRECIO PROMEDIO": "Precio Promedio ($)",
            },
            formats={
                "% MAGERN S/VENTA": "percent",
                "$VENTAS MES ACTUAL": "money",
                "$UTILIDAD MES ACTUAL": "money",
                "COSTO ACTUAL": "money",
                "PRECIO PROMEDIO": "money",
            },
            max_height=520,
        )


# ============================================================
# VER DATOS FILTRADOS
# ============================================================
# VER DATOS ORIGINALES DESPUÉS DE FILTROS
# ============================================================

st.markdown("---")

with st.expander(
    f"🔎 Ver TOP {top_label} datos originales después de filtros",
    expanded=False,
):

    filtered_original_sorted = filtered_df.copy()

    if "$VENTAS MES ACTUAL" in filtered_original_sorted.columns:
        filtered_original_sorted = (
            filtered_original_sorted
            .sort_values(
                "$VENTAS MES ACTUAL",
                ascending=False,
                na_position="last",
            )
            .reset_index(drop=True)
            .head(top_data_n)
        )

    original_column_labels = {
        "NUMERO DE FAMILIA": "Número de Familia",
        "FAMILIA": "Familia",
        "#COD.": "Código",
        "DESCRIPCION": "Descripción",
        "EMPAQ": "Empaq.",
        "#COMPRAS MES ACTUAL": "Compras Mes Actual",
        "FECHA ULT COMPRA": "Fecha Últ. Compra",
        "COSTO DE BODEGA": "Costo de Bodega ($)",
        "COSTO GENERACION": "Costo Generación ($)",
        "COSTO ACTUAL": "Costo Actual ($)",
        "INVENTARIO TOTAL": "Inventario Total",
        "#VENTAS AÑO ACTUAL": "Ventas Año Actual",
        "#VENTAS MES ANTERIOR": "Ventas Mes Anterior",
        "#VENTAS MES ACTUAL": "Ventas Mes Actual",
        "#VENTAS PERDIDAS": "Ventas Perdidas",
        "FECHA ULT VENTAS": "Fecha Últ. Venta",
        "$VENTAS MES ACTUAL": "Ventas Mes Actual ($)",
        "PRECIO PROMEDIO": "Precio Promedio ($)",
        "% MAGERN S/VENTA": "Margen (%)",
        "$UTILIDAD MES ACTUAL": "Utilidad Mes Actual ($)",
        "$GENERA S/VENTAS": "Genera S/Ventas ($)",
    }

    original_formats = {
        "NUMERO DE FAMILIA": "integer",
        "#COMPRAS MES ACTUAL": "number2",
        "COSTO DE BODEGA": "money",
        "COSTO GENERACION": "money",
        "COSTO ACTUAL": "money",
        "INVENTARIO TOTAL": "number2",
        "#VENTAS AÑO ACTUAL": "number2",
        "#VENTAS MES ANTERIOR": "number2",
        "#VENTAS MES ACTUAL": "number2",
        "#VENTAS PERDIDAS": "number2",
        "$VENTAS MES ACTUAL": "money",
        "PRECIO PROMEDIO": "money",
        "% MAGERN S/VENTA": "percent",
        "$UTILIDAD MES ACTUAL": "money",
        "$GENERA S/VENTAS": "money",
    }

    render_executive_html_table(
        filtered_original_sorted,
        column_labels={
            column: original_column_labels.get(column, column)
            for column in filtered_original_sorted.columns
        },
        formats={
            column: original_formats[column]
            for column in filtered_original_sorted.columns
            if column in original_formats
        },
        max_height=650,
        compact=True,
    )



# ============================================================
# PESTAÑA 7 — DESCARGAR INFORMACIÓN
# ============================================================

with tab_descargas:

    st.markdown(
        "### ⬇️ Descargar información"
    )


    st.caption(

        "Descargue los datos filtrados, "
        "el resumen por familia y el detalle "
        "agrupado por producto en Excel o CSV."

    )


    filtered_export = (
        filtered_df.copy()
    )


    family_export = (
        family_df.copy()
    )


    product_export = (
        product_df.copy()
    )


    # ============================================================
    # DATOS FILTRADOS
    # ============================================================

    st.subheader(
        "Datos filtrados"
    )


    d1, d2 = (
        st.columns(2)
    )


    with d1:


        st.download_button(

            "📗 Descargar datos filtrados en Excel",

            data=
            dataframe_to_excel_bytes(

                filtered_export,

                "Datos Filtrados",

            ),

            file_name=
            "CIERRE_PRODUCTOS_FILTRADOS.xlsx",

            mime=(

                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"

            ),

            use_container_width=True,

        )


    with d2:


        st.download_button(

            "📄 Descargar datos filtrados en CSV",

            data=
            dataframe_to_csv_bytes(

                filtered_export

            ),

            file_name=
            "CIERRE_PRODUCTOS_FILTRADOS.csv",

            mime=
            "text/csv",

            use_container_width=True,

        )


    # ============================================================
    # RESUMEN POR FAMILIA
    # ============================================================

    st.subheader(
        "Resumen por familia"
    )


    d3, d4 = (
        st.columns(2)
    )


    with d3:


        st.download_button(

            "📗 Descargar resumen por familia en Excel",

            data=
            dataframe_to_excel_bytes(

                family_export,

                "Resumen Familia",

            ),

            file_name=
            "RESUMEN_POR_FAMILIA.xlsx",

            mime=(

                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"

            ),

            use_container_width=True,

        )


    with d4:


        st.download_button(

            "📄 Descargar resumen por familia en CSV",

            data=
            dataframe_to_csv_bytes(

                family_export

            ),

            file_name=
            "RESUMEN_POR_FAMILIA.csv",

            mime=
            "text/csv",

            use_container_width=True,

        )


    # ============================================================
    # DETALLE POR PRODUCTO
    # ============================================================

    st.subheader(
        "Detalle por producto"
    )


    d5, d6 = (
        st.columns(2)
    )


    with d5:


        st.download_button(

            "📗 Descargar detalle por producto en Excel",

            data=
            dataframe_to_excel_bytes(

                product_export,

                "Detalle Productos",

            ),

            file_name=
            "DETALLE_POR_PRODUCTO.xlsx",

            mime=(

                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"

            ),

            use_container_width=True,

        )


    with d6:


        st.download_button(

            "📄 Descargar detalle por producto en CSV",

            data=
            dataframe_to_csv_bytes(

                product_export

            ),

            file_name=
            "DETALLE_POR_PRODUCTO.csv",

            mime=
            "text/csv",

            use_container_width=True,

        )





# ============================================================
# OVERRIDES FINALES — PREMIUM
# ============================================================

st.markdown(
    """
    <style>
    .alert-summary-card {
        background:linear-gradient(145deg,#132237,#0D1A29) !important;
        border:1px solid #294158 !important;
        border-radius:13px !important;
        box-shadow:0 8px 24px rgba(0,0,0,.15) !important;
    }

    .executive-table-title {
        background:linear-gradient(90deg,#14253A,#0E1B2B) !important;
        border-left:4px solid #D7AE58 !important;
        color:#F0F4F8 !important;
        border-radius:9px !important;
        font-size:.82rem !important;
        font-weight:760 !important;
    }

    .exec-table-shell {
        border:1px solid #263E57 !important;
        background:#0D1A2A !important;
        border-radius:11px !important;
    }

    .exec-th {
        background:#101F31 !important;
        color:#E4C879 !important;
        border-bottom:2px solid rgba(215,174,88,.68) !important;
    }

    .exec-td {
        background:#102033 !important;
        color:#DCE5EE !important;
        border-right:1px solid #263B51 !important;
        border-bottom:1px solid #22364C !important;
    }

    .exec-tr:nth-child(even) .exec-td {
        background:#0E1C2C !important;
    }

    .exec-tr:hover .exec-td {
        background:#182B41 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# AJUSTE FINAL — LEYENDA INTERACTIVA COMPACTA DE FAMILIAS
# ============================================================

st.markdown(
    """
    <style>
    /* Leyenda interactiva de la dona */
    div[data-testid="stRadio"] {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }

    div[data-testid="stRadio"] > div {
        gap: 0 !important;
        row-gap: 0 !important;
    }

    div[data-testid="stRadio"] label {
        min-height: 15px !important;
        padding: 0 !important;
        margin: 0 !important;
        gap: 1px !important;
        align-items: center !important;
    }

    div[data-testid="stRadio"] label p {
        color: #AFC0D1 !important;
        font-family: "Inter","Segoe UI",Arial,sans-serif !important;
        font-size: .54rem !important;
        line-height: 1.02 !important;
        margin: 0 !important;
        padding: 0 !important;
    }

    div[data-testid="stRadio"] label > div:first-child {
        transform: scale(.54) !important;
        transform-origin: left center !important;
        margin-right: -8px !important;
    }

    div[data-testid="stRadio"] label:hover p {
        color: #F0D487 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# ESTILO LEYENDA EJECUTIVA DONA
# ============================================================

st.markdown(
    """
    <style>
    div[data-testid="stButton"] button[kind="secondary"] {
        white-space: normal !important;
    }

    div[data-testid="stButton"] button {
        line-height: 1.10 !important;
    }

    div[data-testid="stButton"] button p {
        white-space: pre-line !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# LEYENDA LATERAL PROFESIONAL — DONA
# ============================================================

st.markdown(
    """
    <style>

    .donut-legend-heading {
        padding: 4px 2px 7px 2px;
        border-bottom: 1px solid #22384F;
        margin-bottom: 5px;
    }

    .donut-legend-title {
        color: #D7AE58;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .68rem;
        font-weight: 850;
        letter-spacing: .045em;
    }

    .donut-legend-subtitle {
        color: #657C92;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .52rem;
        margin-top: 2px;
    }

    /*
    Leyenda radio:
    - compacta
    - nombres completos
    - una fila visual limpia por familia
    */

    div[data-testid="stRadio"] {
        margin: 0 !important;
        padding: 0 !important;
    }

    div[data-testid="stRadio"] > div {
        gap: 0 !important;
        row-gap: 0 !important;
    }

    div[data-testid="stRadio"] label {
        min-height: 20px !important;
        margin: 0 !important;
        padding: 1px 2px !important;
        border-radius: 6px !important;
        align-items: center !important;
    }

    div[data-testid="stRadio"] label:hover {
        background: #14253A !important;
    }

    div[data-testid="stRadio"] label p {
        color: #B8C6D4 !important;
        font-family: "Inter","Segoe UI",Arial,sans-serif !important;
        font-size: .59rem !important;
        line-height: 1.05 !important;
        margin: 0 !important;
        padding: 0 !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
    }

    div[data-testid="stRadio"] label:hover p {
        color: #F0D487 !important;
    }

    div[data-testid="stRadio"] label > div:first-child {
        transform: scale(.62) !important;
        transform-origin: center !important;
        margin-right: -5px !important;
    }

    .donut-family-summary {
        margin-top: 9px;
        padding: 10px 11px;
        border: 1px solid #28425B;
        border-left: 3px solid #27C99B;
        border-radius: 10px;
        background:
            linear-gradient(
                145deg,
                #132338,
                #0D1B2B
            );
    }

    .donut-family-summary-label {
        color: #6F8399;
        font-size: .50rem;
        font-weight: 800;
        letter-spacing: .05em;
        margin-bottom: 4px;
    }

    .donut-family-summary-name {
        color: #F1D17C;
        font-size: .70rem;
        font-weight: 820;
        line-height: 1.12;
        margin-bottom: 8px;
    }

    .donut-family-summary-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 5px 7px;
    }

    .donut-family-summary-grid div {
        min-width: 0;
        padding: 5px 6px;
        border-radius: 7px;
        background: #102034;
        border: 1px solid #213950;
    }

    .donut-family-summary-grid span {
        display: block;
        color: #6F8398;
        font-size: .46rem;
        margin-bottom: 2px;
    }

    .donut-family-summary-grid b {
        display: block;
        color: #DCE7F1;
        font-size: .57rem;
        line-height: 1.05;
        overflow-wrap: anywhere;
    }

    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# ESTILO — PESTAÑA VENTAS Y UTILIDAD POR FAMILIA
# ============================================================

st.markdown(
    """
    <style>
    .family-analysis-legend-head {
        padding: 4px 2px 7px 2px;
        margin-bottom: 4px;
        border-bottom: 1px solid #22384F;
    }

    .family-analysis-legend-title {
        color: #D7AE58;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .66rem;
        font-weight: 850;
        letter-spacing: .045em;
    }

    .family-analysis-legend-note {
        color: #667C91;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .50rem;
        margin-top: 2px;
    }

    div[data-testid="stRadio"] label p {
        overflow-wrap: anywhere !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# ESTILO EJECUTIVO — DONA FAMILIAS
# ============================================================

st.markdown(
    """
    <style>

    .exec-family-legend-header {
        padding: 5px 2px 8px 2px;
        margin-bottom: 4px;
        border-bottom: 1px solid #263E56;
    }

    .exec-family-legend-title {
        color: #D7AE58;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .70rem;
        font-weight: 850;
        letter-spacing: .045em;
    }

    .exec-family-legend-note {
        color: #6D8297;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .49rem;
        margin-top: 2px;
    }

    /*
    Leyenda interactiva:
    más ejecutiva, compacta y legible.
    */
    div[data-testid="stRadio"] {
        padding: 0 !important;
        margin: 0 !important;
    }

    div[data-testid="stRadio"] > div {
        gap: 0 !important;
        row-gap: 0 !important;
    }

    div[data-testid="stRadio"] label {
        min-height: 20px !important;
        padding: 1px 4px !important;
        margin: 0 !important;
        border-radius: 6px !important;
        align-items: center !important;
    }

    div[data-testid="stRadio"] label:hover {
        background: #14263A !important;
    }

    div[data-testid="stRadio"] label p {
        color: #B9C7D5 !important;
        font-family: "Inter","Segoe UI",Arial,sans-serif !important;
        font-size: .57rem !important;
        line-height: 1.03 !important;
        margin: 0 !important;
        padding: 0 !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
    }

    div[data-testid="stRadio"] label:hover p {
        color: #F0D487 !important;
    }

    div[data-testid="stRadio"] label > div:first-child {
        transform: scale(.60) !important;
        transform-origin: center !important;
        margin-right: -5px !important;
    }

    .exec-family-detail {
        margin-top: 8px;
        padding: 10px 11px;
        border: 1px solid #29445E;
        border-left: 3px solid #27C99B;
        border-radius: 10px;
        background:
            linear-gradient(
                145deg,
                #132338,
                #0D1B2B
            );
    }

    .exec-family-detail-eyebrow {
        color: #6F8499;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .47rem;
        font-weight: 820;
        letter-spacing: .055em;
    }

    .exec-family-detail-name {
        color: #F0D487;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .73rem;
        line-height: 1.12;
        font-weight: 840;
        margin: 4px 0 7px 0;
    }

    .exec-family-share-row {
        display: flex;
        justify-content: space-between;
        gap: 7px;
        align-items: center;
        border-top: 1px solid #263D53;
        border-bottom: 1px solid #263D53;
        padding: 5px 0;
        margin-bottom: 7px;
    }

    .exec-family-share-row span {
        color: #74899E;
        font-size: .47rem;
    }

    .exec-family-share-row b {
        color: #43B9E6;
        font-size: .61rem;
    }

    .exec-family-detail-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 5px;
    }

    .exec-family-detail-grid div {
        padding: 5px 6px;
        background: #102034;
        border: 1px solid #213950;
        border-radius: 7px;
        min-width: 0;
    }

    .exec-family-detail-grid span {
        display: block;
        color: #6F8498;
        font-size: .66rem;
        font-weight: 650;
        margin-bottom: 4px;
    }

    .exec-family-detail-grid b {
        display: block;
        color: #DCE7F1;
        font-size: .56rem;
        line-height: 1.05;
        overflow-wrap: anywhere;
    }

    .exec-family-empty {
        margin-top: 10px;
        padding: 13px 11px;
        text-align: center;
        color: #71869A;
        font-family: "Inter","Segoe UI",Arial,sans-serif;
        font-size: .54rem;
        line-height: 1.35;
        border: 1px dashed #294158;
        border-radius: 10px;
        background: #0D1A2A;
    }

    .exec-family-empty-icon {
        color: #D7AE58;
        font-size: 1rem;
        margin-bottom: 4px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# FICHA EJECUTIVA — FAMILIA
# ============================================================

st.markdown(
    """
    <style>
    .family-meeting-card {
        margin-top: 10px;
        padding: 18px 18px;
        border-radius: 10px;
        border: 1px solid #29445E;
        border-left: 4px solid #D7AE58;
        background:
            linear-gradient(
                145deg,
                #132338,
                #0D1A2A
            );
        box-shadow:
            0 9px 24px rgba(0,0,0,.13);
    }

    .family-meeting-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 18px;
        padding-bottom: 13px;
        margin-bottom: 13px;
        border-bottom: 1px solid #263E55;
    }

    .family-meeting-eyebrow {
        color: #73889D;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .86rem;
        font-weight: 820;
        letter-spacing: .06em;
    }

    .family-meeting-name {
        color: #F0D487;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: 1.38rem;
        font-weight: 850;
        line-height: 1.18;
        margin-top: 3px;
    }

    .family-meeting-share {
        color: #43B9E6;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: 1.72rem;
        font-weight: 880;
        text-align: right;
        white-space: nowrap;
    }

    .family-meeting-share span {
        display: block;
        color: #71869B;
        font-size: .76rem;
        font-weight: 650;
        margin-top: 1px;
    }

    .family-meeting-grid {
        display: grid;
        grid-template-columns:
            repeat(3, minmax(0,1fr));
        gap: 9px;
    }

    .family-meeting-grid div {
        min-width: 0;
        padding: 11px 12px;
        border-radius: 8px;
        background: #102034;
        border: 1px solid #213950;
    }

    .family-meeting-grid span {
        display: block;
        color: #71869A;
        font-size: .66rem;
        margin-bottom: 2px;
    }

    .family-meeting-grid b {
        display: block;
        color: #E3EBF3;
        font-size: 1.02rem;
        font-weight: 780;
        line-height: 1.15;
        overflow-wrap: anywhere;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# TABLA DE PRODUCTOS — FAMILIA SELECCIONADA
# ============================================================

st.markdown(
    """
    <style>
    .executive-table-title {
        margin-top: 12px !important;
        margin-bottom: 5px !important;
        padding: 9px 12px !important;
        background:
            linear-gradient(
                90deg,
                #14253A,
                #0E1B2B
            ) !important;
        border: 1px solid #294158 !important;
        border-left: 4px solid #D7AE58 !important;
        border-radius: 9px !important;
        color: #F1F5F9 !important;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif !important;
        font-size: .66rem !important;
        font-weight: 780 !important;
        letter-spacing: .01em !important;
    }

    .exec-table-shell.exec-table-compact {
        border-color: #294158 !important;
        background: #0C1928 !important;
    }

    .exec-table-compact .exec-th {
        font-size: .62rem !important;
        padding: 8px 9px !important;
        color: #E7CB7B !important;
        background: #101F31 !important;
    }

    .exec-table-compact .exec-td {
        font-size: .61rem !important;
        padding: 7px 9px !important;
        line-height: 1.15 !important;
    }

    .exec-table-compact .exec-tr:hover .exec-td {
        background: #193149 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# BOTÓN CENTRAL ENTRE PARTICIPACIÓN Y UTILIDAD
# ============================================================

st.markdown(
    """
    <style>
    /*
    El botón central se mantiene compacto para no competir
    visualmente con las dos gráficas.
    */
    div[data-testid="stButton"] button[title="Borrar filtro de familia"] {
        min-width: 42px !important;
        min-height: 42px !important;
        padding: 0 !important;
        border-radius: 50% !important;
        border: 1px solid #3A526B !important;
        background: linear-gradient(145deg,#17283C,#102033) !important;
        color: #F0D487 !important;
        font-size: 1rem !important;
        box-shadow: 0 7px 18px rgba(0,0,0,.18) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# KEY METRICS — FAMILIA SELECCIONADA
# ============================================================

st.markdown(
    """
    <style>
    .family-kpi-section-title {
        margin-top: 8px;
        margin-bottom: 7px;
        color: #D7AE58;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .68rem;
        font-weight: 850;
        letter-spacing: .06em;
    }

    .family-kpi-card {
        position: relative;
        overflow: hidden;
        min-height: 96px;
        padding: 12px 13px 10px 13px;
        border-radius: 11px;
        border: 1px solid #29445E;
        background:
            linear-gradient(
                145deg,
                #14253A,
                #0E1B2B
            );
        box-shadow:
            0 8px 20px rgba(0,0,0,.14);
    }

    .family-kpi-card::before {
        content: "";
        position: absolute;
        left: 11px;
        right: 11px;
        top: 0;
        height: 2px;
        border-radius: 0 0 4px 4px;
        background: var(--kpi-accent);
    }

    .family-kpi-label {
        color: #8194A9;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .50rem;
        font-weight: 820;
        letter-spacing: .055em;
        margin-bottom: 7px;
    }

    .family-kpi-value {
        color: var(--kpi-accent);
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .80rem;
        font-weight: 880;
        line-height: 1.05;
        letter-spacing: -.02em;
    }

    .family-kpi-unit {
        color: #71869A;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .48rem;
        margin-top: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# KEY METRICS — FAMILIA SELECCIONADA
# ============================================================

st.markdown(
    """
    <style>
    .family-selected-kpi-head {
        margin-top: 10px;
        margin-bottom: 7px;
        color: #D7AE58;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .68rem;
        font-weight: 850;
        letter-spacing: .055em;
        text-transform: uppercase;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# KPI COMPACTOS — VENTAS Y UTILIDAD POR FAMILIA
# ============================================================

st.markdown(
    """
    <style>
    .family-compact-kpi {
        position: relative;
        overflow: hidden;
        min-height: 88px;
        height: 88px;
        padding: 12px 15px 10px 15px;
        border-radius: 11px;
        border: 1px solid #29445E;
        background:
            linear-gradient(
                145deg,
                #14253A,
                #0E1B2B
            );
        box-shadow:
            0 7px 18px rgba(0,0,0,.13);
        box-sizing: border-box;
    }

    .family-compact-kpi::before {
        content: "";
        position: absolute;
        left: 12px;
        right: 12px;
        top: 0;
        height: 2px;
        background: var(--family-kpi-accent);
        border-radius: 0 0 4px 4px;
    }

    .family-compact-kpi-label {
        color: #B7C4D1;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .56rem;
        font-weight: 820;
        letter-spacing: .045em;
        margin-bottom: 7px;
    }

    .family-compact-kpi-value {
        color: var(--family-kpi-accent);
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .98rem;
        font-weight: 880;
        line-height: 1.02;
        letter-spacing: -.02em;
    }

    .family-compact-kpi-note {
        color: #71869A;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .46rem;
        margin-top: 5px;
    }

    .family-selected-kpi-head {
        margin-top: 8px !important;
        margin-bottom: 5px !important;
        font-size: .62rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# TABLA AMPLIA — PRODUCTOS DE FAMILIA
# ============================================================

st.markdown(
    """
    <style>
    .family-products-wide-title {
        display: flex;
        align-items: flex-end;
        justify-content: space-between;
        gap: 18px;
        padding: 11px 14px;
        margin: 4px 0 8px 0;
        border: 1px solid #29445E;
        border-left: 4px solid #D7AE58;
        border-radius: 10px;
        background:
            linear-gradient(
                90deg,
                #14253A,
                #0D1B2B
            );
    }

    .family-products-wide-title strong {
        display: block;
        color: #F2F6FA;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .90rem;
        line-height: 1.15;
        margin-top: 3px;
    }

    .family-products-wide-eyebrow {
        display: block;
        color: #D7AE58;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .52rem;
        font-weight: 850;
        letter-spacing: .07em;
    }

    .family-products-wide-count {
        color: #7F94A9;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        font-size: .56rem;
        white-space: nowrap;
    }

    .family-products-wide-shell {
        width: 100%;
        max-height: 680px;
        overflow: auto;
        border: 1px solid #294158;
        border-radius: 10px;
        background: #0C1928;
        box-shadow:
            0 10px 28px rgba(0,0,0,.15);
    }

    .family-products-wide-table {
        width: 100%;
        min-width: 1160px;
        table-layout: fixed;
        border-collapse: collapse;
        font-family:
            "Inter","Segoe UI",Arial,sans-serif;
        color: #E2EAF2;
    }

    .family-products-wide-table thead th {
        position: sticky;
        top: 0;
        z-index: 3;
        padding: 11px 12px;
        background: #102033;
        color: #E7C978;
        border-right: 1px solid #294057;
        border-bottom: 2px solid rgba(215,174,88,.72);
        font-size: .68rem;
        font-weight: 820;
        line-height: 1.12;
        text-align: left;
        vertical-align: middle;
    }

    .family-products-wide-table tbody td {
        padding: 10px 12px;
        background: #102033;
        color: #DCE5EE;
        border-right: 1px solid #263B51;
        border-bottom: 1px solid #22364C;
        font-size: .68rem;
        line-height: 1.22;
        vertical-align: middle;
    }

    .family-products-wide-table tbody tr:nth-child(even) td {
        background: #0E1C2C;
    }

    .family-products-wide-table tbody tr:hover td {
        background: #193149;
    }

    .family-products-wide-table .product-cell {
        color: #F1F5F9;
        font-weight: 650;
        white-space: normal;
        overflow-wrap: break-word;
    }

    .family-products-wide-table .code-cell {
        color: #AFC1D2;
        font-weight: 760;
        white-space: nowrap;
    }

    .family-products-wide-table .num-cell {
        text-align: right;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
    }

    .family-products-wide-table .margin-cell {
        color: #72DDBA;
        font-weight: 760;
    }

    .family-products-wide-table th:last-child,
    .family-products-wide-table td:last-child {
        border-right: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# IDENTIDAD CORPORATIVA — DICASA S.A.
# ============================================================

st.markdown(
    """
    <style>
    .dash-company {
        display: inline-block;
        margin-bottom: 5px;
        color: #D7AE58;
        font-family:
            "Inter",
            "Segoe UI Variable",
            "Segoe UI",
            Arial,
            sans-serif;
        font-size: .72rem;
        line-height: 1;
        font-weight: 880;
        letter-spacing: .16em;
        text-transform: uppercase;
        text-shadow:
            0 1px 10px rgba(215,174,88,.12);
    }

    .dash-title {
        font-size: 1.62rem !important;
        font-weight: 850 !important;
        letter-spacing: -.035em !important;
        color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# HEADER CORPORATIVO DICASA S.A.
# ============================================================

st.markdown(
    """
    <style>
    .dicasa-corporate-header {
        min-height: 108px !important;
        padding: 14px 20px !important;
        border: 1px solid #233C55 !important;
        background:
            radial-gradient(
                circle at 13% 50%,
                rgba(57,112,166,.10),
                transparent 26%
            ),
            linear-gradient(
                118deg,
                #12243A 0%,
                #0B1829 58%,
                #08121F 100%
            ) !important;
        box-shadow:
            0 15px 36px rgba(0,0,0,.22),
            inset 0 1px 0 rgba(255,255,255,.03) !important;
    }

    .dicasa-corporate-header {
        display: grid !important;
        grid-template-columns: minmax(0, 1fr) auto minmax(0, 1fr) !important;
        align-items: center !important;
        column-gap: 14px !important;
    }

    .dicasa-header-logo-zone {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        min-width: 0;
    }

    .dicasa-header-copy {
        min-width: 0;
        text-align: center !important;
        justify-self: center;
        width: auto !important;
    }

    .dicasa-header-badge-zone {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        min-width: 0;
    }

    .dicasa-logo-frame {
        width: 84px;
        height: 84px;
        flex: 0 0 84px;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow: hidden;
        border-radius: 15px;
        background: #F7F8FA;
        border: 1px solid rgba(126,166,205,.55);
        box-shadow:
            0 8px 22px rgba(0,0,0,.28),
            0 0 0 3px rgba(45,92,137,.10);
    }

    .dicasa-logo-image {
        width: 100%;
        height: 100%;
        object-fit: cover;
        display: block;
    }

    .dicasa-title-line {
        display: flex;
        align-items: baseline;
        justify-content: center !important;
        flex-wrap: wrap;
        gap: 9px;
        line-height: 1.08;
        text-align: center !important;
    }

    .dicasa-dashboard-title {
        color: #FFFFFF;
        font-family:
            "Aptos Display",
            "Segoe UI Variable Display",
            "Inter",
            "Segoe UI",
            Arial,
            sans-serif;
        font-size: 1.72rem;
        font-weight: 820;
        letter-spacing: -.028em;
        line-height: 1.02;
        text-shadow:
            0 1px 14px rgba(255,255,255,.04);
    }

    .dicasa-title-separator {
        color: #7B8FA4;
        font-size: 1.38rem;
        font-weight: 450;
        margin: 0 2px;
    }

    .dicasa-company-name {
        color: #D9B45F;
        font-family:
            "Aptos Display",
            "Segoe UI Variable Display",
            "Inter",
            "Segoe UI",
            Arial,
            sans-serif;
        font-size: 1.50rem;
        font-weight: 880;
        letter-spacing: .065em;
        line-height: 1.02;
        text-transform: uppercase;
        text-shadow:
            0 1px 14px rgba(217,180,95,.15);
    }

    .dicasa-subtitle {
        margin-top: 8px !important;
        color: #FFFFFF !important;
        font-family:
            "Aptos",
            "Segoe UI Variable",
            "Inter",
            "Segoe UI",
            Arial,
            sans-serif !important;
        font-size: .70rem !important;
        font-weight: 500 !important;
        letter-spacing: .01em;
    }

    @media (max-width: 900px) {
        .dicasa-corporate-header {
            grid-template-columns: minmax(0, 1fr) auto minmax(0, 1fr) !important;
            column-gap: 9px !important;
            align-items: center !important;
        }

        .dicasa-logo-frame {
            width: 68px;
            height: 68px;
            flex-basis: 68px;
        }

        .dicasa-dashboard-title {
            font-size: 1.34rem;
        }

        .dicasa-company-name {
            font-size: 1.12rem;
        }

        .dash-badge {
            font-size: .58rem !important;
            padding: 7px 10px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# KPI PRINCIPALES — TAMAÑO ESTABLE, SIN REDUCCIÓN POSTERIOR
# ============================================================

st.markdown(
    """
    <style>
    /* Este bloque se carga al final y FIJA los tamaños definitivos.
       Evita que estilos posteriores o media queries reduzcan los KPI. */
    .exec-kpi,
    .leader-card {
        height: 158px !important;
        min-height: 158px !important;
        padding: 18px 18px 15px !important;
    }

    .exec-kpi-label,
    .leader-label {
        font-size: 16px !important;
        line-height: 1.18 !important;
        letter-spacing: .025em !important;
        margin-bottom: 8px !important;
        overflow: visible !important;
        white-space: normal !important;
        word-break: normal !important;
    }

    .exec-kpi-value {
        font-size: 20px !important;
        line-height: 1.06 !important;
        letter-spacing: -.015em !important;
        white-space: nowrap !important;
        overflow: visible !important;
        text-overflow: clip !important;
        word-break: keep-all !important;
    }

    .leader-name {
        font-size: 18px !important;
        line-height: 1.18 !important;
        margin-bottom: 8px !important;
        display: -webkit-box !important;
        -webkit-line-clamp: 3 !important;
        -webkit-box-orient: vertical !important;
        overflow: hidden !important;
        word-break: normal !important;
        color:#E8C96C !important;
        font-weight:840 !important;
    }

    .leader-code {
        font-size: 15px !important;
        line-height: 1.20 !important;
        margin-bottom: 5px !important;
        white-space: normal !important;
        color:#C9DCF4 !important;
        font-weight:760 !important;
    }

    .leader-value {
        font-size: 20px !important;
        line-height: 1.18 !important;
        white-space: nowrap !important;
    }

    .exec-kpi-note {
        font-size: 15px !important;
        line-height: 1.18 !important;
        padding-top: 8px !important;
        color:#BFD7F2 !important;
        font-weight:760 !important;
    }

    /* Mantener el pequeño espacio horizontal sin comprimir tipografía */
    div[data-testid="stHorizontalBlock"] {
        gap: .55rem !important;
    }

    div[data-testid="stColumn"] > div {
        min-width: 0 !important;
    }

    /* IMPORTANTE:
       no reducir la tipografía KPI en pantallas <=1500px. */
    @media (max-width: 1500px) {
        .exec-kpi-label,
        .leader-label {
            font-size: 16px !important;
        }

        .exec-kpi-value {
            font-size: 20px !important;
        }

        .leader-name {
            font-size: 18px !important;
        }

        .leader-code {
            font-size: 15px !important;
        }

        .leader-value {
            font-size: 20px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# AJUSTE DE ESPACIADO — CONTENIDO MÁS CERCA DEL ENCABEZADO
# ============================================================

st.markdown(
    """
    <style>
    /* Reducir espacio superior general del contenido */
    .block-container {
        padding-top: .35rem !important;
    }

    /* Acercar el contenido inmediatamente posterior al header */
    .dash-header {
        margin-bottom: 5px !important;
    }

    /* Reducir espacios innecesarios entre bloques principales */
    div[data-testid="stVerticalBlock"] {
        gap: .45rem !important;
    }

    /* Menos separación en separadores horizontales */
    hr {
        margin-top: .45rem !important;
        margin-bottom: .45rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# ============================================================
# SEPARACIÓN SUAVE — VENTAS Y UTILIDAD POR FAMILIA
# ============================================================

st.markdown(
    """
    <style>
    /* Mantener el dashboard compacto, pero con aire visual */
    .family-selected-kpi-head {
        margin-top: 10px !important;
        margin-bottom: 5px !important;
    }

    .family-compact-kpi {
        margin-bottom: 2px !important;
    }
    

/* ==========================================================
   TÍTULOS KPI — ANÁLISIS DE PRODUCTOS
   Regla específica para evitar que estilos globales los reduzcan
   ========================================================== */
.product-analysis-kpi-title {
    color: #79D8FF !important;
    font-size: .70rem !important;
    line-height: 1.18 !important;
    font-weight: 900 !important;
    letter-spacing: .025em !important;
    margin-bottom: 12px !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
    word-break: normal !important;
}
</style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.markdown("---")


st.caption(

    "Dashboard de Cierre de Productos · "
    "Todos los indicadores se calculan "
    "exclusivamente con los datos del archivo Excel cargado."

)
