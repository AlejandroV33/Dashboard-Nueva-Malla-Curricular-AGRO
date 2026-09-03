
from __future__ import annotations

import io
import json
import math
import re
import unicodedata
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from rapidfuzz import fuzz, process
    RAPIDFUZZ_AVAILABLE = True
except ImportError:
    from difflib import SequenceMatcher
    RAPIDFUZZ_AVAILABLE = False

try:
    from streamlit_searchbox import st_searchbox
    SEARCHBOX_AVAILABLE = True
except ImportError:
    SEARCHBOX_AVAILABLE = False


# ============================================================
# AGRO — Constructor interactivo de nueva malla curricular
# Modelo simbólico relacional sobre 5 libros Excel auditados.
#
# PRINCIPIOS:
# 1) Base Maestra v1.2 = catálogo curricular canónico.
# 2) EPN Dataset = baseline operativo EPN + carga/prerrequisitos.
# 3) Matriz Comparativa = benchmark internacional.
# 4) Matriz Brechas = diagnóstico por familia AGRO-NORM.
# 5) CAEE = evidencia externa / EUR-ACE / stakeholders.
#
# NO:
# - convertir NR en 0;
# - mezclar EPN con denominadores del benchmark;
# - declarar que una materia "hace cumplir" EUR-ACE;
# - comparar semestres extranjeros por número nominal.
# ============================================================


st.set_page_config(
    page_title="AGRO · Constructor Curricular",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded",
)


FILES = {
    "master": "Base_Maestra_AGRO_v1_2_CAPAS_EVIDENCIA_NORMALIZADAS_AUDITADAS.xlsx",
    "epn": "EPN_AGRO_Dataset_Comparativo_Integrable_v1_0.xlsx",
    "benchmark": "Matriz_Comparativa_AGRO_v1_0.xlsx",
    "gaps": "Matriz_Brechas_EPN_Agro_vs_Benchmark_v1_0.xlsx",
    "caee": "CAEE_AGRO_v1_0_NORMALIZADA_AUDITADA.xlsx",
}

# Hojas mínimas: se excluyen manifiestos, auditorías y resúmenes no necesarios
# para el uso interactivo del dashboard.
SHEETS = {
    "master_courses": ("master", "09_Asignaturas", ["ID_Asignatura", "ID_Familia_Principal_Normalizada"]),
    "master_contents": ("master", "11_Contenidos", ["ID_Contenido", "ID_Asignatura", "Contenido_original"]),
    "master_curriculum_institution": ("master", "04_Curriculo_Institucion", ["ID_Curriculo_Unico", "Universidad"]),
    "family_dict": ("master", "33_Diccionario_Familias_Norm", ["ID_Area", "ID_Familia"]),
    "epn_courses": ("epn", "04_Asignaturas_60", ["ID_Asignatura", "Nombre_espanol"]),
    "epn_load": ("epn", "05_Carga_52", ["ID_Asignatura", "Créditos"]),
    "epn_prereq": ("epn", "08_Prerrequisitos", ["ID_Relacion", "ID_Asignatura_objetivo"]),
    "benchmark_universes": ("benchmark", "02_Universos", ["Codigo_Dataset", "ID_Curriculo_Unico"]),
    "benchmark_presence": ("benchmark", "04_Presencia_Curriculos", ["ID_Area", "ID_Familia"]),
    "benchmark_position": ("benchmark", "06_Posicion_Core", ["ID_Area", "ID_Familia"]),
    "benchmark_depth": ("benchmark", "07_Profundidad_Core", ["ID_Area", "ID_Familia"]),
    "benchmark_areas": ("benchmark", "08_Resumen_Areas", ["ID_Area", "Area"]),
    "gaps": ("gaps", "02_Cobertura_Familias", ["ID_Area", "ID_Familia", "Accion_Preliminar"]),
    "stakeholder_signals": ("gaps", "08_Senales_Stakeholders", ["ID", "Tema", "Familias_relacionadas"]),
    "caee_topics": ("caee", "02_Taxonomia_Temas", ["Tema_ID", "Tema_normalizado"]),
    "eurace": ("caee", "03_EURACE_Criterios", ["Criterio", "Calificación"]),
    "caee_coverage": ("caee", "14_Cobertura_Temas", ["Tema_ID", "N_evidencias_elegibles"]),
}

# 7 ejes ejecutivos. No sustituyen AGRO-NORM: son una vista de presentación.
AXES = [
    "Bases científico-cuantitativas",
    "Ingeniería y procesos",
    "Agroalimentos, calidad y producto",
    "Digitalización y automatización",
    "Sostenibilidad y circularidad",
    "Práctica e integración",
    "Gestión, innovación y profesión",
]

# Contribución por área a ejes ejecutivos (0..1).
# Es una capa DERIVADA de visualización; AGRO-NORM sigue siendo la taxonomía canónica.
AREA_AXIS_WEIGHTS: dict[str, dict[str, float]] = {
    "A01": {"Bases científico-cuantitativas": 1.0, "Digitalización y automatización": 0.25},
    "A02": {"Bases científico-cuantitativas": 1.0},
    "A03": {"Bases científico-cuantitativas": 0.85, "Agroalimentos, calidad y producto": 0.20},
    "A04": {"Ingeniería y procesos": 1.0},
    "A05": {"Ingeniería y procesos": 1.0, "Agroalimentos, calidad y producto": 0.20},
    "A06": {"Agroalimentos, calidad y producto": 1.0},
    "A07": {"Agroalimentos, calidad y producto": 1.0},
    "A08": {"Agroalimentos, calidad y producto": 1.0},
    "A09": {"Ingeniería y procesos": 0.80, "Digitalización y automatización": 0.65},
    "A10": {"Agroalimentos, calidad y producto": 0.65, "Sostenibilidad y circularidad": 0.15},
    "A11": {"Sostenibilidad y circularidad": 1.0},
    "A12": {"Gestión, innovación y profesión": 0.90},
    "A13": {"Agroalimentos, calidad y producto": 0.65, "Gestión, innovación y profesión": 0.25},
    "A14": {"Gestión, innovación y profesión": 1.0, "Agroalimentos, calidad y producto": 0.35},
    "A15": {"Práctica e integración": 1.0, "Gestión, innovación y profesión": 0.35},
    "A16": {"Gestión, innovación y profesión": 0.80},
}

# Tipología ejecutiva DERIVADA para presentación; no sustituye AGRO-NORM.
COURSE_TYPES = [
    "Básicas", "Ciencias agro/biológicas", "Ingeniería y procesos",
    "Agroalimentos y calidad", "Digital", "Sostenibilidad",
    "Gestión/administración", "Social/profesional",
    "Práctica/integración", "Optativas", "Otros",
]

TYPE_COLORS = {
    "Básicas": "#F6C945",
    "Ciencias agro/biológicas": "#9BCB65",
    "Ingeniería y procesos": "#4C78A8",
    "Agroalimentos y calidad": "#59A14F",
    "Digital": "#8F6FB8",
    "Sostenibilidad": "#4FB6A6",
    "Gestión/administración": "#F28E2B",
    "Social/profesional": "#D77FB3",
    "Práctica/integración": "#7F9DB9",
    "Optativas": "#A7A7A7",
    "Otros": "#C7C7C7",
}

AREA_TYPE_MAP = {
    "A01": "Básicas", "A02": "Básicas",
    "A03": "Ciencias agro/biológicas",
    "A04": "Ingeniería y procesos", "A05": "Ingeniería y procesos",
    "A06": "Agroalimentos y calidad", "A07": "Agroalimentos y calidad",
    "A08": "Agroalimentos y calidad", "A09": "Ingeniería y procesos",
    "A10": "Agroalimentos y calidad", "A11": "Sostenibilidad",
    "A12": "Gestión/administración", "A13": "Agroalimentos y calidad",
    "A14": "Gestión/administración", "A15": "Práctica/integración",
    "A16": "Social/profesional",
}

DIGITAL_FAMILIES = {"F0108", "F0905", "F0906"}

# Reglas de refuerzo por familia.
FAMILY_AXIS_OVERRIDES: dict[str, dict[str, float]] = {
    "F0108": {"Digitalización y automatización": 1.0},  # Ciencia de datos/IA
    "F0905": {"Digitalización y automatización": 1.0},  # Automatización/control
    "F0906": {"Digitalización y automatización": 1.0},  # Industria 4.0
    "F1102": {"Sostenibilidad y circularidad": 1.0},    # ACV/ecodiseño
    "F1103": {"Sostenibilidad y circularidad": 1.0},    # Residuos/valorización
    "F1105": {"Sostenibilidad y circularidad": 1.0},    # Circularidad
    "F1106": {"Sostenibilidad y circularidad": 1.0},    # Bioenergía
    "F1401": {"Gestión, innovación y profesión": 1.0},  # Innovación producto
    "F1403": {"Gestión, innovación y profesión": 1.0},  # I+D
    "F1404": {"Gestión, innovación y profesión": 1.0},  # DOE I+D
    "F1505": {"Práctica e integración": 1.0},           # Internship
    "F1507": {"Práctica e integración": 1.0},           # Taller/lab integrador
}

# Puente DERIVADO área/familia -> CAEE-NORM.
# Sirve para mostrar contribución potencial, NO para reescribir evidencia fuente.
AREA_TOPIC_MAP: dict[str, set[str]] = {
    "A01": {"T01", "T02", "T05", "T17"},
    "A02": {"T01", "T02", "T17"},
    "A03": {"T01", "T17"},
    "A04": {"T01", "T02", "T03", "T17"},
    "A05": {"T02", "T03", "T17"},
    "A06": {"T01", "T04", "T11", "T17"},
    "A07": {"T03", "T04", "T11"},
    "A08": {"T04", "T16"},
    "A09": {"T03", "T05", "T19"},
    "A10": {"T18", "T06"},
    "A11": {"T06", "T16", "T17"},
    "A12": {"T09", "T15", "T07"},
    "A13": {"T11", "T17", "T15"},
    "A14": {"T11", "T17", "T09"},
    "A15": {"T10", "T17", "T07", "T09"},
    "A16": {"T07", "T08", "T16", "T20"},
}

FAMILY_TOPIC_OVERRIDES: dict[str, set[str]] = {
    "F0108": {"T05"},
    "F0804": {"T04"},
    "F0805": {"T04"},
    "F0806": {"T04"},
    "F0807": {"T04"},
    "F0905": {"T05"},
    "F0906": {"T05"},
    "F1102": {"T06"},
    "F1103": {"T06"},
    "F1105": {"T06"},
    "F1106": {"T06"},
    "F1401": {"T11", "T17"},
    "F1402": {"T11", "T17"},
    "F1403": {"T11", "T17"},
    "F1404": {"T11", "T17"},
    "F1505": {"T10"},
    "F1507": {"T10", "T17"},
    "F1602": {"T07"},
    "F1603": {"T07", "T16"},
}

# Criterios EUR-ACE sensibles al diseño curricular.
# El estado oficial se lee del Excel. Solo se calcula "contribución potencial".
EURACE_TOPIC_MAP: dict[str, set[str]] = {
    "2.2": {"T01", "T03", "T04"},
    "2.3": {"T03", "T07", "T10", "T17"},
    "2.4": {"T01", "T02"},
    "2.5": {"T06", "T07", "T16"},
    "2.6": {"T03", "T04", "T05", "T10", "T17"},
    "2.7": {"T03", "T10", "T17"},
    "3.4": {"T10", "T17"},
    "3.6": {"T08", "T10"},
    "5.1": {"T01", "T03", "T10", "T17"},
    "5.2": {"T02", "T03", "T10", "T17"},
}

PRIORITY_SCORE = {
    "Crítica": 100,
    "Muy alta": 90,
    "Alta": 75,
    "Media": 50,
    "Baja": 25,
}

ACTION_ORDER = {
    "AÑADIR": 5,
    "REFORZAR": 4,
    "INTEGRAR": 3,
    "MANTENER": 2,
    "RECONSIDERAR/ELIMINAR": 1,
    "NO PRIORIZAR": 0,
}


def norm_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).lower()


def is_nr(value: Any) -> bool:
    return norm_text(value) in {"", "nr", "nan", "none", "na", "n/a"}


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    drop_cols = [
        c for c in out.columns
        if norm_text(c) == "index" or norm_text(c).startswith("unnamed:")
    ]
    if drop_cols:
        out = out.drop(columns=drop_cols, errors="ignore")
    return out


def detect_header_row(path: Path, sheet_name: str, required_tokens: Iterable[str], probe_rows: int = 18) -> int:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=probe_rows, dtype=object)
    tokens = [norm_text(t) for t in required_tokens]
    best_i, best_score = 0, -1
    for i, row in raw.iterrows():
        cells = [norm_text(v) for v in row.tolist() if not is_nr(v)]
        score = sum(1 for token in tokens if any(token == c or token in c for c in cells))
        if score > best_score:
            best_i, best_score = int(i), score
        if score >= max(1, min(2, len(tokens))):
            return int(i)
    return best_i


@st.cache_data(show_spinner=False)
def read_sheet(path_str: str, sheet_name: str, required_tokens: tuple[str, ...]) -> pd.DataFrame:
    path = Path(path_str)
    header_row = detect_header_row(path, sheet_name, required_tokens)
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row, dtype=object)
    df = clean_columns(df)
    df = df.dropna(how="all")
    return df


def numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.replace({"NR": np.nan, "": np.nan}), errors="coerce")


def yn(value: Any) -> bool | None:
    t = norm_text(value)
    if t in {"si", "sí", "yes", "true", "1"}:
        return True
    if t in {"no", "false", "0"}:
        return False
    return None


def parse_depth(value: Any) -> float:
    if is_nr(value):
        return np.nan
    m = re.search(r"([0-4])", str(value))
    return float(m.group(1)) if m else np.nan


def parse_semester(value: Any) -> float:
    if is_nr(value):
        return np.nan
    m = re.search(r"(\d+)", str(value))
    return float(m.group(1)) if m else np.nan


def split_family_ids(row: pd.Series) -> list[str]:
    ids: list[str] = []
    for col in [
        "ID_Familia_Principal_Normalizada",
        "ID_Familia_Secundaria_1",
        "ID_Familia_Secundaria_2",
        "ID_Familia_Secundaria_3",
    ]:
        val = row.get(col)
        if not is_nr(val):
            ids.append(str(val).strip())
    return list(dict.fromkeys(ids))


def slim_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    existing = [c for c in columns if c in df.columns]
    return df[existing].copy()


@st.cache_data(show_spinner="Cargando y normalizando los cinco Excel…")
def load_model(data_dir_str: str) -> dict[str, pd.DataFrame]:
    data_dir = Path(data_dir_str)
    paths = {key: data_dir / filename for key, filename in FILES.items()}

    missing = [str(path) for path in paths.values() if not path.exists()]
    if missing:
        raise FileNotFoundError(
            "Faltan archivos:\n" + "\n".join(f"- {m}" for m in missing)
        )

    raw: dict[str, pd.DataFrame] = {}
    for name, (book_key, sheet_name, tokens) in SHEETS.items():
        raw[name] = read_sheet(str(paths[book_key]), sheet_name, tuple(tokens))

    # --------------------------------------------------------
    # DIM / FACT: catálogo canónico de asignaturas
    # --------------------------------------------------------
    course_cols = [
        "ID_Asignatura",
        "ID_Curriculo_Unico",
        "Codigo_Dataset_Canonico",
        "Nombre_original",
        "Nombre_espanol",
        "Nombre_normalizado",
        "Nombre_asignatura_normalizado",
        "Asignatura_normalizada",
        "Universidad",
        "Institucion",
        "Nombre_Universidad",
        "Pais",
        "Codigo_asignatura",
        "Creditos_originales",
        "Sistema_creditos",
        "Tipo_asignatura_original",
        "Obligatoria",
        "Optativa_o_electiva",
        "Tiene_laboratorio_explicito",
        "Laboratorio_independiente",
        "Descripcion_original",
        "Objetivos_originales",
        "Resultados_aprendizaje_texto",
        "Elegible_Contenidos",
        "Elegible_Carga",
        "ID_Area_Principal_Normalizada",
        "Area_Principal_Normalizada",
        "ID_Familia_Principal_Normalizada",
        "Familia_Principal_Normalizada",
        "ID_Familia_Secundaria_1",
        "Familia_Secundaria_1",
        "ID_Familia_Secundaria_2",
        "Familia_Secundaria_2",
        "ID_Familia_Secundaria_3",
        "Familia_Secundaria_3",
        "Nivel_Profundidad",
        "Tipo_Experiencia_Normalizado",
        "Componente_Experimental_Normalizado",
        "Componente_Computacional_Normalizado",
        "Numero_Periodo_Comparable",
        "Avance_Curricular_Medio_pct",
        "Etapa_Curricular_Normalizada",
        "Confianza_Normalizacion",
        "Estado_Normalizacion_Final",
    ]
    catalog = slim_columns(raw["master_courses"], course_cols)

    if "Creditos_originales" in catalog:
        catalog["creditos_fuente"] = numeric(catalog["Creditos_originales"])
    else:
        catalog["creditos_fuente"] = np.nan

    if "Numero_Periodo_Comparable" in catalog:
        catalog["semestre_fuente"] = numeric(catalog["Numero_Periodo_Comparable"])
    else:
        catalog["semestre_fuente"] = np.nan

    if "Avance_Curricular_Medio_pct" in catalog:
        catalog["avance_pct"] = numeric(catalog["Avance_Curricular_Medio_pct"])
    else:
        catalog["avance_pct"] = np.nan

    catalog["depth_n"] = catalog.get("Nivel_Profundidad", pd.Series(index=catalog.index, dtype=object)).map(parse_depth)
    catalog["familias"] = catalog.apply(split_family_ids, axis=1)
    catalog["nombre_display"] = catalog.apply(
        lambda r: (
            r.get("Nombre_espanol")
            if not is_nr(r.get("Nombre_espanol"))
            else r.get("Nombre_original")
        ),
        axis=1,
    )
    catalog["source_label"] = catalog.apply(
        lambda r: f"{r.get('nombre_display', 'NR')} · {r.get('Codigo_Dataset_Canonico', 'NR')} · {r.get('Familia_Principal_Normalizada', 'sin familia')}",
        axis=1,
    )

    # Universidad(es) de origen. Para programas conjuntos se agregan los nombres
    # en una sola cadena y así no se duplica ninguna asignatura al hacer el merge.
    ci = raw["master_curriculum_institution"].copy()
    if {"ID_Curriculo_Unico", "Universidad"}.issubset(ci.columns):
        ci = ci[["ID_Curriculo_Unico", "Universidad"]].dropna(subset=["ID_Curriculo_Unico"]).copy()
        ci["Universidad"] = ci["Universidad"].fillna("").astype(str).str.strip()
        ci_agg = (
            ci.groupby("ID_Curriculo_Unico", as_index=False)["Universidad"]
              .agg(lambda values: " / ".join(dict.fromkeys(v for v in values if v)))
              .rename(columns={"Universidad": "Universidades_Origen"})
        )
        catalog = catalog.merge(ci_agg, on="ID_Curriculo_Unico", how="left")

    # Contenidos: solo campos que alimentan comparación cualitativa.
    contents = slim_columns(
        raw["master_contents"],
        [
            "ID_Contenido",
            "ID_Asignatura",
            "ID_Programa",
            "Contenido_original",
            "Contenido_espanol",
            "Subcontenido_original",
            "Orden_aparicion",
            "Tipo_evidencia",
            "Nivel_detalle_fuente",
            "Nivel_profundidad",
            "Componente_experimental",
            "Componente_computacional",
            "Confianza_normalizacion",
        ],
    )
    # Texto canónico para comparación cualitativa. Se genera siempre, pero la
    # función consumidora también tiene fallback por seguridad.
    def _contenido_display(row: pd.Series) -> str:
        for col in ["Contenido_espanol", "Contenido_original", "Subcontenido_original"]:
            value = row.get(col)
            if not is_nr(value):
                return str(value).strip()
        return ""

    contents["contenido_display"] = contents.apply(_contenido_display, axis=1)

    # EPN: baseline operativo.
    epn = raw["epn_courses"].copy()
    epn_load = raw["epn_load"].copy()
    if "ID_Asignatura" in epn_load:
        keep_load = [
            c for c in ["ID_Asignatura", "Créditos", "AC_h", "AP_h", "AA_h", "Total_h"]
            if c in epn_load.columns
        ]
        epn = epn.merge(epn_load[keep_load], on="ID_Asignatura", how="left")
    epn_master = catalog[catalog["Codigo_Dataset_Canonico"].astype(str).str.upper().eq("EPN")].copy()
    enrich_cols = [
        "ID_Asignatura",
        "ID_Area_Principal_Normalizada",
        "Area_Principal_Normalizada",
        "ID_Familia_Principal_Normalizada",
        "Familia_Principal_Normalizada",
        "ID_Familia_Secundaria_1",
        "ID_Familia_Secundaria_2",
        "ID_Familia_Secundaria_3",
        "Nivel_Profundidad",
        "Componente_Experimental_Normalizado",
        "Componente_Computacional_Normalizado",
        "Tipo_Experiencia_Normalizado",
        "Numero_Periodo_Comparable",
        "Avance_Curricular_Medio_pct",
        "Estado_Normalizacion_Final",
    ]
    enrich_cols = [c for c in enrich_cols if c in epn_master.columns]
    epn = epn.merge(epn_master[enrich_cols], on="ID_Asignatura", how="left", suffixes=("", "_master"))

    # Benchmark universos / matrices largas.
    universes = raw["benchmark_universes"].copy()

    # Enriquecimiento ligero del catálogo para búsqueda por universidad/programa.
    # Solo incorpora metadatos descriptivos presentes en 02_Universos; no altera
    # las claves ni los datos curriculares canónicos de la Base Maestra.
    if "Codigo_Dataset" in universes.columns and "Codigo_Dataset_Canonico" in catalog.columns:
        meta_cols = ["Codigo_Dataset"]
        for col in universes.columns:
            nc = norm_text(col)
            if any(token in nc for token in ["universidad", "institucion", "programa", "pais", "country"]):
                meta_cols.append(col)
        meta_cols = list(dict.fromkeys(meta_cols))
        meta = universes[meta_cols].drop_duplicates(subset=["Codigo_Dataset"]).copy()
        rename_meta = {c: f"ref_{c}" for c in meta.columns if c != "Codigo_Dataset"}
        meta = meta.rename(columns=rename_meta)
        catalog = catalog.merge(
            meta,
            left_on="Codigo_Dataset_Canonico",
            right_on="Codigo_Dataset",
            how="left",
        ).drop(columns=["Codigo_Dataset"], errors="ignore")

    codes = (
        universes["Codigo_Dataset"].dropna().astype(str).str.strip().tolist()
        if "Codigo_Dataset" in universes
        else []
    )

    def melt_matrix(df: pd.DataFrame, value_name: str) -> pd.DataFrame:
        id_vars = [c for c in ["ID_Area", "ID_Familia", "Familia"] if c in df.columns]
        value_vars = [c for c in codes if c in df.columns]
        if not value_vars:
            return pd.DataFrame(columns=id_vars + ["Codigo_Dataset", value_name])
        out = df.melt(
            id_vars=id_vars,
            value_vars=value_vars,
            var_name="Codigo_Dataset",
            value_name=value_name,
        )
        out[value_name] = numeric(out[value_name])
        return out

    benchmark_presence = melt_matrix(raw["benchmark_presence"], "presencia_codigo")
    benchmark_position = melt_matrix(raw["benchmark_position"], "posicion_pct")
    benchmark_depth = melt_matrix(raw["benchmark_depth"], "profundidad_n")

    # Brechas: tabla central por familia para decisiones.
    gaps = raw["gaps"].copy()
    for col in [
        "Presencia_Core_pct",
        "Presencia_Core_Obligatoria_pct",
        "Presencia_Master_pct",
        "Peso_Benchmark_Core_pct",
        "Primera_Aparicion_Benchmark_pct",
        "Creditos_Core_Principal_EPN",
        "Peso_Core_EPN_pct",
        "Ratio_Peso_EPN_vs_Benchmark",
        "Primera_Aparicion_EPN_pct",
    ]:
        if col in gaps:
            gaps[col] = numeric(gaps[col])
    if "Profundidad_Benchmark" in gaps:
        gaps["benchmark_depth_n"] = gaps["Profundidad_Benchmark"].map(parse_depth)
    if "Profundidad_EPN" in gaps:
        gaps["epn_depth_n"] = gaps["Profundidad_EPN"].map(parse_depth)

    # CAEE.
    caee_coverage = raw["caee_coverage"].copy()
    for col in ["N_evidencias_elegibles", "N_corrientes"]:
        if col in caee_coverage:
            caee_coverage[col] = numeric(caee_coverage[col])

    # Modelo mínimo devuelto.
    return {
        "catalog": catalog,
        "contents": contents,
        "family_dict": raw["family_dict"],
        "curriculum_institution": raw["master_curriculum_institution"],
        "epn": epn,
        "epn_prereq": raw["epn_prereq"],
        "universes": universes,
        "benchmark_presence": benchmark_presence,
        "benchmark_position": benchmark_position,
        "benchmark_depth": benchmark_depth,
        "benchmark_areas": raw["benchmark_areas"],
        "gaps": gaps,
        "stakeholder_signals": raw["stakeholder_signals"],
        "caee_topics": raw["caee_topics"],
        "eurace": raw["eurace"],
        "caee_coverage": caee_coverage,
    }


def derive_course_type_from_row(row: pd.Series | dict[str, Any]) -> str:
    """Tipología ejecutiva. Optativa tiene prioridad; luego familia/área."""
    get = row.get
    if yn(get("Optativa_o_electiva")) is True:
        return "Optativas"
    tipo_orig = norm_text(get("Tipo_asignatura_original"))
    if any(k in tipo_orig for k in ["optativ", "electiv"]):
        return "Optativas"
    family_id = str(get("ID_Familia_Principal_Normalizada") or get("family_id") or "")
    if family_id in DIGITAL_FAMILIES:
        return "Digital"
    area_id = str(get("ID_Area_Principal_Normalizada") or get("area_id") or "")
    return AREA_TYPE_MAP.get(area_id, "Otros")


def course_search_fields(catalog: pd.DataFrame) -> list[str]:
    """Campos disponibles que aportan a la búsqueda humana de una asignatura."""
    preferred = [
        "nombre_display",
        "Nombre_espanol",
        "Nombre_original",
        "Nombre_normalizado",
        "Nombre_asignatura_normalizado",
        "Asignatura_normalizada",
        "Codigo_asignatura",
        "Codigo_Dataset_Canonico",
        "Universidad",
        "Institucion",
        "Nombre_Universidad",
        "Universidades_Origen",
        "Pais",
        "Familia_Principal_Normalizada",
        "Area_Principal_Normalizada",
    ]
    fields = [c for c in preferred if c in catalog.columns]
    for col in catalog.columns:
        nc = norm_text(col)
        if any(token in nc for token in ["universidad", "institucion", "programa", "pais", "normalizado"]):
            if col not in fields:
                fields.append(col)
    return fields


def normalized_course_name(row: pd.Series) -> str:
    for col in ["Nombre_normalizado", "Nombre_asignatura_normalizado", "Asignatura_normalizada"]:
        value = row.get(col)
        if not is_nr(value):
            return str(value).strip()
    value = row.get("Familia_Principal_Normalizada")
    return "" if is_nr(value) else str(value).strip()


def course_origin_label(row: pd.Series) -> str:
    for col in [
        "Universidades_Origen", "Universidad", "Nombre_Universidad", "Institucion",
        "ref_Universidad", "ref_Nombre_Universidad", "ref_Institucion",
        "ref_Programa",
    ]:
        value = row.get(col)
        if not is_nr(value):
            return str(value).strip()
    return str(row.get("Codigo_Dataset_Canonico", "NR"))


def course_option_label(row: pd.Series, score: float | None = None) -> str:
    """Etiqueta visible limpia del buscador.

    Solo presenta la información curricular solicitada:
    nombre de la materia · universidad/origen · nombre normalizado.
    El ID y la puntuación fuzzy permanecen internos y nunca se muestran.
    """
    name = row.get("nombre_display", "NR")
    origin = course_origin_label(row)
    normalized = normalized_course_name(row) or "NR"
    return f"{name} · {origin} · {normalized}"

def fuzzy_course_matches(
    catalog: pd.DataFrame,
    query: str,
    limit: int | None = 160,
) -> pd.DataFrame:
    """Búsqueda fuzzy sobre todo el catálogo filtrado.

    - Sin texto: devuelve TODAS las opciones disponibles, ordenadas primero por
      semestre/periodo temprano y después por avance curricular y nombre.
    - Con texto: puntúa nombre de materia, universidad, nombre normalizado,
      familia, área, dataset y código; los detalles técnicos de la puntuación
      no se muestran al usuario.
    """
    if catalog.empty:
        return catalog.copy()

    work = catalog.copy()
    q = norm_text(query)
    fields = course_search_fields(work)

    # Columnas auxiliares para un orden estable y curricularmente intuitivo.
    if "semestre_fuente" in work.columns:
        work["_semester_sort"] = pd.to_numeric(work["semestre_fuente"], errors="coerce")
    elif "Numero_Periodo_Comparable" in work.columns:
        work["_semester_sort"] = pd.to_numeric(work["Numero_Periodo_Comparable"], errors="coerce")
    else:
        work["_semester_sort"] = np.nan

    if "avance_pct" in work.columns:
        work["_progress_sort"] = pd.to_numeric(work["avance_pct"], errors="coerce")
    elif "Avance_Curricular_Medio_pct" in work.columns:
        work["_progress_sort"] = pd.to_numeric(work["Avance_Curricular_Medio_pct"], errors="coerce")
    else:
        work["_progress_sort"] = np.nan

    work["_origin_sort"] = work.apply(course_origin_label, axis=1)

    if not q:
        # Cuando el buscador está vacío NO se recorta el catálogo:
        # se muestran todas las opciones y aparecen primero las de etapas tempranas.
        work["_fuzzy_score"] = 100.0
        work = work.sort_values(
            ["_semester_sort", "_progress_sort", "nombre_display", "_origin_sort"],
            ascending=[True, True, True, True],
            na_position="last",
        )
        return work.copy() if limit is None else work.head(limit).copy()

    def score_row(row: pd.Series) -> float:
        best = 0.0
        for col in fields:
            value = row.get(col)
            if is_nr(value):
                continue
            text_value = norm_text(value)
            if not text_value:
                continue

            if q == text_value:
                field_score = 100.0
            elif q in text_value:
                # Una coincidencia literal parcial debe quedar arriba.
                field_score = 99.0
            elif RAPIDFUZZ_AVAILABLE:
                field_score = max(
                    float(fuzz.WRatio(q, text_value)),
                    float(fuzz.partial_ratio(q, text_value)),
                    float(fuzz.token_set_ratio(q, text_value)),
                )
            else:
                field_score = 100.0 * SequenceMatcher(None, q, text_value).ratio()

            best = max(best, field_score)
        return best

    work["_fuzzy_score"] = work.apply(score_row, axis=1)
    work = work.sort_values(
        ["_fuzzy_score", "_semester_sort", "_progress_sort", "nombre_display", "_origin_sort"],
        ascending=[False, True, True, True, True],
        na_position="last",
    )

    if limit is None:
        return work.copy()
    return work.head(limit).copy()

def searchbox_course_selector(catalog: pd.DataFrame) -> pd.Series | None:
    """Buscador único fuzzy con catálogo completo al abrirse.

    La interfaz muestra únicamente:
        Materia · Universidad · Nombre normalizado

    Los IDs y scores fuzzy quedan ocultos. Internamente se conserva un mapa
    etiqueta->fila para resolver la selección sin contaminar la interfaz.
    """
    if catalog.empty:
        st.warning("No hay asignaturas disponibles con los filtros actuales.")
        return None

    # Orden de navegación inicial: primeros semestres/etapas antes que etapas tardías.
    browse_catalog = fuzzy_course_matches(catalog, "", limit=None)

    # Puede haber etiquetas repetidas (misma materia/universidad/nombre normalizado).
    # Conservamos todos los índices y resolvemos por el primer registro en el orden
    # curricular de navegación, sin mostrar un identificador técnico al usuario.
    label_to_indices: dict[str, list[Any]] = {}
    for idx, row in browse_catalog.iterrows():
        label = course_option_label(row)
        label_to_indices.setdefault(label, []).append(idx)

    # Lista completa para el primer clic con el cuadro vacío.
    default_labels = list(label_to_indices.keys())

    def suggest(searchterm: str) -> list[str]:
        # Vacío = TODO el catálogo. Con texto = fuzzy.
        if not str(searchterm or "").strip():
            return default_labels

        matches = fuzzy_course_matches(catalog, searchterm, limit=180)
        labels: list[str] = []
        seen: set[str] = set()
        for _, row in matches.iterrows():
            label = course_option_label(row)
            if label not in seen:
                labels.append(label)
                seen.add(label)
        return labels

    def resolve_label(selected_label: str | None) -> pd.Series | None:
        if not selected_label:
            return None
        indices = label_to_indices.get(str(selected_label), [])
        if not indices:
            # Por seguridad, volver a buscar la etiqueta entre todo el catálogo.
            hit = browse_catalog[
                browse_catalog.apply(course_option_label, axis=1).astype(str).eq(str(selected_label))
            ]
            return hit.iloc[0] if not hit.empty else None
        return catalog.loc[indices[0]]

    if SEARCHBOX_AVAILABLE:
        selected = st_searchbox(
            suggest,
            key="course_fuzzy_searchbox",
            label="Buscar asignatura",
            placeholder="Nombre de materia, universidad o nombre normalizado…",
            help=(
                "Con el campo vacío se muestran todas las asignaturas disponibles, "
                "ordenadas desde las etapas curriculares más tempranas. Al escribir, "
                "la búsqueda fuzzy tolera errores y busca por materia, universidad, "
                "nombre normalizado, familia, área, dataset y código."
            ),
            default_options=default_labels,
        )
        return resolve_label(selected)

    st.warning(
        "Falta `streamlit-searchbox`; se usa un fallback de dos pasos. "
        "Ejecute `pip install -r requirements.txt` para activar el buscador único."
    )
    query = st.text_input(
        "Buscar asignatura",
        placeholder="Nombre de materia, universidad o nombre normalizado…",
        help=(
            "Deje el campo vacío para ver todo el catálogo ordenado por etapa curricular. "
            "Al escribir se aplica búsqueda fuzzy."
        ),
    )
    matches = fuzzy_course_matches(
        catalog,
        query,
        limit=None if not str(query or "").strip() else 180,
    )
    if matches.empty:
        return None

    options: list[str] = []
    seen: set[str] = set()
    for _, row in matches.iterrows():
        label = course_option_label(row)
        if label not in seen:
            options.append(label)
            seen.add(label)

    selected = st.selectbox(
        "Asignaturas disponibles",
        options,
        index=None,
        placeholder="Seleccione una asignatura…",
    )
    return resolve_label(selected)

def selected_course_details(selected_row: pd.Series, contents: pd.DataFrame) -> None:
    course_id = selected_row.get("ID_Asignatura")
    with st.expander("Ver información de la materia seleccionada", expanded=False):
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"**Origen**  \n{course_origin_label(selected_row)}")
        c2.markdown(f"**Tipo visual**  \n{derive_course_type_from_row(selected_row)}")
        c3.markdown(f"**Etapa original**  \n{selected_row.get('Etapa_Curricular_Normalizada', 'NR')}")
        info = {
            "Nombre normalizado": normalized_course_name(selected_row) or "NR",
            "Dataset": selected_row.get("Codigo_Dataset_Canonico", "NR"),
            "Área AGRO-NORM": selected_row.get("Area_Principal_Normalizada", "NR"),
            "Familia AGRO-NORM": selected_row.get("Familia_Principal_Normalizada", "NR"),
            "Profundidad": selected_row.get("Nivel_Profundidad", "NR"),
            "Periodo comparable": selected_row.get("Numero_Periodo_Comparable", "NR"),
            "Avance curricular (%)": selected_row.get("Avance_Curricular_Medio_pct", "NR"),
            "Créditos fuente": selected_row.get("Creditos_originales", "NR"),
            "Sistema de créditos": selected_row.get("Sistema_creditos", "NR"),
            "Laboratorio explícito": selected_row.get("Tiene_laboratorio_explicito", "NR"),
            "Experimental": selected_row.get("Componente_Experimental_Normalizado", "NR"),
            "Computacional": selected_row.get("Componente_Computacional_Normalizado", "NR"),
        }
        st.dataframe(pd.DataFrame([info]), use_container_width=True, hide_index=True)
        for label, col in [
            ("Descripción", "Descripcion_original"),
            ("Objetivos", "Objetivos_originales"),
            ("Resultados de aprendizaje", "Resultados_aprendizaje_texto"),
        ]:
            value = selected_row.get(col)
            if not is_nr(value):
                st.markdown(f"**{label}**")
                st.write(str(value))
        samples = content_samples_for_course_ids(contents, [course_id], limit=10)
        st.markdown("**Contenidos explícitos recuperados**")
        if samples:
            for item in samples:
                st.write("•", item)
        else:
            st.caption("Sin microcontenido explícito recuperado para esta materia.")


def type_distribution_from_courses(courses: list[dict[str, Any]]) -> pd.DataFrame:
    total = sum(float(c.get("credits") or 0) for c in courses)
    rows = [{"Tipo": c.get("course_type") or derive_course_type_from_row(c),
             "Créditos": float(c.get("credits") or 0)} for c in courses]
    if not rows:
        return pd.DataFrame({"Tipo": COURSE_TYPES, "Créditos": 0.0, "Porcentaje": 0.0})
    out = pd.DataFrame(rows).groupby("Tipo", as_index=False)["Créditos"].sum()
    out = pd.DataFrame({"Tipo": COURSE_TYPES}).merge(out, on="Tipo", how="left").fillna({"Créditos": 0.0})
    out["Porcentaje"] = np.where(total > 0, 100 * out["Créditos"] / total, 0.0)
    return out


def type_distribution_from_foreign(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame({"Tipo": COURSE_TYPES, "Conteo": 0.0, "Porcentaje": 0.0})
    work = df.copy()
    work["Tipo"] = work.apply(derive_course_type_from_row, axis=1)
    out = work.groupby("Tipo", as_index=False).size().rename(columns={"size": "Conteo"})
    out = pd.DataFrame({"Tipo": COURSE_TYPES}).merge(out, on="Tipo", how="left").fillna({"Conteo": 0.0})
    total = out["Conteo"].sum()
    out["Porcentaje"] = np.where(total > 0, 100 * out["Conteo"] / total, 0.0)
    return out


def type_radar_figure(ours: pd.DataFrame, foreign: pd.DataFrame, foreign_label: str) -> go.Figure:
    cats = COURSE_TYPES[:-1]
    ours_map = dict(zip(ours["Tipo"], ours["Porcentaje"]))
    foreign_map = dict(zip(foreign["Tipo"], foreign["Porcentaje"]))
    theta = cats + [cats[0]]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=[ours_map.get(x, 0) for x in cats] + [ours_map.get(cats[0], 0)],
        theta=theta, fill="toself", name="Nuestra propuesta", opacity=.45,
    ))
    fig.add_trace(go.Scatterpolar(
        r=[foreign_map.get(x, 0) for x in cats] + [foreign_map.get(cats[0], 0)],
        theta=theta, fill="toself", name=foreign_label, opacity=.32,
    ))
    max_val = max([25.0] + list(ours_map.values()) + list(foreign_map.values()))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, min(100, math.ceil(max_val / 10) * 10 + 10)])),
        margin=dict(l=25, r=25, t=50, b=20), height=470,
        title="Composición por tipos", legend=dict(orientation="h"),
    )
    return fig


def row_to_course(row: pd.Series, semester: int | None = None, credits: float | None = None, origin: str | None = None) -> dict[str, Any]:
    name = row.get("nombre_display", row.get("Nombre_espanol", row.get("Nombre_original", "Materia")))
    source_id = row.get("ID_Asignatura")
    semester_val = semester
    if semester_val is None:
        semester_val = int(row["semestre_fuente"]) if not pd.isna(row.get("semestre_fuente", np.nan)) else None

    source_credits = row.get("creditos_fuente", np.nan)
    if credits is None:
        credits = float(source_credits) if not pd.isna(source_credits) else 3.0

    return {
        "instance_id": str(uuid.uuid4()),
        "source_course_id": source_id,
        "source_dataset": row.get("Codigo_Dataset_Canonico", "CUSTOM"),
        "code": row.get("Codigo_asignatura", ""),
        "name": str(name),
        "semester": semester_val,
        "credits": int(round(float(credits))),
        "course_type": derive_course_type_from_row(row),
        "area_id": row.get("ID_Area_Principal_Normalizada", ""),
        "area": row.get("Area_Principal_Normalizada", ""),
        "family_id": row.get("ID_Familia_Principal_Normalizada", ""),
        "family": row.get("Familia_Principal_Normalizada", ""),
        "family_ids": row.get("familias", split_family_ids(row)),
        "depth_n": None if pd.isna(row.get("depth_n", np.nan)) else float(row.get("depth_n")),
        "experimental": row.get("Componente_Experimental_Normalizado", row.get("Tiene_laboratorio_explicito", "")),
        "computational": row.get("Componente_Computacional_Normalizado", ""),
        "lab_explicit": bool(yn(row.get("Tiene_laboratorio_explicito")) is True),
        "experience_type": row.get("Tipo_Experiencia_Normalizado", ""),
        "source_credits": None if pd.isna(source_credits) else float(source_credits),
        "source_credit_system": row.get("Sistema_creditos", ""),
        "origin": origin or str(row.get("Codigo_Dataset_Canonico", "")),
        "normalization_status": row.get("Estado_Normalizacion_Final", ""),
    }


def initialize_epn_curriculum(model: dict[str, pd.DataFrame], n_semesters: int = 9) -> list[dict[str, Any]]:
    catalog = model["catalog"]
    epn_rows = catalog[catalog["Codigo_Dataset_Canonico"].astype(str).str.upper().eq("EPN")].copy()
    epn_rows = epn_rows[
        epn_rows["semestre_fuente"].notna()
        & epn_rows["semestre_fuente"].between(1, n_semesters)
    ]
    curriculum = [
        row_to_course(row, origin="EPN actual")
        for _, row in epn_rows.iterrows()
    ]
    curriculum.sort(key=lambda x: (x.get("semester") or 99, x.get("name", "")))
    return curriculum


def get_course_axis_weights(course: dict[str, Any]) -> dict[str, float]:
    weights = {axis: 0.0 for axis in AXES}
    area_id = str(course.get("area_id") or "")
    family_id = str(course.get("family_id") or "")

    for axis, w in AREA_AXIS_WEIGHTS.get(area_id, {}).items():
        weights[axis] = max(weights[axis], w)

    for axis, w in FAMILY_AXIS_OVERRIDES.get(family_id, {}).items():
        weights[axis] = max(weights[axis], w)

    # Señales explícitas de práctica/computación.
    experimental = yn(course.get("experimental"))
    computational = yn(course.get("computational"))
    if course.get("lab_explicit") or experimental is True:
        weights["Práctica e integración"] = max(weights["Práctica e integración"], 0.65)
    if computational is True:
        weights["Digitalización y automatización"] = max(weights["Digitalización y automatización"], 0.55)

    return weights


def semester_courses(curriculum: list[dict[str, Any]], semester: int) -> list[dict[str, Any]]:
    return [c for c in curriculum if int(c.get("semester") or -1) == semester]


def semester_axis_scores(curriculum: list[dict[str, Any]], semester: int) -> dict[str, float]:
    courses = semester_courses(curriculum, semester)
    total = sum(float(c.get("credits") or 0) for c in courses)
    if total <= 0:
        return {axis: 0.0 for axis in AXES}

    scores = {axis: 0.0 for axis in AXES}
    for course in courses:
        credits = float(course.get("credits") or 0)
        weights = get_course_axis_weights(course)
        for axis in AXES:
            scores[axis] += credits * weights[axis]

    return {axis: round(100 * scores[axis] / total, 1) for axis in AXES}


def semester_quality_metrics(
    curriculum: list[dict[str, Any]],
    semester: int,
    gaps: pd.DataFrame,
) -> dict[str, Any]:
    courses = semester_courses(curriculum, semester)
    total = sum(float(c.get("credits") or 0) for c in courses)
    if total <= 0:
        return {
            "credits": 0,
            "courses": 0,
            "practice_pct": np.nan,
            "digital_pct": np.nan,
            "benchmark_core": np.nan,
            "stakeholder_priority": np.nan,
            "emerging_pct": np.nan,
            "depth_delta": np.nan,
        }

    practice_credits = 0.0
    digital_credits = 0.0
    family_credit: dict[str, float] = {}

    for c in courses:
        cr = float(c.get("credits") or 0)
        exp = yn(c.get("experimental"))
        comp = yn(c.get("computational"))
        if c.get("lab_explicit") or exp is True or str(c.get("area_id")) == "A15":
            practice_credits += cr
        if comp is True or str(c.get("family_id")) in {"F0108", "F0905", "F0906"}:
            digital_credits += cr
        fam = str(c.get("family_id") or "")
        if fam:
            family_credit[fam] = family_credit.get(fam, 0.0) + cr

    relevant = gaps[gaps["ID_Familia"].astype(str).isin(family_credit.keys())].copy()
    if relevant.empty:
        benchmark_core = stakeholder_priority = emerging_pct = depth_delta = np.nan
    else:
        relevant["w"] = relevant["ID_Familia"].astype(str).map(family_credit).fillna(0)
        valid_presence = relevant["Presencia_Core_pct"].notna()
        benchmark_core = (
            np.average(relevant.loc[valid_presence, "Presencia_Core_pct"], weights=relevant.loc[valid_presence, "w"]) * 100
            if valid_presence.any()
            else np.nan
        )

        relevant["priority_score"] = relevant.get(
            "Prioridad_Stakeholders", pd.Series(index=relevant.index, dtype=object)
        ).map(PRIORITY_SCORE)
        valid_priority = relevant["priority_score"].notna()
        stakeholder_priority = (
            np.average(relevant.loc[valid_priority, "priority_score"], weights=relevant.loc[valid_priority, "w"])
            if valid_priority.any()
            else np.nan
        )

        emerging_fams = set(
            relevant.loc[
                relevant.get("Senal_Emergente", pd.Series(index=relevant.index, dtype=object))
                .fillna("")
                .astype(str)
                .str.strip()
                .ne(""),
                "ID_Familia",
            ].astype(str)
        )
        emerging_credits = sum(family_credit.get(f, 0.0) for f in emerging_fams)
        emerging_pct = 100 * emerging_credits / total

        deltas: list[tuple[float, float]] = []
        for _, r in relevant.iterrows():
            fam = str(r["ID_Familia"])
            bdepth = r.get("benchmark_depth_n", np.nan)
            if pd.isna(bdepth):
                continue
            course_depths = [
                c.get("depth_n")
                for c in courses
                if str(c.get("family_id") or "") == fam and c.get("depth_n") is not None
            ]
            if course_depths:
                deltas.append((float(np.nanmax(course_depths)) - float(bdepth), family_credit.get(fam, 1)))
        depth_delta = (
            np.average([d[0] for d in deltas], weights=[d[1] for d in deltas])
            if deltas
            else np.nan
        )

    return {
        "credits": round(total, 1),
        "courses": len(courses),
        "practice_pct": round(100 * practice_credits / total, 1),
        "digital_pct": round(100 * digital_credits / total, 1),
        "benchmark_core": None if pd.isna(benchmark_core) else round(float(benchmark_core), 1),
        "stakeholder_priority": None if pd.isna(stakeholder_priority) else round(float(stakeholder_priority), 1),
        "emerging_pct": None if pd.isna(emerging_pct) else round(float(emerging_pct), 1),
        "depth_delta": None if pd.isna(depth_delta) else round(float(depth_delta), 2),
    }


def course_topics(course: dict[str, Any]) -> set[str]:
    area = str(course.get("area_id") or "")
    fams = set(str(f) for f in course.get("family_ids", []) if f)
    topics = set(AREA_TOPIC_MAP.get(area, set()))
    for fam in fams:
        topics |= FAMILY_TOPIC_OVERRIDES.get(fam, set())
    return topics


def semester_topics(curriculum: list[dict[str, Any]], semester: int) -> set[str]:
    topics: set[str] = set()
    for course in semester_courses(curriculum, semester):
        topics |= course_topics(course)
    return topics


def eurace_alignment(
    curriculum: list[dict[str, Any]],
    semester: int,
    eurace: pd.DataFrame,
) -> pd.DataFrame:
    topics = semester_topics(curriculum, semester)
    rows = []
    for criterion, required in EURACE_TOPIC_MAP.items():
        official = eurace[eurace["Criterio"].astype(str).str.strip().eq(criterion)]
        if official.empty:
            continue
        r = official.iloc[0]
        covered = required & topics
        ratio = len(covered) / len(required) if required else 0
        if ratio >= 0.75:
            contribution = "Alta"
        elif ratio > 0:
            contribution = "Media"
        else:
            contribution = "Sin impacto directo"

        status = str(r.get("Calificación", "NR"))
        if contribution != "Sin impacto directo":
            interpretation = (
                "Contribuye a una brecha vigente"
                if norm_text(status) in {"parcialmente", "no cumple"}
                else "Refuerza un criterio actualmente cumplido"
            )
        else:
            interpretation = "No se infiere efecto desde este semestre"

        rows.append({
            "Criterio EUR-ACE": criterion,
            "Sección": r.get("Sección", ""),
            "Estado oficial": status,
            "Contribución propuesta": contribution,
            "Temas CAEE vinculados": ", ".join(sorted(covered)) if covered else "—",
            "Lectura": interpretation,
            "Comentario experto": r.get("Comentario_expertos", ""),
        })
    return pd.DataFrame(rows)


def radar_figure(current: dict[str, float], baseline: dict[str, float]) -> go.Figure:
    categories = AXES + [AXES[0]]
    curr = [current[a] for a in AXES] + [current[AXES[0]]]
    base = [baseline[a] for a in AXES] + [baseline[AXES[0]]]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=base,
        theta=categories,
        fill="toself",
        name="EPN actual / baseline",
        opacity=0.35,
    ))
    fig.add_trace(go.Scatterpolar(
        r=curr,
        theta=categories,
        fill="toself",
        name="Propuesta",
        opacity=0.45,
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        margin=dict(l=30, r=30, t=50, b=30),
        height=460,
        title="Perfil de equilibrio curricular del semestre",
        legend=dict(orientation="h"),
    )
    return fig


def delta_figure(current: dict[str, float], baseline: dict[str, float]) -> go.Figure:
    df = pd.DataFrame({
        "Eje": AXES,
        "Delta": [round(current[a] - baseline[a], 1) for a in AXES],
    }).sort_values("Delta")
    fig = px.bar(
        df,
        x="Delta",
        y="Eje",
        orientation="h",
        title="Ganancias / pérdidas frente al baseline",
        labels={"Delta": "Cambio en intensidad (puntos)"},
    )
    fig.add_vline(x=0, line_width=1)
    fig.update_layout(height=420, margin=dict(l=20, r=20, t=50, b=20))
    return fig


def build_semester_family_set(courses: list[dict[str, Any]]) -> set[str]:
    out: set[str] = set()
    for c in courses:
        out |= set(str(x) for x in c.get("family_ids", []) if x and not is_nr(x))
    return out


def foreign_window_courses(
    catalog: pd.DataFrame,
    dataset: str,
    semester: int,
    n_semesters: int,
) -> pd.DataFrame:
    low = (semester - 1) / n_semesters * 100
    high = semester / n_semesters * 100
    df = catalog[
        catalog["Codigo_Dataset_Canonico"].astype(str).eq(dataset)
        & catalog["avance_pct"].notna()
        & catalog["avance_pct"].ge(low)
        & catalog["avance_pct"].lt(high + 1e-9)
    ].copy()
    if "Estado_Normalizacion_Final" in df:
        approved = df["Estado_Normalizacion_Final"].astype(str).str.upper().eq("APROBADO")
        df = df[approved]
    return df


def content_samples_for_course_ids(
    contents: pd.DataFrame,
    course_ids: Iterable[str],
    limit: int = 8,
) -> list[str]:
    """Devuelve muestras de contenido sin asumir que existe `contenido_display`.

    Esta función es deliberadamente defensiva porque las distintas versiones de
    la Base Maestra pueden exponer el texto como Contenido_espanol,
    Contenido_original o Subcontenido_original. NR nunca se transforma en cero
    ni en ausencia curricular.
    """
    ids = {str(x).strip() for x in course_ids if x is not None and str(x).strip()}
    if not ids or contents is None or contents.empty:
        return []

    if "ID_Asignatura" not in contents.columns:
        return []

    c = contents[contents["ID_Asignatura"].astype(str).str.strip().isin(ids)].copy()
    if c.empty:
        return []

    # Fallback robusto: reconstruir la columna derivada si no llegó desde
    # load_model() (por caché de Streamlit o por diferencias de versión Excel).
    if "contenido_display" not in c.columns:
        candidate_cols = [
            col for col in [
                "Contenido_espanol",
                "Contenido_original",
                "Subcontenido_original",
                "Contenido",
                "Tema",
            ]
            if col in c.columns
        ]
        if not candidate_cols:
            return []

        def pick_content(row: pd.Series) -> str:
            for col in candidate_cols:
                value = row.get(col)
                if not is_nr(value):
                    return str(value).strip()
            return ""

        c["contenido_display"] = c.apply(pick_content, axis=1)

    c = c[c["contenido_display"].notna()].copy()
    c = c[~c["contenido_display"].map(is_nr)].copy()

    vals: list[str] = []
    for value in c["contenido_display"].astype(str):
        value = re.sub(r"\s+", " ", value).strip()
        if value and value not in vals:
            vals.append(value)
        if len(vals) >= limit:
            break
    return vals


def compare_semester(
    curriculum: list[dict[str, Any]],
    semester: int,
    n_semesters: int,
    dataset: str,
    model: dict[str, pd.DataFrame],
) -> dict[str, Any]:
    ours = semester_courses(curriculum, semester)
    ours_families = build_semester_family_set(ours)

    foreign = foreign_window_courses(model["catalog"], dataset, semester, n_semesters)
    foreign_families: set[str] = set()
    for _, r in foreign.iterrows():
        foreign_families |= set(split_family_ids(r))

    only_ours = ours_families - foreign_families
    only_foreign = foreign_families - ours_families
    common = ours_families & foreign_families

    gaps = model["gaps"]
    fam_names = dict(zip(gaps["ID_Familia"].astype(str), gaps["Familia_Curricular"].astype(str)))

    ours_course_ids = [
        str(c.get("source_course_id"))
        for c in ours
        if str(c.get("family_id") or "") in only_ours and c.get("source_course_id")
    ]
    foreign_course_ids = foreign[
        foreign["ID_Familia_Principal_Normalizada"].astype(str).isin(only_foreign)
    ]["ID_Asignatura"].astype(str).tolist()

    return {
        "window_low": round((semester - 1) / n_semesters * 100, 1),
        "window_high": round(semester / n_semesters * 100, 1),
        "ours_only": [(f, fam_names.get(f, f)) for f in sorted(only_ours)],
        "foreign_only": [(f, fam_names.get(f, f)) for f in sorted(only_foreign)],
        "common": [(f, fam_names.get(f, f)) for f in sorted(common)],
        "ours_contents": content_samples_for_course_ids(model["contents"], ours_course_ids),
        "foreign_contents": content_samples_for_course_ids(model["contents"], foreign_course_ids),
        "foreign_courses": foreign,
    }


def curriculum_dataframe(curriculum: list[dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for c in curriculum:
        rows.append({
            "Semestre": c.get("semester"),
            "Código": c.get("code"),
            "Asignatura": c.get("name"),
            "Créditos EPN propuestos": c.get("credits"),
            "Área AGRO-NORM": c.get("area"),
            "Familia AGRO-NORM": c.get("family"),
            "Fuente/Origen": c.get("origin"),
            "Dataset fuente": c.get("source_dataset"),
            "ID fuente": c.get("source_course_id"),
            "Profundidad N": c.get("depth_n"),
            "Experimental": c.get("experimental"),
            "Computacional": c.get("computational"),
            "Estado normalización": c.get("normalization_status"),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["Semestre", "Asignatura"], na_position="last")
    return df


def export_excel_bytes(curriculum: list[dict[str, Any]], n_semesters: int, model: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    malla = curriculum_dataframe(curriculum)

    summaries = []
    for sem in range(1, n_semesters + 1):
        metrics = semester_quality_metrics(curriculum, sem, model["gaps"])
        scores = semester_axis_scores(curriculum, sem)
        row = {"Semestre": sem, **metrics, **scores}
        summaries.append(row)
    summary = pd.DataFrame(summaries)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        malla.to_excel(writer, index=False, sheet_name="Malla_Propuesta")
        summary.to_excel(writer, index=False, sheet_name="Resumen_Semestres")
    output.seek(0)
    return output.read()


def format_metric(value: Any, suffix: str = "") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "NR"
    if isinstance(value, (float, np.floating)):
        return f"{value:.1f}{suffix}"
    return f"{value}{suffix}"


def family_badges(courses: list[dict[str, Any]], gaps: pd.DataFrame) -> pd.DataFrame:
    fam_ids = build_semester_family_set(courses)
    if not fam_ids:
        return pd.DataFrame()
    cols = [
        "ID_Familia", "Familia_Curricular", "Estatus_Benchmark",
        "Senal_Emergente", "Presencia_Core_pct", "Profundidad_Benchmark",
        "Prioridad_Stakeholders", "Accion_Preliminar", "Prioridad_Revision",
        "Justificacion_Accion",
    ]
    cols = [c for c in cols if c in gaps.columns]
    out = gaps[gaps["ID_Familia"].astype(str).isin(fam_ids)][cols].copy()
    if "Presencia_Core_pct" in out:
        out["Presencia Core"] = (out["Presencia_Core_pct"] * 100).round(1)
        out = out.drop(columns=["Presencia_Core_pct"])
    return out


def add_custom_course(
    name: str,
    semester: int,
    credits: float,
    family_row: pd.Series,
) -> dict[str, Any]:
    return {
        "instance_id": str(uuid.uuid4()),
        "source_course_id": None,
        "source_dataset": "CUSTOM_EPN",
        "code": "",
        "name": name.strip(),
        "semester": int(semester),
        "credits": int(round(float(credits))),
        "course_type": derive_course_type_from_row({"area_id": family_row.get("ID_Area", ""), "family_id": family_row.get("ID_Familia", "")}),
        "area_id": family_row.get("ID_Area", ""),
        "area": family_row.get("Area", ""),
        "family_id": family_row.get("ID_Familia", ""),
        "family": family_row.get("Familia_Curricular", ""),
        "family_ids": [family_row.get("ID_Familia", "")],
        "depth_n": None,
        "experimental": "",
        "computational": "",
        "lab_explicit": False,
        "experience_type": "",
        "source_credits": None,
        "source_credit_system": "",
        "origin": "Materia propia EPN",
        "normalization_status": "DERIVADO_USUARIO",
    }


# ---------------------------- UI helpers ----------------------------

def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {padding-top: .7rem; padding-bottom: 1.2rem; max-width: 98%;}
        h1 {font-size: 1.75rem !important; margin-bottom: .2rem !important;}
        h2 {font-size: 1.28rem !important;}
        h3 {font-size: 1.02rem !important; margin: .1rem 0 !important;}
        [data-testid="stMetric"] {background: rgba(127,127,127,.055); border: 1px solid rgba(127,127,127,.14); padding: .42rem .6rem; border-radius: .6rem;}
        [data-testid="stMetricValue"] {font-size: 1.15rem;}
        div[data-testid="stVerticalBlockBorderWrapper"] {border-radius: .55rem;}
        div.stButton > button {min-height: 1.75rem; padding: .1rem .35rem; font-size: .78rem;}
        .agro-note {border-left: 4px solid #888; padding: .45rem .7rem; background: rgba(127,127,127,.06); border-radius: .25rem;}
        .course-mini {border-left: 7px solid var(--course-color); background: rgba(127,127,127,.045); border-radius: .42rem; padding: .38rem .48rem .28rem .48rem; margin-bottom: .18rem; line-height: 1.12;}
        .course-mini .name {font-size: .82rem; font-weight: 700; margin-bottom: .14rem;}
        .course-mini .meta {font-size: .66rem; opacity: .78;}
        .semester-head {display:flex; justify-content:space-between; align-items:center; margin-bottom:.2rem;}
        .semester-head strong {font-size:.95rem;}
        .semester-head span {font-size:.68rem; opacity:.75;}
        .type-legend {display:flex; flex-wrap:wrap; gap:.28rem .6rem; margin:.2rem 0 .65rem 0;}
        .type-pill {font-size:.68rem; display:inline-flex; align-items:center; gap:.25rem;}
        .type-dot {width:.7rem; height:.7rem; border-radius:50%; display:inline-block;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_type_legend() -> None:
    pills = []
    for name in COURSE_TYPES:
        if name == "Otros":
            continue
        pills.append(f'<span class="type-pill"><span class="type-dot" style="background:{TYPE_COLORS[name]}"></span>{name}</span>')
    st.markdown('<div class="type-legend">' + ''.join(pills) + '</div>', unsafe_allow_html=True)


def course_card(course: dict[str, Any], n_semesters: int) -> None:
    name = course.get("name", "Materia")
    code = course.get("code") or ""
    credits = int(round(float(course.get("credits") or 0)))
    ctype = course.get("course_type") or derive_course_type_from_row(course)
    color = TYPE_COLORS.get(ctype, TYPE_COLORS["Otros"])
    family = course.get("family") or "Sin familia"
    source = course.get("source_dataset") or "—"
    code_prefix = f"{code} · " if code else ""
    html = (
        f'<div class="course-mini" style="--course-color:{color}">'
        f'<div class="name">{name}</div>'
        f'<div class="meta">{code_prefix}{credits} cr · {ctype}</div>'
        f'<div class="meta">{family} · {source}</div>'
        '</div>'
    )
    st.markdown(html, unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3, gap="small")
    sem = int(course.get("semester") or 1)
    if c1.button("←", key=f"left_{course['instance_id']}", disabled=sem <= 1, help="Semestre anterior", use_container_width=True):
        course["semester"] = sem - 1
        st.rerun()
    if c2.button("×", key=f"rm_{course['instance_id']}", help="Retirar", use_container_width=True):
        st.session_state.curriculum = [c for c in st.session_state.curriculum if c["instance_id"] != course["instance_id"]]
        st.rerun()
    if c3.button("→", key=f"right_{course['instance_id']}", disabled=sem >= n_semesters, help="Semestre siguiente", use_container_width=True):
        course["semester"] = sem + 1
        st.rerun()


def analysis_interpretation_guide() -> None:
    with st.expander("ℹ️ Cómo interpretar radar, barras e indicadores", expanded=False):
        st.markdown(
            """
### ¿Qué significa cada eje del radar?

- **Bases científico-cuantitativas:** representa la intensidad de matemáticas, estadística, física, química, biología y otros fundamentos cuantitativos que sostienen el razonamiento ingenieril.
- **Ingeniería y procesos:** representa balances, termodinámica, fenómenos de transporte, operaciones, diseño, simulación y transformación de procesos agroindustriales.
- **Agroalimentos, calidad y producto:** representa materias primas, tecnología y procesamiento de alimentos, tecnologías sectoriales, inocuidad, calidad, conservación y desarrollo de producto.
- **Digitalización y automatización:** representa programación, análisis de datos, inteligencia artificial, modelado computacional, instrumentación, control, automatización e Industria 4.0.
- **Sostenibilidad y circularidad:** representa ambiente, análisis de ciclo de vida, residuos, valorización, economía circular, energía, bioenergía y producción sostenible.
- **Práctica e integración:** representa laboratorios, talleres, proyectos integradores, práctica preprofesional, capstone y otras experiencias en las que el estudiante aplica e integra conocimientos.
- **Gestión, innovación y profesión:** representa administración, economía, gestión de proyectos, emprendimiento, innovación, I+D, comunicación, ética y competencias profesionales/sociales.

### ¿Cómo leer los gráficos e indicadores?

**Radar de equilibrio curricular.** Cada eje expresa qué proporción de los créditos del semestre contribuye a ese dominio. Una materia puede aportar a más de un eje; por eso los ejes **no tienen que sumar 100%**. Una superficie mayor no significa automáticamente "mejor": permite identificar fortalezas, concentraciones y posibles vacíos.

**Barras de ganancias/pérdidas.** Muestran la diferencia en puntos entre la propuesta y el **baseline EPN del mismo semestre**. Valores positivos indican mayor intensidad relativa; valores negativos, menor intensidad. Deben interpretarse como *trade-offs*, no como una calificación.

**Exposición práctica/digital.** Porcentaje de créditos asociados a experiencias explícitamente experimentales/laboratorio o computacionales/digitales según la evidencia disponible.

**Presencia Core media.** Prevalencia internacional media de las familias AGRO-NORM incluidas. Un valor alto indica que esas familias son comunes entre referentes Core; no determina por sí solo su pertinencia local.

**Prioridad stakeholders.** Señal agregada de prioridad derivada de graduados, empleadores y otras capas de evidencia. `NR` significa evidencia insuficiente, nunca cero.

**Profundidad vs benchmark.** Diferencia en niveles N entre la profundidad propuesta y la referencia internacional de la familia. Positivo = mayor profundidad; negativo = menor. Debe leerse junto con créditos, ubicación y secuencia curricular.
            """
        )

def type_interpretation_guide() -> None:
    with st.expander("ℹ️ Cómo interpretar la comparación por tipos", expanded=False):
        st.markdown(
            """
La tipología por colores es una **vista ejecutiva derivada** de AGRO-NORM; no reemplaza la clasificación auditada.

- **Nuestra propuesta:** porcentaje de los créditos EPN del semestre en cada tipo.
- **Referente internacional:** porcentaje de asignaturas del referente en la misma ventana de avance curricular. Se usa conteo relativo porque los sistemas de créditos internacionales no son directamente equivalentes a los créditos EPN.
- **Diferencia pp:** puntos porcentuales de composición. `+10 pp` significa mayor presencia relativa de ese tipo en nuestra propuesta; `−10 pp`, menor presencia relativa.

El radar sirve para comparar **forma/composición**, no para afirmar que una universidad es globalmente mejor que otra.
            """
        )


def comparison_interpretation_guide() -> None:
    with st.expander("ℹ️ Cómo interpretar la comparación internacional", expanded=False):
        st.markdown(
            """
La comparación usa **avance curricular relativo (%)**, no el número nominal de semestre. Esto permite comparar programas de distinta duración.

- **Familias comunes:** dominios AGRO-NORM presentes en ambos lados de la ventana curricular.
- **Cobertura exclusiva propuesta:** familias que nuestra propuesta cubre en esa etapa y el referente no muestra en la evidencia estructurada disponible.
- **Cobertura exclusiva referente:** familias que el referente cubre en esa etapa y nuestra propuesta no.
- **Microcontenido explícito:** solo aparece cuando fue recuperado directamente de una fuente. Si no hay microcontenido registrado, el dashboard muestra falta de evidencia y **no interpreta ausencia real de enseñanza**.

Las diferencias son señales para discusión curricular; no son, por sí solas, recomendaciones automáticas de añadir o eliminar materias.
            """
        )


def evidence_interpretation_guide() -> None:
    with st.expander("ℹ️ Cómo interpretar CAEE y EUR-ACE", expanded=False):
        st.markdown(
            """
**CAEE-NORM** resume temas provenientes de evidencia externa. El número de evidencias se muestra como trazabilidad y no se transforma automáticamente en un peso de decisión.

**EUR-ACE:** el dashboard conserva el **estado oficial** del criterio y calcula por separado una **contribución potencial** de las materias del semestre. Añadir una asignatura no cambia automáticamente un criterio de `Parcialmente` a `Cumple`.
            """
        )


def ensure_state(model: dict[str, pd.DataFrame], n_semesters: int) -> None:
    if "baseline_curriculum" not in st.session_state:
        st.session_state.baseline_curriculum = initialize_epn_curriculum(model, n_semesters)
    if "curriculum" not in st.session_state:
        st.session_state.curriculum = json.loads(json.dumps(st.session_state.baseline_curriculum, default=str))
    for bucket in ["baseline_curriculum", "curriculum"]:
        for course in st.session_state.get(bucket, []):
            course["credits"] = int(round(float(course.get("credits") or 0)))
            if not course.get("course_type"):
                course["course_type"] = derive_course_type_from_row(course)


# =============================== APP ===============================

inject_css()

st.sidebar.title("AGRO · Datos")
data_dir = st.sidebar.text_input(
    "Carpeta con los cinco Excel",
    value="data",
    help="Coloque los cinco libros aprobados dentro de esta carpeta.",
)
n_semesters = int(st.sidebar.number_input("Número de semestres de la propuesta", min_value=6, max_value=12, value=9, step=1))

try:
    model = load_model(data_dir)
except Exception as exc:
    st.error("No se pudo cargar el modelo.")
    st.code(str(exc))
    st.info(
        "Cree una carpeta `data/` junto al script y copie allí los cinco Excel con sus nombres originales."
    )
    st.stop()

ensure_state(model, n_semesters)

st.sidebar.success("5/5 fuentes cargadas")
st.sidebar.caption(
    f"Catálogo: {len(model['catalog']):,} asignaturas · "
    f"Brechas: {len(model['gaps']):,} familias · "
    f"Temas CAEE: {len(model['caee_topics']):,}"
)

if st.sidebar.button("Restaurar baseline EPN", use_container_width=True):
    st.session_state.curriculum = json.loads(json.dumps(st.session_state.baseline_curriculum, default=str))
    st.rerun()

if st.sidebar.button("Vaciar propuesta", use_container_width=True):
    st.session_state.curriculum = []
    st.rerun()

# Export
excel_bytes = export_excel_bytes(st.session_state.curriculum, n_semesters, model)
st.sidebar.download_button(
    "Exportar propuesta a Excel",
    data=excel_bytes,
    file_name="Malla_AGRO_Propuesta_Dashboard.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

json_bytes = json.dumps(st.session_state.curriculum, ensure_ascii=False, indent=2, default=str).encode("utf-8")
st.sidebar.download_button(
    "Guardar estado JSON",
    data=json_bytes,
    file_name="Malla_AGRO_estado.json",
    mime="application/json",
    use_container_width=True,
)

st.title("Constructor curricular interactivo · Ingeniería Agroindustrial")
st.caption(
    "Modelo simbólico-relacional en Pandas/Streamlit. "
    "EPN permanece fuera de los denominadores del benchmark; NR nunca se interpreta como cero."
)

tab_builder, tab_analysis, tab_types, tab_compare, tab_evidence, tab_model = st.tabs([
    "🧩 Constructor",
    "📊 Análisis por semestre",
    "🎨 Análisis por tipos",
    "🌍 Comparación internacional",
    "🎯 Evidencia / EUR-ACE",
    "🧱 Modelo de datos",
])

# --------------------------- Constructor ---------------------------
with tab_builder:
    st.subheader("Biblioteca de asignaturas")
    catalog = model["catalog"].copy()

    include_review = st.checkbox(
        "Mostrar también registros AGRO-NORM en REVISAR",
        value=True,
        help=(
            "Activado: el buscador incluye APROBADO + REVISAR para no ocultar asignaturas potencialmente útiles. "
            "Desactivado: muestra solo registros con normalización APROBADA. Un registro REVISAR no significa que la materia sea incorrecta; "
            "significa que su clasificación AGRO-NORM requiere validación."
        ),
    )
    if not include_review and "Estado_Normalizacion_Final" in catalog:
        catalog = catalog[catalog["Estado_Normalizacion_Final"].astype(str).str.upper().eq("APROBADO")]

    st.caption(f"Buscador activo sobre {len(catalog):,} registros curriculares.")

    add_mode = st.radio(
        "Tipo de incorporación",
        ["Buscar en referentes", "Crear materia propia EPN"],
        horizontal=True,
        help="Use referentes para reutilizar evidencia existente; use materia propia EPN cuando la asignatura propuesta aún no existe como registro fuente.",
    )

    if add_mode == "Buscar en referentes":
        selected_row = searchbox_course_selector(catalog)

        if selected_row is not None:
            selected_course_details(selected_row, model["contents"])
            a, b, c = st.columns([2, 1, 1])
            with a:
                st.write(f"**Familia:** {selected_row.get('Familia_Principal_Normalizada', 'NR')}")
                st.caption(
                    f"Origen: {course_origin_label(selected_row)} [{selected_row.get('Codigo_Dataset_Canonico', 'NR')}] · "
                    f"Área: {selected_row.get('Area_Principal_Normalizada', 'NR')} · "
                    f"Tipo: {derive_course_type_from_row(selected_row)}"
                )
            with b:
                add_sem = int(st.number_input("Semestre destino", 1, n_semesters, 1, step=1, key="add_sem_ref", help="Semestre en el que se ubicará la materia dentro de la propuesta editable."))
            with c:
                source_cr = selected_row.get("creditos_fuente", np.nan)
                default_cr = int(round(float(source_cr))) if (selected_row.get("Codigo_Dataset_Canonico") == "EPN" and not pd.isna(source_cr)) else 3
                default_cr = min(10, max(1, default_cr))
                add_cr = int(st.number_input("Créditos EPN", 1, 10, default_cr, 1, key="add_cr_ref", help="Créditos enteros de diseño EPN. No se convierten automáticamente desde ECTS u otros sistemas."))
            st.caption("Los créditos extranjeros no se convierten automáticamente; se asignan como decisión de diseño EPN.")
            if st.button("Agregar a la propuesta", type="primary"):
                st.session_state.curriculum.append(row_to_course(selected_row, semester=add_sem, credits=add_cr, origin=f"Referencia {selected_row.get('Codigo_Dataset_Canonico', '')}"))
                st.rerun()

    else:
        gaps = model["gaps"].copy()
        family_options = (gaps["ID_Familia"].astype(str) + " · " + gaps["Familia_Curricular"].astype(str)).tolist()
        custom_name = st.text_input("Nombre de la nueva asignatura")
        family_choice = st.selectbox("Familia AGRO-NORM principal", family_options, index=None)
        c1, c2 = st.columns(2)
        custom_sem = int(c1.number_input("Semestre", 1, n_semesters, 1, step=1, key="custom_sem"))
        custom_cr = int(c2.number_input("Créditos", 1, 10, 3, 1, key="custom_cr"))
        if family_choice and custom_name.strip():
            fam_id = family_choice.split(" · ", 1)[0]
            fam_row = gaps[gaps["ID_Familia"].astype(str).eq(fam_id)].iloc[0]
            st.caption(f"Área: {fam_row.get('Area', 'NR')} · Acción: {fam_row.get('Accion_Preliminar', 'NR')} · Prioridad: {fam_row.get('Prioridad_Revision', 'NR')}")
            if st.button("Crear e incorporar", type="primary"):
                st.session_state.curriculum.append(add_custom_course(custom_name, custom_sem, custom_cr, fam_row))
                st.rerun()

    st.divider()
    h1, h2 = st.columns([3, 1])
    h1.subheader("Malla editable")
    h2.caption("← mover · × retirar · → mover")
    render_type_legend()

    semesters_per_row = 5 if n_semesters >= 8 else 4
    for start in range(1, n_semesters + 1, semesters_per_row):
        cols = st.columns(semesters_per_row, gap="small")
        for offset, col in enumerate(cols):
            sem = start + offset
            if sem > n_semesters:
                continue
            courses = semester_courses(st.session_state.curriculum, sem)
            credits = sum(int(round(float(c.get("credits") or 0))) for c in courses)
            with col:
                with st.container(border=True):
                    status = "✓" if credits == 15 else ("!" if credits > 15 else "")
                    st.markdown(f'<div class="semester-head"><strong>S{sem}</strong><span>{len(courses)} mat · {credits} cr {status}</span></div>', unsafe_allow_html=True)
                    for course in sorted(courses, key=lambda x: x.get("name", "")):
                        course_card(course, n_semesters)


# ------------------------ Análisis por semestre ------------------------
with tab_analysis:
    semester = st.select_slider(
        "Semestre a analizar",
        options=list(range(1, n_semesters + 1)),
        value=1,
        key="analysis_sem",
        help="Selecciona el semestre de la propuesta para calcular indicadores y compararlo con el baseline EPN del mismo periodo.",
    )
    analysis_interpretation_guide()

    curr_scores = semester_axis_scores(st.session_state.curriculum, semester)
    base_scores = semester_axis_scores(st.session_state.baseline_curriculum, semester)
    metrics = semester_quality_metrics(st.session_state.curriculum, semester, model["gaps"])

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Créditos", format_metric(metrics["credits"]), help="Suma de créditos EPN propuestos en el semestre.")
    m2.metric("Exposición práctica", format_metric(metrics["practice_pct"], "%"), help="Porcentaje de créditos con laboratorio, componente experimental o experiencia explícitamente práctica.")
    m3.metric("Exposición digital", format_metric(metrics["digital_pct"], "%"), help="Porcentaje de créditos asociados a componente computacional/digital o familias explícitas de datos, automatización e Industria 4.0.")
    m4.metric("Presencia Core media", format_metric(metrics["benchmark_core"], "%"), help="Prevalencia internacional media de las familias AGRO-NORM del semestre entre referentes Core.")
    m5.metric("Prioridad stakeholders", format_metric(metrics["stakeholder_priority"], "/100"), help="Prioridad agregada de señales externas disponibles para las familias del semestre. NR = no reportable, no cero.")

    c1, c2 = st.columns(2)
    c1.plotly_chart(radar_figure(curr_scores, base_scores), use_container_width=True)
    c2.plotly_chart(delta_figure(curr_scores, base_scores), use_container_width=True)

    st.markdown("#### Lectura técnica del semestre")
    q1, q2, q3 = st.columns(3)
    q1.metric("Peso en familias emergentes", format_metric(metrics["emerging_pct"], "%"))
    q2.metric(
        "Profundidad vs benchmark",
        "NR" if metrics["depth_delta"] is None else f"{metrics['depth_delta']:+.2f} N",
        help="Promedio ponderado: profundidad máxima propuesta menos profundidad benchmark de la familia.",
    )
    q3.metric(
        "Materias del semestre",
        metrics["courses"],
    )

    fam_table = family_badges(
        semester_courses(st.session_state.curriculum, semester),
        model["gaps"],
    )
    if not fam_table.empty:
        st.dataframe(fam_table, use_container_width=True, hide_index=True)
    else:
        st.info("No hay familias AGRO-NORM asignadas en este semestre.")

# -------------------------- Análisis por tipos --------------------------
with tab_types:
    st.subheader("Composición por tipos de materia")
    st.caption(
        "La tipología es una vista ejecutiva derivada de AGRO-NORM. Las optativas se identifican primero; "
        "el resto se clasifica por área/familia. La comparación internacional usa la misma ventana de avance curricular."
    )
    render_type_legend()
    type_interpretation_guide()

    sem_type = st.select_slider("Semestre de la propuesta", options=list(range(1, n_semesters + 1)), value=1, key="type_sem", help="Compara la composición del semestre seleccionado con la etapa curricular equivalente del referente.")
    universes_t = model["universes"].copy()
    if "Grupo" in universes_t:
        preferred_t = universes_t[universes_t["Grupo"].astype(str).str.contains("CORE", na=False)]
        if preferred_t.empty:
            preferred_t = universes_t
    else:
        preferred_t = universes_t

    type_labels = {
        f"{r.get('Codigo_Dataset')} · {r.get('Programa')}": str(r.get("Codigo_Dataset"))
        for _, r in preferred_t.iterrows()
    }
    type_ref_label = st.selectbox("Referente", list(type_labels.keys()), index=0 if type_labels else None, key="type_ref", help="Currículo internacional usado para comparar la distribución relativa por tipos de materia.")

    ours_courses = semester_courses(st.session_state.curriculum, sem_type)
    ours_types = type_distribution_from_courses(ours_courses)

    if type_ref_label:
        type_dataset = type_labels[type_ref_label]
        foreign_df = foreign_window_courses(model["catalog"], type_dataset, sem_type, n_semesters)
        foreign_types = type_distribution_from_foreign(foreign_df)

        c1, c2 = st.columns([1.15, 1])
        c1.plotly_chart(type_radar_figure(ours_types, foreign_types, type_dataset), use_container_width=True)

        merged_t = ours_types[["Tipo", "Porcentaje"]].rename(columns={"Porcentaje": "Nuestra propuesta"}).merge(
            foreign_types[["Tipo", "Porcentaje"]].rename(columns={"Porcentaje": type_dataset}),
            on="Tipo", how="outer",
        ).fillna(0)
        long_t = merged_t.melt(id_vars="Tipo", var_name="Malla", value_name="Porcentaje")
        fig_types = px.bar(
            long_t,
            x="Porcentaje", y="Tipo", color="Malla", barmode="group", orientation="h",
            title="Distribución del semestre por tipo",
            category_orders={"Tipo": list(reversed(COURSE_TYPES[:-1]))},
        )
        fig_types.update_layout(height=470, margin=dict(l=20, r=20, t=50, b=20))
        c2.plotly_chart(fig_types, use_container_width=True)

        merged_t["Diferencia pp"] = merged_t["Nuestra propuesta"] - merged_t[type_dataset]
        st.markdown("#### Diferencias de composición")
        st.dataframe(merged_t.sort_values("Diferencia pp", ascending=False).round(1), use_container_width=True, hide_index=True)

        strongest = merged_t.sort_values("Diferencia pp", ascending=False).iloc[0]
        weakest = merged_t.sort_values("Diferencia pp", ascending=True).iloc[0]
        a, b = st.columns(2)
        a.success(f"Mayor fortaleza relativa: {strongest['Tipo']} ({strongest['Diferencia pp']:+.1f} pp)")
        b.warning(f"Menor presencia relativa: {weakest['Tipo']} ({weakest['Diferencia pp']:+.1f} pp)")
    else:
        st.info("Seleccione un referente para comparar la composición por tipos.")


# ---------------------- Comparación internacional ----------------------
with tab_compare:
    semester_cmp = st.select_slider(
        "Semestre de la propuesta",
        options=list(range(1, n_semesters + 1)),
        value=1,
        key="compare_sem",
        help="El semestre se transforma en una ventana de avance curricular relativo para compararlo con programas de distinta duración.",
    )
    comparison_interpretation_guide()

    universes = model["universes"].copy()
    if "Grupo" in universes:
        default_univ = universes[universes["Grupo"].astype(str).str.contains("CORE", na=False)]
        if default_univ.empty:
            default_univ = universes
    else:
        default_univ = universes

    universe_labels = {
        f"{r.get('Codigo_Dataset')} · {r.get('Programa')}": str(r.get("Codigo_Dataset"))
        for _, r in default_univ.iterrows()
    }
    selected_univ_label = st.selectbox(
        "Universidad / currículo de comparación",
        list(universe_labels.keys()),
        index=0 if universe_labels else None,
        help="Selecciona un currículo de referencia. La comparación se hace por etapa relativa, no por número nominal de semestre.",
    )

    if selected_univ_label:
        dataset = universe_labels[selected_univ_label]
        comp = compare_semester(
            st.session_state.curriculum,
            semester_cmp,
            n_semesters,
            dataset,
            model,
        )
        st.info(
            f"Comparación por avance curricular equivalente: "
            f"{comp['window_low']}%–{comp['window_high']}% del programa. "
            "No se equipara el número nominal de semestre."
        )

        a, b, c = st.columns(3)
        a.metric("Familias comunes", len(comp["common"]), help="Familias AGRO-NORM presentes tanto en nuestra propuesta como en el referente dentro de la misma etapa relativa.")
        b.metric("Cobertura exclusiva propuesta", len(comp["ours_only"]), help="Familias presentes en nuestra propuesta y no observadas en el referente dentro de esa ventana curricular. No implica superioridad automática.")
        c.metric("Cobertura exclusiva referente", len(comp["foreign_only"]), help="Familias observadas en el referente y no presentes en nuestra propuesta dentro de esa ventana curricular. Son candidatas a revisión, no adiciones automáticas.")

        left, right = st.columns(2)
        with left:
            st.markdown("#### Lo que gana / conserva nuestra propuesta")
            if comp["ours_only"]:
                st.dataframe(
                    pd.DataFrame(comp["ours_only"], columns=["ID", "Familia"]),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.caption("Sin familias diferenciales en los datos disponibles.")
            if comp["ours_contents"]:
                with st.expander(f"Ver microcontenido explícito ({len(comp['ours_contents'])})", expanded=False):
                    st.caption("Se muestran únicamente contenidos recuperados explícitamente de las fuentes disponibles.")
                    for item in comp["ours_contents"]:
                        st.write("•", item)
            else:
                st.caption("Sin microcontenido explícito recuperado para esas diferencias.")

        with right:
            st.markdown("#### Lo que aparece en el referente y no en la propuesta")
            if comp["foreign_only"]:
                st.dataframe(
                    pd.DataFrame(comp["foreign_only"], columns=["ID", "Familia"]),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.caption("Sin familias diferenciales en los datos disponibles.")
            if comp["foreign_contents"]:
                with st.expander(f"Ver microcontenido explícito ({len(comp['foreign_contents'])})", expanded=False):
                    st.caption("Se muestran únicamente contenidos recuperados explícitamente de las fuentes disponibles.")
                    for item in comp["foreign_contents"]:
                        st.write("•", item)
            else:
                st.caption(
                    "No hay microcontenido explícito disponible. "
                    "Esto NO se interpreta como ausencia real de contenidos."
                )

        with st.expander("Ver asignaturas del referente dentro de la misma ventana de avance"):
            cols = [
                c for c in [
                    "Nombre_espanol",
                    "Familia_Principal_Normalizada",
                    "Avance_Curricular_Medio_pct",
                    "Nivel_Profundidad",
                    "Componente_Experimental_Normalizado",
                    "Componente_Computacional_Normalizado",
                ]
                if c in comp["foreign_courses"].columns
            ]
            st.dataframe(comp["foreign_courses"][cols], use_container_width=True, hide_index=True)

# ----------------------- Evidencia / EUR-ACE -----------------------
with tab_evidence:
    semester_ev = st.select_slider(
        "Semestre",
        options=list(range(1, n_semesters + 1)),
        value=1,
        key="evidence_sem",
        help="Semestre de la propuesta cuya contribución potencial a temas CAEE y criterios EUR-ACE se analizará.",
    )
    evidence_interpretation_guide()

    topics = semester_topics(st.session_state.curriculum, semester_ev)
    coverage = model["caee_coverage"].copy()
    topic_name_col = "Tema" if "Tema" in coverage.columns else "Tema_normalizado"
    topic_table = coverage[coverage["Tema_ID"].astype(str).isin(topics)].copy()

    st.markdown("### Temas CAEE a los que contribuye el semestre")
    st.caption(
        "Puente analítico DERIVADO desde AGRO-NORM hacia CAEE-NORM. "
        "La cobertura de CAEE es descriptiva; N de evidencias no se usa como ponderación automática."
    )
    if not topic_table.empty:
        show_cols = [
            c for c in [
                "Tema_ID", topic_name_col, "N_evidencias_elegibles",
                "EUR_ACE", "Seguimiento", "Encuesta_131", "Focus_Group",
                "Mercado_Laboral", "N_corrientes",
            ]
            if c in topic_table.columns
        ]
        st.dataframe(topic_table[show_cols], use_container_width=True, hide_index=True)
    else:
        st.info("No se identificaron temas CAEE vinculados a las materias del semestre.")

    st.markdown("### EUR-ACE: estado oficial vs contribución de la propuesta")
    eur = eurace_alignment(st.session_state.curriculum, semester_ev, model["eurace"])
    st.dataframe(eur, use_container_width=True, hide_index=True)
    st.markdown(
        '<div class="agro-note"><b>Regla:</b> el dashboard nunca recalifica el criterio EUR-ACE. '
        'Solo indica si el diseño curricular propuesto aporta evidencia potencial a un criterio sensible al currículo.</div>',
        unsafe_allow_html=True,
    )

    st.markdown("### Señales de stakeholders relacionadas con las familias del semestre")
    sem_fams = build_semester_family_set(semester_courses(st.session_state.curriculum, semester_ev))
    sig = model["stakeholder_signals"].copy()
    if "Familias_relacionadas" in sig.columns:
        mask = sig["Familias_relacionadas"].fillna("").astype(str).apply(
            lambda txt: any(f in {x.strip() for x in txt.split(";")} for f in sem_fams)
        )
        sig = sig[mask]
    if not sig.empty:
        show_cols = [
            c for c in ["Tema", "Fuente", "Indicador", "Valor", "Prioridad", "Implicacion", "Familias_relacionadas"]
            if c in sig.columns
        ]
        st.dataframe(sig[show_cols], use_container_width=True, hide_index=True)
    else:
        st.caption("No hay señales stakeholder explícitamente vinculadas a estas familias.")

# --------------------------- Modelo de datos ---------------------------
with tab_model:
    st.subheader("Modelo relacional simbólico mínimo")
    st.code(
        """
DIM_CURSO  (Base Maestra / 09_Asignaturas)
   PK ID_Asignatura
   FK Codigo_Dataset_Canonico
   FK ID_Area_Principal_Normalizada
   FK ID_Familia_Principal_Normalizada
       │
       ├── 1:N ── FACT_CONTENIDO (11_Contenidos)
       │
       └── N:1 ── DIM_FAMILIA / AGRO-NORM
                       │
                       ├── 1:1 ── FACT_BRECHA (Matriz Brechas / 02_Cobertura_Familias)
                       │
                       └── DERIVADO ── BRIDGE_FAMILIA_TEMA_CAEE
                                           │
                                           └── N:1 DIM_TEMA_CAEE

DIM_CURRICULO_REFERENTE (Matriz Comparativa / 02_Universos)
       │
       ├── FACT_PRESENCIA_FAMILIA
       ├── FACT_POSICION_FAMILIA
       └── FACT_PROFUNDIDAD_FAMILIA

FACT_EPN_BASELINE
   EPN / 04_Asignaturas_60
   + EPN / 05_Carga_52
   + EPN / 08_Prerrequisitos

FACT_EURACE
   CAEE / 03_EURACE_Criterios
       │
       └── estado oficial (inmutable en dashboard)

FACT_PROPUESTA
   session_state de Streamlit
   [instance_id, source_course_id, semestre, créditos EPN, familia, área, origen]
        """.strip(),
        language="text",
    )

    st.markdown("#### Campos realmente utilizados")
    used = pd.DataFrame([
        ["Base Maestra", "09_Asignaturas", "nombre, universidad, familia/área, profundidad, laboratorio, computación, avance %, estado normalización", "Catálogo y comparación"],
        ["Base Maestra", "11_Contenidos", "ID_Asignatura + contenido explícito", "Diferencias cualitativas"],
        ["EPN", "04_Asignaturas_60 + 05_Carga_52", "semestre, créditos, estructura y carga", "Baseline"],
        ["EPN", "08_Prerrequisitos", "relación asignatura→requisito", "Validación futura de secuencia"],
        ["Matriz Comparativa", "02_Universos", "currículo, grupo, nivel", "Selector de referentes"],
        ["Matriz Comparativa", "04/06/07", "presencia, posición, profundidad", "Benchmark cuantitativo"],
        ["Matriz Brechas", "02_Cobertura_Familias", "brecha, acción, prioridad, benchmark, EPN", "Diagnóstico y scoring"],
        ["Matriz Brechas", "08_Senales_Stakeholders", "tema, indicador, prioridad, familias", "Evidencia local"],
        ["CAEE", "02/14", "taxonomía y cobertura por corrientes", "Convergencia de evidencia"],
        ["CAEE", "03_EURACE_Criterios", "criterio, calificación, comentario", "Alineación acreditación"],
    ], columns=["Libro", "Hoja", "Campos", "Uso"])
    st.dataframe(used, use_container_width=True, hide_index=True)

    st.markdown("#### Exclusiones deliberadas")
    st.write(
        "No se cargan en la interfaz hojas de manifiesto, hashes, auditorías de cierre, diccionarios completos, "
        "evidencias raw de 492 filas ni mapeos raw de 673 filas, porque no aportan decisión visual directa. "
        "Se preservan en los Excel como trazabilidad y pueden incorporarse después en un modo de auditoría."
    )