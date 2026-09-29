# -*- coding: utf-8 -*-
"""
KPI ENTREGA DE INFORMES — Seguimiento de cierre de backlog
ADEMINSAC · Contrato N° 3700002135 · Refinería La Pampilla

Modelo de datos (archivo canónico data/matriz_kpi_informes.xlsx)
  INFORMES : maestro, una fila por informe (G1 / G2 / G3)
  ENTREGAS : registro diario de entregas (FECHA · N° · GRUPO · GP_AD)
  PLAN     : meta diaria del plan de entrega
  NOTAS    : notas de conciliación generadas al importar
También importa y exporta el formato de seguimiento original
(INFORMES G1 G2 G3 · PENDIENTES G3 · Plan de entrega · Avance Diario).
"""
from __future__ import annotations

import base64
import io
import math
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# ════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="KPI Entrega de Informes | ADEMINSAC", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

DATA_DIR = Path(__file__).parent / "data"
DATA_FILE = DATA_DIR / "matriz_kpi_informes.xlsx"
TZ = ZoneInfo("America/Lima")
PREFIJO = "ADEMINSAC-FIAB-RLP"

COLS_INF = ["Código", "N° Informe", "Adicional", "Grupo", "Plan", "TAG / Circuito", "Restricción",
            "Fecha Ingreso", "Estado", "Fecha Entrega", "Tipo (GP_AD)", "Observaciones"]
COLS_ENT = ["Fecha", "N° Informe", "Código", "Grupo", "Tipo (GP_AD)", "Observaciones"]
COLS_PLAN = ["Fecha", "Meta Diaria"]

GRUPOS = ["G1", "G2", "G3"]
DESC_GRUPO = {"G1": "Pendientes hasta Inf. 1110", "G2": "Pendientes post Inf. 1110",
              "G3": "Pendientes con restricción (Repsol)"}
# Códigos GP_AD usados en «Avance Diario». Complete la descripción según su nomenclatura.
TIPOS_GPAD = {"EQ": "EQ", "L": "L", "E": "E", "NP": "NP", "PSV": "PSV"}
PLANES = ["Plan de equipos", "Plan de circuitos", "HORAS"]
RESTRICCIONES = ["CORREGIR PSAIM", "PENDIENTE DESCARGA PSAIM", "UT MANUAL", "INSPECCION",
                 "Pend. Isométrico (Circuitos)"]
OBS_MARCA = "Marcado como entregado en la matriz"
OBS_HIST = "sin detalle diario"

AZUL, AZUL_CL, NARANJA, GRIS, ROJO, VERDE_AZ = "#1F4E79", "#5DADE2", "#E67E22", "#A6ACAF", "#B03A2E", "#148F77"
COLOR_GRUPO = {"G1": AZUL, "G2": AZUL_CL, "G3": NARANJA, "Reentrega": GRIS}

st.markdown("""
<style>
[data-testid="stMetric"]{background:rgba(31,78,121,.06);border:1px solid rgba(31,78,121,.18);
  border-radius:10px;padding:10px 14px;}
[data-testid="stMetricLabel"] p{font-weight:600;}
.hdr{padding:.4rem 0 .2rem 0;border-bottom:3px solid #1F4E79;margin-bottom:.8rem;}
.hdr h1{font-size:1.65rem;margin:0;color:#1F4E79;} .hdr p{margin:.1rem 0 0 0;opacity:.75;}
</style>""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# FECHAS Y FERIADOS (PERÚ)
# ════════════════════════════════════════════════════════════════════
def hoy_lima() -> date:
    return datetime.now(TZ).date()


def _pascua(y: int) -> date:
    a, b, c = y % 19, y // 100, y % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    return date(y, (h + l - 7 * m + 114) // 31, ((h + l - 7 * m + 114) % 31) + 1)


@st.cache_data
def feriados_peru(y0: int, y1: int) -> np.ndarray:
    fijos = [(1, 1), (5, 1), (6, 7), (6, 29), (7, 23), (7, 28), (7, 29), (8, 6),
             (8, 30), (10, 8), (11, 1), (12, 8), (12, 9), (12, 25)]
    dias = []
    for y in range(y0, y1 + 1):
        dias += [date(y, m, d) for m, d in fijos]
        p = _pascua(y)
        dias += [p - timedelta(days=3), p - timedelta(days=2)]
    return np.array(dias, dtype="datetime64[D]")


HOL = feriados_peru(2018, hoy_lima().year + 2)


def es_habil(d) -> bool:
    return bool(np.is_busday(np.datetime64(pd.Timestamp(d).date()), holidays=HOL))


def habiles_entre(a: pd.Series, b: pd.Timestamp) -> pd.Series:
    out = pd.Series(np.nan, index=a.index, dtype="float")
    m = a.notna()
    if m.any():
        out[m] = np.busday_count(a[m].values.astype("datetime64[D]"),
                                 np.datetime64(pd.Timestamp(b).date()), holidays=HOL)
    return out


def a_fecha(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s).dt.normalize()
    num = pd.to_numeric(s, errors="coerce")
    out = pd.to_datetime(s.astype(str).where(s.notna(), None), errors="coerce", dayfirst=True, format="mixed")
    serial = num.notna() & num.between(20000, 80000)
    if serial.any():
        out[serial] = pd.to_datetime(num[serial], unit="D", origin="1899-12-30")
    return out.dt.normalize()


# ════════════════════════════════════════════════════════════════════
# CÓDIGOS DE INFORME
# ════════════════════════════════════════════════════════════════════
def normalizar_codigo(txt) -> str:
    t = re.sub(r"\s+", "", str(txt)).upper().replace("FLAB", "FIAB")   # corrige «FlAB»
    return "" if t in ("", "NAN", "NONE") else t


def numero_de(cod) -> float:
    m = re.search(r"RLP-(\d+)", str(cod))
    return float(m.group(1)) if m else np.nan


def construir_codigo(n, anio: int, adicional: bool = False) -> str:
    return f"{PREFIJO}-{int(n)}-{anio}" + ("-ADICIONAL" if adicional else "")


def _limpiar_texto(s: pd.Series) -> pd.Series:
    return (s.astype(object).where(s.notna(), "").astype(str).str.strip().str.lstrip("`")
            .replace({"nan": "", "None": "", "NaT": "", "<NA>": ""}))


# ════════════════════════════════════════════════════════════════════
# ESTANDARIZACIÓN
# ════════════════════════════════════════════════════════════════════
def std_inf(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in COLS_INF:
        if c not in df.columns:
            df[c] = None
    df = df[COLS_INF]
    for c in ["Fecha Ingreso", "Fecha Entrega"]:
        df[c] = a_fecha(df[c])
    for c in [c for c in COLS_INF if c not in ("N° Informe", "Fecha Ingreso", "Fecha Entrega")]:
        df[c] = _limpiar_texto(df[c])
    df["Código"] = df["Código"].map(normalizar_codigo)
    n = pd.to_numeric(df["N° Informe"], errors="coerce")
    falta = df["Código"].eq("") & n.notna()
    df.loc[falta, "Código"] = [
        construir_codigo(num, (f.year if pd.notna(f) else hoy_lima().year), ad == "Sí")
        for num, f, ad in zip(n[falta], df.loc[falta, "Fecha Ingreso"], df.loc[falta, "Adicional"])]
    df = df[df["Código"] != ""].copy()
    df["N° Informe"] = df["Código"].map(numero_de).fillna(n).astype("Int64")
    df["Adicional"] = np.where(df["Código"].str.contains("ADICIONAL"), "Sí", "No")
    df["Grupo"] = df["Grupo"].str.upper()
    df["Tipo (GP_AD)"] = df["Tipo (GP_AD)"].str.upper()
    est = df["Estado"].str.lower()
    df["Estado"] = np.where(est.str.startswith("entreg"), "Entregado",
                            np.where(est.eq("") & df["Fecha Entrega"].notna(), "Entregado", "Pendiente"))
    df = df.drop_duplicates("Código", keep="last")
    return df.sort_values(["Grupo", "N° Informe", "Código"]).reset_index(drop=True)


def std_ent(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in COLS_ENT:
        if c not in df.columns:
            df[c] = None
    df = df[COLS_ENT]
    df["Fecha"] = a_fecha(df["Fecha"])
    for c in ["Código", "Grupo", "Tipo (GP_AD)", "Observaciones"]:
        df[c] = _limpiar_texto(df[c])
    df["Código"] = df["Código"].map(normalizar_codigo)
    df["N° Informe"] = pd.to_numeric(df["N° Informe"], errors="coerce").fillna(df["Código"].map(numero_de)).astype("Int64")
    df["Grupo"] = df["Grupo"].str.upper()
    df["Tipo (GP_AD)"] = df["Tipo (GP_AD)"].str.upper()
    df = df[df["Fecha"].notna() & (df["N° Informe"].notna() | df["Código"].ne(""))]
    return df.sort_values(["Fecha"], kind="stable").reset_index(drop=True)


def std_plan(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in COLS_PLAN:
        if c not in df.columns:
            df[c] = None
    df = df[COLS_PLAN]
    df["Fecha"] = a_fecha(df["Fecha"])
    df["Meta Diaria"] = pd.to_numeric(df["Meta Diaria"], errors="coerce").fillna(0.0)
    df = df[df["Fecha"].notna()].groupby("Fecha", as_index=False)["Meta Diaria"].last()
    return df.sort_values("Fecha").reset_index(drop=True)


def resolver_codigo(inf: pd.DataFrame, n, grupo: str) -> str | None:
    if pd.isna(n):
        return None
    c = inf[inf["N° Informe"] == int(n)]
    if c.empty:
        return None
    for cond in [(c["Grupo"] == grupo) & (c["Estado"] != "Entregado"), c["Grupo"] == grupo,
                 c["Estado"] != "Entregado", pd.Series(True, index=c.index)]:
        s = c[cond]
        if len(s):
            return s.sort_values("Adicional").iloc[0]["Código"]      # prioriza el no adicional
    return None


def sincronizar(inf: pd.DataFrame, ent: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Completa códigos del registro diario y actualiza estado / fecha de entrega en el maestro."""
    inf, ent = std_inf(inf), std_ent(ent)
    nuevos = []
    for i in ent.index:
        cod = ent.at[i, "Código"]
        if not cod:
            cod = resolver_codigo(inf, ent.at[i, "N° Informe"], ent.at[i, "Grupo"])
            if cod is None:
                cod = construir_codigo(ent.at[i, "N° Informe"], ent.at[i, "Fecha"].year)
            ent.at[i, "Código"] = cod
        if cod not in set(inf["Código"]) and cod not in {r["Código"] for r in nuevos}:
            nuevos.append({"Código": cod, "Grupo": ent.at[i, "Grupo"] or "G1", "Fecha Ingreso": ent.at[i, "Fecha"],
                           "Observaciones": "Incorporado desde el registro de entregas (no figuraba en las listas)"})
    if nuevos:
        inf = std_inf(pd.concat([inf, pd.DataFrame(nuevos)], ignore_index=True))
    gmap = inf.set_index("Código")["Grupo"]
    vacio = ent["Grupo"].eq("")
    ent.loc[vacio, "Grupo"] = ent.loc[vacio, "Código"].map(gmap).fillna("")
    if len(ent):
        prim = ent.sort_values("Fecha").groupby("Código").agg(F=("Fecha", "min"), T=("Tipo (GP_AD)", "last"))
        m = inf["Código"].isin(prim.index)
        inf.loc[m, "Fecha Entrega"] = inf.loc[m, "Código"].map(prim["F"])
        inf.loc[m, "Estado"] = "Entregado"
        sin_tipo = m & inf["Tipo (GP_AD)"].eq("")
        inf.loc[sin_tipo, "Tipo (GP_AD)"] = inf.loc[sin_tipo, "Código"].map(prim["T"])
        marca = m & inf["Observaciones"].str.startswith(OBS_MARCA)
        inf.loc[marca, "Observaciones"] = ""
    return std_inf(inf), std_ent(ent)


# ════════════════════════════════════════════════════════════════════
# IMPORTACIÓN DEL FORMATO ORIGINAL (INFORMES G1 G2 G3 · Avance Diario …)
# ════════════════════════════════════════════════════════════════════
def _rgb(c) -> str | None:
    try:
        return c.fill.fgColor.rgb if c.fill is not None and c.fill.fill_type else None
    except Exception:  # noqa: BLE001
        return None


def _color_entregado(ws) -> str:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "entregado" in c.value.lower() and c.column > 1:
                col = _rgb(ws.cell(c.row, c.column - 1))
                if isinstance(col, str):
                    return col
    return "FFFFFF00"


def _eval_simple(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and v.startswith("="):
        expr = v[1:].lstrip("+")
        if re.fullmatch(r"[\d\.\s\+\-\*/\(\)]+", expr):
            try:
                return float(eval(expr, {"__builtins__": {}}, {}))  # noqa: S307 (solo aritmética)
            except Exception:  # noqa: BLE001
                return None
    return None


def importar_formato_original(contenido: bytes):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(contenido))
    wv = load_workbook(io.BytesIO(contenido), data_only=True)
    notas: list[str] = []
    hoja = next(s for s in wb.sheetnames if re.search(r"INFORMES\s*G1", s, re.I))
    ws = wb[hoja]
    color = _color_entregado(ws)

    snaps = []
    for col in range(1, ws.max_column + 1):
        g = str(ws.cell(2, col).value or "").strip().upper()
        if g not in GRUPOS:
            continue
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(ws.cell(1, col).value or ""))
        fecha = pd.Timestamp(int(m[3]), int(m[2]), int(m[1])) if m else pd.Timestamp(hoy_lima())
        items = []
        for r in range(3, ws.max_row + 1):
            c = ws.cell(r, col)
            v = c.value
            if v is None or (isinstance(v, str) and (v.startswith("=") or "RLP" not in v.upper())):
                continue
            items.append((v, _rgb(c) == color))
        snaps.append((g, fecha, items))

    reg: dict[str, dict] = {}
    for g in GRUPOS:
        ss = sorted([s for s in snaps if s[0] == g], key=lambda s: s[1])
        previos, f_prev = set(), None
        for k, (_, fecha, items) in enumerate(ss):
            ultimo = k == len(ss) - 1
            actuales = set()
            for v, pint in items:
                cod = normalizar_codigo(v) if isinstance(v, str) else construir_codigo(int(v), fecha.year)
                actuales.add(cod)
                if cod not in reg:
                    reg[cod] = {"Código": cod, "Grupo": g, "Fecha Ingreso": fecha, "Estado": "Pendiente",
                                "Fecha Entrega": pd.NaT, "Observaciones": ""}
                elif reg[cod]["Grupo"] != g:
                    notas.append(f"{cod} figura en las listas {reg[cod]['Grupo']} y {g}; se conservó {reg[cod]['Grupo']}.")
                if ultimo and pint:
                    reg[cod].update({"Estado": "Entregado", "Fecha Entrega": fecha,
                                     "Observaciones": f"{OBS_MARCA} (fecha referencial de corte)"})
            salieron = previos - actuales
            for cod in salieron:
                if reg[cod]["Estado"] != "Entregado":
                    reg[cod].update({"Estado": "Entregado", "Fecha Entrega": fecha,
                                     "Observaciones": f"Entregado entre {f_prev + timedelta(days=1):%d/%m} y "
                                                      f"{fecha:%d/%m} ({OBS_HIST}; fecha referencial)"})
            if salieron:
                notas.append(f"{g}: {len(salieron)} informes de la lista al {f_prev:%d/%m} ya no figuran al "
                             f"{fecha:%d/%m}; se registraron como entregados en ese periodo (fecha referencial "
                             f"{fecha:%d/%m}, {OBS_HIST}).")
            nuevos = actuales - previos if previos else set()
            if nuevos:
                notas.append(f"{g}: {len(nuevos)} informes se incorporaron a la lista al {fecha:%d/%m} "
                             f"({', '.join(sorted(str(int(numero_de(c))) for c in nuevos))}).")
            previos, f_prev = actuales, fecha
    typo = sum(1 for s in snaps for v, _ in s[2] if isinstance(v, str) and "FLAB" in v.upper())
    if typo:
        notas.append(f"Se corrigieron {typo} códigos escritos «FlAB» (L minúscula) por «FIAB».")

    # Detalle de pendientes G3 / restricciones
    for s in wb.sheetnames:
        if s == hoja or not re.search(r"PENDIENTE", s, re.I):
            continue
        for row in wb[s].iter_rows(values_only=True):
            if not row or not isinstance(row[0], str) or "RLP" not in row[0].upper():
                continue
            cod = normalizar_codigo(row[0])
            det = {"Plan": row[1] if len(row) > 1 else "", "TAG / Circuito": str(row[2] or "").lstrip("`") if len(row) > 2 else "",
                   "Restricción": row[3] if len(row) > 3 else ""}
            if cod not in reg:
                reg[cod] = {"Código": cod, "Grupo": "G3", "Fecha Ingreso": pd.NaT, "Estado": "Pendiente",
                            "Fecha Entrega": pd.NaT, "Observaciones": f"Tomado de la hoja «{s}»"}
            reg[cod].update({k: v for k, v in det.items() if v})

    # Registro de entregas diarias
    log = []
    for s in wb.sheetnames:
        if not re.search(r"AVANCE", s, re.I):
            continue
        wsa = wb[s]
        hdr, cols = None, {}
        for r in range(1, min(wsa.max_row, 40) + 1):
            vals = {str(wsa.cell(r, c).value or "").strip().upper(): c for c in range(1, wsa.max_column + 1)}
            if "FECHA" in vals and "CODIGO" in vals:
                hdr, cols = r, vals
                break
        if not hdr:
            continue
        for r in range(hdr + 1, wsa.max_row + 1):
            f = wsa.cell(r, cols["FECHA"]).value
            n = wsa.cell(r, cols["CODIGO"]).value
            if not isinstance(f, datetime) or n is None:
                continue
            cod = normalizar_codigo(n) if isinstance(n, str) and "RLP" in n.upper() else ""
            log.append({"Fecha": f, "N° Informe": numero_de(cod) if cod else pd.to_numeric(n, errors="coerce"),
                        "Código": cod, "Grupo": str(wsa.cell(r, cols["GRUPO"]).value or "") if "GRUPO" in cols else "",
                        "Tipo (GP_AD)": str(wsa.cell(r, cols["GP_AD"]).value or "") if "GP_AD" in cols else "",
                        "Observaciones": ""})

    # Plan de entrega
    plan_rows, real_prev = [], {}
    for s in wb.sheetnames:
        if not re.search(r"PLAN", s, re.I):
            continue
        pv, pf = wv[s], wb[s]
        fila_f = max(range(1, pv.max_row + 1),
                     key=lambda r: sum(isinstance(pv.cell(r, c).value, datetime) for c in range(1, pv.max_column + 1)))
        fila_p = fila_r = None
        for r in range(1, pv.max_row + 1):
            t = str(pv.cell(r, 1).value or "").lower()
            if "entrega" in t and "acumul" not in t and "proyect" in t:
                fila_p = r
            if "entrega" in t and "acumul" not in t and "real" in t:
                fila_r = r
        if not fila_p:
            continue
        for c in range(2, pv.max_column + 1):
            f = pv.cell(fila_f, c).value
            if not isinstance(f, datetime):
                continue
            v = pv.cell(fila_p, c).value
            v = _eval_simple(v if v is not None else pf.cell(fila_p, c).value)
            if v is not None:
                plan_rows.append({"Fecha": f, "Meta Diaria": v})
            if fila_r:
                rv = pv.cell(fila_r, c).value
                rv = _eval_simple(rv if rv is not None else pf.cell(fila_r, c).value)
                if rv is not None:
                    real_prev[pd.Timestamp(f).normalize()] = rv
        break

    inf = std_inf(pd.DataFrame(list(reg.values())))
    ent = std_ent(pd.DataFrame(log, columns=COLS_ENT))
    inf, ent = sincronizar(inf, ent)
    plan = std_plan(pd.DataFrame(plan_rows, columns=COLS_PLAN))

    # Verificaciones del plan original
    if real_prev and len(ent):
        gm = inf.set_index("Código")["Grupo"]
        tot = ent.groupby("Fecha").size()
        g1 = ent[ent["Código"].map(gm) == "G1"].groupby("Fecha").size()
        dias = [d for d in tot.index if d in real_prev]
        if dias and all(abs(real_prev[d] - tot[d]) < .01 for d in dias) and any(tot[d] != g1.get(d, 0) for d in dias):
            notas.append("La fila «Entregas G1-real» del plan original suma las entregas diarias de todos los grupos "
                         f"(G1+G2+G3), no solo G1; entre el {dias[0]:%d/%m} y el {dias[-1]:%d/%m} sobreestima el avance "
                         f"G1 en {int(sum(tot[d] - g1.get(d, 0) for d in dias))} informes. El tablero calcula el real por grupo.")
        antes = [d for d in real_prev if len(ent) and d < ent["Fecha"].min()]
        pl = plan.set_index("Fecha")["Meta Diaria"]
        if antes and all(abs(real_prev[d] - pl.get(d, -1)) < .01 for d in antes):
            notas.append(f"Los valores «reales» del {min(antes):%d/%m} al {max(antes):%d/%m} replican el proyectado; "
                         "no existe detalle diario de ese periodo.")
    rep = len(ent) - ent["Código"].nunique()
    if rep:
        notas.append(f"El registro diario contiene {rep} entrega(s) repetida(s) del mismo informe (reentregas).")
    return inf, ent, plan, notas


def es_formato_original(contenido: bytes) -> bool:
    try:
        return any(re.search(r"INFORMES\s*G1", s, re.I) for s in pd.ExcelFile(io.BytesIO(contenido)).sheet_names)
    except Exception:  # noqa: BLE001
        return False


def importar_canonico(contenido: bytes):
    xl = pd.ExcelFile(io.BytesIO(contenido))
    hojas = {s.upper(): s for s in xl.sheet_names}
    inf = xl.parse(hojas["INFORMES"]) if "INFORMES" in hojas else pd.DataFrame(columns=COLS_INF)
    ent = xl.parse(hojas["ENTREGAS"]) if "ENTREGAS" in hojas else pd.DataFrame(columns=COLS_ENT)
    plan = xl.parse(hojas["PLAN"]) if "PLAN" in hojas else pd.DataFrame(columns=COLS_PLAN)
    notas = xl.parse(hojas["NOTAS"]).iloc[:, 0].dropna().astype(str).tolist() if "NOTAS" in hojas else []
    inf, ent = sincronizar(inf, ent)
    return inf, ent, std_plan(plan), notas


# ════════════════════════════════════════════════════════════════════
# EXPORTACIÓN
# ════════════════════════════════════════════════════════════════════
def _formatear_hoja(ws, df: pd.DataFrame):
    from openpyxl.styles import Font, PatternFill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(df.columns, start=1):
        largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)])
        ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max(11, largo + 2), 48)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")


def a_excel(hojas: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="DD/MM/YYYY", date_format="DD/MM/YYYY") as w:
        for nombre, df in hojas.items():
            df.to_excel(w, sheet_name=nombre, index=False)
            _formatear_hoja(w.sheets[nombre], df)
    return buf.getvalue()


def excel_canonico(inf, ent, plan, notas) -> bytes:
    return a_excel({"INFORMES": inf, "ENTREGAS": ent, "PLAN": plan,
                    "NOTAS": pd.DataFrame({"Notas de conciliación": notas or [""]})})


def excel_formato_seguimiento(inf, ent, plan, corte: pd.Timestamp, alcance: list[str]) -> bytes:
    """Regenera el formato de trabajo habitual con valores recalculados."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    amarillo = PatternFill("solid", fgColor="FFFF00")
    crema = PatternFill("solid", fgColor="FFFFCC")
    neg = Font(bold=True)
    wb = Workbook()
    ini_log = ent["Fecha"].min() if len(ent) else corte

    ws = wb.active
    ws.title = "INFORMES G1 G2 G3"
    for j, g in enumerate(GRUPOS):
        col = 1 + j * 2
        sub = inf[(inf["Grupo"] == g) & ((inf["Estado"] == "Pendiente") |
                                         (inf["Fecha Entrega"] >= ini_log) | inf["Fecha Entrega"].isna())]
        sub = sub.sort_values(["Adicional", "N° Informe"])
        c = ws.cell(1, col, f"PENDIENTES AL {corte:%d/%m/%Y}\n{DESC_GRUPO[g].upper()}")
        c.alignment = Alignment(wrap_text=True)
        c.font = neg
        ws.cell(2, col, g).font = neg
        ws.cell(3, col, f"=COUNTA({ws.cell(4, col).column_letter}4:{ws.cell(4, col).column_letter}{max(4, 3 + len(sub))})").fill = crema
        for i, (_, r) in enumerate(sub.iterrows(), start=4):
            cc = ws.cell(i, col, r["Código"])
            if r["Estado"] == "Entregado":
                cc.fill = amarillo
        ws.column_dimensions[ws.cell(1, col).column_letter].width = 42
    ws.row_dimensions[1].height = 45
    ws.cell(3, 8, "Leyenda").font = neg
    ws.cell(4, 8).fill = amarillo
    ws.cell(4, 9, "Informe Entregado")

    w3 = wb.create_sheet("PENDIENTES G3")
    w3.cell(1, 1, "Pendientes con restricción").font = neg
    for j, h in enumerate(["CÓDIGO", "PLAN", "TAG / CIRCUITO", "RESTRICCIÓN"], start=1):
        w3.cell(2, j, h).font = neg
    p3 = inf[(inf["Estado"] == "Pendiente") & ((inf["Grupo"] == "G3") | inf["Restricción"].ne(""))]
    for i, (_, r) in enumerate(p3.iterrows(), start=3):
        for j, k in enumerate(["Código", "Plan", "TAG / Circuito", "Restricción"], start=1):
            w3.cell(i, j, r[k])
    for j, wdt in enumerate([42, 18, 26, 30], start=1):
        w3.column_dimensions[w3.cell(1, j).column_letter].width = wdt

    wp = wb.create_sheet("Plan de entrega")
    wp.cell(1, 1, f"Plan de entrega de informes · alcance {' + '.join(alcance)}").font = neg
    if len(plan):
        d0 = plan["Fecha"].min() - pd.Timedelta(days=1)
        b0 = int(backlog_en(inf[inf["Grupo"].isin(alcance)], d0))
        reales = entregas_por_dia(inf[inf["Grupo"].isin(alcance)])
        acp = b0
        filas = ["Entregas proyectadas", "Backlog proyectado", "Entregas reales", "Backlog real"]
        for i, t in enumerate(filas, start=4):
            wp.cell(i, 1, t).font = neg
        wp.cell(3, 1, f"Backlog inicial al {d0:%d/%m/%Y}: {b0}")
        for j, (_, r) in enumerate(plan.iterrows(), start=2):
            f = r["Fecha"]
            wp.cell(3, j, f.to_pydatetime()).number_format = "DD/MM"
            acp = max(0, acp - r["Meta Diaria"])
            wp.cell(4, j, r["Meta Diaria"])
            wp.cell(5, j, round(acp, 2))
            if f <= corte:
                wp.cell(6, j, int(reales.get(f, 0)))
                wp.cell(7, j, backlog_en(inf[inf["Grupo"].isin(alcance)], f))
        wp.column_dimensions["A"].width = 26
        wp.cell(9, 1, "Backlog real incluye informes incorporados al alcance después de la línea base. "
                      "Entregas sin detalle diario se muestran en su fecha referencial de corte.")

    wa = wb.create_sheet("Avance Diario")
    wa.cell(1, 2, "Detalle de entregas diarias").font = neg
    for j, h in enumerate(["FECHA", "CODIGO", "GRUPO", "GP_AD"], start=2):
        wa.cell(3, j, h).font = neg
    prev = None
    for i, (_, r) in enumerate(ent.iterrows(), start=4):
        c = wa.cell(i, 2, r["Fecha"].to_pydatetime())
        c.number_format = "DD/MM/YYYY"
        if r["Fecha"] != prev:
            c.fill = amarillo
            prev = r["Fecha"]
        wa.cell(i, 3, int(r["N° Informe"]) if pd.notna(r["N° Informe"]) else r["Código"])
        wa.cell(i, 4, r["Grupo"])
        wa.cell(i, 5, r["Tipo (GP_AD)"])
    wa.column_dimensions["B"].width = 13

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════
# PERSISTENCIA
# ════════════════════════════════════════════════════════════════════
def _secret(*keys):
    try:
        v = st.secrets
        for k in keys:
            v = v[k]
        return v
    except Exception:  # noqa: BLE001
        return None


def subir_github(contenido: bytes, mensaje: str) -> tuple[bool, str]:
    token, repo = _secret("github", "token"), _secret("github", "repo")
    if not token or not repo:
        return False, ""
    import requests
    rama = _secret("github", "branch") or "main"
    ruta = _secret("github", "path") or "data/matriz_kpi_informes.xlsx"
    url = f"https://api.github.com/repos/{repo}/contents/{ruta}"
    hdr = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
    try:
        r = requests.get(url, headers=hdr, params={"ref": rama}, timeout=20)
        payload = {"message": mensaje, "content": base64.b64encode(contenido).decode(), "branch": rama}
        if r.status_code == 200:
            payload["sha"] = r.json().get("sha")
        r = requests.put(url, headers=hdr, json=payload, timeout=30)
        return r.status_code in (200, 201), f"GitHub respondió {r.status_code}"
    except Exception as e:  # noqa: BLE001
        return False, f"GitHub: {e}"


def guardar(inf, ent, plan, motivo: str, notas: list[str] | None = None) -> None:
    inf, ent = sincronizar(inf, ent)
    plan = std_plan(plan)
    notas = st.session_state.get("notas", []) if notas is None else notas
    st.session_state.update(inf=inf, ent=ent, plan=plan, notas=notas)
    contenido = excel_canonico(inf, ent, plan, notas)
    msgs = []
    try:
        DATA_DIR.mkdir(exist_ok=True)
        DATA_FILE.write_bytes(contenido)
        msgs.append("✅ Matriz guardada en el servidor")
    except Exception as e:  # noqa: BLE001
        msgs.append(f"⚠️ No se pudo escribir el archivo: {e}")
    ok, info = subir_github(contenido, f"{motivo} · {datetime.now(TZ):%d/%m/%Y %H:%M}")
    if ok:
        msgs.append("✅ Respaldo sincronizado en GitHub")
    elif info:
        msgs.append(f"⚠️ {info}")
    st.session_state.ultimo_guardado = datetime.now(TZ)
    st.session_state.flash = f"**{motivo}** · " + " · ".join(msgs)


def cargar_inicial():
    if DATA_FILE.exists():
        try:
            return importar_canonico(DATA_FILE.read_bytes())
        except Exception:  # noqa: BLE001
            pass
    return (std_inf(pd.DataFrame(columns=COLS_INF)), std_ent(pd.DataFrame(columns=COLS_ENT)),
            std_plan(pd.DataFrame(columns=COLS_PLAN)), [])


def puede_editar() -> bool:
    return (not _secret("edicion", "password")) or st.session_state.get("editor_ok", False)


# ════════════════════════════════════════════════════════════════════
# CÁLCULOS DE KPI
# ════════════════════════════════════════════════════════════════════
def backlog_en(inf_s: pd.DataFrame, t: pd.Timestamp) -> int:
    ing = inf_s["Fecha Ingreso"]
    fe = inf_s["Fecha Entrega"]
    entregado_t = inf_s["Estado"].eq("Entregado") & (fe.isna() | (fe <= t))
    return int(((ing.isna() | (ing <= t)) & ~entregado_t).sum())


def entregas_por_dia(inf_s: pd.DataFrame) -> pd.Series:
    e = inf_s[inf_s["Estado"].eq("Entregado") & inf_s["Fecha Entrega"].notna()]
    return e.groupby("Fecha Entrega").size()


def serie_real(inf_s: pd.DataFrame, desde: pd.Timestamp, corte: pd.Timestamp) -> pd.DataFrame:
    ev = set(inf_s["Fecha Ingreso"].dropna()) | set(inf_s["Fecha Entrega"].dropna())
    ev = sorted({desde, corte} | {d for d in ev if desde < d <= corte})
    return pd.DataFrame({"Fecha": ev, "Pendientes": [backlog_en(inf_s, t) for t in ev]})


def serie_plan(plan: pd.DataFrame, b0: float) -> pd.DataFrame:
    if not len(plan):
        return pd.DataFrame(columns=["Fecha", "Plan"])
    d0 = plan["Fecha"].min() - pd.Timedelta(days=1)
    acum = (b0 - plan["Meta Diaria"].cumsum()).clip(lower=0)
    return pd.DataFrame({"Fecha": [d0] + plan["Fecha"].tolist(), "Plan": [b0] + acum.tolist()})


def kpis(inf, ent, plan, alcance, corte, n_vel):
    s = inf[inf["Grupo"].isin(alcance)]
    r = {"total": len(s), "entregados": int((s["Estado"] == "Entregado").sum())}
    r["pendientes"] = backlog_en(s, corte)
    r["avance"] = r["entregados"] / r["total"] * 100 if r["total"] else np.nan
    d0 = plan["Fecha"].min() - pd.Timedelta(days=1) if len(plan) else (s["Fecha Ingreso"].min() if len(s) else corte)
    r["d0"], r["b0"] = d0, backlog_en(s, d0) if pd.notna(d0) else r["total"]
    fe = s["Fecha Entrega"]
    r["real_acum"] = int((s["Estado"].eq("Entregado") & (fe > d0) & (fe <= corte)).sum())
    r["plan_acum"] = float(plan.loc[(plan["Fecha"] > d0) & (plan["Fecha"] <= corte), "Meta Diaria"].sum()) if len(plan) else np.nan
    r["spi"] = r["real_acum"] / r["plan_acum"] * 100 if r["plan_acum"] else np.nan
    pl = serie_plan(plan, r["b0"])
    fin = pl[pl["Plan"] <= 0.001]["Fecha"]
    r["fin_plan"] = fin.min() if len(fin) else (plan["Fecha"].max() if len(plan) else pd.NaT)
    # velocidad: primeras entregas del alcance registradas en el log, últimos n días hábiles
    ini = pd.Timestamp(np.busday_offset(np.datetime64(corte.date()), -(n_vel - 1), roll="backward", holidays=HOL))
    cods = set(s["Código"])
    prim = ent[ent["Código"].isin(cods)].sort_values("Fecha").drop_duplicates("Código")
    r["vel"] = len(prim[(prim["Fecha"] >= ini) & (prim["Fecha"] <= corte)]) / n_vel
    r["vel_ini"] = ini
    if r["vel"] > 0 and r["pendientes"] > 0:
        dias = math.ceil(r["pendientes"] / r["vel"])
        r["cierre"] = pd.Timestamp(np.busday_offset(np.datetime64(corte.date()), dias, roll="forward", holidays=HOL))
    elif r["pendientes"] == 0:
        r["cierre"] = corte
    else:
        r["cierre"] = pd.NaT
    r["meta_prom"] = float(plan.loc[plan["Meta Diaria"] >= 1, "Meta Diaria"].tail(15).median()) if len(plan) else np.nan
    r["crec_alcance"] = int(((s["Fecha Ingreso"] > d0) & (s["Fecha Ingreso"] <= corte)).sum()) if pd.notna(d0) else 0
    return r


def calidad(inf, ent, plan) -> dict[str, pd.DataFrame]:
    gm = inf.set_index("Código")["Grupo"]
    out = {}
    dup = ent[ent.duplicated("Código", keep=False)].sort_values(["Código", "Fecha"])
    out["Reentregas (mismo informe registrado más de una vez)"] = dup
    e = ent.assign(**{"Grupo en maestro": ent["Código"].map(gm)})
    out["Grupo del registro diario ≠ grupo en listas"] = e[e["Grupo"].ne(e["Grupo en maestro"]) & e["Grupo"].ne("")]
    out["Incorporados desde el registro diario (no figuraban en listas)"] = inf[inf["Observaciones"].str.startswith("Incorporado")]
    out["Entregados con fecha referencial (sin detalle diario)"] = inf[inf["Observaciones"].str.contains("referencial")]
    out["Pendientes G3 sin restricción registrada"] = inf[(inf["Grupo"] == "G3") & (inf["Estado"] == "Pendiente") & inf["Restricción"].eq("")]
    if len(plan):
        p = plan.copy()
        p["Día"] = p["Fecha"].dt.day_name(locale=None)
        p["Hábil"] = p["Fecha"].map(es_habil)
        out["Plan con meta en feriado / fin de semana"] = p[(p["Meta Diaria"] > 0) & ~p["Hábil"]][["Fecha", "Meta Diaria"]]
    return out


# ════════════════════════════════════════════════════════════════════
# GRÁFICOS
# ════════════════════════════════════════════════════════════════════
def estilo(fig: go.Figure, h: int = 380, titulo: str | None = None) -> go.Figure:
    fig.update_layout(template="plotly_white", height=h, margin=dict(l=10, r=10, t=60 if titulo else 25, b=10),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
                      font=dict(family="Segoe UI, Arial, sans-serif", size=12),
                      title=dict(text=titulo, font=dict(size=15, color=AZUL)) if titulo else None)
    return fig


def mostrar(fig):
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})


def sin_datos(msg="No hay datos suficientes para este gráfico."):
    st.info(msg, icon="ℹ️")


def fig_burndown(inf_s, plan, k, corte, titulo):
    fig = go.Figure()
    pl = serie_plan(plan, k["b0"])
    if len(pl):
        fig.add_scatter(x=pl["Fecha"], y=pl["Plan"], name="Plan", mode="lines",
                        line=dict(color=GRIS, width=2.5, dash="dash"))
    desde = k["d0"] if pd.notna(k["d0"]) else inf_s["Fecha Ingreso"].min()
    rl = serie_real(inf_s, desde, corte)
    fig.add_scatter(x=rl["Fecha"], y=rl["Pendientes"], name="Real", mode="lines+markers",
                    line=dict(color=AZUL, width=3), marker=dict(size=7),
                    hovertemplate="%{x|%d/%m}: %{y} pendientes<extra></extra>")
    fig.add_annotation(x=rl["Fecha"].iloc[-1], y=rl["Pendientes"].iloc[-1], text=f"<b>{rl['Pendientes'].iloc[-1]}</b>",
                       showarrow=False, yshift=14, font=dict(color=AZUL))
    if pd.notna(k["cierre"]) and k["pendientes"] > 0:
        fig.add_scatter(x=[corte, k["cierre"]], y=[k["pendientes"], 0], mode="lines+markers",
                        name=f"Proyección al ritmo actual ({k['vel']:.1f}/día)",
                        line=dict(color=NARANJA, width=2, dash="dot"))
    fig.add_vline(x=corte, line_color="#17202A", line_width=1, line_dash="dot")
    fig.update_layout(yaxis_title="Informes pendientes", xaxis_title=None, hovermode="x unified")
    fig.update_xaxes(tickformat="%d/%m")
    return estilo(fig, 400, titulo)


# ════════════════════════════════════════════════════════════════════
# FORMULARIOS
# ════════════════════════════════════════════════════════════════════
def form_entregas(prefix: str, corte: pd.Timestamp):
    if not puede_editar():
        st.warning("🔒 Ingrese la clave de edición en la barra lateral para registrar entregas.")
        return
    inf, ent = st.session_state.inf, st.session_state.ent
    reent = st.checkbox("Incluir informes ya entregados (registrar reentrega)", key=f"{prefix}_reent")
    base = inf if reent else inf[inf["Estado"] == "Pendiente"]
    etiqueta = {r["Código"]: f"{int(r['N° Informe']) if pd.notna(r['N° Informe']) else ''} · {r['Grupo']} · "
                f"{r['Código'].replace(PREFIJO + '-', '')}" + (f" · {r['TAG / Circuito']}" if r["TAG / Circuito"] else "")
                for _, r in base.iterrows()}
    tipos = sorted(set(TIPOS_GPAD) | set(ent["Tipo (GP_AD)"].unique()) - {""})
    with st.form(f"{prefix}_fe", clear_on_submit=True):
        c1, c2 = st.columns([1, 3])
        fecha = c1.date_input("Fecha de entrega", value=hoy_lima(), format="DD/MM/YYYY", key=f"{prefix}_f")
        cods = c2.multiselect("Informes entregados (escriba el N° para buscar)", list(etiqueta),
                              format_func=lambda c: etiqueta[c], key=f"{prefix}_c")
        c3, c4 = st.columns([1, 3])
        tipo = c3.selectbox("Tipo (GP_AD)", tipos, key=f"{prefix}_t")
        obs = c4.text_input("Observaciones", key=f"{prefix}_o")
        ok = st.form_submit_button("📬 Registrar entregas", type="primary")
    if ok:
        if not cods:
            st.error("Seleccione al menos un informe.")
            return
        gm = inf.set_index("Código")["Grupo"]
        nuevas = pd.DataFrame([{"Fecha": pd.Timestamp(fecha), "N° Informe": numero_de(c), "Código": c,
                                "Grupo": gm.get(c, ""), "Tipo (GP_AD)": tipo, "Observaciones": obs} for c in cods])
        guardar(inf, pd.concat([ent, nuevas], ignore_index=True), st.session_state.plan,
                f"{len(cods)} entrega(s) registradas el {fecha:%d/%m/%Y}")
        st.rerun()


def form_informe(prefix: str):
    if not puede_editar():
        st.warning("🔒 Ingrese la clave de edición en la barra lateral para añadir informes.")
        return
    inf = st.session_state.inf
    with st.form(f"{prefix}_fi", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
        n = c1.number_input("N° de informe *", min_value=1, step=1, value=None, key=f"{prefix}_n")
        adic = c2.selectbox("Adicional", ["No", "Sí"], key=f"{prefix}_a")
        grupo = c3.selectbox("Grupo *", GRUPOS, key=f"{prefix}_g")
        f_ing = c4.date_input("Fecha de ingreso al backlog", value=hoy_lima(), format="DD/MM/YYYY", key=f"{prefix}_fi")
        c5, c6, c7 = st.columns(3)
        plan = c5.selectbox("Plan", [""] + sorted(set(PLANES) | set(inf["Plan"]) - {""}), key=f"{prefix}_p")
        tag = c6.text_input("TAG / Circuito", key=f"{prefix}_tag")
        restr = c7.selectbox("Restricción (si aplica)", [""] + sorted(set(RESTRICCIONES) | set(inf["Restricción"]) - {""}),
                             key=f"{prefix}_r")
        obs = st.text_input("Observaciones", key=f"{prefix}_ob")
        ok = st.form_submit_button("➕ Añadir al backlog", type="primary")
    if ok:
        if not n:
            st.error("Ingrese el N° de informe.")
            return
        cod = construir_codigo(n, f_ing.year, adic == "Sí")
        if cod in set(inf["Código"]):
            st.error(f"El informe {cod} ya existe en la matriz.")
            return
        fila = {"Código": cod, "Grupo": grupo, "Plan": plan, "TAG / Circuito": tag, "Restricción": restr,
                "Fecha Ingreso": pd.Timestamp(f_ing), "Estado": "Pendiente", "Observaciones": obs}
        guardar(pd.concat([inf, pd.DataFrame([fila])], ignore_index=True), st.session_state.ent,
                st.session_state.plan, f"Informe {cod} añadido a {grupo}")
        st.rerun()


def carga_matriz(prefix: str, compacto: bool = False):
    if not puede_editar():
        st.warning("🔒 Ingrese la clave de edición en la barra lateral para subir una matriz.")
        return
    up = st.file_uploader("Matriz (.xlsx) — formato original o exportado por este tablero", type=["xlsx", "xlsm"],
                          key=f"{prefix}_up")
    if not up:
        return
    contenido = up.getvalue()
    try:
        original = es_formato_original(contenido)
        inf, ent, plan, notas = (importar_formato_original if original else importar_canonico)(contenido)
    except Exception as e:  # noqa: BLE001
        st.error(f"No se pudo interpretar el archivo: {e}")
        return
    st.caption(("Formato de seguimiento original detectado" if original else "Formato del tablero detectado") +
               f" · {len(inf)} informes · {len(ent)} entregas diarias · {len(plan)} días de plan")
    resumen = inf.groupby(["Grupo", "Estado"]).size().unstack(fill_value=0)
    st.dataframe(resumen, use_container_width=True)
    if notas:
        with st.expander(f"📝 Notas de conciliación ({len(notas)})"):
            for nt in notas:
                st.markdown(f"- {nt}")
    modo = st.radio("Modo de carga", ["Reemplazar matriz completa", "Actualizar / anexar"], key=f"{prefix}_modo",
                    horizontal=not compacto,
                    help="«Actualizar» agrega informes nuevos, actualiza los existentes por código, suma las "
                         "entregas no registradas y reemplaza el plan en las fechas incluidas.")
    if st.button("✅ Confirmar carga", type="primary", key=f"{prefix}_ok"):
        if modo.startswith("Reemplazar"):
            guardar(inf, ent, plan, "Matriz reemplazada", notas)
        else:
            s = st.session_state
            i2 = pd.concat([s.inf, inf]).drop_duplicates("Código", keep="last")
            e2 = pd.concat([s.ent, ent]).drop_duplicates(["Fecha", "Código"], keep="last")
            p2 = pd.concat([s.plan, plan]).drop_duplicates("Fecha", keep="last")
            guardar(i2, e2, p2, "Matriz actualizada por carga de archivo", list(dict.fromkeys(s.notas + notas)))
        st.rerun()



# ════════════════════════════════════════════════════════════════════
# PLAN VS REAL (avance acumulado · plan diario · hitos de control)
# ════════════════════════════════════════════════════════════════════
C_PLAN, C_REAL, C_G2, C_G3 = "#A9A9A9", "#D35400", "#148F77", "#1F4E79"
C_SOMBRA = "rgba(214,204,186,0.28)"
DIAS_ES = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre",
            "Octubre", "Noviembre", "Diciembre"]


def _vline(fig, x, texto=None, color="#17202A"):
    fig.add_shape(type="line", x0=x, x1=x, y0=0, y1=1, yref="paper", line=dict(color=color, width=1.2, dash="dot"))
    if texto:
        fig.add_annotation(x=x, y=1, yref="paper", text=texto, showarrow=False, xanchor="left", yanchor="top",
                           xshift=6, font=dict(size=12, color="#17202A"))


def acum_real(inf_s: pd.DataFrame, d0: pd.Timestamp, t: pd.Timestamp) -> int:
    fe = inf_s["Fecha Entrega"]
    return int((inf_s["Estado"].eq("Entregado") & (fe > d0) & (fe <= t)).sum())


def acum_plan(plan: pd.DataFrame, d0: pd.Timestamp, t: pd.Timestamp) -> float:
    return float(plan.loc[(plan["Fecha"] > d0) & (plan["Fecha"] <= t), "Meta Diaria"].sum())


def fig_avance_acumulado(inf_s, plan, K, corte, alcance_txt):
    d0, b0 = K["d0"], K["b0"]
    fin = plan["Fecha"].max()
    pa = pd.concat([pd.DataFrame({"Fecha": [d0], "Plan": [0.0]}),
                    plan.assign(Plan=plan["Meta Diaria"].cumsum())[["Fecha", "Plan"]]])
    fe = inf_s.loc[inf_s["Estado"].eq("Entregado"), "Fecha Entrega"].dropna()
    fechas = sorted({d0, corte} | {d for d in fe if d0 < d <= corte})
    ra = pd.DataFrame({"Fecha": fechas, "Real": [acum_real(inf_s, d0, t) for t in fechas]})
    p_c, r_c = acum_plan(plan, d0, corte), acum_real(inf_s, d0, corte)

    fig = go.Figure()
    if corte < fin:
        fig.add_vrect(x0=corte, x1=fin, fillcolor=C_SOMBRA, line_width=0, layer="below")
    fig.add_scatter(x=pa["Fecha"], y=pa["Plan"], name="Proyectado", mode="lines", line=dict(color=C_PLAN, width=3),
                    hovertemplate="%{x|%d/%m} · programado %{y:.0f}<extra></extra>")
    fig.add_scatter(x=ra["Fecha"], y=ra["Real"], name="Real", mode="lines", fill="tozeroy",
                    fillcolor="rgba(211,84,0,0.12)", line=dict(color=C_REAL, width=3.5),
                    hovertemplate="%{x|%d/%m} · real %{y}<extra></extra>")
    fig.add_hline(y=b0, line_color="#566573", line_dash="dot", line_width=1)
    fig.add_annotation(x=d0, y=b0, text=f"Meta: {b0} informes", showarrow=False, xanchor="left", yanchor="bottom",
                       font=dict(size=12))
    if K["total"] != b0:
        fig.add_hline(y=K["total"], line_color=C_PLAN, line_dash="dot", line_width=1)
        fig.add_annotation(x=fin, y=K["total"], text=f"Alcance actual: {K['total']}", showarrow=False,
                           xanchor="right", yanchor="bottom", font=dict(size=11, color="#566573"))
    _vline(fig, corte, f"Corte {corte:%d/%m}")
    fig.add_scatter(x=[corte], y=[p_c], mode="markers", showlegend=False, hoverinfo="skip",
                    marker=dict(size=11, color="white", line=dict(color=C_PLAN, width=2.5)))
    fig.add_scatter(x=[corte], y=[r_c], mode="markers", showlegend=False, hoverinfo="skip",
                    marker=dict(size=10, color=C_REAL, line=dict(color="white", width=1.5)))
    fig.add_scatter(x=[fin], y=[pa["Plan"].iloc[-1]], mode="markers", showlegend=False, hoverinfo="skip",
                    marker=dict(size=10, color="white", line=dict(color=C_PLAN, width=2.5)))
    arriba = p_c >= r_c
    fig.add_annotation(x=corte, y=p_c, text=f"Plan: {p_c:.0f}", showarrow=False, xanchor="right", xshift=-8,
                       yshift=12 if arriba else -12, font=dict(color="#566573", size=12))
    fig.add_annotation(x=corte, y=r_c, text=f"<b>Real: {r_c}</b>", showarrow=False, xanchor="left", xshift=8,
                       yshift=-12 if arriba else 12, font=dict(color=C_REAL, size=12))
    ticks = sorted({d0, fin, corte} | set(pd.date_range(d0, fin, freq="W-MON")))
    fig.update_xaxes(tickvals=ticks, tickformat="%d/%m", showgrid=False, range=[d0 - pd.Timedelta(days=1), fin + pd.Timedelta(days=1)])
    fig.update_yaxes(rangemode="tozero", gridcolor="rgba(0,0,0,.08)")
    fig.update_layout(hovermode="x unified", legend=dict(x=1, xanchor="right"))
    return estilo(fig, 400, f"Avance acumulado {alcance_txt}: proyectado vs real"), p_c, r_c


def tarjeta(valor: str, texto: str, color: str) -> str:
    return (f"<div style='margin:0 0 1.1rem 0'><div style='font-size:2.6rem;font-weight:700;line-height:1;"
            f"color:{color}'>{valor}</div><div style='font-size:.85rem;opacity:.85'>{texto}</div></div>")


def fig_plan_diario(inf, plan, alcance, corte, meta):
    """Barras por día hábil: plan del alcance vs entregas reales (primera entrega) por grupo."""
    dias_plan = set(plan.loc[plan["Meta Diaria"] > 0, "Fecha"])
    reales = inf[inf["Estado"].eq("Entregado") & inf["Fecha Entrega"].notna()]
    ref = reales["Observaciones"].str.contains("referencial")
    det, hist = reales[~ref], reales[ref & reales["Grupo"].isin(alcance)]
    fin = plan["Fecha"].max()
    extra = {d for d in det["Fecha Entrega"] if plan["Fecha"].min() <= d <= min(corte, fin) and d not in dias_plan}
    dias = sorted(dias_plan | extra)
    idx = {d: i for i, d in enumerate(dias)}
    pmap = plan.set_index("Fecha")["Meta Diaria"]

    # Entregas sin detalle diario: se distribuyen como promedio en los días del plan del periodo
    prom = pd.Series(0.0, index=dias)
    for f_ref, n in hist.groupby("Fecha Entrega").size().items():
        periodo = [d for d in dias if d in dias_plan and d <= f_ref and d not in set(det["Fecha Entrega"])]
        periodo = [d for d in periodo if d > max([x for x in det["Fecha Entrega"] if x < f_ref], default=pd.Timestamp.min)]
        if periodo:
            prom[periodo] += n / len(periodo)

    fig = go.Figure()
    x = list(range(len(dias)))
    fig.add_bar(x=x, y=[pmap.get(d, 0) for d in dias], name=f"Plan {'+'.join(alcance)}", marker_color="#C9C9C9",
                offsetgroup="plan", customdata=[f"{d:%d/%m}" for d in dias],
                hovertemplate="%{customdata} · plan %{y:.1f}<extra></extra>")
    base = np.zeros(len(dias))
    if prom.sum():
        fig.add_bar(x=x, y=prom.values, base=base.copy(), name="Real (promedio, sin detalle diario)", offsetgroup="real",
                    marker=dict(color="rgba(211,84,0,0.45)", pattern=dict(shape="/", fgcolor=C_REAL)),
                    hovertemplate="promedio %{y:.1f}<extra></extra>")
        base += prom.values
    for g, col in [("G1", C_REAL), ("G2", C_G2), ("G3", C_G3)]:
        v = det[det["Grupo"] == g].groupby("Fecha Entrega").size().reindex(dias, fill_value=0).values
        v = np.where(np.array(dias) <= corte, v, 0)
        if v.sum():
            fig.add_bar(x=x, y=v, base=base.copy(), name=f"Real {g}", offsetgroup="real", marker_color=col,
                        hovertemplate=f"Real {g}: %{{y}}<extra></extra>")
            base += v
    tot_det = det.groupby("Fecha Entrega").size().reindex(dias, fill_value=0)
    for i, d in enumerate(dias):
        if d <= corte and tot_det[d] > 0:
            fig.add_annotation(x=i + 0.2, y=base[i], text=f"<b>{int(tot_det[d])}</b>", showarrow=False, yshift=9,
                               font=dict(size=12))
    if pd.notna(meta):
        fig.add_hline(y=meta, line_color="#566573", line_dash="dash", line_width=1.2)
        fig.add_scatter(x=[None], y=[None], mode="lines", name=f"Meta {meta:.0f}/día",
                        line=dict(color="#566573", dash="dash"))
    futuros = [i for i, d in enumerate(dias) if d > corte]
    if futuros:
        fig.add_vrect(x0=futuros[0] - 0.5, x1=len(dias) - 0.5, fillcolor=C_SOMBRA, line_width=0, layer="below")
        fig.add_annotation(x=futuros[0] - 0.4, y=1, yref="paper", text=f"<b>Programado a partir del {dias[futuros[0]]:%d/%m}</b>",
                           showarrow=False, xanchor="left", yanchor="top", font=dict(size=12))
    ticktext = [(f"<i>{d:%d}*</i>" if d in extra else f"{d:%d}") for d in dias]
    fig.update_xaxes(tickvals=x, ticktext=ticktext, showgrid=False, range=[-0.7, len(dias) - 0.3])
    vistos = set()
    for i, d in enumerate(dias):
        if d.month not in vistos:
            vistos.add(d.month)
            fig.add_annotation(x=i - 0.4, y=-0.11, yref="paper", text=f"<b>{MESES_ES[d.month - 1]}</b>",
                               showarrow=False, xanchor="left", font=dict(size=12))
    fig.update_yaxes(title_text="Informes", gridcolor="rgba(0,0,0,.08)", rangemode="tozero")
    fig.update_layout(barmode="group", bargap=0.25, bargroupgap=0.05)
    fig = estilo(fig, 440, "Plan diario por día hábil vs entregas reales")
    fig.update_layout(margin=dict(b=55))
    return fig, bool(extra)


def tabla_hitos(inf_s, plan, K, corte, fechas):
    d0, b0 = K["d0"], K["b0"]
    filas = []
    for f in sorted(set(fechas)):
        prog = acum_plan(plan, d0, f)
        r = {"Fecha de corte": f"{DIAS_ES[f.weekday()]} {f:%d/%m}", "Acumulado programado": round(prog),
             "Pendiente programado": round(b0 - prog), "% avance programado": prog / b0 * 100 if b0 else np.nan,
             "Acumulado real": None, "Pendiente real": None, "% avance real": None, "Brecha": None, "Estado": "⏳ Por ejecutar",
             "_corte": f == corte}
        if f <= corte:
            real = acum_real(inf_s, d0, f)
            r.update({"Acumulado real": real, "Pendiente real": backlog_en(inf_s, f),
                      "% avance real": real / b0 * 100 if b0 else np.nan, "Brecha": real - round(prog),
                      "Estado": "✅ Cumple" if real >= round(prog) else "⚠️ Atraso"})
        filas.append(r)
    return pd.DataFrame(filas)


def acciones_rapidas(prefix: str, corte):
    st.divider()
    c1, c2 = st.columns(2)
    with c1.expander("📬 Registrar entregas"):
        form_entregas(prefix + "e", corte)
    with c2.expander("➕ Añadir informe al backlog"):
        form_informe(prefix + "i")


# ════════════════════════════════════════════════════════════════════
# ESTADO INICIAL
# ════════════════════════════════════════════════════════════════════
if "inf" not in st.session_state:
    st.session_state.inf, st.session_state.ent, st.session_state.plan, st.session_state.notas = cargar_inicial()
INF, ENT, PLAN = st.session_state.inf, st.session_state.ent, st.session_state.plan
HOY = pd.Timestamp(hoy_lima())

# ════════════════════════════════════════════════════════════════════
# BARRA LATERAL
# ════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Parámetros")
    corte_def = min(ENT["Fecha"].max(), HOY) if len(ENT) else HOY
    corte = pd.Timestamp(st.date_input("Fecha de corte", value=corte_def.date(), format="DD/MM/YYYY",
                                       help="Por defecto, la última fecha con entregas registradas."))
    op_alc = {"G1 (plan de entrega)": ["G1"], "G1 + G2": ["G1", "G2"], "G1 + G2 + G3": GRUPOS}
    alcance = op_alc[st.radio("Alcance del plan", list(op_alc), index=0)]
    n_vel = st.slider("Ventana de productividad (días hábiles)", 3, 15, 5,
                      help="Base para la velocidad actual y la fecha estimada de cierre.")

    st.markdown("### 📁 Matriz")
    pwd = _secret("edicion", "password")
    if pwd and not st.session_state.get("editor_ok"):
        clave = st.text_input("Clave de edición", type="password")
        if clave == pwd:
            st.session_state.editor_ok = True
            st.rerun()
        elif clave:
            st.error("Clave incorrecta")
    elif pwd:
        st.success("🔓 Edición habilitada")
    ts = st.session_state.get("ultimo_guardado")
    st.caption(f"{len(INF)} informes · {len(ENT)} entregas registradas"
               + (f" · guardado {ts:%d/%m %H:%M}" if ts else ""))
    st.download_button("⬇️ Descargar matriz (.xlsx)", excel_canonico(INF, ENT, PLAN, st.session_state.notas),
                       file_name=f"MATRIZ_KPI_ENTREGA_INFORMES_{HOY:%Y%m%d}.xlsx", use_container_width=True)
    with st.expander("⬆️ Subir matriz actualizada"):
        carga_matriz("sb", compacto=True)

# ════════════════════════════════════════════════════════════════════
# ENCABEZADO
# ════════════════════════════════════════════════════════════════════
st.markdown(f"""<div class="hdr"><h1>📊 KPI · Entrega de Informes</h1>
<p>ADEMINSAC · Contrato N° 3700002135 · Refinería La Pampilla · Corte al {corte:%d/%m/%Y} ·
alcance del plan {' + '.join(alcance)}</p></div>""", unsafe_allow_html=True)
if "flash" in st.session_state:
    st.success(st.session_state.pop("flash"))
if INF.empty:
    st.info("Aún no hay datos. Suba la matriz **KPI_ENTREGA_DE_INFORMES.xlsx** (formato original) desde la barra lateral "
            "o la pestaña **Gestión de Matriz**.", icon="📂")
    carga_matriz("inicio")
    st.stop()

K = kpis(INF, ENT, PLAN, alcance, corte, n_vel)
S = INF[INF["Grupo"].isin(alcance)]
PEND = INF[INF["Estado"] == "Pendiente"].copy()
PEND["Antigüedad (días háb.)"] = habiles_entre(PEND["Fecha Ingreso"], corte)
RESTR = PEND[(PEND["Grupo"] == "G3") | PEND["Restricción"].ne("")]
ENT_C = ENT[ENT["Fecha"] <= corte].copy()
ENT_C["Grupo maestro"] = ENT_C["Código"].map(INF.set_index("Código")["Grupo"])
ENT_C["Reentrega"] = ENT_C.duplicated("Código", keep="first")
CAL = calidad(INF, ENT, PLAN)
n_incons = sum(len(v) for k, v in CAL.items() if not k.startswith("Entregados con fecha"))

tabs = st.tabs(["📊 Resumen Ejecutivo", "🎯 Plan vs Real", "📈 Avance y Productividad", "🚧 Backlog y Restricciones",
                "🔍 Calidad de Datos", "🗂️ Gestión de Matriz"])

# ─────────────────────────── 1. RESUMEN ───────────────────────────
with tabs[0]:
    b_total = backlog_en(INF, corte)
    b_prev = backlog_en(INF, corte - pd.Timedelta(days=7))
    c = st.columns(4)
    c[0].metric("Backlog pendiente (todos los grupos)", f"{b_total}", f"{b_total - b_prev:+d} vs hace 7 días",
                delta_color="inverse")
    c[1].metric(f"Avance {' + '.join(alcance)}", f"{K['avance']:.1f}%" if pd.notna(K["avance"]) else "—",
                f"{K['entregados']} de {K['total']} entregados", delta_color="off")
    c[2].metric("Cumplimiento del plan (real / plan acumulado)", f"{K['spi']:.0f}%" if pd.notna(K["spi"]) else "—",
                f"{K['real_acum'] - K['plan_acum']:+.0f} informes vs plan" if pd.notna(K["spi"]) else None)
    c[3].metric(f"Productividad (últimos {n_vel} días háb.)", f"{K['vel']:.1f} inf/día",
                f"{K['vel'] - K['meta_prom']:+.1f} vs meta {K['meta_prom']:.0f}" if pd.notna(K["meta_prom"]) else None)
    c = st.columns(4)
    if pd.notna(K["cierre"]) and pd.notna(K["fin_plan"]):
        desf = int(np.busday_count(np.datetime64(K["fin_plan"].date()), np.datetime64(K["cierre"].date()), holidays=HOL))
        c[0].metric("Cierre estimado al ritmo actual", f"{K['cierre']:%d/%m/%Y}",
                    f"{desf:+d} días háb. vs plan ({K['fin_plan']:%d/%m})", delta_color="inverse")
    else:
        c[0].metric("Cierre estimado al ritmo actual", "—", "sin entregas recientes", delta_color="off")
    c[1].metric("Pendientes con restricción externa", f"{len(RESTR)}",
                f"{len(RESTR) / len(PEND) * 100:.0f}% del backlog" if len(PEND) else None, delta_color="off")
    c[2].metric("Entregas registradas (detalle diario)", f"{len(ENT_C)}",
                f"{int(ENT_C['Reentrega'].sum())} reentregas", delta_color="off")
    c[3].metric("Antigüedad máxima de un pendiente", f"{int(PEND['Antigüedad (días háb.)'].max())} días háb."
                if PEND["Antigüedad (días háb.)"].notna().any() else "—")

    g1, g2 = st.columns([2, 1])
    with g1:
        mostrar(fig_burndown(S, PLAN, K, corte, f"Curva de cierre del backlog {' + '.join(alcance)}: plan vs real"))
        st.caption("Tramo sin detalle diario (entre listas de corte) unido en línea recta. La proyección naranja "
                   "extiende la productividad de la ventana seleccionada.")
    with g2:
        g = INF.groupby(["Grupo", "Estado"]).size().unstack(fill_value=0).reindex(columns=["Entregado", "Pendiente"], fill_value=0)
        g["Total"] = g.sum(axis=1)
        fig = go.Figure()
        fig.add_bar(x=g.index, y=g["Entregado"], name="Entregado", marker_color=AZUL, text=g["Entregado"])
        fig.add_bar(x=g.index, y=g["Pendiente"], name="Pendiente", marker_color=NARANJA, text=g["Pendiente"])
        for grp, r in g.iterrows():
            fig.add_annotation(x=grp, y=r["Total"], text=f"<b>{r['Entregado'] / r['Total'] * 100:.0f}%</b>",
                               showarrow=False, yshift=12)
        fig.update_layout(barmode="stack", yaxis_title="N° informes")
        mostrar(estilo(fig, 400, "Estado por grupo (% avance)"))

    st.markdown("##### 🧠 Lectura gerencial")
    notas = []
    if pd.notna(K["spi"]):
        est = "por encima" if K["spi"] >= 100 else "por debajo"
        notas.append(f"El avance real acumulado del alcance {' + '.join(alcance)} ({K['real_acum']} informes desde el "
                     f"{K['d0'] + pd.Timedelta(days=1):%d/%m}) está **{est} del plan** ({K['plan_acum']:.0f}): "
                     f"cumplimiento de **{K['spi']:.0f}%**.")
    if K.get("crec_alcance"):
        pl_hoy = serie_plan(PLAN, K["b0"])
        pl_hoy = pl_hoy[pl_hoy["Fecha"] <= corte]["Plan"].iloc[-1] if len(pl_hoy) else np.nan
        notas.append(f"El alcance creció en **{K['crec_alcance']}** informes desde la línea base del "
                     f"{K['d0']:%d/%m} ({K['b0']} → {K['total']}); por ello, aun entregando más de lo planificado, "
                     f"el backlog real (**{K['pendientes']}**) se compara con **{pl_hoy:.0f}** del plan a la fecha.")
    if pd.notna(K["cierre"]) and pd.notna(K["fin_plan"]):
        notas.append(f"Con **{K['vel']:.1f} informes/día** (meta {K['meta_prom']:.0f}), los **{K['pendientes']} pendientes** "
                     f"se cerrarían hacia el **{K['cierre']:%d/%m/%Y}** frente al **{K['fin_plan']:%d/%m/%Y}** del plan"
                     + (" — se requiere reforzar capacidad o priorizar." if K["cierre"] > K["fin_plan"] else "."))
        if K["cierre"] > K["fin_plan"] and K["pendientes"]:
            dias_rest = max(1, int(np.busday_count(np.datetime64(corte.date()), np.datetime64(K["fin_plan"].date()), holidays=HOL)))
            notas.append(f"Para cumplir la fecha del plan se necesitan **{K['pendientes'] / dias_rest:.1f} informes/día hábil**.")
    if len(RESTR):
        top = RESTR["Restricción"].replace("", "Sin restricción registrada").value_counts()
        notas.append(f"**{len(RESTR)}** pendientes dependen de gestiones externas; la principal es **{top.index[0]}** "
                     f"({top.iloc[0]} informes, {top.iloc[0] / len(RESTR) * 100:.0f}%). Conviene tratarla como punto "
                     "fijo en la reunión de seguimiento con Repsol.")
    ven = ENT_C[(ENT_C["Fecha"] >= K["vel_ini"]) & ENT_C["Código"].isin(S["Código"]) & ~ENT_C["Reentrega"]]
    fin_sem = ven[~ven["Fecha"].map(es_habil)]
    if len(fin_sem):
        solo_hab = (len(ven) - len(fin_sem)) / n_vel
        notas.append(f"De las entregas del alcance en la ventana, **{len(fin_sem)}** se hicieron en fin de semana o "
                     f"feriado (días sin meta). Sin ese esfuerzo adicional la productividad sería de **{solo_hab:.1f} "
                     f"informes/día hábil**" + (", por debajo de la meta: el cumplimiento actual depende de horas extra."
                                                  if pd.notna(K["meta_prom"]) and solo_hab < K["meta_prom"] else "."))
    if n_incons:
        notas.append(f"Se detectaron **{n_incons}** registros por conciliar (ver pestaña *Calidad de Datos*).")
    for nt in notas:
        st.markdown(f"- {nt}")
    acciones_rapidas("t1", corte)

# ─────────────────────────── 2. PLAN VS REAL ───────────────────────────
with tabs[1]:
    alc_txt = " + ".join(alcance)
    if PLAN.empty:
        sin_datos("Registre el plan de entrega (Gestión de Matriz › Plan de entrega) para ver el control plan vs real.")
    else:
        st.markdown(f"<div style='font-size:.8rem;letter-spacing:.08em;color:{C_REAL};font-weight:700'>AVANCE ACUMULADO</div>",
                    unsafe_allow_html=True)
        g1, g2 = st.columns([4, 1])
        with g1:
            fig, p_c, r_c = fig_avance_acumulado(S, PLAN, K, corte, alc_txt)
            mostrar(fig)
        with g2:
            brecha = r_c - p_c
            cumpl = r_c / p_c * 100 if p_c else np.nan
            st.markdown("<div style='height:3.2rem'></div>" +
                        tarjeta(f"{p_c:.0f}", f"Acumulado programado al {corte:%d/%m}", "#7F8C8D") +
                        tarjeta(f"{r_c}", f"Acumulado real al {corte:%d/%m}", C_REAL) +
                        tarjeta(f"{brecha:+.0f}".replace("-", "−"),
                                f"Brecha · {cumpl:.1f}% de cumplimiento" if pd.notna(cumpl) else "Brecha", "#1C2833"),
                        unsafe_allow_html=True)
        st.caption("El real acumulado cuenta informes del alcance entregados desde el inicio del plan (primera entrega). "
                   "El tramo sin detalle diario se une en línea recta hasta la fecha de corte de las listas.")

        st.markdown(f"<div style='font-size:.8rem;letter-spacing:.08em;color:{C_REAL};font-weight:700;margin-top:1rem'>"
                    "ENTREGAS DIARIAS</div>", unsafe_allow_html=True)
        fig, hay_extra = fig_plan_diario(INF, PLAN, alcance, corte, K["meta_prom"])
        mostrar(fig)
        st.caption("Real por grupo = primera entrega de cada informe (sin reentregas). Las barras rayadas distribuyen en "
                   "promedio las entregas sin detalle diario."
                   + (" Los días marcados con * son fines de semana o feriados sin meta en los que hubo entregas." if hay_extra else ""))

        st.markdown(f"<div style='font-size:.8rem;letter-spacing:.08em;color:{C_REAL};font-weight:700;margin-top:1rem'>"
                    "HITOS DE CONTROL</div>", unsafe_allow_html=True)
        st.markdown(f"#### Puntos de control semanales · {alc_txt}")
        fin_p = K["fin_plan"] if pd.notna(K["fin_plan"]) else PLAN["Fecha"].max()
        viernes = [d for d in PLAN["Fecha"] if d.weekday() == 4 and d > K["d0"] and d <= fin_p]
        opciones_h = sorted(set(PLAN["Fecha"]) | {corte})
        defecto = sorted({corte, fin_p} | {v for v in viernes if v >= corte - pd.Timedelta(days=14)})
        with st.expander("Configurar puntos de control"):
            sel_h = st.multiselect("Fechas de control", opciones_h, default=[d for d in defecto if d in opciones_h],
                                   format_func=lambda d: f"{DIAS_ES[d.weekday()]} {d:%d/%m/%Y}", key="hitos_sel")
        th = tabla_hitos(S, PLAN, K, corte, sel_h or defecto)
        marca = th.pop("_corte")
        sty = (th.style
               .apply(lambda r: ["background-color: rgba(214,204,186,.45); font-weight:600"
                                 if marca[r.name] else "" for _ in r], axis=1)
               .apply(lambda c: [("color:#B03A2E;font-weight:600" if isinstance(v, (int, float)) and pd.notna(v) and v < 0
                                  else "color:#148F77;font-weight:600" if isinstance(v, (int, float)) and pd.notna(v) else "")
                                 for v in c], subset=["Brecha"])
               .format({"% avance programado": "{:.1f}%", "% avance real": lambda v: "—" if pd.isna(v) else f"{v:.1f}%",
                        "Acumulado real": lambda v: "—" if pd.isna(v) else f"{int(v)}",
                        "Pendiente real": lambda v: "—" if pd.isna(v) else f"{int(v)}",
                        "Brecha": lambda v: "—" if pd.isna(v) else f"{int(v):+d}"}))
        st.dataframe(sty, hide_index=True, use_container_width=True)
        ini8 = PLAN[PLAN["Meta Diaria"] >= (K["meta_prom"] or 1)]["Fecha"].min() if pd.notna(K["meta_prom"]) else None
        st.caption((f"Programa: {K['meta_prom']:.0f} informes por día hábil a partir del {ini8:%d/%m}. " if ini8 is not None and pd.notna(ini8) else "")
                   + f"Línea base {K['b0']} informes al {K['d0']:%d/%m}. Pendiente real incluye informes incorporados "
                   "al alcance después de la línea base. El control real se actualiza en cada corte.")
    acciones_rapidas("tpv", corte)

# ─────────────────────── 3. AVANCE Y PRODUCTIVIDAD ───────────────────────
with tabs[2]:
    if ENT_C.empty:
        sin_datos("No hay entregas diarias registradas hasta la fecha de corte.")
    else:
        dia = ENT_C.assign(Serie=np.where(ENT_C["Reentrega"], "Reentrega", ENT_C["Grupo maestro"].fillna("G1")))
        d = dia.groupby(["Fecha", "Serie"]).size().reset_index(name="N")
        fig = px.bar(d, x="Fecha", y="N", color="Serie", color_discrete_map=COLOR_GRUPO, text="N",
                     category_orders={"Serie": ["G1", "G2", "G3", "Reentrega"]})
        if len(PLAN):
            p = PLAN[(PLAN["Fecha"] >= ENT_C["Fecha"].min()) & (PLAN["Fecha"] <= corte)]
            fig.add_scatter(x=p["Fecha"], y=p["Meta Diaria"], name="Meta diaria", mode="lines+markers",
                            line=dict(color="#17202A", dash="dash", shape="hv"))
        fig.update_layout(barmode="stack", yaxis_title="Informes entregados", xaxis_title=None)
        fig.update_xaxes(tickformat="%a %d/%m", dtick=86400000)
        mostrar(estilo(fig, 390, "Entregas diarias por grupo vs meta"))

        g1, g2 = st.columns(2)
        with g1:
            if len(PLAN):
                d0 = K["d0"]
                pl = PLAN[PLAN["Fecha"] > d0].copy()
                pl["Plan acumulado"] = pl["Meta Diaria"].cumsum()
                rp = entregas_por_dia(S)
                rp = rp[(rp.index > d0) & (rp.index <= corte)].cumsum()
                fig = go.Figure()
                fig.add_scatter(x=pl["Fecha"], y=pl["Plan acumulado"], name="Plan acumulado", mode="lines",
                                line=dict(color=GRIS, dash="dash", width=2.5))
                fig.add_scatter(x=rp.index, y=rp.values, name="Real acumulado", mode="lines+markers",
                                line=dict(color=AZUL, width=3))
                fig.add_hline(y=K["b0"], line_color=NARANJA, line_dash="dot",
                              annotation_text=f"Alcance inicial {K['b0']}", annotation_position="top left")
                fig.update_xaxes(tickformat="%d/%m")
                fig.update_layout(yaxis_title="Informes entregados", hovermode="x unified")
                mostrar(estilo(fig, 380, f"Curva S: entregas acumuladas {' + '.join(alcance)}"))
            else:
                sin_datos("Registre el plan de entrega para ver la curva S.")
        with g2:
            t = ENT_C.groupby(["Tipo (GP_AD)", "Grupo maestro"]).size().reset_index(name="N")
            t["Tipo (GP_AD)"] = t["Tipo (GP_AD)"].replace("", "Sin tipo")
            orden = t.groupby("Tipo (GP_AD)")["N"].sum().sort_values().index.tolist()
            fig = px.bar(t, x="N", y="Tipo (GP_AD)", color="Grupo maestro", orientation="h", text="N",
                         color_discrete_map=COLOR_GRUPO, category_orders={"Tipo (GP_AD)": orden[::-1]})
            fig.update_layout(barmode="stack", xaxis_title="Entregas", yaxis_title=None)
            mostrar(estilo(fig, 380, "Mix de entregas por tipo (GP_AD)"))

        st.markdown("##### 📉 Estabilidad de la productividad diaria (gráfico I-MR)")
        hab = ENT_C.groupby("Fecha").size()
        rango = pd.date_range(ENT_C["Fecha"].min(), corte)
        hab = hab.reindex(rango, fill_value=0)
        hab = hab[[es_habil(x) for x in hab.index]]
        if len(hab) >= 5:
            x = hab.to_numpy(dtype=float)
            media, mr = x.mean(), np.abs(np.diff(x)).mean()
            lsc, lic = media + 2.66 * mr, max(0.0, media - 2.66 * mr)
            fig = go.Figure()
            fig.add_scatter(x=hab.index, y=x, mode="lines+markers+text", text=x.astype(int), textposition="top center",
                            name="Entregas por día hábil", line=dict(color=AZUL_CL), marker=dict(color=AZUL, size=8))
            fuera = (x > lsc) | (x < lic)
            fig.add_scatter(x=hab.index[fuera], y=x[fuera], mode="markers", name="Fuera de control",
                            marker=dict(color=ROJO, size=13, symbol="diamond"))
            for y, n, col, dsh in [(media, "Media", "#17202A", "solid"), (lsc, "LSC", NARANJA, "dash"), (lic, "LIC", NARANJA, "dash")]:
                fig.add_hline(y=y, line_color=col, line_dash=dsh, annotation_text=f"{n} {y:.1f}", annotation_position="right")
            if pd.notna(K["meta_prom"]):
                fig.add_hline(y=K["meta_prom"], line_color=VERDE_AZ, line_dash="dot",
                              annotation_text=f"Meta {K['meta_prom']:.0f}", annotation_position="left")
            fig.update_xaxes(tickformat="%a %d/%m")
            fig.update_layout(yaxis_title="Informes / día")
            mostrar(estilo(fig, 360))
            st.caption(f"Media {media:.1f} · LSC {lsc:.1f} · LIC {lic:.1f} (n={len(x)} días hábiles). Con pocos días los "
                       "límites son orientativos; si la meta cae fuera de la banda, el proceso actual no la alcanza de "
                       "forma sostenida y requiere un cambio de capacidad o método.")
        else:
            sin_datos("Se requieren al menos 5 días hábiles con registro para el gráfico de control.")

        st.markdown("##### 📋 Plan vs real por día")
        if len(PLAN):
            real_d = ENT_C.groupby("Fecha").size()
            real_s = entregas_por_dia(S)
            tab = PLAN[(PLAN["Fecha"] >= ENT_C["Fecha"].min()) & (PLAN["Fecha"] <= corte)].copy()
            tab["Real (todos)"] = tab["Fecha"].map(real_d).fillna(0).astype(int)
            tab[f"Real {'+'.join(alcance)}"] = tab["Fecha"].map(real_s).fillna(0).astype(int)
            tab["Desvío alcance"] = tab[f"Real {'+'.join(alcance)}"] - tab["Meta Diaria"]
            st.dataframe(tab, hide_index=True, use_container_width=True,
                         column_config={"Fecha": st.column_config.DateColumn(format="ddd DD/MM/YYYY")})
    acciones_rapidas("t2", corte)

# ─────────────────────── 3. BACKLOG Y RESTRICCIONES ───────────────────────
with tabs[3]:
    if PEND.empty:
        st.success("No hay informes pendientes. ✅")
    else:
        g1, g2 = st.columns(2)
        with g1:
            b = PEND.copy()
            b["Antigüedad"] = pd.cut(b["Antigüedad (días háb.)"].fillna(-1), [-np.inf, -0.5, 5, 10, 20, np.inf],
                                     labels=["Sin fecha", "0–5", "6–10", "11–20", ">20"]).astype(str)
            a = b.groupby(["Antigüedad", "Grupo"]).size().reset_index(name="N")
            fig = px.bar(a, x="Antigüedad", y="N", color="Grupo", text="N", color_discrete_map=COLOR_GRUPO,
                         category_orders={"Antigüedad": ["0–5", "6–10", "11–20", ">20", "Sin fecha"], "Grupo": GRUPOS})
            fig.update_layout(barmode="stack", xaxis_title="Días hábiles desde el ingreso al backlog", yaxis_title="Pendientes")
            mostrar(estilo(fig, 380, "Antigüedad del backlog pendiente"))
        with g2:
            if len(RESTR):
                p = RESTR["Restricción"].replace("", "Sin restricción registrada").value_counts().rename_axis("Restricción").reset_index(name="N")
                p["Acum"] = p["N"].cumsum() / p["N"].sum() * 100
                vital = p["Acum"].shift(fill_value=0) < 80
                fig = make_subplots(specs=[[{"secondary_y": True}]])
                fig.add_bar(x=p["Restricción"], y=p["N"], text=p["N"], name="Pendientes",
                            marker_color=np.where(vital, NARANJA, GRIS))
                fig.add_scatter(x=p["Restricción"], y=p["Acum"], name="% acumulado", mode="lines+markers",
                                line=dict(color=AZUL, width=2.5), secondary_y=True)
                fig.add_scatter(x=p["Restricción"], y=[80] * len(p), name="80%", mode="lines",
                                line=dict(color="#17202A", dash="dot"), secondary_y=True)
                fig.update_yaxes(title_text="Pendientes", secondary_y=False)
                fig.update_yaxes(range=[0, 105], showgrid=False, secondary_y=True, title_text="% acumulado")
                mostrar(estilo(fig, 380, "Pareto de restricciones externas (G3)"))
            else:
                sin_datos("No hay pendientes con restricción registrada.")

        tm = PEND.assign(Plan=PEND["Plan"].replace("", "Sin plan"),
                         Restricción=PEND["Restricción"].replace("", "Sin restricción"))
        fig = px.treemap(tm, path=[px.Constant("Backlog pendiente"), "Grupo", "Plan", "Restricción"],
                         color="Grupo", color_discrete_map={**COLOR_GRUPO, "(?)": "#EAF2F8"})
        fig.update_traces(textinfo="label+value")
        mostrar(estilo(fig, 420, "Composición del backlog pendiente: grupo › plan › restricción"))

        st.markdown("##### 📋 Detalle de pendientes")
        c1, c2 = st.columns([1, 2])
        fg = c1.multiselect("Grupo", GRUPOS, default=GRUPOS, key="bk_g")
        q = c2.text_input("Buscar (N°, TAG, restricción…)", key="bk_q")
        v = PEND[PEND["Grupo"].isin(fg)]
        if q:
            v = v[v.astype(str).apply(lambda r: r.str.contains(q, case=False)).any(axis=1)]
        st.dataframe(v[["Código", "Grupo", "Plan", "TAG / Circuito", "Restricción", "Fecha Ingreso",
                        "Antigüedad (días háb.)", "Observaciones"]].sort_values("Antigüedad (días háb.)", ascending=False),
                     hide_index=True, use_container_width=True,
                     column_config={"Fecha Ingreso": st.column_config.DateColumn(format="DD/MM/YYYY")})
    acciones_rapidas("t3", corte)

# ─────────────────────────── 4. CALIDAD DE DATOS ───────────────────────────
with tabs[4]:
    st.markdown("La confiabilidad del KPI depende de la conciliación entre las listas por grupo, el registro diario y "
                "el plan. Estas verificaciones se recalculan con cada actualización.")
    cols = st.columns(3)
    for i, (k, v) in enumerate(CAL.items()):
        cols[i % 3].metric(k, len(v))
    for k, v in CAL.items():
        if len(v):
            with st.expander(f"{k} ({len(v)})"):
                st.dataframe(v, hide_index=True, use_container_width=True,
                             column_config={c: st.column_config.DateColumn(format="DD/MM/YYYY")
                                            for c in ["Fecha", "Fecha Ingreso", "Fecha Entrega"] if c in v.columns})
    if st.session_state.notas:
        st.markdown("##### 📝 Notas de conciliación de la última importación")
        for nt in st.session_state.notas:
            st.markdown(f"- {nt}")

# ─────────────────────────── 5. GESTIÓN DE MATRIZ ───────────────────────────
with tabs[5]:
    s1, s2, s3, s4, s5, s6, s7 = st.tabs(["📬 Registrar entregas", "➕ Nuevo informe", "✏️ Maestro de informes",
                                          "📅 Registro diario", "🗓️ Plan de entrega", "⬆️ Subir matriz", "⬇️ Descargar"])
    with s1:
        form_entregas("gm_e", corte)
    with s2:
        form_informe("gm_i")

    def _editor(df, key, cfg, motivo, cual):
        if not puede_editar():
            st.warning("🔒 Ingrese la clave de edición en la barra lateral para modificar la matriz.")
            st.dataframe(df, hide_index=True, use_container_width=True)
            return
        ed = st.data_editor(df, num_rows="dynamic", hide_index=True, use_container_width=True,
                            column_config=cfg, key=key, height=480)
        c1, c2, _ = st.columns([1, 1, 3])
        if c1.button("💾 Guardar cambios", type="primary", key=f"{key}_g"):
            s = st.session_state
            args = {"inf": s.inf, "ent": s.ent, "plan": s.plan}
            args[cual] = ed
            guardar(args["inf"], args["ent"], args["plan"], motivo)
            st.rerun()
        if c2.button("↩️ Descartar", key=f"{key}_d"):
            st.session_state.pop(key, None)
            st.rerun()

    fecha_cfg = lambda c: st.column_config.DateColumn(c, format="DD/MM/YYYY")  # noqa: E731
    with s3:
        st.caption("Una fila por informe. El estado y la fecha de entrega se actualizan automáticamente desde el "
                   "registro diario; los entregados sin registro diario conservan su fecha referencial.")
        _editor(INF, "ed_inf", {
            "Grupo": st.column_config.SelectboxColumn(options=GRUPOS),
            "Estado": st.column_config.SelectboxColumn(options=["Pendiente", "Entregado"]),
            "Plan": st.column_config.SelectboxColumn(options=[""] + sorted(set(PLANES) | set(INF["Plan"]) - {""})),
            "Restricción": st.column_config.SelectboxColumn(options=[""] + sorted(set(RESTRICCIONES) | set(INF["Restricción"]) - {""})),
            "Adicional": st.column_config.TextColumn(disabled=True),
            "N° Informe": st.column_config.NumberColumn(disabled=True),
            "Fecha Ingreso": fecha_cfg("Fecha Ingreso"), "Fecha Entrega": fecha_cfg("Fecha Entrega"),
        }, "Maestro de informes actualizado", "inf")
    with s4:
        st.caption("Registro de entregas diarias (equivale a «Avance Diario»). Puede ingresar solo el N° de informe: "
                   "el código se completa automáticamente priorizando el grupo indicado.")
        _editor(ENT, "ed_ent", {
            "Fecha": fecha_cfg("Fecha"),
            "Grupo": st.column_config.SelectboxColumn(options=GRUPOS),
            "Tipo (GP_AD)": st.column_config.SelectboxColumn(options=sorted(set(TIPOS_GPAD) | set(ENT["Tipo (GP_AD)"]) - {""})),
            "N° Informe": st.column_config.NumberColumn(step=1),
        }, "Registro diario actualizado", "ent")
    with s5:
        st.caption("Meta diaria de entregas del plan. Use 0 en fines de semana y feriados.")
        _editor(PLAN, "ed_plan", {"Fecha": fecha_cfg("Fecha"),
                                  "Meta Diaria": st.column_config.NumberColumn(min_value=0.0, step=1.0, format="%.2f")},
                "Plan de entrega actualizado", "plan")
    with s6:
        st.markdown("Suba el archivo **KPI_ENTREGA_DE_INFORMES.xlsx** en su formato original (hojas *INFORMES G1 G2 G3*, "
                    "*PENDIENTES G3*, *Plan de entrega* y *Avance Diario*) o la matriz exportada por este tablero. "
                    "En el formato original, el relleno amarillo indica «Informe Entregado».")
        carga_matriz("gm")
    with s7:
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ Matriz del tablero (INFORMES · ENTREGAS · PLAN)", excel_canonico(INF, ENT, PLAN, st.session_state.notas),
                           file_name=f"MATRIZ_KPI_ENTREGA_INFORMES_{HOY:%Y%m%d}.xlsx", use_container_width=True)
        c2.download_button("⬇️ Formato de seguimiento (G1 G2 G3 · Plan · Avance Diario)",
                           excel_formato_seguimiento(INF, ENT, PLAN, corte, alcance),
                           file_name=f"KPI_ENTREGA_DE_INFORMES_{corte:%Y%m%d}.xlsx", use_container_width=True)
        st.caption("El formato de seguimiento regenera las hojas habituales con valores recalculados (el real del "
                   "plan corresponde solo al alcance seleccionado). Es un reporte: para volver a subir datos use la "
                   "matriz del tablero, que conserva el historial completo.")

st.caption("Backlog = informes ingresados a la fecha y no entregados. Cumplimiento del plan = entregas reales del alcance "
           "desde el inicio del plan ÷ meta acumulada a la fecha de corte. Días hábiles: lunes a viernes sin feriados del Perú.")
