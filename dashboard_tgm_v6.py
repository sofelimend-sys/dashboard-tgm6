"""
Panel TGM · Transvisa E.I.R.L.  (v6)
Baja eficacia del cumplimiento de mantenimiento preventivo · 5 causas raíz · ADDIE · TPM · SMED · ISOS · DMAIC

Los datos se leen de `datos_tgm.json` (generado con `extraer_datos_excel.py`); el panel no inventa ni simula valores.

Ejecutar:   streamlit run dashboard_tgm_v6.py
"""
import base64
import io
import json
import os
import re
import textwrap

import streamlit as st

import numpy as np
import pandas as pd
import plotly.graph_objects as go

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_JSON = os.path.join(AQUI, "datos_tgm.json")
RUTA_LOGO = os.path.join(AQUI, "logo_b64.txt")
LOGO_B64 = open(RUTA_LOGO).read().strip() if os.path.exists(RUTA_LOGO) else ""


def _icono():
    try:
        from PIL import Image
        return Image.open(io.BytesIO(base64.b64decode(LOGO_B64)))
    except Exception:
        return None


st.set_page_config(page_title="TGM | Mantenimiento preventivo", page_icon=_icono() or "🚛", layout="wide")


def _ancho(fn):
    """Compatibilidad: width='stretch' (Streamlit reciente) o use_container_width=True (versiones anteriores)."""
    def w(*a, **k):
        if k.pop("use_container_width", None):
            try:
                return fn(*a, width="stretch", **k)
            except TypeError:
                return fn(*a, use_container_width=True, **k)
        return fn(*a, **k)
    return w


_st_button, _st_download = st.button, st.download_button
st.button, st.download_button = _ancho(_st_button), _ancho(_st_download)


@st.cache_data(show_spinner=False)
def cargar_datos(mtime):
    with open(RUTA_JSON, encoding="utf-8") as f:
        return json.load(f)


if not os.path.exists(RUTA_JSON):
    st.error("No se encontró `datos_tgm.json`. Ejecuta primero:  python extraer_datos_excel.py T3_-_Equipo_-_06.xlsm datos_tgm.json")
    st.stop()
D = cargar_datos(os.path.getmtime(RUTA_JSON))

# ======================= PALETA CORPORATIVA (sin cambios: azul marino + rojo) =======================
NAVY, NAVY2, AZUL, AZUL_CLARO = "#0B1F4B", "#173A7A", "#2457A6", "#6B93D6"
ROJO, ROJO_OSC, GRIS, NEGRO = "#C8102E", "#8E0C22", "#5B6475", "#000000"
SEM_COL = {"verde": "#2E9E5B", "amarillo": "#F2B01E", "rojo": ROJO}
SEM_TINT = {"verde": "#D6EFE0", "amarillo": "#FCEFC7", "rojo": "#F6CBD2"}
SEM_TXT = {"verde": "Cumple", "amarillo": "En proceso", "rojo": "Crítico"}

META = {
    "ADDIE": dict(color=NAVY, corto="Capacitación"),
    "TPM": dict(color=AZUL_CLARO, corto="Disponibilidad"),
    "SMED": dict(color=ROJO_OSC, corto="Setup"),
    "ISOS": dict(color=AZUL, corto="Tiempos muertos"),
    "DMAIC": dict(color=ROJO, corto="Reprocesos"),
}
CAUSAS = D["causas"]
POR_ID = {c["id"]: c for c in CAUSAS}
POR_COD = {c["codigo"]: c for c in CAUSAS}
IDS = [c["id"] for c in CAUSAS]
MESES = D["meses"]
SEM = D["semaforo"]["indicadores"]
FL = D["flujo"]
PB = D["problema"]
PC = "P. CENTRAL"
FUENTES = ["Serie base", "Serie de semaforización"]


# ======================= FORMATO =======================
def soles(x, dec=0):
    return f"S/ {x:,.{dec}f}"


def f_val(x, u, unidad=True):
    if x is None:
        return "N/D"
    if u == "%":
        return f"{x * 100:.1f}%"
    if u == "min/mes":
        return f"{x:,.1f}" + (" min/mes" if unidad else "")
    return f"{x:,.2f}" + (" h/mes" if unidad else "")


def num_u(x, u):
    """Valor y unidad por separado (la unidad se muestra en tamaño menor)."""
    if x is None:
        return "N/D", ""
    if u == "%":
        return f"{x * 100:.1f}%", ""
    if u == "min/mes":
        return f"{x:,.1f}", "min/mes"
    return f"{x:,.2f}", "h/mes"


def val_html(x, u):
    n, un = num_u(x, u)
    return n + (f"<small>{un}</small>" if un else "")


def pct_signo(m):
    return f"{'+' if m >= 0 else '−'}{abs(m) * 100:.1f}%"


def disp(x, u):
    return np.asarray(x, dtype=float) * (100 if u == "%" else 1)


def etq_u(u):
    return "%" if u == "%" else u


def mayor(c):
    return str(c["sentido"]).lower().startswith("mayor")


def cumple(valor, est, es_mayor):
    return valor >= est - 1e-12 if es_mayor else valor <= est + 1e-12


def estado_sem(cod, valor):
    """Verde si alcanza la meta; rojo si está en la línea base o peor; amarillo entre ambas."""
    s = SEM[cod]
    es_mayor = str(s["sentido"]).lower().startswith("mayor")
    meta, base = s["meta"], s["linea_base"]
    if es_mayor:
        return "verde" if valor >= meta - 1e-9 else ("rojo" if valor <= base + 1e-6 else "amarillo")
    return "verde" if valor <= meta + 1e-9 else ("rojo" if valor >= base - 1e-6 else "amarillo")


def logro(valor, est, es_mayor):
    return valor / est if es_mayor else est / valor


def mejora_pct(a, t, es_mayor):
    return (t - a) / a if es_mayor else (a - t) / a


def nn(c):
    return c["id"].lstrip("Cc").zfill(2)


def nombre(c):
    return META[c["codigo"]]["corto"]


def titulo_pb():
    t = str(PB["titulo"])
    t = t.capitalize() if t.isupper() else t
    return re.sub(r"\btgm\b", "TGM", t, flags=re.I)


def md(html):
    """Markdown/HTML en una sola línea (evita que líneas en blanco o sangrías rompan el bloque)."""
    st.markdown(re.sub(r"\s*\n\s*", " ", html), unsafe_allow_html=True)


# ======================= ESTADO (selección múltiple de causas) =======================
S = st.session_state
S.setdefault("pagina", "panel")
S.setdefault("sel", [])            # ids de causas elegidas; vacío = se muestran las 5
S.setdefault("nonce", 0)
S.setdefault("meses_sel", list(S.get("_u_meses_sel", [])))      # vacío = todo el año
S.setdefault("fuente", FUENTES[0])


def ir(p):
    S.pagina = p


def alternar(cid):
    s = set(S.sel)
    s.symmetric_difference_update({cid})
    S.sel = [i for i in IDS if i in s]


def todas():
    S.sel = list(IDS)


def ninguna():
    S.sel = []


def activas():
    """Causas que muestra el panel: las elegidas o, sin selección, las cinco."""
    return [c for c in CAUSAS if c["id"] in S.sel] or list(CAUSAS)


def resumen_sel():
    cs = activas()
    if len(cs) == len(CAUSAS):
        return "Visión global", "Las cinco causas raíz"
    if len(cs) == 1:
        c = cs[0]
        return f"{nn(c)} · {nombre(c)} ({c['codigo']})", f"Indicador: {c['indicador_formula'].replace(chr(10), ' ')}"
    return f"{len(cs)} causas seleccionadas", " · ".join(f"{nn(c)} {c['codigo']}" for c in cs)


def selector(opciones, key, tipo="pills"):
    """Selector de opción única que nunca queda vacío."""
    if S.get(key) not in opciones:
        S[key] = S.get("_u_" + key, opciones[0])

    def _fix():
        if S.get(key) is None:
            S[key] = S.get("_u_" + key, opciones[0])
        S["_u_" + key] = S[key]
    fn = getattr(st, "segmented_control" if tipo == "seg" else "pills", None)
    if fn is not None:
        fn(key, opciones, selection_mode="single", key=key, on_change=_fix, label_visibility="collapsed")
    else:
        st.radio(key, opciones, key=key, horizontal=True, label_visibility="collapsed")
    S["_u_" + key] = S[key]
    return S[key]


# ======================= SERIES Y CÁLCULOS =======================
def serie(cod, esc):
    if S.fuente == FUENTES[1] and cod in SEM:
        return list(SEM[cod]["mensual_asis" if esc == "AS-IS" else "mensual_tobe"])
    return list(D["mensual_base"][cod][esc])


def idx_meses():
    sel = [m for m in S.meses_sel if m in MESES]
    return [MESES.index(m) for m in (sel or MESES)]


def calc(c, idx=None):
    idx = idx_meses() if idx is None else idx
    cod, u, may = c["codigo"], c["unidad"], mayor(c)
    a = np.array(serie(cod, "AS-IS"))[idx]
    t = np.array(serie(cod, "TO-BE"))[idx]
    if len(idx) == 12:
        if S.fuente == FUENTES[1]:
            as_is, to_be = SEM[cod]["anual_asis"], SEM[cod]["anual_tobe"]
        else:
            as_is, to_be = c["as_is"], c["to_be"]
    else:
        as_is, to_be = float(a.mean()), float(t.mean())
    est = c["estandar"]
    return dict(a=a, t=t, as_is=as_is, to_be=to_be, est=est, may=may, u=u, idx=idx,
                mejora=mejora_pct(as_is, to_be, may), logro_as=logro(as_is, est, may), logro_to=logro(to_be, est, may),
                cumple_to=cumple(to_be, est, may), cumple_as=cumple(as_is, est, may),
                sem_as=estado_sem(cod, as_is), sem_to=estado_sem(cod, to_be),
                brecha_as=(as_is - est) * (1 if may else -1), brecha_to=(to_be - est) * (1 if may else -1))


def meses_txt(idx):
    return "Todo el año" if len(idx) == 12 else ", ".join(MESES[i] for i in idx)


def inversion_de(cod):
    return next(x for x in FL["inversion_por_causa"] if x["propuesta"] == cod)


def eval_de(cod):
    return next(x for x in FL["eval_causas"] if x["propuesta"] == cod)


def corto(nombre_):
    t = str(nombre_)
    if ":" in t:
        t = t.split(":", 1)[1]
    elif ". " in t[:4]:
        t = t.split(". ", 1)[1]
    elif " - " in t[:6]:
        t = t.split(" - ", 1)[1]
    siglas = {"ISOS", "ERP", "OT", "AS-IS", "TO-BE", "SMED", "TPM", "(BPMN)"}
    out = [p if p.upper() in siglas else p.lower() for p in t.strip().split()]
    s = " ".join(out)
    return s[:1].upper() + s[1:]


# ======================= ESTILOS =======================
CSS = r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Semi+Condensed:wght@500;600;700&display=swap');
:root, html, body, .stApp { color-scheme: light !important; }
:root { --txt: 'Barlow','Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif; --disp: 'Barlow Semi Condensed','Barlow','Arial Narrow','Segoe UI',Arial,sans-serif; }
html { font-size: 16px !important; }
.stApp { background: linear-gradient(135deg,#EAF0FA 0%,#FFFFFF 52%,#FCEFF1 100%) !important; }
[data-testid="stHeader"], header[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stAppDeployButton"], .stDeployButton, #MainMenu, footer { display: none !important; }
.block-container { padding: 2.2rem 3rem 4rem; max-width: 1480px; }

/* ---- Tipografía base ---- */
.stApp, .stApp div, .stApp p, .stApp li, .stApp label, .stApp input, .stApp textarea, .stApp button, .stApp table, .stApp th, .stApp td,
.stApp [data-baseweb], .stApp span:not([data-testid="stIconMaterial"]) { font-family: var(--txt) !important; }
.stApp [data-testid="stIconMaterial"], .stApp .material-symbols-rounded { font-family: "Material Symbols Rounded" !important; }
.stApp * { color: __NEGRO__ !important; -webkit-text-fill-color: currentColor !important; }
%M% p, %M% li { font-size: 1rem; line-height: 1.6; }
.stApp [data-testid="stCaptionContainer"] * { font-size: .88rem !important; color: __GRIS__ !important; }

/* ---- Barra lateral ---- */
section[data-testid="stSidebar"] { border-right: 4px solid __ROJO__; }
section[data-testid="stSidebar"] > div { background: linear-gradient(180deg,#0B1F4B 0%,#0E2A66 100%) !important; }
.stApp section[data-testid="stSidebar"] * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; }
[data-testid="stSidebarUserContent"] { padding: 1.2rem 1rem 2rem !important; }
%M% .brand { text-align: center; margin-bottom: .4rem; }
%M% .brand-logo { background: #FFFFFF; border-radius: 14px; padding: 12px 16px; margin-bottom: 14px; border-bottom: 4px solid __ROJO__; box-shadow: 0 4px 12px rgba(0,0,0,.25); }
%M% .brand-logo img { width: 100%; max-width: 170px; display: block; margin: 0 auto; }
%M% .brand-n { font-family: var(--disp) !important; font-weight: 700; font-size: 1.1rem; letter-spacing: .08em; line-height: 1.2; white-space: nowrap; }
%M% .brand-s { font-size: .86rem; letter-spacing: .1em; margin-top: 2px; }
%M% .sb-lab { font-family: var(--disp) !important; font-weight: 600; font-size: .74rem; letter-spacing: .22em; text-transform: uppercase; margin: 1.5rem 0 .45rem; padding-bottom: 7px; border-bottom: 1px solid rgba(255,255,255,.22); }
.stApp section[data-testid="stSidebar"] .stButton button { width: 100%; justify-content: flex-start; text-align: left; background: transparent !important; border: none !important; box-shadow: none !important;
    border-radius: 8px; padding: .5rem .8rem; min-height: 2.45rem; transition: background .15s; }
.stApp section[data-testid="stSidebar"] .stButton button p { font-size: .98rem !important; font-weight: 500 !important; text-align: left; margin: 0; }
.stApp section[data-testid="stSidebar"] .stButton button:hover { background: rgba(255,255,255,.10) !important; }
.stApp section[data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"] { background: rgba(255,255,255,.14) !important; box-shadow: inset 3px 0 0 __ROJO__ !important; }
.stApp section[data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"] p { font-weight: 700 !important; }
/* casillas de las causas */
.stApp section[data-testid="stSidebar"] [class*="st-key-sbc_"] button::before { content: ""; flex: none; width: 17px; height: 17px; margin-right: .75rem; border: 1.5px solid rgba(255,255,255,.65); border-radius: 4px; box-sizing: border-box; }
.stApp section[data-testid="stSidebar"] [class*="st-key-sbc_"] button[data-testid="stBaseButton-primary"]::before { content: "✓"; background: __ROJO__; border-color: __ROJO__; color: #FFFFFF; font-size: 11px; font-weight: 700; line-height: 14px; text-align: center; }
.stApp section[data-testid="stSidebar"] [class*="st-key-sbm_"] button { justify-content: center; border: 1px solid rgba(255,255,255,.35) !important; min-height: 2.1rem; }
.stApp section[data-testid="stSidebar"] [class*="st-key-sbm_"] button p { font-size: .88rem !important; text-align: center; }
.stApp section[data-testid="stSidebar"] [data-testid="stRadio"] label p { font-size: .95rem !important; }

/* ---- Cabecera de página y secciones ---- */
%M% .cab { padding: 2px 0 20px; margin-bottom: 4px; }
%M% .cab-over { font-family: var(--disp) !important; font-weight: 600; font-size: .8rem; letter-spacing: .22em; text-transform: uppercase; color: __ROJO__ !important; }
%M% .cab h1 { font-family: var(--disp) !important; font-weight: 700; font-size: 2.65rem; line-height: 1.04; letter-spacing: -.005em; margin: .4rem 0 .55rem; padding: 0; color: __NAVY__ !important; }
%M% .cab-sub { font-size: 1.1rem; line-height: 1.5; margin: 0; max-width: 62rem; color: __GRIS__ !important; }
%M% .cab-meta { margin-top: .9rem; font-size: .92rem; font-weight: 500; }
%M% .cab-meta span + span::before { content: ""; display: inline-block; width: 4px; height: 4px; border-radius: 50%; background: __ROJO__; margin: 0 .85rem; vertical-align: middle; }
%M% .sec { display: flex; align-items: baseline; gap: .8rem; margin: 2.7rem 0 .9rem; }
%M% .sec::after { content: ""; flex: 1; height: 1px; background: #CFD8E8; transform: translateY(-.35rem); }
%M% .sec h2 { font-family: var(--disp) !important; font-weight: 600; font-size: 1.65rem; line-height: 1.15; margin: 0; padding: 0; color: __NAVY__ !important; }
%M% .sec .sec-n { font-family: var(--disp) !important; font-weight: 600; font-size: .95rem; letter-spacing: .12em; color: __ROJO__ !important; }
%M% .sec-s { font-size: .98rem; margin: -.55rem 0 1.05rem; color: __GRIS__ !important; }
%M% .sub3 { font-family: var(--disp) !important; font-weight: 600; font-size: 1rem; letter-spacing: .1em; text-transform: uppercase; margin: 1.5rem 0 .65rem; color: __NAVY__ !important; }

/* ---- Indicadores (sin recuadros) ---- */
%M% .stat { position: relative; padding-top: 15px; min-height: 124px; }
%M% .stat::before { content: ""; position: absolute; top: 0; left: 0; width: 36px; height: 3px; background: var(--ac); }
%M% .stat-l { font-family: var(--disp) !important; font-weight: 600; font-size: .78rem; letter-spacing: .14em; text-transform: uppercase; color: __GRIS__ !important; }
%M% .stat-v { font-family: var(--disp) !important; font-weight: 700; font-size: clamp(1.7rem, 1.2rem + 1.1vw, 2.45rem); line-height: 1.08; margin: .28rem 0 .3rem; white-space: nowrap; color: __NAVY__ !important; font-variant-numeric: tabular-nums; }
%M% .stat-v small { font-size: .42em; font-weight: 600; letter-spacing: .04em; margin-left: .35em; color: __GRIS__ !important; }
%M% .stat-s { font-size: .9rem; line-height: 1.4; color: __GRIS__ !important; }

/* ---- Estados del semáforo ---- */
%M% .sem { display: inline-flex; align-items: center; gap: .45rem; padding: .18rem .72rem .2rem; border-radius: 999px; background: var(--t); border: 1px solid var(--c); font-family: var(--disp) !important; font-weight: 600; font-size: .78rem; letter-spacing: .09em; text-transform: uppercase; white-space: nowrap; line-height: 1.3; }
%M% .sem i { width: 8px; height: 8px; border-radius: 50%; background: var(--c); display: inline-block; }

/* ---- Tarjetas de causa (botones reales: seleccionar / quitar) ---- */
.stApp [class*="st-key-card_"] { position: relative; gap: 0 !important; }
.stApp [class*="st-key-cardbtn_"] { position: absolute !important; inset: 0; z-index: 4; margin: 0 !important; width: 100% !important; height: 100% !important; }
.stApp [class*="st-key-cardbtn_"] > div, .stApp [class*="st-key-cardbtn_"] [data-testid="stButton"] { width: 100% !important; height: 100% !important; }
.stApp [class*="st-key-cardbtn_"] button { position: absolute !important; inset: 0 !important; width: 100% !important; height: 100% !important; min-height: 0 !important; opacity: 0; cursor: pointer; border-radius: 14px; }
%M% .cc { position: relative; height: 238px; box-sizing: border-box; background: #FFFFFF; border: 1px solid #D5DCE8; border-radius: 14px; padding: 18px 18px 15px;
    display: flex; flex-direction: column; justify-content: space-between; box-shadow: 0 2px 8px rgba(11,31,75,.07); transition: transform .18s, box-shadow .18s, border-color .18s; overflow: hidden; }
%M% .cc::before { content: ""; position: absolute; left: 0; right: 0; top: 0; height: 5px; background: var(--ac); }
.stApp [class*="st-key-card_"]:hover .cc { transform: translateY(-3px); box-shadow: 0 10px 22px rgba(11,31,75,.15); }
%M% .cc.on { border-color: __ROJO__; background: #FFF8F9; box-shadow: 0 0 0 2px __ROJO__, 0 12px 24px rgba(200,16,46,.16); }
%M% .cc-top { display: flex; justify-content: space-between; align-items: center; }
%M% .cc-n { font-family: var(--disp) !important; font-weight: 700; font-size: 1.05rem; letter-spacing: .14em; }
%M% .cc-chk { width: 22px; height: 22px; border-radius: 6px; border: 2px solid #AEBBD6; box-sizing: border-box; display: flex; align-items: center; justify-content: center; background: #FFFFFF; }
%M% .cc.on .cc-chk { background: __ROJO__; border-color: __ROJO__; }
%M% .cc.on .cc-chk::after { content: "✓"; color: #FFFFFF !important; font-size: 13px; font-weight: 700; line-height: 1; }
%M% .cc-nm { font-family: var(--disp) !important; font-weight: 600; font-size: clamp(1.05rem, .9rem + .3vw, 1.25rem); line-height: 1.1; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: __NAVY__ !important; }
%M% .cc-cod { font-size: .76rem; font-weight: 600; letter-spacing: .16em; text-transform: uppercase; color: __GRIS__ !important; }
%M% .cc-v { font-family: var(--disp) !important; font-weight: 700; font-size: 2rem; line-height: 1; white-space: nowrap; font-variant-numeric: tabular-nums; }
%M% .cc-v small { font-size: .4em; font-weight: 600; margin-left: .35em; color: __GRIS__ !important; }
%M% .cc-r { display: flex; justify-content: space-between; align-items: baseline; font-size: .86rem; line-height: 1.5; color: __GRIS__ !important; }
%M% .cc-r b { font-weight: 600; color: __NEGRO__ !important; }
%M% .cc-r b small { font-size: .85em; font-weight: 500; color: __GRIS__ !important; margin-left: .25em; }
%M% .cc-foot { display: flex; justify-content: space-between; align-items: center; }
.stApp [class*="st-key-ctl_"] button { min-height: 2.2rem; padding: .2rem .9rem; }
%M% .selbar { font-size: 1rem; padding-top: .35rem; }
%M% .selbar b { font-weight: 700; }

/* ---- Comparativo (lista de lectura, sin cuadrícula) ---- */
%M% .rk { margin-top: .2rem; }
%M% .rk-h, %M% .rk-f { display: grid; grid-template-columns: 52px minmax(190px,2.1fr) minmax(240px,2.7fr) 120px minmax(150px,1.2fr); gap: 0 22px; align-items: center; }
%M% .rk-h { font-family: var(--disp) !important; font-weight: 600; font-size: .74rem; letter-spacing: .14em; text-transform: uppercase; color: __GRIS__ !important; padding: 0 12px 8px; border-bottom: 2px solid __NAVY__; }
%M% .rk-f { padding: 15px 12px; border-bottom: 1px solid #DCE3EE; transition: opacity .2s, background .2s; border-radius: 0; }
%M% .rk-f.off { opacity: .42; }
%M% .rk-f.on { background: rgba(200,16,46,.06); border-bottom-color: rgba(200,16,46,.35); }
%M% .rk-f.pc { background: rgba(11,31,75,.045); }
%M% .rk-n { font-family: var(--disp) !important; font-weight: 700; font-size: 1.5rem; color: var(--ac) !important; }
%M% .rk-nm { font-family: var(--disp) !important; font-weight: 600; font-size: 1.1rem; line-height: 1.2; color: __NAVY__ !important; }
%M% .rk-id { font-size: .88rem; color: __GRIS__ !important; line-height: 1.35; margin-top: 2px; }

%M% .rk-vals { display: flex; gap: 20px; margin-bottom: 8px; white-space: nowrap; }
%M% .rk-vals i { display: block; font-style: normal; font-size: .7rem; letter-spacing: .1em; text-transform: uppercase; color: __GRIS__ !important; }
%M% .rk-vals b { font-weight: 700; font-size: .98rem; }
%M% .rk-bar { position: relative; height: 14px; }
%M% .rk-bar .t { position: absolute; left: 0; right: 0; top: 5px; height: 4px; border-radius: 4px; background: #E3E8F1; }
%M% .rk-bar .a { position: absolute; left: 0; top: 2px; height: 4px; border-radius: 4px; background: #AEBBD6; }
%M% .rk-bar .b { position: absolute; left: 0; top: 7px; height: 5px; border-radius: 4px; }
%M% .rk-bar .m { position: absolute; top: 0; width: 2px; height: 14px; background: __NEGRO__; }
%M% .rk-mj { font-family: var(--disp) !important; font-weight: 700; font-size: 1.35rem; line-height: 1.1; margin-bottom: 5px; }
%M% .rk-e { text-align: right; font-size: .9rem; line-height: 1.5; color: __GRIS__ !important; }
%M% .rk-e b { font-weight: 600; font-variant-numeric: tabular-nums; color: __NEGRO__ !important; }

/* ---- Notas, lecturas y fichas ---- */
%M% .nota { background: rgba(255,255,255,.75); border: 1px solid #DCE3EE; border-radius: 10px; padding: 12px 16px; font-size: .96rem; line-height: 1.55; margin: 10px 0; }
%M% .nota.roja { background: #FBE3E7; border-color: rgba(200,16,46,.35); }
%M% .lect { background: #FFFFFF; border: 1px solid #DCE3EE; border-radius: 12px; padding: 16px 20px; font-size: 1.02rem; line-height: 1.65; box-shadow: 0 2px 8px rgba(11,31,75,.05); }
%M% .lect b { font-weight: 700; }
%M% .lect .eti, %M% .nota .eti { font-family: var(--disp) !important; font-weight: 600; font-size: .76rem; letter-spacing: .14em; text-transform: uppercase; color: __ROJO__ !important; display: block; margin-bottom: 3px; }
%M% .kv { display: grid; grid-template-columns: 160px 1fr; gap: 9px 22px; margin-top: 12px; }
%M% .kv dt { font-family: var(--disp) !important; font-weight: 600; font-size: .78rem; letter-spacing: .12em; text-transform: uppercase; color: __GRIS__ !important; padding-top: .18rem; }
%M% .kv dd { margin: 0; font-size: 1rem; line-height: 1.55; }
%M% .etapa h4 { font-family: var(--disp) !important; font-weight: 600; font-size: 1.25rem; margin: 0 0 2px; padding: 0; color: __NAVY__ !important; }
%M% .tag { display: inline-block; font-size: .84rem; font-weight: 600; padding: .15rem .7rem; border-radius: 999px; background: #E1EAF8; margin: 0 .4rem .4rem 0; }
%M% .regla { display: grid; grid-template-columns: 120px 1fr auto; gap: 14px; align-items: center; padding: 11px 0; border-bottom: 1px solid #DCE3EE; font-size: .96rem; }
%M% .regla.on { background: rgba(200,16,46,.05); }
%M% .ficha { box-sizing: border-box; height: 150px; display: flex; flex-direction: column; justify-content: space-between; padding: 16px 18px; background: #FFFFFF; border: 1px solid #D5DCE8; border-top: 5px solid var(--ac); border-radius: 14px; box-shadow: 0 2px 8px rgba(11,31,75,.07); }
%M% .ficha .fn { font-family: var(--disp) !important; font-weight: 600; font-size: 1.1rem; color: __NAVY__ !important; }
%M% .ficha .fv { font-family: var(--disp) !important; font-weight: 700; font-size: 1.9rem; line-height: 1; white-space: nowrap; }
%M% .ficha .fv small { font-size: .4em; font-weight: 600; margin-left: .35em; color: __GRIS__ !important; }
%M% .ficha .fs { font-size: .86rem; color: __GRIS__ !important; }
%M% .ficha-pie { display: flex; justify-content: space-between; align-items: center; gap: 8px; }

/* ---- Listados de detalle (sin cuadrícula, sin celdas) ---- */
%M% .lst { overflow: auto; margin: 4px 0 12px; }
%M% table.lst-t { width: 100%; border-collapse: collapse; border: none; font-size: .96rem; margin: 0; }
%M% table.lst-t th { font-family: var(--disp) !important; font-weight: 600; font-size: .74rem; letter-spacing: .12em; text-transform: uppercase; text-align: left; color: __GRIS__ !important;
    padding: 0 .9rem 9px 0; border: none; border-bottom: 2px solid __NAVY__; background: transparent !important; white-space: nowrap; }
%M% table.lst-t td { padding: .62rem .9rem .62rem 0; border: none; border-bottom: 1px solid #E3E8F1; background: transparent !important; line-height: 1.4; vertical-align: top; }
%M% table.lst-t .num { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; padding-right: 1.1rem; }
%M% table.lst-t tr:last-child td { border-bottom: none; }
%M% table.lst-t tr.tot td { font-weight: 700; border-top: 2px solid __NAVY__; border-bottom: none; }
%M% table.lst-t tr.sel td { background: rgba(200,16,46,.06) !important; font-weight: 600; }

/* ---- Controles del área principal ---- */
.stApp [data-testid="stMain"] .stButton button { border-radius: 10px; border: 1.5px solid __ROJO__; background: #FFFFFF !important; font-weight: 600; transition: all .16s; box-shadow: none; }
.stApp [data-testid="stMain"] .stButton button:hover { background: #FBE3E7 !important; }
.stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"] { background: __ROJO__ !important; border-color: __ROJO_OSC__; }
.stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"] * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; }
.stApp [data-testid="stMain"] .stButton button p { font-size: .95rem !important; font-weight: 600 !important; }
[data-testid="stBaseButton-pills"], [data-testid="stBaseButton-segmented_control"] { background: #FFFFFF !important; border: 1.5px solid #AEBBD6 !important; border-radius: 999px !important; min-height: 2.2rem; padding: .2rem 1rem !important; transition: all .15s; }
[data-testid="stBaseButton-segmented_control"] { border-radius: 10px !important; }
[data-testid="stBaseButton-pills"]:hover, [data-testid="stBaseButton-segmented_control"]:hover { border-color: __ROJO__ !important; }
[data-testid="stBaseButton-pillsActive"], [data-testid="stBaseButton-segmented_controlActive"] { background: __ROJO__ !important; border: 1.5px solid __ROJO_OSC__ !important; border-radius: 999px !important; min-height: 2.2rem; padding: .2rem 1rem !important; }
[data-testid="stBaseButton-segmented_controlActive"] { border-radius: 10px !important; }
.stApp [data-testid="stBaseButton-pillsActive"] *, .stApp [data-testid="stBaseButton-segmented_controlActive"] * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; }
[data-testid^="stBaseButton-pills"] p, [data-testid^="stBaseButton-segmented"] p { font-size: .94rem !important; font-weight: 600 !important; }
.stApp [data-testid="stMain"] [data-testid="stRadio"] label p { font-size: .98rem !important; font-weight: 600 !important; }
div[data-baseweb="select"] > div { background: #FFFFFF !important; }
.stApp [data-testid="stWidgetLabel"] p { font-family: var(--disp) !important; font-weight: 600; font-size: .78rem !important; letter-spacing: .14em; text-transform: uppercase; color: __GRIS__ !important; }
%M% .lab { font-family: var(--disp) !important; font-weight: 600; font-size: .78rem; letter-spacing: .14em; text-transform: uppercase; color: __GRIS__ !important; margin: 0 0 .4rem; }
div[data-baseweb="tab-list"] { gap: 1.6rem; border-bottom: 1px solid #CFD8E8; }
button[data-baseweb="tab"] { background: transparent !important; padding: .55rem 0 !important; height: auto; }
button[data-baseweb="tab"] p { font-family: var(--disp) !important; font-size: 1.05rem !important; font-weight: 600 !important; letter-spacing: .02em; }
button[data-baseweb="tab"][aria-selected="true"] { border-bottom: 3px solid __ROJO__ !important; }
div[data-baseweb="tab-highlight"] { background-color: __ROJO__ !important; }
div[data-baseweb="tab-border"] { background-color: transparent !important; }
[data-testid="stPlotlyChart"] { background: #FFFFFF; border: 1px solid #E0E6F0; border-radius: 14px; box-shadow: 0 3px 12px rgba(11,31,75,.07); padding: 6px 8px; }
[data-testid="stExpander"] { background: transparent !important; border: none !important; border-top: 1px solid #CFD8E8 !important; border-radius: 0 !important; }
[data-testid="stExpander"] details, [data-testid="stExpander"] summary { background: transparent !important; }
[data-testid="stExpander"] summary p { font-family: var(--disp) !important; font-size: 1.05rem; font-weight: 600; }

/* ---- Adaptable ---- */
@media (max-width: 1100px) { .block-container { padding: 1.8rem 1.6rem 3rem; } %M% .cab h1 { font-size: 2.1rem; } %M% .rk-h { display: none; }
    %M% .rk-f { grid-template-columns: 44px 1fr; row-gap: 10px; } %M% .rk-e { text-align: left; } }
@media (max-width: 700px)  { .block-container { padding: 1.2rem .9rem 3rem; } %M% .cab h1 { font-size: 1.75rem; } %M% .sec h2 { font-size: 1.35rem; } %M% .kv { grid-template-columns: 1fr; gap: 2px; } %M% .regla { grid-template-columns: 1fr; } }
@media (min-width: 2000px) { html { font-size: 18px !important; } .block-container { max-width: 2100px; } }

%M% .fk { position: relative; height: 196px; box-sizing: border-box; background: #FFFFFF; border: 1px solid #D5DCE8; border-radius: 16px; padding: 15px 18px 14px 22px; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; transition: transform .18s, box-shadow .18s, border-color .18s; box-shadow: 0 2px 8px rgba(11,31,75,.06); }
%M% .fk::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 6px; background: var(--ac); }
%M% .fk.on { border-color: __ROJO__; background: #FFF8F9; box-shadow: 0 0 0 2px __ROJO__, 0 12px 24px rgba(200,16,46,.16); }
%M% .fk-top { display: flex; align-items: center; gap: 10px; }
%M% .fk-ic { width: 34px; height: 34px; border-radius: 10px; background: var(--tn); display: inline-flex; align-items: center; justify-content: center; font-size: 1.05rem; }
%M% .fk-l { font-family: var(--disp) !important; font-weight: 600; font-size: .82rem; letter-spacing: .14em; text-transform: uppercase; color: __GRIS__ !important; }
%M% .fk-ck { margin-left: auto; width: 20px; height: 20px; border-radius: 50%; border: 2px solid #AEBBD6; box-sizing: border-box; display: flex; align-items: center; justify-content: center; background: #FFFFFF; }
%M% .fk.on .fk-ck { background: __ROJO__; border-color: __ROJO__; }
%M% .fk.on .fk-ck::after { content: "✓"; color: #FFFFFF !important; font-size: 12px; font-weight: 700; line-height: 1; }
%M% .fk-v { font-family: var(--disp) !important; font-weight: 700; font-size: clamp(1.55rem, 1.15rem + .9vw, 2.15rem); line-height: 1; white-space: nowrap; color: __NAVY__ !important; font-variant-numeric: tabular-nums; }
%M% .fk-sp { height: 38px; display: flex; align-items: stretch; gap: 3px; }
%M% .fk-sp .b { flex: 1; position: relative; }
%M% .fk-sp .b i { position: absolute; left: 0; right: 0; border-radius: 3px; }
%M% .fk-seg { height: 14px; display: flex; border-radius: 99px; overflow: hidden; gap: 2px; margin: 12px 0; }
%M% .fk-seg i { display: block; height: 100%; }
%M% .fk-hh { display: flex; flex-direction: column; gap: 6px; padding: 2px 0; }
%M% .fk-h { display: flex; align-items: center; gap: 8px; font-size: .74rem; font-weight: 600; }
%M% .fk-h span { flex: 0 0 62px; color: __GRIS__ !important; }
%M% .fk-h div { flex: 1; height: 9px; background: #E7ECF4; border-radius: 99px; overflow: hidden; }
%M% .fk-h div i { display: block; height: 100%; border-radius: 99px; }
%M% .fk-h b { flex: 0 0 auto; min-width: 52px; text-align: right; font-weight: 700; }
%M% .fk-foot { display: flex; align-items: center; }
%M% .fk-b { display: inline-flex; align-items: center; padding: .16rem .7rem .18rem; border-radius: 999px; border: 1px solid var(--c); font-family: var(--disp) !important; font-weight: 600; font-size: .8rem; letter-spacing: .05em; background: #FFFFFF; }
%M% .ff { background: #FFFFFF; border: 1px solid #D5DCE8; border-top: 6px solid var(--ac); border-radius: 16px; padding: 22px 24px 20px; min-height: 340px; display: flex; flex-direction: column; gap: 12px; box-shadow: 0 2px 10px rgba(11,31,75,.06); }
%M% .ff-l { font-family: var(--disp) !important; font-weight: 600; font-size: .85rem; letter-spacing: .16em; text-transform: uppercase; color: __GRIS__ !important; }
%M% .ff-v { font-family: var(--disp) !important; font-weight: 700; font-size: clamp(2rem, 1.4rem + 1.5vw, 3rem); line-height: 1.05; color: __NAVY__ !important; font-variant-numeric: tabular-nums; word-break: break-word; }
%M% .ff-t { font-size: 1.02rem; line-height: 1.5; }
%M% .ff-f { margin-top: auto; padding-top: 12px; border-top: 1px dashed #B9C5DE; font-size: .85rem; font-weight: 600; color: __GRIS__ !important; }
.stApp [class*="st-key-fk_"] { position: relative; gap: 0 !important; }
.stApp [class*="st-key-fkbtn_"] { position: absolute !important; inset: 0; z-index: 4; margin: 0 !important; width: 100% !important; height: 100% !important; }
.stApp [class*="st-key-fkbtn_"] > div, .stApp [class*="st-key-fkbtn_"] [data-testid="stButton"] { width: 100% !important; height: 100% !important; }
.stApp [class*="st-key-fkbtn_"] button { position: absolute !important; inset: 0 !important; width: 100% !important; height: 100% !important; min-height: 0 !important; opacity: 0; cursor: pointer; border-radius: 16px; }
.stApp [class*="st-key-fk_"]:hover .fk { transform: translateY(-3px); box-shadow: 0 10px 22px rgba(11,31,75,.15); border-color: __ROJO__; }
@media (max-width: 900px) { %M% .fk { height: auto; min-height: 170px; gap: 8px; } %M% .ff { min-height: 0; } }

/* ---- Estados de los botones: al soltar (focus/active) conservan su color ---- */
.stApp section[data-testid="stSidebar"] .stButton button:focus, .stApp section[data-testid="stSidebar"] .stButton button:focus-visible,
.stApp section[data-testid="stSidebar"] .stButton button:active, .stApp section[data-testid="stSidebar"] .stButton button:focus:not(:active) {
    background: rgba(255,255,255,.10) !important; border: none !important; outline: none !important; box-shadow: none !important; color: #FFFFFF !important; }
.stApp section[data-testid="stSidebar"] .stButton button:focus *, .stApp section[data-testid="stSidebar"] .stButton button:active * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; }
.stApp section[data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"]:focus, .stApp section[data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"]:active,
.stApp section[data-testid="stSidebar"] .stButton button[data-testid="stBaseButton-primary"]:focus-visible { background: rgba(255,255,255,.14) !important; box-shadow: inset 3px 0 0 __ROJO__ !important; }
.stApp section[data-testid="stSidebar"] [class*="st-key-sbm_"] button:focus, .stApp section[data-testid="stSidebar"] [class*="st-key-sbm_"] button:active { border: 1px solid rgba(255,255,255,.35) !important; }

.stApp [data-testid="stMain"] .stButton button:focus, .stApp [data-testid="stMain"] .stButton button:focus-visible, .stApp [data-testid="stMain"] .stButton button:active,
.stApp [data-testid="stMain"] .stButton button:focus:not(:active) { background: #FFFFFF !important; border: 1.5px solid __ROJO__ !important; outline: none !important; box-shadow: none !important; color: __NEGRO__ !important; }
.stApp [data-testid="stMain"] .stButton button:focus *, .stApp [data-testid="stMain"] .stButton button:active * { color: __NEGRO__ !important; -webkit-text-fill-color: __NEGRO__ !important; }
.stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"]:focus, .stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"]:active,
.stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"]:focus-visible { background: __ROJO__ !important; border-color: __ROJO_OSC__ !important; }
.stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"]:focus *, .stApp [data-testid="stMain"] .stButton button[data-testid="stBaseButton-primary"]:active * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; }
.stApp [data-testid="stMain"] .stDownloadButton button:focus, .stApp [data-testid="stMain"] .stDownloadButton button:active { background: #FFFFFF !important; color: __NEGRO__ !important; outline: none !important; box-shadow: none !important; }

[data-testid="stBaseButton-pills"]:focus, [data-testid="stBaseButton-pills"]:active, [data-testid="stBaseButton-segmented_control"]:focus, [data-testid="stBaseButton-segmented_control"]:active
    { background: #FFFFFF !important; border-color: #AEBBD6 !important; outline: none !important; box-shadow: none !important; }
.stApp [data-testid="stBaseButton-pills"]:focus *, .stApp [data-testid="stBaseButton-pills"]:active *, .stApp [data-testid="stBaseButton-segmented_control"]:focus *, .stApp [data-testid="stBaseButton-segmented_control"]:active * { color: __NEGRO__ !important; -webkit-text-fill-color: __NEGRO__ !important; }
[data-testid="stBaseButton-pillsActive"]:focus, [data-testid="stBaseButton-pillsActive"]:active, [data-testid="stBaseButton-segmented_controlActive"]:focus, [data-testid="stBaseButton-segmented_controlActive"]:active
    { background: __ROJO__ !important; border-color: __ROJO_OSC__ !important; outline: none !important; box-shadow: none !important; }
button[data-baseweb="tab"]:focus, button[data-baseweb="tab"]:active, button[data-baseweb="tab"]:hover { background: transparent !important; outline: none !important; box-shadow: none !important; }
button[data-baseweb="tab"]:focus p, button[data-baseweb="tab"]:active p { color: __NEGRO__ !important; -webkit-text-fill-color: __NEGRO__ !important; }
</style>
"""
_PRE = '.stApp [data-testid="stMarkdownContainer"]'
for _k, _v in dict(NEGRO=NEGRO, NAVY=NAVY, ROJO=ROJO, ROJO_OSC=ROJO_OSC, GRIS=GRIS).items():
    CSS = CSS.replace(f"__{_k}__", _v)
CSS = CSS.replace("%M%", _PRE)
st.markdown(CSS, unsafe_allow_html=True)

FF = "Barlow, 'Segoe UI', Arial, sans-serif"
FD = "'Barlow Semi Condensed', Barlow, 'Segoe UI', Arial, sans-serif"


# ======================= COMPONENTES =======================
def cabecera(titulo, sub, meta=()):
    m = "".join(f"<span>{x}</span>" for x in meta)
    md(f'<div class="cab"><div class="cab-over">Transvisa · TGM 2025</div><h1>{titulo}</h1><p class="cab-sub">{sub}</p>'
       + (f'<div class="cab-meta">{m}</div>' if m else "") + "</div>")


def seccion(titulo, sub=None, n=None):
    num = f'<span class="sec-n">{n}</span>' if n else ""
    md(f'<div class="sec">{num}<h2>{titulo}</h2></div>' + (f'<p class="sec-s">{sub}</p>' if sub else ""))


def subtitulo(texto):
    md(f'<div class="sub3">{texto}</div>')


def etiqueta(texto):
    md(f'<div class="lab">{texto}</div>')


def stat(titulo, valor, sub, color=NAVY):
    md(f'<div class="stat" style="--ac:{color};"><div class="stat-l">{titulo}</div><div class="stat-v">{valor}</div><div class="stat-s">{sub}</div></div>')


kpi = stat


def chip_sem(est, texto=None):
    return f'<span class="sem" style="--c:{SEM_COL[est]};--t:{SEM_TINT[est]};"><i></i>{texto or SEM_TXT[est]}</span>'


def nota(texto, roja=False, etq=None):
    e = f'<span class="eti">{etq}</span>' if etq else ""
    md(f'<div class="nota {"roja" if roja else ""}">{e}{texto}</div>')


def lectura(texto, etq=None, color=None):
    e = f'<span class="eti">{etq}</span>' if etq else ""
    est = f' style="border-left:4px solid {color};"' if color else ""
    md(f'<div class="lect"{est}>{e}{texto}</div>')


def tabla(df, alto=None, num=(), filas_sel=(), total_ultima=False):
    cab = "".join(f'<th class="{"num" if c in num else ""}">{c}</th>' for c in df.columns)
    filas, n = "", len(df)
    for i, row in enumerate(df.itertuples(index=False)):
        cls = "sel" if i in filas_sel else ("tot" if total_ultima and i == n - 1 else "")
        filas += f'<tr class="{cls}">' + "".join(f'<td class="{"num" if c in num else ""}">{v}</td>' for c, v in zip(df.columns, row)) + "</tr>"
    est = f' style="max-height:{alto}px;"' if alto else ""
    md(f'<div class="lst"{est}><table class="lst-t"><thead><tr>{cab}</tr></thead><tbody>{filas}</tbody></table></div>')


def tarjeta_causa(c, K, sel):
    m = META[c["codigo"]]
    u = c["unidad"]
    return (f'<div class="cc {"on" if sel else ""}" style="--ac:{m["color"]};">'
            f'<div><div class="cc-top"><span class="cc-n" style="color:{m["color"]} !important;">{nn(c)}</span><span class="cc-chk"></span></div>'
            f'<div class="cc-nm">{m["corto"]}</div><div class="cc-cod">{c["codigo"]}</div></div>'
            f'<div class="cc-v">{val_html(K["to_be"], u)}</div>'
            f'<div><div class="cc-r"><span>Actual</span><b>{val_html(K["as_is"], u)}</b></div><div class="cc-r"><span>Meta</span><b>{val_html(c["estandar"], u)}</b></div></div>'
            f'<div class="cc-foot">{chip_sem(K["sem_to"])}</div></div>')


def ranking(idx):
    """Comparativo de las causas y del problema central; resalta las elegidas y atenúa el resto."""
    hay = bool(S.sel)

    def fila(n, color, nm, ident, a, t, meta, la, lt, mej, sem, inv, ah, cls):
        esc = 1.5
        wa, wt, wm = min(la, esc) / esc * 100, min(lt, esc) / esc * 100, 100 / esc
        return (f'<div class="rk-f {cls}" style="--ac:{color};"><div class="rk-n">{n}</div>'
                f'<div><div class="rk-nm">{nm}</div><div class="rk-id">{ident}</div></div>'
                f'<div><div class="rk-vals"><span><i>Actual</i><b>{a}</b></span><span><i>Con la mejora</i><b>{t}</b></span><span><i>Meta</i><b>{meta}</b></span></div>'
                f'<div class="rk-bar"><div class="t"></div><div class="a" style="width:{wa:.1f}%;"></div><div class="b" style="width:{wt:.1f}%;background:{SEM_COL[sem]};"></div><div class="m" style="left:{wm:.1f}%;"></div></div></div>'
                f'<div><div class="rk-mj">{mej}</div>{chip_sem(sem)}</div>'
                f'<div class="rk-e">Inversión <b>{inv}</b><br>Ahorro anual <b>{ah}</b></div></div>')
    cab = '<div class="rk-h"><span></span><span>Causa raíz</span><span>Avance frente a la meta</span><span>Mejora</span><span style="text-align:right;">Impacto (S/)</span></div>'
    filas = ""
    g = D["conclusiones"]["general"]
    nm = "Problema central"
    filas += fila("PC", NAVY, nm, "Cumplimiento de OT de mantenimiento preventivo", f_val(PB["as_is"], "%"), f_val(PB["to_be"], "%"), f_val(PB["estandar"], "%"),
                  PB["as_is"] / PB["estandar"], PB["to_be"] / PB["estandar"], pct_signo(g["variacion"]), estado_sem(PC, PB["to_be"]),
                  f"{FL['inversion_total']:,.0f}", f"{FL['eval_total']['beneficio']:,.0f}", "pc")
    for c in CAUSAS:
        K = calc(c, idx)
        m = META[c["codigo"]]
        nm = f'{nombre(c)} · {c["codigo"]}'
        cls = ("on" if c["id"] in S.sel else "off") if hay else ""
        filas += fila(nn(c), m["color"], nm, c["indicador"], f_val(K["as_is"], c["unidad"]), f_val(K["to_be"], c["unidad"]), f_val(c["estandar"], c["unidad"]),
                      K["logro_as"], K["logro_to"], pct_signo(K["mejora"]), K["sem_to"], f"{inversion_de(c['codigo'])['inversion']:,.0f}", f"{eval_de(c['codigo'])['beneficio']:,.0f}", cls)
    md(f'<div class="rk">{cab}{filas}</div>')


# ======================= GRÁFICOS =======================
def estilo(fig, alto=None, leyenda_h=True):
    fig.update_layout(template="plotly_white", font=dict(family=FF, size=13, color=NEGRO), paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                      title_font=dict(family=FD, size=19, color=NAVY), title_x=0.01, title_xanchor="left", title_y=.965, title_yanchor="top",
                      margin=dict(l=58, r=30, t=84, b=50),
                      legend=dict(font=dict(size=12, color=NEGRO), bgcolor="rgba(255,255,255,0)",
                                  **(dict(orientation="h", yanchor="bottom", y=1.015, xanchor="left", x=0) if leyenda_h else {})),
                      hoverlabel=dict(bgcolor="#FFFFFF", font=dict(family=FF, color=NEGRO, size=13), bordercolor=NAVY))
    eje = dict(tickfont=dict(size=12, color=NEGRO), title_font=dict(size=13, color=GRIS), gridcolor="#E7ECF4", linecolor="#9AA5B8", zerolinecolor="#9AA5B8")
    fig.update_xaxes(**eje)
    fig.update_yaxes(**eje)
    if alto:
        fig.update_layout(height=alto)
    return fig


def grafico(fig, key, seleccion=False):
    """Dibuja el gráfico. Con seleccion=True, un clic sobre una causa la añade o la quita de la selección."""
    kw = dict(theme=None, key=f"{key}_{S.nonce}" if seleccion else key)
    if seleccion:
        try:
            try:
                ev = st.plotly_chart(fig, width="stretch", on_select="rerun", selection_mode="points", **kw)
            except TypeError:
                ev = st.plotly_chart(fig, use_container_width=True, on_select="rerun", selection_mode="points", **kw)
        except TypeError:
            st.plotly_chart(fig, **kw)
            return
        try:
            pts = ev["selection"]["points"]
        except Exception:
            pts = []
        if pts:
            cd = pts[0].get("customdata")
            cid = cd[0] if isinstance(cd, (list, tuple)) else cd
            if cid in POR_ID:
                S.nonce += 1          # reinicia la selección interna del gráfico: cada clic es un evento nuevo
                alternar(cid)
                st.rerun()
        return
    try:
        st.plotly_chart(fig, width="stretch", **kw)
    except TypeError:
        try:
            st.plotly_chart(fig, use_container_width=True, **kw)
        except TypeError:
            st.plotly_chart(fig, **kw)


def fig_termometro(cod, u, valores, titulo=None, alto=430):
    """Tubo con zonas rojo/amarillo/verde, llenado hasta el valor y marca de la meta. valores = [(etiqueta, valor), ...]."""
    s = SEM[cod]
    mn, mx, base, meta = s["escala_min"], s["escala_max"], s["linea_base"], s["meta"]
    may = str(s["sentido"]).lower().startswith("mayor")
    k = 100 if u == "%" else 1
    mn, mx, base, meta = mn * k, mx * k, base * k, meta * k
    zonas = [(mn, base, "rojo"), (base, meta, "amarillo"), (meta, mx, "verde")] if may else [(mn, meta, "verde"), (meta, base, "amarillo"), (base, mx, "rojo")]
    fig = go.Figure()
    span = mx - mn
    n = len(valores)
    centros = [0.0] if n == 1 else [-1.0, 1.0]
    ancho = 0.28
    for (etq, val), cx in zip(valores, centros):
        v_ = val * k
        for z0, z1, est in zonas:
            fig.add_shape(type="rect", x0=cx - ancho, x1=cx + ancho, y0=z0, y1=z1, fillcolor=SEM_TINT[est], line=dict(width=0), layer="below")
        est_v = estado_sem(cod, val)
        v_cl = min(max(v_, mn), mx)
        fig.add_shape(type="rect", x0=cx - ancho * .55, x1=cx + ancho * .55, y0=mn, y1=v_cl, fillcolor=SEM_COL[est_v], line=dict(width=0), layer="above")
        fig.add_shape(type="rect", x0=cx - ancho, x1=cx + ancho, y0=mn, y1=mx, line=dict(color=NAVY, width=2.5), fillcolor="rgba(0,0,0,0)", layer="above")
        fig.add_trace(go.Scatter(x=[cx], y=[mn - span * .09], mode="markers", marker=dict(size=40, color=SEM_COL[est_v], line=dict(color=NAVY, width=3)),
                                 hoverinfo="skip", showlegend=False))
        fig.add_trace(go.Scatter(x=[cx], y=[v_], mode="markers", marker=dict(size=13, color=NAVY, symbol="diamond", line=dict(color="#FFFFFF", width=1.5)),
                                 name=etq, showlegend=False,
                                 hovertemplate=f"<b>{etq}</b><br>Valor: {f_val(val, u)}<br>Estado: {SEM_TXT[est_v]}<extra></extra>"))
        fig.add_annotation(x=cx, y=mn - span * .19, yanchor="top", text=f"<b>{etq}</b><br><b>{f_val(val, u, False)}</b><br>{SEM_TXT[est_v]}", showarrow=False, font=dict(size=13, color=NEGRO))
    xr = 0.9 if n == 1 else 2.0
    fig.add_shape(type="line", x0=-xr, x1=xr, y0=meta, y1=meta, line=dict(color=NEGRO, width=3, dash="dash"))
    fig.add_annotation(x=xr, y=meta, text=f"<b>Meta {f_val(meta / k, u, False)}{'' if u == '%' else ' ' + u}</b>", showarrow=False, xanchor="right", yanchor="bottom",
                       font=dict(size=12, color=NEGRO), bgcolor="rgba(255,255,255,.85)")
    fig.update_xaxes(visible=False, range=[-xr - .1, xr + .1])
    fig.update_yaxes(range=[mn - span * .42, mx + span * .03], title=etq_u(u), showgrid=False)
    fig.update_layout(title=titulo or "Termómetro de cumplimiento", showlegend=False)
    return estilo(fig, alto, leyenda_h=False)


def fig_bullet(c, K):
    u, est, may = c["unidad"], c["estandar"], mayor(c)
    k = 100 if u == "%" else 1
    fig = go.Figure()
    for etq, val, col in (("As-Is", K["as_is"], ROJO), ("To-Be", K["to_be"], NAVY)):
        fig.add_trace(go.Bar(y=[etq], x=[val * k], orientation="h", marker_color=col, width=.5, name=etq, text=[f_val(val, u)], textposition="outside",
                             textfont=dict(size=14, color=NEGRO), cliponaxis=False, hovertemplate=f"<b>{etq}</b>: {f_val(val, u)}<extra></extra>"))
    fig.add_vline(x=est * k, line_dash="dash", line_color=NEGRO, line_width=3,
                  annotation_text=f"Estándar {f_val(est, u)}", annotation_position="top", annotation_font=dict(size=13, color=NEGRO))
    fig.update_layout(title=f"Comparación con el estándar ({'mayor' if may else 'menor'} es mejor)", showlegend=False, xaxis_title=etq_u(u), yaxis_title="")
    fig.update_xaxes(rangemode="tozero")
    return estilo(fig, 300, leyenda_h=False)


def fig_mensual(c, K, esc, tendencia):
    u, est = c["unidad"], c["estandar"]
    x = [MESES[i] for i in K["idx"]]
    a, t = disp(K["a"], u), disp(K["t"], u)
    fig = go.Figure()
    xs = np.arange(len(x))
    if esc in ("Ambos", "As-Is"):
        fig.add_trace(go.Scatter(x=x, y=a, mode="lines+markers+text", name="As-Is 2025", text=[f"{v:,.1f}" for v in a], textposition="bottom center",
                                 textfont=dict(size=11, color=NEGRO), line=dict(color=ROJO, width=3), marker=dict(size=10, color=ROJO),
                                 hovertemplate="<b>%{x}</b><br>As-Is: %{y:,.2f}<extra></extra>"))
    if esc in ("Ambos", "To-Be"):
        fig.add_trace(go.Scatter(x=x, y=t, mode="lines+markers+text", name="To-Be (con la mejora)", text=[f"{v:,.1f}" for v in t], textposition="top center",
                                 textfont=dict(size=11, color=NEGRO), line=dict(color=NAVY, width=3), marker=dict(size=10, color=NAVY),
                                 hovertemplate="<b>%{x}</b><br>To-Be: %{y:,.2f}<extra></extra>"))
    if tendencia and len(x) >= 3:
        for ser_, col, nm_, vis in ((a, ROJO, "Tendencia As-Is", esc in ("Ambos", "As-Is")), (t, NAVY, "Tendencia To-Be", esc in ("Ambos", "To-Be"))):
            if vis:
                p = np.polyfit(xs, ser_, 1)
                fig.add_trace(go.Scatter(x=x, y=np.polyval(p, xs), mode="lines", name=f"{nm_} ({p[0]:+.2f}/mes)", line=dict(color=col, width=2, dash="dot"), hoverinfo="skip"))
    fig.add_hline(y=float(disp(est, u)), line_dash="dot", line_color=GRIS, line_width=2, annotation_text=f"Estándar: {f_val(est, u)}",
                  annotation_position="top right", annotation_font=dict(color=NEGRO, size=12))
    fig.update_layout(title=f"Evolución mensual · {c['indicador']}", yaxis_title=etq_u(u), xaxis_title="")
    return estilo(fig, 420)


def fig_brecha(c, K, esc):
    u, est, sig = c["unidad"], c["estandar"], (1 if mayor(c) else -1)
    x = [MESES[i] for i in K["idx"]]
    fig = go.Figure()
    for nm_, arr, col, vis in (("Brecha As-Is", K["a"], ROJO, esc in ("Ambos", "As-Is")), ("Brecha To-Be", K["t"], NAVY, esc in ("Ambos", "To-Be"))):
        if vis:
            b = (disp(arr, u) - float(disp(est, u))) * sig
            fig.add_trace(go.Bar(x=x, y=b, name=nm_, marker_color=col, text=[f"{v:+,.1f}" for v in b], textposition="outside", textfont=dict(size=11, color=NEGRO),
                                 cliponaxis=False, hovertemplate="<b>%{x}</b><br>" + nm_ + ": %{y:+,.2f}<extra></extra>"))
    fig.add_hline(y=0, line_color=NEGRO, line_width=2)
    fig.update_layout(title="Brechas frente al estándar por mes (positivo = cumple · negativo = brecha)", barmode="group", yaxis_title=f"Brecha ({etq_u(u)})", xaxis_title="")
    return estilo(fig, 400)


def etq_causa(c):
    return f"{nn(c)} · {c['codigo']}"


def fig_all_logro(idx, cs):
    """Logro de la meta As-Is vs To-Be de las causas en pantalla (clic = añadir / quitar de la selección)."""
    fig = go.Figure()
    ids = [c["id"] for c in cs]
    nm = [etq_causa(c) for c in cs]
    Ks = [calc(c, idx) for c in cs]
    la = [K["logro_as"] * 100 for K in Ks]
    lt = [K["logro_to"] * 100 for K in Ks]
    for y_, a_, t_ in zip(nm, la, lt):
        fig.add_trace(go.Scatter(x=[a_, t_], y=[y_, y_], mode="lines", line=dict(color="#9AA5B8", width=5), hoverinfo="skip", showlegend=False))
    fig.add_trace(go.Scatter(x=la, y=nm, mode="markers+text", name="As-Is", marker=dict(size=19, color=ROJO), text=[f"{v:.0f}%" for v in la], textposition="middle left",
                             textfont=dict(size=12, color=NEGRO), customdata=ids, hovertemplate="<b>%{y}</b><br>Logro As-Is: %{x:.1f}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=lt, y=nm, mode="markers+text", name="To-Be", marker=dict(size=19, color=NAVY), text=[f"{v:.0f}%" for v in lt], textposition="middle right",
                             textfont=dict(size=12, color=NEGRO), customdata=ids, hovertemplate="<b>%{y}</b><br>Logro To-Be: %{x:.1f}%<extra></extra>"))
    fig.add_vline(x=100, line_dash="dash", line_color=NEGRO, line_width=3, annotation_text="Meta = 100%", annotation_font=dict(size=12, color=NEGRO))
    fig.update_layout(title="Logro de la meta: As-Is → To-Be (clic para añadir o quitar una causa)", xaxis_title="% de logro de la meta", yaxis_title="")
    fig.update_yaxes(autorange="reversed")
    return estilo(fig, 400)


def fig_mejora_barras(idx):
    Ks = [calc(c, idx) for c in CAUSAS]
    ids = [c["id"] for c in CAUSAS]
    cols = [META[c["codigo"]]["color"] for c in CAUSAS]
    op = [1.0 if (not S.sel or c["id"] in S.sel) else .3 for c in CAUSAS]
    fig = go.Figure(go.Bar(x=[etq_causa(c) for c in CAUSAS], y=[K["mejora"] * 100 for K in Ks], marker=dict(color=cols, opacity=op, line=dict(color=NEGRO, width=1)),
                           text=[f"{K['mejora'] * 100:.1f}%" for K in Ks], textposition="outside", textfont=dict(size=13, color=NEGRO), customdata=ids, cliponaxis=False,
                           hovertemplate="<b>%{x}</b><br>Mejora del indicador: %{y:.1f}%<extra></extra>"))
    fig.update_layout(title="Mejora esperada por causa (To-Be frente a As-Is) · clic para añadir o quitar", yaxis_title="% de mejora", xaxis_title="", showlegend=False)
    return estilo(fig, 380, leyenda_h=False)


def fig_semaforo_mensual(esc, cs_items):
    """Estado mensual de cada indicador: un punto por mes con el color del semáforo."""
    fig = go.Figure()
    etqs = [e for e, _, _, _ in cs_items]
    for est in ("verde", "amarillo", "rojo"):
        xs_, ys_, tx_ = [], [], []
        for etq, cod, vals, u in cs_items:
            for m_, v in zip(MESES, vals):
                if estado_sem(cod, v) == est:
                    xs_.append(m_); ys_.append(etq); tx_.append(f_val(v, u, False))
        fig.add_trace(go.Scatter(x=xs_, y=ys_, mode="markers+text", name=SEM_TXT[est], text=tx_, textposition="bottom center", textfont=dict(size=10, color=GRIS),
                                 marker=dict(size=22, color=SEM_COL[est], line=dict(color="#FFFFFF", width=2)),
                                 hovertemplate="<b>%{y}</b> · %{x}<br>Valor: %{text}<extra></extra>"))
    fig.update_yaxes(autorange="reversed", categoryorder="array", categoryarray=etqs, showgrid=True, gridcolor="#EEF1F7")
    fig.update_xaxes(categoryorder="array", categoryarray=MESES, showgrid=False)
    fig.update_layout(title=f"Estado mensual de los indicadores · {esc}")
    return estilo(fig, 90 + 74 * len(cs_items))


LOGO_HTML = f'<img src="data:image/png;base64,{LOGO_B64}">' if LOGO_B64 else ""


# ======================= BARRA LATERAL =======================
PAGINAS = [("panel", "Panel interactivo"), ("central", "Problema central"), ("semaforo", "Semáforo"),
           ("flujo", "Flujo financiero")]
with st.sidebar:
    md(f'<div class="brand"><div class="brand-logo">{LOGO_HTML}</div><div class="brand-n">TRANSVISA E.I.R.L.</div><div class="brand-s">TGM · Año 2025</div></div>')
    md('<div class="sb-lab">Navegación</div>')
    for pid, etq in PAGINAS:
        st.button(etq, on_click=ir, args=(pid,), key="nav_" + pid, type="primary" if S.pagina == pid else "secondary")
    md('<div class="sb-lab">Causas raíz</div>')
    for c in CAUSAS:
        st.button(f"{nn(c)}   {nombre(c)}", on_click=alternar, args=(c["id"],), key="sbc_" + c["id"],
                  type="primary" if c["id"] in S.sel else "secondary", help="Añadir o quitar de la selección")
    b1, b2 = st.columns(2, gap="small")
    with b1:
        st.button("Todas", on_click=todas, key="sbm_todas")
    with b2:
        st.button("Ninguna", on_click=ninguna, key="sbm_ninguna")
    md('<div class="sb-lab">Serie mensual</div>')
    st.radio("Serie mensual", FUENTES, key="fuente", label_visibility="collapsed")


# ======================= PANEL: COMPONENTES POR PESTAÑA =======================
def panel_metodologia(cs):
    MET = D["metodologias"]
    if len(cs) > 1:
        diag = sorted({d for cc in cs for d in MET[cc["codigo"]].get("diagnostico", [])})
        cols = ["Metodología / herramienta de mejora"] + diag + ["Simulación Monte Carlo (Python)"]
        fig = go.Figure()
        for cc in cs:
            y = f"{nn(cc)} · {nombre(cc)}"
            fig.add_trace(go.Scatter(x=[cols[0]], y=[y], mode="markers+text", text=[f"<b>{cc['codigo']}</b>"], textposition="middle center", textfont=dict(size=13, color="#FFFFFF"),
                                     marker=dict(size=46, color=META[cc["codigo"]]["color"], line=dict(color=NEGRO, width=1.5)), customdata=[cc["id"]],
                                     hovertemplate=f"<b>{cc['herramienta']}</b><br>{cc['objetivo']}<extra></extra>", showlegend=False))
            for dcol in diag:
                ok = dcol in MET[cc["codigo"]].get("diagnostico", [])
                fig.add_trace(go.Scatter(x=[dcol], y=[y], mode="markers+text", text=["✔" if ok else "—"], textposition="middle center", textfont=dict(size=16, color="#FFFFFF" if ok else GRIS),
                                         marker=dict(size=34, color=AZUL if ok else "#EEF1F7", line=dict(color=NAVY, width=1)), customdata=[cc["id"]], showlegend=False,
                                         hovertemplate=f"<b>{dcol}</b><br>{'Aplicado' if ok else 'No registrado'} en {cc['codigo']}<extra></extra>"))
            ok = bool(MET[cc["codigo"]].get("simulacion"))
            fig.add_trace(go.Scatter(x=[cols[-1]], y=[y], mode="markers+text", text=["✔" if ok else "—"], textposition="middle center", textfont=dict(size=16, color="#FFFFFF" if ok else GRIS),
                                     marker=dict(size=34, color=ROJO if ok else "#EEF1F7", line=dict(color=NAVY, width=1)), customdata=[cc["id"]], showlegend=False,
                                     hovertemplate=f"<b>Simulación de resultados To-Be</b><br>{MET[cc['codigo']].get('simulacion') or 'No registrada'}<extra></extra>"))
        fig.update_layout(title="Herramientas aplicadas a cada causa raíz (clic en una fila para añadirla o quitarla)")
        fig.update_xaxes(side="top", tickangle=0, tickfont=dict(size=11))
        fig.update_yaxes(autorange="reversed")
        grafico(estilo(fig, 150 + 72 * len(cs), leyenda_h=False), "mat_met", seleccion=True)
        filas = [[nn(cc), cc["motivo"], f"<b>{cc['codigo']}</b> · {cc['herramienta'].split(chr(10))[0]}", cc["tipo"], len(MET[cc["codigo"]]["etapas"]), cc["objetivo"]] for cc in cs]
        tabla(pd.DataFrame(filas, columns=["N.º", "Causa raíz", "Metodología / herramienta", "Tipo", "Etapas", "Objetivo específico"]), num=("Etapas",))
        nota("Los ✔ indican las herramientas de diagnóstico y la simulación aplicadas en cada causa. Elige una sola causa para ver sus etapas de implementación.")
        return
    c = cs[0]
    cod = c["codigo"]
    m = META[cod]
    M = MET[cod]
    barrera = next((b for b in D["barreras"] if b["causa"].split()[0:2] == c["motivo"].split()[0:2] or cod in b["causa"]), None)
    k1, k2 = st.columns([1.15, 1], gap="large")
    with k1:
        lectura(f'<b>{nn(c)} · {cod}</b> — {c["herramienta"].replace(chr(10), " ")}<br><b>Objetivo específico:</b> {c["objetivo"]}<br><b>Causa raíz:</b> {c["causa_raiz"]}<br>'
                f'<b>Indicador:</b> {c["indicador_formula"].replace(chr(10), " ")}', etq="Metodología", color=m["color"])
        nota(c["articulo"], etq="Artículo base")
    with k2:
        tags = "".join(f'<span class="tag">{d}</span>' for d in M.get("diagnostico", []))
        lectura(f'{tags}<br><b>Proyección del To-Be:</b> {corto(M.get("simulacion") or "—")}', etq="Herramientas de diagnóstico")
        ei = inversion_de(cod)
        nota(f'<b>{soles(ei["inversion"])}</b> ({ei["pct"] * 100:.1f}% del total sin imprevistos) · activo fijo {soles(ei["activo_fijo"])} · intangible y capital de trabajo {soles(ei["intangible"])}', etq="Inversión de la propuesta")
        if barrera:
            nota(f'{barrera["barrera"]} (indicador: {barrera["herramienta"]})', roja=True, etq="Barrera que genera en el cumplimiento de OT")
    etapas = M["etapas"]
    nombres = [f"{i + 1}. {corto(e['nombre'])}" for i, e in enumerate(etapas)]
    subtitulo(f"Ruta de implementación de {cod} · {len(etapas)} {'componentes' if cod == 'ISOS' else 'etapas'}")
    sel = selector(nombres, f"etapa_{cod}")
    pos = nombres.index(sel)
    fig = go.Figure()
    n = len(etapas)
    fig.add_trace(go.Scatter(x=list(range(n)), y=[0] * n, mode="lines", line=dict(color="#9AA5B8", width=5), hoverinfo="skip", showlegend=False))
    fig.add_trace(go.Scatter(x=list(range(n)), y=[0] * n, mode="markers+text", text=[str(i + 1) for i in range(n)], textfont=dict(size=15, color="#FFFFFF"),
                             marker=dict(size=[52 if i == pos else 38 for i in range(n)], color=[ROJO if i == pos else m["color"] for i in range(n)], line=dict(color=NEGRO, width=2)),
                             hovertext=[corto(e["nombre"]) for e in etapas], hoverinfo="text", showlegend=False))
    for i, e in enumerate(etapas):
        fig.add_annotation(x=i, y=-0.55, text=("<b>" if i == pos else "") + "<br>".join(textwrap.wrap(corto(e["nombre"]), 18)) + ("</b>" if i == pos else ""),
                           showarrow=False, font=dict(size=11, color=NEGRO), yanchor="top")
    fig.update_xaxes(visible=False, range=[-.6, n - .4])
    fig.update_yaxes(visible=False, range=[-2.1, .8])
    fig.update_layout(title=f"Etapas de {cod}")
    grafico(estilo(fig, 250, leyenda_h=False), f"stepper_{cod}")
    e = etapas[pos]
    det = ""
    for lab, key in (("Pregunta que responde", "pregunta"), ("Qué se hace", "que_se_hace"), ("Entregables", "entregables"), ("Descripción", "descripcion"),
                     ("Objetivo en TGM", "objetivo"), ("Herramientas", "herramientas"), ("Fecha", "fecha"), ("Inicio", "inicio"), ("Fin", "fin"), ("Responsable", "responsable")):
        if e.get(key):
            det += f"<dt>{lab}</dt><dd>{e[key]}</dd>"
    md(f'<div class="lect etapa" style="border-left:4px solid {m["color"]};"><h4>{e["nombre"]}</h4><dl class="kv">{det}</dl></div>')
    if cod == "TPM" and M.get("pilares"):
        with st.expander("Alcance de los 8 pilares del TPM"):
            tabla(pd.DataFrame([[p["pilar"], p["aplica"], p["motivo"]] for p in M["pilares"]], columns=["Pilar del TPM", "¿Se aplica?", "Motivo"]))
            if M.get("nota"):
                st.caption(M["nota"])


def _vals_esc(K, esc):
    vals = [("As-Is", K["as_is"]), ("To-Be", K["to_be"])]
    return vals[:1] if esc == "As-Is" else (vals[1:] if esc == "To-Be" else vals)


def tab_kpi(cs, idx):
    esc = S.escenario
    if len(cs) > 1:
        cols = st.columns(len(cs), gap="small")
        for col, cc in zip(cols, cs):
            K = calc(cc, idx)
            with col:
                grafico(fig_termometro(cc["codigo"], cc["unidad"], _vals_esc(K, esc), etq_causa(cc), 430), f"term_all_{cc['id']}")
        nota("Zonas rojo / amarillo / verde según los límites de semaforización de cada indicador; la línea discontinua marca la meta.")
        return
    c = cs[0]
    K = calc(c, idx)
    u = c["unidad"]
    nota(c["causa_raiz"], etq="Causa raíz")
    k = st.columns(5, gap="medium")
    with k[0]: stat("Situación actual (As-Is)", val_html(K["as_is"], u), f"Promedio · {meses_txt(idx)}", ROJO)
    with k[1]: stat("Con la mejora (To-Be)", val_html(K["to_be"], u), f"Aplicando {c['codigo']}", NAVY)
    with k[2]: stat("Estándar / meta", val_html(c["estandar"], u), f"{'Mayor' if K['may'] else 'Menor'} es mejor", GRIS)
    with k[3]: stat("Mejora esperada", pct_signo(K["mejora"]), "Frente al As-Is", AZUL)
    with k[4]: stat("Estado del To-Be", "Cumple" if K["cumple_to"] else "No cumple", f"Semáforo: {SEM_TXT[K['sem_to']]}", SEM_COL[K["sem_to"]])
    g1, g2 = st.columns([1, 1.25], gap="large")
    with g1:
        s = SEM[c["codigo"]]
        activos = {"As-Is": K["sem_as"], "To-Be": K["sem_to"]}
        filas = ""
        for est, txt in (("verde", s["verde"]), ("amarillo", s["amarillo"]), ("rojo", s["rojo"])):
            marca = " · ".join(f"<b>{k_}</b>" for k_, v_ in activos.items() if v_ == est and (esc == "Ambos" or esc == k_))
            filas += f'<div class="regla {"on" if marca else ""}"><div>{chip_sem(est)}</div><div>{txt}</div><div>{marca}</div></div>'
        subtitulo(f"Semáforo de cumplimiento · {etq_u(u)}")
        md(f'<div>{filas}</div>')
        st.write("")
        lectura(f'As-Is {chip_sem(K["sem_as"])} &nbsp;→&nbsp; To-Be {chip_sem(K["sem_to"])}<br>'
                + ("El To-Be alcanza el estándar." if K["cumple_to"] else f"El To-Be aún no alcanza el estándar: brecha de {f_val(abs(K['to_be'] - c['estandar']), u)}."))
    with g2:
        grafico(fig_termometro(c["codigo"], u, _vals_esc(K, esc), f"Termómetro · {c['indicador']}", 430), f"term_{c['id']}")


def tab_antes_despues(cs, idx):
    esc = S.escenario
    if len(cs) > 1:
        grafico(fig_all_logro(idx, cs), "logro_all", seleccion=True)
        filas = []
        for cc in cs:
            K = calc(cc, idx)
            filas.append([nn(cc), cc["indicador"], f_val(K["as_is"], cc["unidad"]), f_val(K["to_be"], cc["unidad"]), f_val(cc["estandar"], cc["unidad"]),
                          f"{K['logro_as'] * 100:.0f}% → {K['logro_to'] * 100:.0f}%", pct_signo(K["mejora"])])
        tabla(pd.DataFrame(filas, columns=["N.º", "Indicador", "As-Is", "To-Be", "Estándar", "Logro de la meta", "Mejora"]), num=("As-Is", "To-Be", "Estándar", "Logro de la meta", "Mejora"))
        nota("Logro de la meta = valor ÷ estándar (mayor es mejor) o estándar ÷ valor (menor es mejor).")
        return
    c = cs[0]
    K = calc(c, idx)
    u = c["unidad"]
    tend = st.toggle("Mostrar línea de tendencia lineal", value=False, key="tend_ad")
    g1, g2 = st.columns([1.6, 1], gap="large")
    with g1:
        grafico(fig_mensual(c, K, esc, tend), f"mens_{c['id']}")
    with g2:
        grafico(fig_bullet(c, K), f"bullet_{c['id']}")
        lectura(f'Pasa de <b>{f_val(K["as_is"], u)}</b> (As-Is) a <b>{f_val(K["to_be"], u)}</b> (To-Be): <b>{pct_signo(K["mejora"])}</b> de mejora.')
    det = pd.DataFrame({"Mes": [MESES[i] for i in idx], "As-Is": [f_val(v, u, False) for v in K["a"]], "To-Be": [f_val(v, u, False) for v in K["t"]],
                        "Estándar": [f_val(c["estandar"], u, False)] * len(idx)})
    if esc == "As-Is":
        det = det.drop(columns=["To-Be"])
    elif esc == "To-Be":
        det = det.drop(columns=["As-Is"])
    with st.expander("Detalle mensual"):
        tabla(det, alto=320, num=tuple(det.columns[1:]))


def tab_brechas(cs, idx):
    esc = S.escenario
    if len(cs) > 1:
        Ks = [calc(cc, idx) for cc in cs]
        fig = go.Figure()
        for nm_, key, col in (("Brecha As-Is", "as_is", ROJO), ("Brecha To-Be", "to_be", NAVY)):
            if esc == "Ambos" or nm_.endswith(esc):
                ys = [((K[key] - cc["estandar"]) / cc["estandar"]) * (1 if K["may"] else -1) * 100 for K, cc in zip(Ks, cs)]
                fig.add_trace(go.Bar(x=[etq_causa(cc) for cc in cs], y=ys, name=nm_, marker_color=col, text=[f"{v:+.1f}%" for v in ys], textposition="outside",
                                     cliponaxis=False, textfont=dict(size=12, color=NEGRO), customdata=[cc["id"] for cc in cs], hovertemplate="<b>%{x}</b><br>%{y:+.1f}% del estándar<extra></extra>"))
        fig.add_hline(y=0, line_color=NEGRO, line_width=2)
        fig.update_layout(title="Brecha frente al estándar, como % del estándar (positivo = cumple · negativo = brecha)", barmode="group", yaxis_title="% del estándar", xaxis_title="")
        grafico(estilo(fig, 400), "brecha_all", seleccion=True)
        filas = []
        for cc, K in zip(cs, Ks):
            k_ = 100 if cc["unidad"] == "%" else 1
            filas.append([nn(cc), cc["indicador"], f_val(cc["estandar"], cc["unidad"]), f"{K['brecha_as'] * k_:+,.1f} {etq_u(cc['unidad'])}",
                          f"{K['brecha_to'] * k_:+,.1f} {etq_u(cc['unidad'])}", "Sí" if K["cumple_to"] else "No"])
        tabla(pd.DataFrame(filas, columns=["N.º", "Indicador", "Estándar", "Brecha As-Is", "Brecha To-Be", "To-Be cumple"]), num=("Estándar", "Brecha As-Is", "Brecha To-Be"))
        return
    c = cs[0]
    K = calc(c, idx)
    u = c["unidad"]
    k_ = 100 if u == "%" else 1
    g1, g2 = st.columns([1, 1.7], gap="large")
    with g1:
        grafico(fig_bullet(c, K), f"bullet2_{c['id']}")
        s1, s2 = st.columns(2, gap="medium")
        with s1: stat("Brecha As-Is", f"{K['brecha_as'] * k_:+,.1f}<small>{etq_u(u)}</small>", "Negativo = no cumple", ROJO)
        with s2: stat("Brecha To-Be", f"{K['brecha_to'] * k_:+,.1f}<small>{etq_u(u)}</small>", "Positivo = cumple", NAVY if K["cumple_to"] else ROJO)
    with g2:
        grafico(fig_brecha(c, K, esc), f"brecha_{c['id']}")
    sig = 1 if K["may"] else -1
    filas = [[MESES[i], f_val(K["a"][j], u, False), f_val(K["t"][j], u, False), f_val(c["estandar"], u, False),
              f"{(float(disp(K['a'][j], u)) - float(disp(c['estandar'], u))) * sig:+,.1f}", f"{(float(disp(K['t'][j], u)) - float(disp(c['estandar'], u))) * sig:+,.1f}",
              "Sí" if cumple(K["t"][j], c["estandar"], K["may"]) else "No"] for j, i in enumerate(K["idx"])]
    with st.expander("Brechas por mes"):
        tabla(pd.DataFrame(filas, columns=["Mes", "As-Is", "To-Be", "Estándar", "Brecha As-Is", "Brecha To-Be", "To-Be cumple"]), alto=340, num=("As-Is", "To-Be", "Estándar", "Brecha As-Is", "Brecha To-Be"))


def tab_tendencias(cs, idx):
    esc = S.escenario
    tend = st.toggle("Mostrar línea de tendencia lineal", value=True, key="tend_tt")
    x = [MESES[i] for i in idx]
    xs = np.arange(len(x))
    una = len(cs) == 1
    fig = go.Figure()
    filas = []
    for cc in cs:
        K = calc(cc, idx)
        col = META[cc["codigo"]]["color"]
        la = np.array([logro(v, cc["estandar"], K["may"]) for v in K["a"]]) * 100
        lt = np.array([logro(v, cc["estandar"], K["may"]) for v in K["t"]]) * 100
        if esc in ("Ambos", "As-Is"):
            fig.add_trace(go.Scatter(x=x, y=la, mode="lines+markers", name=f"{cc['codigo']} As-Is", line=dict(color=col, width=2, dash="dot"), marker=dict(size=7),
                                     hovertemplate=f"<b>{cc['codigo']} As-Is</b> %{{x}}: %{{y:.0f}}% de la meta<extra></extra>"))
        if esc in ("Ambos", "To-Be"):
            fig.add_trace(go.Scatter(x=x, y=lt, mode="lines+markers+text" if una else "lines+markers", name=f"{cc['codigo']} To-Be", line=dict(color=col, width=3.5),
                                     marker=dict(size=9), text=[f"{v:.0f}%" for v in lt] if una else None, textposition="top center", textfont=dict(size=11, color=NEGRO),
                                     hovertemplate=f"<b>{cc['codigo']} To-Be</b> %{{x}}: %{{y:.0f}}% de la meta<extra></extra>"))
        pa = np.polyfit(xs, la, 1)[0] if len(x) >= 3 else float("nan")
        pt = np.polyfit(xs, lt, 1)[0] if len(x) >= 3 else float("nan")
        if tend and len(x) >= 3:
            if esc in ("Ambos", "As-Is"):
                fig.add_trace(go.Scatter(x=x, y=np.polyval(np.polyfit(xs, la, 1), xs), mode="lines", line=dict(color=col, width=1.5, dash="longdash"), name=f"Tendencia {cc['codigo']} As-Is", hoverinfo="skip", showlegend=una))
            if esc in ("Ambos", "To-Be"):
                fig.add_trace(go.Scatter(x=x, y=np.polyval(np.polyfit(xs, lt, 1), xs), mode="lines", line=dict(color=col, width=1.5, dash="longdash"), name=f"Tendencia {cc['codigo']} To-Be", hoverinfo="skip", showlegend=una))
        filas.append([nn(cc), cc["codigo"], f"{pa:+.2f} p.p./mes", f"{pt:+.2f} p.p./mes", f"{np.mean(la):.0f}%", f"{np.mean(lt):.0f}%"])
    fig.add_hline(y=100, line_dash="dash", line_color=NEGRO, line_width=2, annotation_text="Meta = 100%", annotation_font=dict(size=12, color=NEGRO))
    fig.update_layout(title="Tendencia mensual del logro de la meta (100% = alcanza el estándar)", yaxis_title="% de logro de la meta", xaxis_title="")
    grafico(estilo(fig, 440), f"tend_{'_'.join(c['id'] for c in cs)}")
    tabla(pd.DataFrame(filas, columns=["N.º", "Metodología", "Pendiente As-Is", "Pendiente To-Be", "Logro medio As-Is", "Logro medio To-Be"]),
          num=("Pendiente As-Is", "Pendiente To-Be", "Logro medio As-Is", "Logro medio To-Be"))
    nota("Las pendientes son la recta de mínimos cuadrados de los valores mensuales: un cálculo descriptivo, no un dato adicional.")


def tab_objetivos(cs, idx):
    conc = D["conclusiones"]
    if len(cs) > 1:
        g = conc["general"]
        Ks = [calc(cc, idx) for cc in cs]
        n_ok = sum(K["cumple_to"] for K in Ks)
        if len(cs) == len(CAUSAS):
            k = st.columns(4, gap="medium")
            with k[0]: stat("Objetivo general", f"{g['to_be'] * 100:.1f}%", f"As-Is {g['as_is'] * 100:.1f}% · VR {g['estandar'] * 100:.0f}% · {g['resultado']}", NAVY)
            with k[1]: stat("Objetivos específicos cumplidos", f"{n_ok} de {len(cs)}", "To-Be frente al estándar", NAVY if n_ok == len(cs) else ROJO)
            with k[2]: stat("Brecha técnica", f"{PB['brecha_tecnica'] * 100:.1f}<small>p.p.</small>", "BT = VR − VIP", ROJO)
            with k[3]: stat("Variación del cumplimiento OT", f"+{g['variacion'] * 100:.1f}%", "To-Be frente a As-Is", AZUL)
        else:
            k = st.columns(2, gap="medium")
            with k[0]: stat("Objetivos específicos cumplidos", f"{n_ok} de {len(cs)}", "To-Be frente al estándar", NAVY if n_ok == len(cs) else ROJO)
            fuera = ", ".join(cc["codigo"] for cc, K in zip(cs, Ks) if not K["cumple_to"]) or "Ninguna"
            with k[1]: stat("Causas por debajo de la meta", fuera, "Con la mejora aplicada", ROJO if fuera != "Ninguna" else NAVY)
        fig = go.Figure(go.Bar(x=[etq_causa(cc) for cc in cs], y=[K["logro_to"] * 100 for K in Ks],
                               marker=dict(color=[SEM_COL[K["sem_to"]] for K in Ks], line=dict(color=NEGRO, width=1)), customdata=[cc["id"] for cc in cs],
                               text=[f"{K['logro_to'] * 100:.0f}%" for K in Ks], textposition="outside", textfont=dict(size=13, color=NEGRO), cliponaxis=False,
                               hovertemplate="<b>%{x}</b><br>Logro To-Be: %{y:.1f}% de la meta<extra></extra>"))
        fig.add_hline(y=100, line_dash="dash", line_color=NEGRO, line_width=3, annotation_text="Meta = 100%", annotation_font=dict(size=12, color=NEGRO))
        fig.update_layout(title="Cumplimiento de objetivos específicos (color = semáforo del To-Be · clic para añadir o quitar)", yaxis_title="% de logro de la meta", showlegend=False)
        grafico(estilo(fig, 380, leyenda_h=False), "obj_all", seleccion=True)
        if len(cs) == len(CAUSAS):
            lectura(g["texto"], etq="Conclusión general")
        else:
            for cc in cs:
                lectura(conc["especificas"][cc["codigo"]]["texto"], etq=f"{nn(cc)} · {nombre(cc)} ({cc['codigo']})", color=META[cc["codigo"]]["color"])
                st.write("")
        return
    c = cs[0]
    K = calc(c, idx)
    e = conc["especificas"][c["codigo"]]
    u = c["unidad"]
    n_ok = int(sum(cumple(v, c["estandar"], K["may"]) for v in K["t"]))
    g1, g2 = st.columns([1, 1.3], gap="large")
    with g1:
        fig = go.Figure(go.Indicator(mode="gauge+number+delta", value=round(K["logro_to"] * 100, 1), number=dict(suffix="%", font=dict(family=FD, size=44, color=NEGRO)),
                                     delta=dict(reference=round(K["logro_as"] * 100, 1), relative=False, valueformat=".1f", suffix=" pts vs As-Is", font=dict(size=14)),
                                     title=dict(text="Logro de la meta (To-Be)", font=dict(family=FD, size=18, color=NAVY)),
                                     gauge=dict(axis=dict(range=[0, max(120, K["logro_to"] * 100 * 1.1)], tickfont=dict(size=12, color=NEGRO)), bar=dict(color=SEM_COL[K["sem_to"]]),
                                                steps=[dict(range=[0, K["logro_as"] * 100], color=SEM_TINT["rojo"]), dict(range=[K["logro_as"] * 100, 100], color="#EEF1F7")],
                                                threshold=dict(line=dict(color=NEGRO, width=5), thickness=.85, value=100))))
        fig.update_layout(margin=dict(t=90, b=10))
        grafico(estilo(fig, 330, leyenda_h=False), f"gauge_{c['id']}")
        nota(f'{chip_sem(K["sem_to"])} &nbsp; Resultado: <b>{e["resultado"]}</b> · Meses que cumplen el estándar: <b>{n_ok} de {len(idx)}</b>')
    with g2:
        k = st.columns(3, gap="medium")
        with k[0]: stat("As-Is", val_html(e["as_is"], u), "Situación actual", ROJO)
        with k[1]: stat("To-Be", val_html(e["to_be"], u), "Con la mejora", NAVY)
        with k[2]: stat("Estándar", val_html(e["estandar"], u), "Meta", GRIS)
        lectura(f'<b>Objetivo específico:</b> {c["objetivo"]}<br><br>{e["texto"]}', etq="Conclusión")


def tab_economia(cs, idx):
    if len(cs) > 1:
        inv = [inversion_de(cc["codigo"]) for cc in cs]
        ev = [eval_de(cc["codigo"]) for cc in cs]
        fig = go.Figure()
        nm_ = [etq_causa(cc) for cc in cs]
        ids = [cc["id"] for cc in cs]
        fig.add_trace(go.Bar(x=nm_, y=[i["inversion"] for i in inv], name="Inversión (año 0)", marker_color=NAVY, customdata=ids, text=[f"{i['inversion']:,.0f}" for i in inv], textposition="outside", textfont=dict(size=11, color=NEGRO)))
        fig.add_trace(go.Bar(x=nm_, y=[e_["beneficio"] for e_ in ev], name="Beneficio económico anual", marker_color=ROJO, customdata=ids, text=[f"{e_['beneficio']:,.0f}" for e_ in ev], textposition="outside", textfont=dict(size=11, color=NEGRO)))
        fig.update_layout(title="Inversión frente a beneficio anual por causa (S/) · clic para añadir o quitar", barmode="group", yaxis_title="S/", xaxis_title="")
        grafico(estilo(fig, 400), "eco_all", seleccion=True)
        filas = [[nn(cc), cc["codigo"], f"{i['inversion']:,.0f}", f"{e_['costo_asis']:,.0f}", f"{e_['costo_tobe']:,.0f}", f"{e_['beneficio']:,.0f}", f"{e_['pct'] * 100:.1f}%"] for cc, i, e_ in zip(cs, inv, ev)]
        if len(cs) == len(CAUSAS):
            filas.append(["", "Total", f"{sum(i['inversion'] for i in inv):,.0f}", f"{FL['eval_total']['costo_asis']:,.0f}", f"{FL['eval_total']['costo_tobe']:,.0f}", f"{FL['eval_total']['beneficio']:,.0f}", "100.0%"])
        else:
            filas.append(["", "Total selección", f"{sum(i['inversion'] for i in inv):,.0f}", f"{sum(e_['costo_asis'] for e_ in ev):,.0f}", f"{sum(e_['costo_tobe'] for e_ in ev):,.0f}",
                          f"{sum(e_['beneficio'] for e_ in ev):,.0f}", f"{sum(e_['pct'] for e_ in ev) * 100:.1f}%"])
        tabla(pd.DataFrame(filas, columns=["N.º", "Metodología", "Inversión (S/)", "Costo anual As-Is (S/)", "Costo anual To-Be (S/)", "Beneficio anual (S/)", "% del beneficio"]),
              num=("Inversión (S/)", "Costo anual As-Is (S/)", "Costo anual To-Be (S/)", "Beneficio anual (S/)", "% del beneficio"), total_ultima=True)
        return
    c = cs[0]
    cod = c["codigo"]
    ev, ei = eval_de(cod), inversion_de(cod)
    k = st.columns(5, gap="medium")
    with k[0]: stat("Inversión de la propuesta", soles(ei["inversion"]), f"{ei['pct'] * 100:.1f}% del total (sin imprevistos)", NAVY)
    with k[1]: stat("Costo anual As-Is", soles(ev["costo_asis"]), "Situación actual", ROJO)
    with k[2]: stat("Costo anual To-Be", soles(ev["costo_tobe"]), "Con la mejora", NAVY)
    with k[3]: stat("Beneficio anual", soles(ev["beneficio"]), "As-Is − To-Be", AZUL)
    with k[4]: stat("% del beneficio total", f"{ev['pct'] * 100:.1f}%", f"de {soles(FL['eval_total']['beneficio'])}", ROJO)
    cm = FL["costos_mensuales"].get(cod)
    if cm:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=MESES, y=cm["asis"], name="Costo As-Is", marker_color=ROJO, text=[f"{v:,.0f}" for v in cm["asis"]], textposition="outside", textfont=dict(size=10, color=NEGRO), cliponaxis=False))
        fig.add_trace(go.Bar(x=MESES, y=cm["tobe"], name="Costo To-Be", marker_color=NAVY, text=[f"{v:,.0f}" for v in cm["tobe"]], textposition="outside", textfont=dict(size=10, color=NEGRO), cliponaxis=False))
        fig.add_trace(go.Scatter(x=MESES, y=cm["ahorro"], name="Ahorro mensual", mode="lines+markers", line=dict(color=AZUL_CLARO, width=3), marker=dict(size=8)))
        fig.update_layout(title=f"Costo mensual As-Is frente a To-Be · {cod} (S/)", barmode="group", yaxis_title="S/", xaxis_title="")
        grafico(estilo(fig, 400), f"costo_{cod}")
    else:
        aa = FL["asistencia_addie"]
        to_be = [a + r for a, r in zip(aa["asistieron"], aa["recuperadas"])]
        fig = go.Figure()
        fig.add_trace(go.Bar(x=MESES, y=aa["asistieron"], name="Asistentes As-Is", marker_color=ROJO, text=aa["asistieron"], textposition="outside", textfont=dict(size=11, color=NEGRO)))
        fig.add_trace(go.Bar(x=MESES, y=to_be, name="Asistentes To-Be (As-Is + recuperadas)", marker_color=NAVY, text=[f"{v:.1f}" for v in to_be], textposition="outside", textfont=dict(size=11, color=NEGRO)))
        fig.add_hline(y=aa["convocados"][0], line_dash="dot", line_color=GRIS, annotation_text=f"Convocados: {aa['convocados'][0]}", annotation_font=dict(size=12, color=NEGRO))
        fig.update_layout(title="Asistencia a capacitaciones: As-Is frente a To-Be", barmode="group", yaxis_title="colaboradores", xaxis_title="")
        grafico(estilo(fig, 400), "asis_addie")
        nota(f"ADDIE registra un <b>beneficio económico directo de {soles(aa['beneficio_directo'])}</b>: su aporte es habilitar a los técnicos. Se recuperan {aa['total_recuperadas']:.1f} asistencias al año.")
    a1, a2 = st.columns(2, gap="large")
    with a1:
        subtitulo(f"Ítems de inversión · {cod}")
        it = [[x["concepto"], x["cantidad"], f"{x['unitario']:,.0f}", f"{x['total']:,.0f}", x["clase"]] for x in FL["inversion_items"] if x["propuesta"] == cod]
        tabla(pd.DataFrame(it, columns=["Concepto", "Cant.", "Costo unit. (S/)", "Total (S/)", "Clase"]), alto=300, num=("Cant.", "Costo unit. (S/)", "Total (S/)"))
    with a2:
        subtitulo(f"Costos anuales de sostenimiento · {cod}")
        so = [[x["concepto"], x["cantidad"], f"{x['unitario']:,.2f}", f"{x['monto']:,.0f}", x["desde"]] for x in FL["sostenimiento"] if x["propuesta"] == cod]
        if so:
            tabla(pd.DataFrame(so, columns=["Concepto", "Cant./año", "Unit. (S/)", "Monto anual (S/)", "Desde el año"]), alto=300, num=("Cant./año", "Unit. (S/)", "Monto anual (S/)"))
        else:
            st.caption("Esta propuesta no registra costos de sostenimiento.")


# ======================= PÁGINA: PANEL INTERACTIVO =======================
def pagina_panel():
    cabecera("Panel interactivo", titulo_pb(),
             ["Elige una o varias causas y todo el panel se actualiza", f"Referencia: {PB['vr'] * 100:.0f}% de cumplimiento de OT", "Comparación As-Is · To-Be"])
    F = FL["kpi"]
    seccion("Problema central", f"{PB['indicador'].splitlines()[0].strip()} · valor de referencia VR = {PB['vr'] * 100:.0f}%", "01")
    r1 = st.columns(3, gap="large")
    with r1[0]: stat("Brecha técnica", f"{PB['brecha_tecnica'] * 100:.1f}<small>p.p.</small>", f"BT = VR − VIP = {PB['vr'] * 100:.0f}% − {PB['vip'] * 100:.1f}%", ROJO)
    with r1[1]: stat("Cumplimiento de OT", f"{PB['as_is'] * 100:.1f}% → {PB['to_be'] * 100:.1f}%", f"Situación actual → con la mejora · VR {PB['vr'] * 100:.0f}%", NAVY)
    with r1[2]: stat("Impacto económico", soles(PB["impacto_economico"]), f"{PB['pct_ingresos'] * 100:.1f}% de los ingresos ({soles(PB['ingresos_anuales'])})", ROJO)
    md('<div style="height:14px"></div>')
    r2 = st.columns(3, gap="large")
    with r2[0]: stat("Inversión total", soles(F["inversion"]), "5 propuestas · año 0", NAVY)
    with r2[1]: stat("Ahorro anual", soles(FL["eval_total"]["beneficio"]), f"Costo {soles(FL['eval_total']['costo_asis'])} → {soles(FL['eval_total']['costo_tobe'])}", AZUL)
    with r2[2]: stat("VAN", soles(F["van"]), f"TIR {F['tir'] * 100:.1f}% · CPPC {F['cppc'] * 100:.2f}%", NAVY)

    seccion("Causas raíz", "Cada tarjeta es un botón: haz clic para añadirla a la selección y otro clic para quitarla. Puedes combinar las que quieras.", "02")
    cs = activas()
    k_sel = len(S.sel)
    s1, s2, s3 = st.columns([6, 1.1, 1.1], gap="small")
    with s1:
        md(f'<div class="selbar"><b>{k_sel} de {len(CAUSAS)}</b> causas seleccionadas</div>' if k_sel else '<div class="selbar">Sin selección: se muestran las <b>5 causas</b></div>')
    with s2:
        st.button("Todas", key="ctl_todas", on_click=todas, use_container_width=True)
    with s3:
        st.button("Ninguna", key="ctl_ninguna", on_click=ninguna, use_container_width=True)
    idx = idx_meses()
    cols = st.columns(5, gap="small")
    for col, c in zip(cols, CAUSAS):
        K = calc(c, idx)
        on = c["id"] in S.sel
        with col:
            with st.container(key="card_" + c["id"]):
                md(tarjeta_causa(c, K, on))
                st.button("Quitar de la selección" if on else "Añadir a la selección", key="cardbtn_" + c["id"], on_click=alternar, args=(c["id"],),
                          help=f"{nombre(c)} ({c['codigo']}): clic para {'quitarla de' if on else 'añadirla a'} la selección")

    seccion("Periodo y escenario", "Los meses y el escenario se aplican a todos los indicadores del panel.", "03")
    f1, f2 = st.columns([3.4, 1.6], gap="large")
    with f1:
        etiqueta("Meses (sin selección = todo el año)")
        try:
            st.pills("Meses", MESES, selection_mode="multi", key="meses_sel", label_visibility="collapsed")
        except Exception:
            st.multiselect("Meses", MESES, key="meses_sel", label_visibility="collapsed")
        S["_u_meses_sel"] = list(S.meses_sel)
    with f2:
        etiqueta("Escenario")
        selector(["Ambos", "As-Is", "To-Be"], "escenario", "seg")
    idx = idx_meses()

    seccion("Comparativo", "Las causas elegidas se resaltan; el resto se atenúa. La barra muestra el avance frente a la meta (la línea vertical es el 100%).", "04")
    ranking(idx)
    st.write("")
    grafico(fig_mejora_barras(idx), "mejora_click", seleccion=True)

    titulo_sel, sub_sel = resumen_sel()
    seccion(f"Detalle · {titulo_sel}", sub_sel, "05")
    tabs = st.tabs(["KPI y semáforo", "Antes y después", "Estándar y brechas", "Tendencias", "Objetivos", "Herramienta y metodología", "Economía"])
    with tabs[0]: tab_kpi(cs, idx)
    with tabs[1]: tab_antes_despues(cs, idx)
    with tabs[2]: tab_brechas(cs, idx)
    with tabs[3]: tab_tendencias(cs, idx)
    with tabs[4]: tab_objetivos(cs, idx)
    with tabs[5]: panel_metodologia(cs)
    with tabs[6]: tab_economia(cs, idx)


# ======================= PÁGINA: PROBLEMA CENTRAL =======================
def pagina_central():
    cabecera("Problema central", titulo_pb(), ["26 quincenas en 2025", f"Referencia: {PB['vr'] * 100:.0f}% de cumplimiento de OT", "Vinculado a las 5 causas raíz"])
    Q = pd.DataFrame(D["quincenas"])
    Q["lab"] = Q["periodo"].str.replace("SEMANA ", "S", regex=False).str.replace(" y ", "-", regex=False)
    n = len(Q)
    seccion("Cumplimiento por quincena", "Filtra el rango de quincenas y la vista del gráfico.", "01")
    f1, f2, f3 = st.columns([2.2, 3.2, 1.2], gap="large")
    with f1:
        rango = st.select_slider("Rango de quincenas", options=list(range(1, n + 1)), value=(1, n), key="rango_q", format_func=lambda i: Q["lab"][i - 1])
    with f2:
        etiqueta("Vista")
        vista = selector(["% cumplimiento", "OT programadas vs ejecutadas", "Costo por quincena"], "vista_q")
    with f3:
        tend = st.toggle("Tendencia lineal", value=True, key="tend_q")
    q = Q.iloc[rango[0] - 1: rango[1]].reset_index(drop=True)
    vip = q["cumplimiento"].mean()
    st.write("")
    k = st.columns(5, gap="medium")
    with k[0]: stat("Cumplimiento de OT (VIP)", f"{vip * 100:.1f}%", f"Promedio de {len(q)} quincenas · VR {PB['vr'] * 100:.0f}%", ROJO)
    with k[1]: stat("Brecha técnica", f"{(PB['vr'] - vip) * 100:.1f}<small>p.p.</small>", "BT = VR − VIP", ROJO_OSC)
    with k[2]: stat("OT ejecutadas / programadas", f"{int(q['ejecutadas'].sum())} / {int(q['programadas'].sum())}", "En el rango elegido", AZUL)
    with k[3]: stat("Costo del periodo", soles(q["costo_total"].sum()), f"Tracto {soles(q['costo_tracto'].sum())} · Tolva/cist. {soles(q['costo_tolva'].sum())}", NAVY)
    with k[4]: stat("Quincenas ≥ VR", f"{int((q['cumplimiento'] >= PB['vr']).sum())} de {len(q)}", f"Alcanzan el {PB['vr'] * 100:.0f}%", NAVY)
    st.write("")
    x = q["lab"]
    fig = go.Figure()
    if vista == "% cumplimiento":
        y = q["cumplimiento"] * 100
        fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers+text", name="% cumplimiento", text=[f"{v:.0f}" for v in y], textposition="top center", textfont=dict(size=10, color=NEGRO),
                                 line=dict(color=ROJO, width=3), marker=dict(size=9, color=NAVY), hovertemplate="<b>%{x}</b><br>%{y:.1f}%<extra></extra>"))
        if tend and len(q) >= 3:
            p = np.polyfit(np.arange(len(q)), y, 1)
            fig.add_trace(go.Scatter(x=x, y=np.polyval(p, np.arange(len(q))), mode="lines", name=f"Tendencia ({p[0]:+.2f} p.p./quincena)", line=dict(color=NAVY, dash="dot", width=2)))
        fig.add_hline(y=PB["vr"] * 100, line_dash="dash", line_color=NAVY, line_width=2, annotation_text=f"VR {PB['vr'] * 100:.0f}%", annotation_font=dict(size=12, color=NEGRO))
        fig.update_layout(title="% de cumplimiento de OT por quincena", yaxis_title="%", yaxis_range=[0, 112])
    elif vista.startswith("OT"):
        fig.add_trace(go.Bar(x=x, y=q["programadas"], name="Programadas", marker_color="#ADB5BD", text=q["programadas"], textposition="outside", textfont=dict(size=10, color=NEGRO)))
        fig.add_trace(go.Bar(x=x, y=q["ejecutadas"], name="Ejecutadas", marker_color=ROJO, text=q["ejecutadas"], textposition="outside", textfont=dict(size=10, color=NEGRO)))
        fig.update_layout(title="OT preventivas programadas frente a ejecutadas", barmode="group", yaxis_title="N.º de OT")
    else:
        fig.add_trace(go.Bar(x=x, y=q["costo_tracto"], name="Tracto", marker_color=NAVY))
        fig.add_trace(go.Bar(x=x, y=q["costo_tolva"], name="Tolva / cisterna", marker_color=ROJO))
        fig.update_layout(title="Costo por quincena (S/)", barmode="stack", yaxis_title="S/")
    fig.update_layout(xaxis_title="")
    grafico(estilo(fig, 430), "q_main")
    if (Q["costo_total"] == 0).any():
        nota("La última quincena (S50-51) figura con costo 0 y se muestra tal como está registrada.")

    seccion("Cómo se vincula el problema con las causas", "Cada causa crea una barrera que impide ejecutar la OT. Las causas elegidas en el panel se resaltan.", "02")
    filas = []
    for cc, b in zip(CAUSAS, D["barreras"]):
        K = calc(cc, idx_meses())
        filas.append([nn(cc), cc["motivo"], b["barrera"], cc["codigo"], f"{f_val(K['as_is'], cc['unidad'])} → {f_val(K['to_be'], cc['unidad'])}", chip_sem(K["sem_to"])])
    sel = [i for i, cc in enumerate(CAUSAS) if cc["id"] in S.sel]
    tabla(pd.DataFrame(filas, columns=["N.º", "Causa raíz", "Barrera en el cumplimiento de OT", "Herramienta", "As-Is → To-Be", "Semáforo"]), filas_sel=sel)
    c1, c2 = st.columns(2, gap="large")
    o = FL["ot"]
    with c1:
        fig = go.Figure()
        ms = [m["mes"][:3] for m in o["meses"]]
        fig.add_trace(go.Bar(x=ms, y=[m["asis"] for m in o["meses"]], name="OT ejecutadas As-Is", marker_color=ROJO))
        fig.add_trace(go.Bar(x=ms, y=[m["tobe"] for m in o["meses"]], name="OT ejecutadas To-Be", marker_color=NAVY))
        fig.add_trace(go.Scatter(x=ms, y=[m["programadas"] for m in o["meses"]], name="Programadas", mode="lines+markers", line=dict(color=GRIS, dash="dot")))
        fig.update_layout(title="OT ejecutadas por mes: As-Is frente a To-Be", barmode="group", yaxis_title="N.º de OT")
        grafico(estilo(fig, 380), "ot_mes")
    with c2:
        s = SEM[PC]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=MESES, y=np.array(s["mensual_asis"]) * 100, name="As-Is", mode="lines+markers", line=dict(color=ROJO, width=3)))
        fig.add_trace(go.Scatter(x=MESES, y=np.array(s["mensual_tobe"]) * 100, name="To-Be", mode="lines+markers", line=dict(color=NAVY, width=3)))
        fig.add_hline(y=s["meta"] * 100, line_dash="dash", line_color=NEGRO, annotation_text=f"Meta {s['meta'] * 100:.0f}%", annotation_font=dict(size=12, color=NEGRO))
        fig.update_layout(title="Cumplimiento de OT mensual", yaxis_title="%")
        grafico(estilo(fig, 380), "pc_mensual")
    s = SEM[PC]
    nota(f"Cumplimiento As-Is: <b>{PB['as_is'] * 100:.2f}%</b> como promedio de quincenas, frente a <b>{s['anual_asis'] * 100:.2f}%</b> como promedio de los 12 meses. "
         "El panel usa el primero como cifra oficial.")
    lectura(f'{o["total"]["adicionales"]:.1f} OT adicionales al año · ingreso incremental {soles(o["ingreso_incremental"])} · costo de ejecutarlas {soles(o["costo_adicional"])} · '
            f'margen incremental {soles(o["margen_incremental"])}.', etq="Resultado económico de cerrar la brecha")


# ======================= PÁGINA: SEMÁFORO =======================
def pagina_semaforo():
    cabecera("Semáforo y termómetros", "Estado de cada indicador frente a su meta y a su línea base", ["Cumple", "En proceso", "Crítico"])
    etiqueta("Escenario")
    esc = selector(["To-Be", "As-Is"], "esc_sem", "seg")
    idx = idx_meses()
    items = [(PC, "Problema central", PB["as_is"] if esc == "As-Is" else PB["to_be"], "%")]
    for c in CAUSAS:
        K = calc(c, idx)
        items.append((c["codigo"], f"{nn(c)} · {nombre(c)}", K["as_is"] if esc == "As-Is" else K["to_be"], c["unidad"]))
    est = [estado_sem(cod, v) for cod, _, v, _ in items]
    seccion("Resumen", None, "01")
    k = st.columns(3, gap="large")
    for col, e in zip(k, ("verde", "amarillo", "rojo")):
        with col:
            stat(SEM_TXT[e], str(est.count(e)), f"de {len(items)} indicadores · {esc}", SEM_COL[e])
    ex = D["semaforo"]
    nota(f"Semaforización oficial del To-Be anual: {ex['n_cumple']} cumplen · {ex['n_proceso']} en proceso · {ex['n_critico']} críticos.")
    seccion("Indicadores", None, "02")
    cols = st.columns(3, gap="medium")
    for i, ((cod, nm_, v, u), e) in enumerate(zip(items, est)):
        with cols[i % 3]:
            s = SEM[cod]
            md(f'<div class="ficha" style="--ac:{SEM_COL[e]};"><div class="fn">{nm_}</div><div class="fv">{val_html(v, u)}</div>'
               f'<div class="ficha-pie"><span class="fs">Meta {f_val(s["meta"], u)} · Base {f_val(s["linea_base"], u)}</span>{chip_sem(e)}</div></div>')
            st.write("")
    seccion("Termómetros", None, "03")
    cols = st.columns(3, gap="medium")
    for i, (cod, nm_, v, u) in enumerate(items):
        with cols[i % 3]:
            grafico(fig_termometro(cod, u, [(esc, v)], nm_, 400), f"sem_term_{cod}")
    seccion("Estado mensual", "Cada punto aplica las reglas del semáforo al valor del mes.", "04")
    lista = [("PC · Cumplimiento OT", PC, SEM[PC]["mensual_asis" if esc == "As-Is" else "mensual_tobe"], "%")]
    for c in CAUSAS:
        lista.append((etq_causa(c), c["codigo"], serie(c["codigo"], "AS-IS" if esc == "As-Is" else "TO-BE"), c["unidad"]))
    grafico(fig_semaforo_mensual(esc, lista), "heat_sem")


# ======================= PÁGINA: FLUJO FINANCIERO =======================
BOTONES_FIN = ["Inversión", "COK", "CPPC", "Flujo de caja", "VAN", "TIR", "B/C", "Recuperación"]
NOMBRES_FIN = {"Inversión": "Inversión del proyecto", "COK": "Costo de oportunidad del capital (COK)", "CPPC": "Costo ponderado del capital (CPPC)", "Flujo de caja": "Flujo de caja",
               "VAN": "Valor actual neto (VAN)", "TIR": "Tasa interna de retorno (TIR)", "B/C": "Relación beneficio / costo", "Recuperación": "Recuperación de la inversión"}
ICONOS_FIN = {"Inversión": "🧾", "COK": "🏦", "CPPC": "⚖️", "Flujo de caja": "💵", "VAN": "📈", "TIR": "🔁", "B/C": "🎁", "Recuperación": "⏳"}
COLOR_FIN = {"Inversión": NAVY, "COK": AZUL, "CPPC": AZUL_CLARO, "Flujo de caja": AZUL, "VAN": NAVY, "TIR": ROJO, "B/C": ROJO_OSC, "Recuperación": ROJO}
TINTE_FIN = {"Inversión": "#E3E9F6", "COK": "#E3ECF9", "CPPC": "#E9F0FB", "Flujo de caja": "#E3ECF9", "VAN": "#E3E9F6", "TIR": "#FBE3E7", "B/C": "#F6DDE1", "Recuperación": "#FBE3E7"}
S.setdefault("fin_ind", "VAN")


def fin_flujos():
    return [FL["detalle_flujo"]["flujo_anio0"]] + list(FL["detalle_flujo"]["flujo"])


def elegir_fin(nombre_):
    S.fin_ind = nombre_


# ---------- mini-gráficos de las tarjetas (HTML puro) ----------
def sp_barras(vals, color, signed=False, colores=None):
    mx = max(abs(v) for v in vals) or 1
    out = ""
    for i, v in enumerate(vals):
        col = colores[i] if colores else (color if v >= 0 else ROJO)
        h = max(abs(v) / mx, .05)
        pos = (f"bottom:50%;height:{h * 50:.1f}%" if v >= 0 else f"top:50%;height:{h * 50:.1f}%") if signed else f"bottom:0;height:{h * 100:.1f}%"
        out += f'<span class="b"><i style="{pos};background:{col};"></i></span>'
    return f'<div class="fk-sp">{out}</div>'


def sp_seg(partes, colores):
    tot = sum(partes) or 1
    return '<div class="fk-seg">' + "".join(f'<i style="width:{p / tot * 100:.2f}%;background:{c};"></i>' for p, c in zip(partes, colores)) + "</div>"


def sp_comp(a, b):
    """Dos barras horizontales comparadas: a=(rótulo, valor, texto, color), b=igual."""
    mx = max(a[1], b[1]) or 1
    fila = lambda x: f'<div class="fk-h"><span>{x[0]}</span><div><i style="width:{max(x[1] / mx, .04) * 100:.1f}%;background:{x[3]};"></i></div><b>{x[2]}</b></div>'
    return f'<div class="fk-hh">{fila(a)}{fila(b)}</div>'


def chip_neutro(texto, color=AZUL):
    return f'<span class="fk-b" style="--c:{color};">{texto}</span>'


def tarjeta_fin(nombre_, valor, badge, spark, on):
    c = COLOR_FIN[nombre_]
    return (f'<div class="fk {"on" if on else ""}" style="--ac:{c};--tn:{TINTE_FIN[nombre_]};"><div class="fk-top"><span class="fk-ic">{ICONOS_FIN[nombre_]}</span>'
            f'<span class="fk-l">{nombre_}</span><span class="fk-ck"></span></div><div class="fk-v">{valor}</div>{spark}<div class="fk-foot">{badge}</div></div>')


def fig_donut(etiquetas, valores, colores, centro, titulo, sub=None):
    fig = go.Figure(go.Pie(labels=etiquetas, values=valores, hole=.64, sort=False, direction="clockwise", marker=dict(colors=colores, line=dict(color="#FFFFFF", width=3)),
                           textinfo="percent", textposition="outside", textfont=dict(size=14, color=NEGRO), hovertemplate="<b>%{label}</b><br>S/ %{value:,.0f} · %{percent}<extra></extra>"))
    fig.add_annotation(text=f"<b>{centro}</b>" + (f"<br><span style='font-size:12px'>{sub}</span>" if sub else ""), showarrow=False, font=dict(size=22, color=NAVY))
    fig.update_layout(title=titulo, showlegend=True)
    estilo(fig, 420, leyenda_h=False)
    fig.update_layout(legend=dict(orientation="v", x=1.0, y=.5, font=dict(size=13, color=NEGRO)))
    return fig


def selector_anio(sel, opciones, anios, defecto):
    key = f"fin_anio_{sel}"
    if S.get(key) not in opciones:
        S[key] = defecto
    return st.select_slider("Año", options=opciones, key=key, format_func=lambda i: f"Año {i} · {anios[i]}", label_visibility="collapsed")


def pagina_flujo():
    F = FL
    cabecera("Flujo financiero", "Evaluación económica del proyecto a 10 años", [f"Horizonte de {F['parametros'][2]['valor']} años", f"Año base {F['anios'][0]}", f"Inversión {soles(F['inversion_total'])}"])
    anios = F["anios"]
    fl = fin_flujos()
    cppc, cok, tir, bcr = F["cppc"]["valor"], F["cok"]["cok"], F["tir"]["valor"], F["bc"]["valor"]
    vp = [F["detalle_flujo"]["flujo_anio0"]] + list(F["detalle_flujo"]["valor_presente"])
    vp_ac = list(np.cumsum(vp))
    rec = F["recuperacion"]
    van = F["van"]["valor"]
    d = F["detalle_flujo"]
    cs = list(activas()) if S.sel else list(CAUSAS)
    codigos = {c["codigo"] for c in cs}
    if S.get("fin_ind") not in BOTONES_FIN:
        S.fin_ind = "VAN"
    sel = S.fin_ind

    # ------------------------------------------------------------------ tarjetas clicables
    seccion("Indicadores", "Cada tarjeta es un botón: haz clic para ver su gráfico y su detalle.", "01")
    inv_c = [inversion_de(c["codigo"])["inversion"] for c in CAUSAS]
    cp = F["cppc"]
    ck = F["cok"]
    tiles = {
        "Inversión": (soles(F["inversion_total"]), chip_neutro("5 propuestas", NAVY), sp_barras(inv_c, NAVY, colores=[META[c["codigo"]]["color"] for c in CAUSAS])),
        "COK": (f"{cok * 100:.2f}%", chip_neutro("Modelo CAPM"), sp_seg([x["valor"] for x in ck["aportes"]], [NAVY, AZUL, AZUL_CLARO, ROJO])),
        "CPPC": (f"{cppc * 100:.2f}%", chip_neutro("Tasa de descuento"), sp_seg([cp["pasivos"]["pond"], cp["patrimonio"]["pond"]], [AZUL_CLARO, NAVY])),
        "Flujo de caja": (soles(sum(fl[1:])), chip_neutro("Acumulado 10 años"), sp_barras(fl[1:], AZUL, signed=True)),
        "VAN": (soles(van), chip_sem("verde" if van > 0 else "rojo", "Rentable" if van > 0 else "No rentable"), sp_barras(vp_ac, NAVY, signed=True)),
        "TIR": (f"{tir * 100:.1f}%", chip_sem("verde" if tir > cppc else "rojo", "Supera al CPPC" if tir > cppc else "Bajo el CPPC"),
                sp_comp(("TIR", tir, f"{tir * 100:.0f}%", ROJO), ("CPPC", cppc, f"{cppc * 100:.0f}%", AZUL_CLARO))),
        "B/C": (f"{bcr:.2f}", chip_sem("verde" if bcr > 1 else "rojo", "Rentable" if bcr > 1 else "No rentable"),
                sp_comp(("Beneficio", F["van"]["suma_vp"], f"{F['van']['suma_vp'] / 1000:,.0f} mil", ROJO_OSC), ("Costo", F["van"]["inversion"], f"{F['van']['inversion'] / 1000:,.0f} mil", GRIS))),
        "Recuperación": (F["kpi"]["recuperacion_txt"], chip_neutro(f"Simple: {rec['simple_txt']}", ROJO), sp_barras(rec["acum_simple"], ROJO, signed=True)),
    }
    for fila_ in (BOTONES_FIN[:4], BOTONES_FIN[4:]):
        cols = st.columns(4, gap="medium")
        for col, nm_ in zip(cols, fila_):
            v_, b_, s_ = tiles[nm_]
            with col:
                with st.container(key="fk_" + nm_):
                    md(tarjeta_fin(nm_, v_, b_, s_, nm_ == sel))
                    st.button(f"Ver {nm_}", key="fkbtn_" + nm_, on_click=elegir_fin, args=(nm_,), help=f"Ver el detalle de {NOMBRES_FIN[nm_]}")
        md('<div style="height:6px"></div>')

    # ------------------------------------------------------------------ ficha + gráfico del indicador elegido
    seccion(NOMBRES_FIN[sel], None, "02")
    hechos, extra, det, fig = [], None, None, None
    if sel == "COK":
        c = ck
        valor, veredicto, frase = f"{cok * 100:.2f}%", chip_neutro("Capital propio"), "Rentabilidad mínima que se le exige al aporte propio."
        formula = "COK = Rf + βL × (Rm − Rf) + riesgo país + prima por tamaño"
        comp = ck["aportes"]
        fig = go.Figure(go.Waterfall(x=["Tasa libre de riesgo", "Prima de mercado × βL", "Riesgo país", "Prima por tamaño", "COK"], measure=["relative"] * 4 + ["total"], y=[x["valor"] * 100 for x in comp] + [0],
                                     text=[f"{x['valor'] * 100:.2f}%" for x in comp] + [f"{cok * 100:.2f}%"], textposition="outside", textfont=dict(size=13, color=NEGRO),
                                     increasing=dict(marker=dict(color=AZUL)), totals=dict(marker=dict(color=NAVY)), connector=dict(line=dict(color=GRIS))))
        fig.update_layout(title="Cómo se forma el COK (CAPM)", yaxis_title="%", showlegend=False)
        hechos = [("Tasa libre de riesgo", f"{c['rf'] * 100:.2f}%", None), ("Beta apalancada (βL)", f"{c['beta_l']:.3f}", None), ("Prima de mercado", f"{c['prima_mercado'] * 100:.2f}%", None),
                  ("Riesgo país", f"{c['riesgo_pais'] * 100:.2f}%", None), ("Prima por tamaño", f"{c['prima_tamano'] * 100:.2f}%", None)]
        det = pd.DataFrame([[x["concepto"], f"{x['valor']:.4f}" if x["valor"] > 0.5 else f"{x['valor'] * 100:.2f}%"] for x in c["componentes"]], columns=["Componente", "Valor"])
    elif sel == "CPPC":
        valor, veredicto = f"{cppc * 100:.2f}%", chip_neutro("Tasa de corte")
        frase = f"Se descuenta con {cp['pasivos']['part'] * 100:.0f}% de deuda y {cp['patrimonio']['part'] * 100:.0f}% de capital propio."
        formula = "CPPC = %Deuda × Kd + %Patrimonio × COK"
        fig = fig_donut(["Deuda", "Patrimonio"], [cp["pasivos"]["pond"], cp["patrimonio"]["pond"]], [AZUL_CLARO, NAVY], f"{cppc * 100:.2f}%", "Aporte de cada fuente al CPPC", "CPPC")
        fig.update_traces(hovertemplate="<b>%{label}</b><br>Aporte: %{value:.4f}<extra></extra>")
        hechos = [("Deuda", f"{cp['pasivos']['costo'] * 100:.2f}%", f"{cp['pasivos']['part'] * 100:.0f}% · aporta {cp['pasivos']['pond'] * 100:.2f}%"),
                  ("Patrimonio (COK)", f"{cp['patrimonio']['costo'] * 100:.2f}%", f"{cp['patrimonio']['part'] * 100:.0f}% · aporta {cp['patrimonio']['pond'] * 100:.2f}%"),
                  ("Monto financiado", soles(cp["pasivos"]["monto"]), f"Aporte propio {soles(cp['patrimonio']['monto'])}")]
        det = pd.DataFrame([["Pasivos (préstamo bancario)", f"{cp['pasivos']['monto']:,.0f}", f"{cp['pasivos']['part'] * 100:.0f}%", f"{cp['pasivos']['costo'] * 100:.2f}%", f"{cp['pasivos']['pond'] * 100:.2f}%"],
                            ["Patrimonio (aporte propio)", f"{cp['patrimonio']['monto']:,.0f}", f"{cp['patrimonio']['part'] * 100:.0f}%", f"{cp['patrimonio']['costo'] * 100:.2f}%", f"{cp['patrimonio']['pond'] * 100:.2f}%"],
                            ["Total", f"{cp['total']['monto']:,.0f}", f"{cp['total']['part'] * 100:.0f}%", "", f"{cp['total']['pond'] * 100:.2f}%"]], columns=["Fuente", "Monto (S/)", "Participación", "Costo", "Costo ponderado"])
    elif sel == "VAN":
        v = F["van"]
        valor, veredicto = soles(van), chip_sem("verde" if van > 0 else "rojo", "Rentable" if van > 0 else "No rentable")
        frase, formula = f"El valor presente de los flujos supera la inversión en {soles(van)}.", f"VAN = −{soles(v['inversion'])} + {soles(v['suma_vp'])}  (CPPC {cppc * 100:.2f}%)"
        fig = go.Figure(go.Waterfall(x=["Inversión"] + [f"Año {i}" for i in range(1, len(vp))] + ["VAN"], measure=["relative"] * len(vp) + ["total"], y=vp + [0],
                                     text=[f"{x / 1000:,.0f}k" for x in vp] + [f"{van / 1000:,.0f}k"], textposition="outside", textfont=dict(size=11, color=NEGRO),
                                     increasing=dict(marker=dict(color=AZUL)), decreasing=dict(marker=dict(color=ROJO)), totals=dict(marker=dict(color=NAVY)), connector=dict(line=dict(color=GRIS))))
        fig.update_layout(title="Cómo se construye el VAN (S/)", yaxis_title="S/", showlegend=False)
        i_ = selector_anio(sel, list(range(len(fl))), anios, 1)
        hechos = [("Flujo de caja", soles(fl[i_]), f"Año {i_} · {anios[i_]}"), ("Valor presente", soles(vp[i_]), f"Descontado al CPPC"), ("VP acumulado", soles(vp_ac[i_]), "Incluye la inversión inicial")]
        det = pd.DataFrame([[i, anios[i], f"{fl[i]:,.0f}", f"{vp[i]:,.0f}", f"{vp_ac[i]:,.0f}"] for i in range(len(fl))], columns=["Año", "Calendario", "Flujo de caja (S/)", "Valor presente (S/)", "VP acumulado (S/)"])
    elif sel == "TIR":
        valor, veredicto = f"{tir * 100:.2f}%", chip_sem("verde" if tir > cppc else "rojo", "Supera al CPPC" if tir > cppc else "Bajo el CPPC")
        frase, formula = f"Supera en {(tir - cppc) * 100:.1f} p.p. a la tasa de corte ({cppc * 100:.2f}%).", "TIR = tasa que hace el VAN igual a cero"
        rs = np.linspace(0, 1.3, 66)
        npv = [sum(f_ / (1 + r) ** t for t, f_ in enumerate(fl)) for r in rs]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=rs * 100, y=npv, mode="lines", name="VAN según la tasa", line=dict(color=NAVY, width=3.5), fill="tozeroy", fillcolor="rgba(36,87,166,.10)", hovertemplate="Tasa %{x:.0f}%<br>VAN S/ %{y:,.0f}<extra></extra>"))
        fig.add_hline(y=0, line_color=NEGRO, line_width=2)
        for nm_, r, col in (("CPPC", cppc, AZUL_CLARO), ("TIR", tir, ROJO)):
            fig.add_vline(x=r * 100, line_dash="dash", line_color=col, line_width=3, annotation_text=f"{nm_} {r * 100:.1f}%", annotation_font=dict(size=12, color=NEGRO))
        fig.update_layout(title="El VAN llega a cero en la TIR", xaxis_title="Tasa de descuento (%)", yaxis_title="VAN (S/)", showlegend=False)
        hechos = [("TIR", f"{tir * 100:.2f}%", None), ("CPPC", f"{cppc * 100:.2f}%", "Tasa de corte"), ("COK", f"{cok * 100:.2f}%", "Capital propio"), ("TEA del préstamo", f"{F['financiamiento']['tea'] * 100:.2f}%", None)]
        det = pd.DataFrame([["TIR", f"{tir * 100:.2f}%"], ["CPPC (tasa de corte)", f"{cppc * 100:.2f}%"], ["COK", f"{cok * 100:.2f}%"], ["TEA del préstamo", f"{F['financiamiento']['tea'] * 100:.2f}%"], ["TIR − CPPC", f"{(tir - cppc) * 100:.2f} p.p."]], columns=["Tasa", "Valor"])
    elif sel == "B/C":
        valor, veredicto = f"{bcr:.2f}", chip_sem("verde" if bcr > 1 else "rojo", "Rentable" if bcr > 1 else "No rentable")
        frase, formula = f"Cada S/ 1.00 invertido devuelve S/ {bcr:.2f} en valor presente.", "B/C = valor presente de los flujos ÷ inversión"
        fig = go.Figure(go.Bar(x=["Inversión (año 0)", "Valor presente de los flujos"], y=[F["van"]["inversion"], F["van"]["suma_vp"]], marker=dict(color=[ROJO, NAVY]), width=.45,
                               text=[soles(F["van"]["inversion"]), soles(F["van"]["suma_vp"])], textposition="outside", textfont=dict(size=15, color=NEGRO)))
        fig.update_layout(title="Beneficio frente a costo (S/)", yaxis_title="S/", showlegend=False)
        i_ = selector_anio(sel, list(range(1, len(fl))), anios, 1)
        hechos = [("Valor presente del año", soles(vp[i_]), f"Año {i_} · {anios[i_]}"), ("Peso en el beneficio total", f"{vp[i_] / F['van']['suma_vp'] * 100:.1f}%", None), ("Beneficio total", soles(F["van"]["suma_vp"]), f"Inversión {soles(F['van']['inversion'])}")]
        det = pd.DataFrame([[i, anios[i], f"{vp[i]:,.0f}"] for i in range(1, len(vp))], columns=["Año", "Calendario", "Valor presente del flujo (S/)"])
    elif sel == "Recuperación":
        valor, veredicto = F["kpi"]["recuperacion_txt"], chip_neutro(f"Simple: {rec['simple_txt']}", ROJO)
        frase, formula = f"Descontada: {rec['desc_anios']:.2f} años ({rec['desc_txt']}). Simple: {rec['simple_anios']:.2f} años ({rec['simple_txt']}).", "Año en que el flujo acumulado deja de ser negativo"
        lab = [f"Año {i}" for i in range(len(fl))]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=lab, y=rec["acum_simple"], mode="lines+markers", name="Sin descontar", line=dict(color=NAVY, width=3.5), marker=dict(size=9), fill="tozeroy", fillcolor="rgba(11,31,75,.08)"))
        fig.add_trace(go.Scatter(x=lab, y=rec["acum_desc"], mode="lines+markers", name="Descontado al CPPC", line=dict(color=ROJO, width=3.5, dash="dot"), marker=dict(size=9)))
        fig.add_hline(y=0, line_color=NEGRO, line_width=2)
        fig.update_layout(title="Flujo acumulado frente a la inversión (S/)", yaxis_title="S/")
        i_ = selector_anio(sel, list(range(len(fl))), anios, 1)
        ok_ = rec["acum_simple"][i_] >= 0
        hechos = [("Acumulado sin descontar", soles(rec["acum_simple"][i_]), f"Año {i_} · {anios[i_]}"), ("Acumulado descontado", soles(rec["acum_desc"][i_]), None), ("Estado", "Recuperada" if ok_ else "Pendiente", "Inversión cubierta" if ok_ else "Aún por recuperar")]
        det = pd.DataFrame([[i, anios[i], f"{rec['acum_simple'][i]:,.0f}", f"{rec['acum_desc'][i]:,.0f}"] for i in range(len(fl))], columns=["Año", "Calendario", "Flujo acumulado (S/)", "VP acumulado (S/)"])
    elif sel == "Inversión":
        inv = F["inversion_por_causa"]
        tot_sel = sum(inversion_de(c["codigo"])["inversion"] for c in cs)
        valor, veredicto = soles(tot_sel if S.sel else F["inversion_total"]), chip_neutro("Selección actual" if S.sel else "Total con imprevistos", NAVY)
        frase, formula = f"Financiada {F['financiamiento']['part_prestamo'] * 100:.0f}% con préstamo y {F['financiamiento']['part_aporte'] * 100:.0f}% con aporte propio.", None
        fig = fig_donut([etq_causa(c) for c in cs], [inversion_de(c["codigo"])["inversion"] for c in cs], [META[c["codigo"]]["color"] for c in cs], soles(tot_sel), "Cómo se reparte la inversión", "sin imprevistos")
        hechos = [("Subtotal", soles(F["inversion_subtotal"]), None), ("Imprevistos", soles(F["inversion_imprevistos"]), None), ("Préstamo", soles(F["financiamiento"]["prestamo"]), f"TEA {F['financiamiento']['tea'] * 100:.2f}%"),
                  ("Aporte propio", soles(F["financiamiento"]["aporte"]), f"{F['financiamiento']['part_aporte'] * 100:.0f}% de la inversión")]
        items = [x for x in F["inversion_items"] if x["propuesta"] in codigos]
        det = pd.DataFrame([[x["propuesta"], x["concepto"], x["cantidad"], f"{x['unitario']:,.0f}", f"{x['total']:,.0f}", x["clase"]] for x in items], columns=["Propuesta", "Concepto", "Cant.", "Unit. (S/)", "Total (S/)", "Clase"])
    else:  # Flujo de caja
        valor, veredicto = soles(sum(fl[1:])), chip_neutro("Acumulado 10 años")
        frase, formula = "Suma de los flujos de caja de los años 1 a 10.", None
        lab = [f"Año {i}" for i in range(1, 11)]
        fig = go.Figure()
        fig.add_trace(go.Bar(x=lab, y=d["ingresos"], name="Ingresos", marker_color=AZUL_CLARO))
        fig.add_trace(go.Bar(x=lab, y=[-g for g in d["gastos"]], name="Gastos de operación", marker_color=ROJO_OSC))
        fig.add_trace(go.Scatter(x=lab, y=fl[1:], name="Flujo de caja", mode="lines+markers", line=dict(color=NAVY, width=3.5), marker=dict(size=9)))
        fig.update_layout(title="Ingresos, gastos y flujo de caja por año (S/)", barmode="relative", yaxis_title="S/")
        i_ = selector_anio(sel, list(range(1, 11)), anios, 1)
        hechos = [("Ingresos", soles(d["ingresos"][i_ - 1]), f"Año {i_} · {anios[i_]}"), ("Gastos de operación", soles(d["gastos"][i_ - 1]), None), ("Utilidad neta", soles(d["utilidad_neta"][i_ - 1]), None), ("Flujo de caja", soles(fl[i_]), None)]
        nombres_ = [("Ingresos", "ingresos"), ("Gastos de operación y mantenimiento", "gastos"), ("Depreciación y amortización", "depreciacion"), ("Interés pagado", "interes"), ("Utilidad neta", "utilidad_neta"),
                    ("(+) Depreciaciones", "mas_dep"), ("(−) Pago principal de la deuda", "pago_principal"), ("Impuesto a la renta", "impuesto"), ("(−) Reinversión", "reinversion"), ("(+) Valor residual", "residual"), ("Flujo de caja", "flujo")]
        det = pd.DataFrame([[n] + [f"{v:,.0f}" for v in d[k_]] for n, k_ in nombres_], columns=["Concepto"] + [f"Año {i}" for i in range(1, 11)])

    c_ac = COLOR_FIN[sel]
    g1, g2 = st.columns([1, 2.1], gap="large")
    with g1:
        md(f'<div class="ff" style="--ac:{c_ac};"><div class="ff-l">{sel}</div><div class="ff-v">{valor}</div><div class="ff-b">{veredicto}</div><div class="ff-t">{frase}</div>'
           + (f'<div class="ff-f">{formula}</div>' if formula else "") + "</div>")
    with g2:
        grafico(estilo(fig, 400), f"fin_fig_{sel}")

    # ------------------------------------------------------------------ detalle compacto
    if hechos:
        subtitulo("Detalle")
        cols = st.columns(len(hechos), gap="large")
        for col, (t_, v_, s_) in zip(cols, hechos):
            with col:
                stat(t_, v_, s_ or "&nbsp;", c_ac)
    with st.expander("Ver la tabla completa"):
        tabla(det, alto=420, num=tuple(c for c in det.columns if c not in ("Componente", "Concepto", "Propuesta", "Tasa", "Calendario", "Fuente")))


# ======================= ENRUTADOR =======================
{"panel": pagina_panel, "central": pagina_central, "semaforo": pagina_semaforo, "flujo": pagina_flujo}[S.pagina]()
