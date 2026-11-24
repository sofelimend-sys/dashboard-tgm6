"""
Extrae del libro T3_-_Equipo_-_06.xlsm TODOS los datos que usa el dashboard y los guarda en datos_tgm.json.
No calcula ni asume nada: lee los valores guardados en las celdas (data_only=True) y registra de qué hoja/celda salen.

Uso:   python extraer_datos_excel.py  [ruta_al_xlsm]  [ruta_salida_json]
"""
import json
import re
import sys
import warnings

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

warnings.filterwarnings("ignore")
SRC = sys.argv[1] if len(sys.argv) > 1 else "T3_-_Equipo_-_06.xlsm"
OUT = sys.argv[2] if len(sys.argv) > 2 else "datos_tgm.json"

wb = load_workbook(SRC, data_only=True, read_only=True)
_cache = {}


def hoja(nombre):
    """Carga una hoja completa en un dict {coordenada: valor} (solo celdas con contenido)."""
    if nombre not in _cache:
        d = {}
        for row in wb[nombre].iter_rows():
            for c in row:
                if c.value is not None and str(c.value).strip() != "":
                    d[c.coordinate] = c.value
        _cache[nombre] = d
    return _cache[nombre]


def v(h, ref, default=None):
    return hoja(h).get(ref, default)


def fila(h, fila_, c1, c2):
    """Valores de una fila entre dos columnas (letras)."""
    return [v(h, f"{L(i)}{fila_}") for i in range(CI(c1), CI(c2) + 1)]


def limpia(t):
    return re.sub(r"\n+", " ", str(t)).strip() if t is not None else None


D = {"fuente": SRC.split("/")[-1]}

# ============================== PLANTEAMIENTO ==============================
P = "Planteamiento"
D["problema"] = dict(
    titulo=v(P, "H2"), indicador=v(P, "R2"), objetivo_general=v(P, "V2"),
    brecha_tecnica=v(P, "K8"), vip=v(P, "J10"), vr=v(P, "J13"),
    impacto_economico=v(P, "J16"), pct_ingresos=v(P, "J19"), ingresos_anuales=v(P, "J21"),
    as_is=v(P, "Z4"), to_be=v(P, "AA4"), estandar=v(P, "AB4"), van_planteamiento=v(P, "AD4"),
    ref=dict(brecha="Planteamiento!K8", vip="Planteamiento!J10", vr="Planteamiento!J13", impacto="Planteamiento!J16",
             pct="Planteamiento!J19", ingresos="Planteamiento!J21"))

# Cada objetivo ocupa un bloque de 4 filas: 10, 14, 18, 22, 26 (orden de Planteamiento = orden de los motivos)
CODIGOS = ["ADDIE", "TPM", "SMED", "ISOS", "DMAIC"]
FILAS_P = [10, 14, 18, 22, 26]
causas = []
for i, (cod, r) in enumerate(zip(CODIGOS, FILAS_P), start=1):
    causas.append(dict(
        id=f"C{i}", codigo=cod, motivo=limpia(v(P, f"N{r}")), causa_raiz=limpia(v(P, f"P{r}")),
        indicador_formula=limpia(v(P, f"R{r}")), herramienta=limpia(v(P, f"T{r}")),
        objetivo=limpia(v(P, f"V{r}")), articulo=limpia(v(P, f"X{r}")),
        as_is_plan=v(P, f"Z{r}"), to_be_plan=v(P, f"AA{r}"), estandar_plan=v(P, f"AB{r}"),
        ref=f"Planteamiento!N{r}:AB{r}"))

# ============================== BASE DE DATOS - DASHBOARD ==============================
B = "BASE DE DATOS - DASHBOARD"
# 2. Parámetros por causa raíz (filas 10-14, columnas AG..AX)
for i, r in enumerate(range(10, 15)):
    c = causas[i]
    assert v(B, f"AG{r}") == c["codigo"], (v(B, f"AG{r}"), c["codigo"])
    c.update(indicador=v(B, f"AK{r}"), unidad=v(B, f"AL{r}"), sentido=v(B, f"AM{r}"),
             tipo=v(B, f"AJ{r}"),
             as_is=v(B, f"AN{r}"), to_be=v(B, f"AO{r}"), estandar=v(B, f"AP{r}"),
             prom_mensual_asis=v(B, f"AQ{r}"), factor_mensual=v(B, f"AR{r}"), mejora=v(B, f"AS{r}"),
             inversion=v(B, f"AT{r}"), costo_asis=v(B, f"AU{r}"), costo_tobe=v(B, f"AV{r}"),
             ahorro=v(B, f"AW{r}"), pct_ahorro=v(B, f"AX{r}"))
D["causas"] = causas
D["nota_to_be_mensual"] = v(B, "AG16")
D["totales_base"] = dict(inversion=v(B, "AT15"), costo_asis=v(B, "AU15"), costo_tobe=v(B, "AV15"), ahorro=v(B, "AW15"))

# 1. Tabla mensual (filas 10-189): N=ID ... AE=Logro de la meta
mensual = {c["codigo"]: {"AS-IS": [], "TO-BE": [], "ESTÁNDAR": [], "cumple_tobe": [], "costo": {"AS-IS": [], "TO-BE": []},
                         "ahorro_tobe": [], "logro": {"AS-IS": [], "TO-BE": []}} for c in causas}
for r in range(10, 190):
    cod, esc = v(B, f"S{r}"), v(B, f"Y{r}")
    if cod not in mensual:
        continue
    m = mensual[cod]
    m[esc].append(v(B, f"Z{r}"))
    if esc == "TO-BE":
        m["cumple_tobe"].append(v(B, f"AB{r}"))
        m["ahorro_tobe"].append(v(B, f"AD{r}"))
    if esc in ("AS-IS", "TO-BE"):
        m["costo"][esc].append(v(B, f"AC{r}"))
        m["logro"][esc].append(v(B, f"AE{r}"))
for cod, m in mensual.items():
    assert len(m["AS-IS"]) == len(m["TO-BE"]) == len(m["ESTÁNDAR"]) == 12, cod
D["mensual_base"] = mensual
D["meses"] = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Set", "Oct", "Nov", "Dic"]

# ============================== SEMAFORIZACIÓN ==============================
S = "SEMAFORIZACIÓN DE INDICADORES"
sem = {}
for r, cod in zip(range(75, 81), ["P. CENTRAL"] + CODIGOS):
    sem[cod] = dict(
        indicador=v(S, f"M{r}"), tipo=v(S, f"AA{r}"), unidad=v(S, f"AF{r}"), sentido=v(S, f"AJ{r}"),
        linea_base=v(S, f"AO{r}"), meta=v(S, f"AS{r}"), escala_min=v(S, f"AW{r}"), escala_max=v(S, f"BA{r}"),
        rojo=v(S, f"BE{r}"), amarillo=v(S, f"BJ{r}"), verde=v(S, f"BO{r}"))
cols_mes = [L(CI("AA") + 3 * k) for k in range(12)]  # AA, AD, AG, ... BH
filas_sem = {"P. CENTRAL": (83, 84)}
for k, cod in enumerate(CODIGOS):
    filas_sem[cod] = (85 + 2 * k, 86 + 2 * k)
for cod, (ra, rt) in filas_sem.items():
    sem[cod]["mensual_asis"] = [v(S, f"{c}{ra}") for c in cols_mes]
    sem[cod]["mensual_tobe"] = [v(S, f"{c}{rt}") for c in cols_mes]
    sem[cod]["anual_asis"] = v(S, f"BK{ra}")
    sem[cod]["anual_tobe"] = v(S, f"BK{rt}")
for r, cod in zip(range(101, 107), ["P. CENTRAL"] + CODIGOS):
    sem[cod]["estado"] = v(S, f"V{r}")
    sem[cod]["variacion_txt"] = v(S, f"AK{r}")
D["semaforo"] = dict(indicadores=sem, n_cumple=v(S, "AH7"), n_proceso=v(S, "AX7"), n_critico=v(S, "BN7"),
                     ref="SEMAFORIZACIÓN DE INDICADORES!J74:BO106")

# ============================== CUMPLIMIENTO OT POR QUINCENA (Análisis Cuantitativo) ==============================
A = "Análisis Cuantitativo"
quin = []
for r in range(216, 242):
    quin.append(dict(periodo=v(A, f"L{r}"), programadas=v(A, f"M{r}"), ejecutadas=v(A, f"N{r}"), cumplimiento=v(A, f"O{r}"),
                     costo_tracto=v(A, f"P{r}"), costo_tolva=v(A, f"Q{r}"), costo_total=v(A, f"R{r}")))
D["quincenas"] = quin
D["quincenas_total"] = dict(programadas=v(A, "M242"), ejecutadas=v(A, "N242"), cumplimiento=v(A, "O242"),
                            costo_tracto=v(A, "P242"), costo_tolva=v(A, "Q242"), costo_total=v(A, "R242"))
D["quincenas_ref"] = "Análisis Cuantitativo!K216:R242 (= DESARROLLO DEL PROBLEMA CENTRAL!H57:K83)"

# ============================== FLUJO CAJA ==============================
F = "Flujo Caja"
fc = {}
fc["kpi"] = dict(van=v(F, "H8"), tir=v(F, "J8"), cppc=v(F, "L8"), bc=v(F, "N8"), recuperacion_txt=v(F, "P8"), inversion=v(F, "R8"),
                 txt_van=v(F, "H9"), txt_tir=v(F, "J9"), txt_cppc=v(F, "L9"), txt_bc=v(F, "N9"), txt_rec=v(F, "P9"), txt_inv=v(F, "R9"))
fc["parametros"] = []
for r in range(14, 50):
    if v(F, f"I{r}") is not None:
        fc["parametros"].append(dict(concepto=v(F, f"H{r}"), valor=v(F, f"I{r}"), unidad=v(F, f"J{r}"), fuente=v(F, f"K{r}"), ref=f"Flujo Caja!H{r}"))
fc["anios"] = [v(F, f"{L(CI('I') + k)}52") for k in range(11)]  # 2026..2036
fc["inversion_items"] = [dict(concepto=v(F, f"H{r}"), propuesta=v(F, f"I{r}"), cantidad=v(F, f"J{r}"), unitario=v(F, f"K{r}"),
                              total=v(F, f"L{r}"), clase=v(F, f"M{r}")) for r in range(57, 87)]
fc["inversion_subtotal"], fc["inversion_imprevistos"], fc["inversion_total"] = v(F, "L87"), v(F, "L88"), v(F, "L89")
fc["inversion_por_causa"] = [dict(causa=v(F, f"H{r}"), propuesta=v(F, f"I{r}"), inversion=v(F, f"J{r}"), pct=v(F, f"K{r}"),
                                  activo_fijo=v(F, f"L{r}"), intangible=v(F, f"M{r}")) for r in range(92, 97)]
fc["eval_causas"] = [dict(causa=v(F, f"H{r}"), propuesta=v(F, f"I{r}"), costo_asis=v(F, f"J{r}"), costo_tobe=v(F, f"K{r}"),
                          beneficio=v(F, f"L{r}"), pct=v(F, f"M{r}")) for r in range(208, 213)]
fc["eval_total"] = dict(costo_asis=v(F, "J213"), costo_tobe=v(F, "K213"), beneficio=v(F, "L213"))
meses_cols = [L(CI("I") + k) for k in range(12)]  # I..T
fc["costos_mensuales"] = {
    "TPM": dict(asis=[v(F, f"{c}135") for c in meses_cols], tobe=[v(F, f"{c}137") for c in meses_cols], ahorro=[v(F, f"{c}138") for c in meses_cols],
                detalle=[dict(concepto=v(F, f"H{r}"), valores=[v(F, f"{c}{r}") for c in meses_cols]) for r in (132, 133, 134)]),
    "SMED": dict(asis=[v(F, f"{c}155") for c in meses_cols], tobe=[v(F, f"{c}157") for c in meses_cols], ahorro=[v(F, f"{c}158") for c in meses_cols]),
    "ISOS": dict(asis=[v(F, f"{c}175") for c in meses_cols], tobe=[v(F, f"{c}177") for c in meses_cols], ahorro=[v(F, f"{c}178") for c in meses_cols]),
    "DMAIC": dict(asis=[v(F, f"{c}195") for c in meses_cols], tobe=[v(F, f"{c}197") for c in meses_cols], ahorro=[v(F, f"{c}198") for c in meses_cols]),
}
fc["asistencia_addie"] = dict(convocados=[v(F, f"{c}113") for c in meses_cols], asistieron=[v(F, f"{c}114") for c in meses_cols],
                              cas_asis=[v(F, f"{c}115") for c in meses_cols], cas_tobe=[v(F, f"{c}117") for c in meses_cols],
                              recuperadas=[v(F, f"{c}118") for c in meses_cols], total_recuperadas=v(F, "U118"),
                              beneficio_directo=v(F, "I122"))
fc["indicadores_anuales"] = dict(fallas=(v(F, "I128"), v(F, "J128")), setup_h=(v(F, "I149"), v(F, "J149")),
                                 tiempo_muerto_h=(v(F, "I169"), v(F, "J169")), equipos_2do_mant=(v(F, "I189"), v(F, "J189")))
fc["ot"] = dict(precio_ot=v(F, "I225"), costo_ot=v(F, "I223"), margen=v(F, "I224"),
                meses=[dict(mes=v(F, f"H{r}"), programadas=v(F, f"I{r}"), asis=v(F, f"J{r}"), tobe=v(F, f"K{r}"), adicionales=v(F, f"L{r}"),
                            ingreso_asis=v(F, f"M{r}"), ingreso_tobe=v(F, f"N{r}"), incremental=v(F, f"O{r}")) for r in range(228, 240)],
                total=dict(programadas=v(F, "I240"), asis=v(F, "J240"), tobe=v(F, "K240"), adicionales=v(F, "L240"),
                           ingreso_asis=v(F, "M240"), ingreso_tobe=v(F, "N240"), incremental=v(F, "O240")),
                ingreso_incremental=v(F, "I242"), costo_adicional=v(F, "I243"), margen_incremental=v(F, "I244"))
fc["sostenimiento"] = [dict(concepto=v(F, f"H{r}"), propuesta=v(F, f"I{r}"), cantidad=v(F, f"J{r}"), unitario=v(F, f"K{r}"),
                            monto=v(F, f"L{r}"), desde=v(F, f"M{r}")) for r in range(255, 264)]
fc["sostenimiento_total"] = v(F, "L264")
cols_anio = [L(CI("J") + k) for k in range(10)]  # J..S
fc["depreciacion"] = [dict(concepto=v(F, f"H{r}"), base=v(F, f"I{r}"), anuales=[v(F, f"{c}{r}") for c in cols_anio], libros=v(F, f"T{r}"))
                      for r in range(271, 276)]
fc["reinversion"] = [v(F, f"{c}276") for c in cols_anio]
fc["valor_residual"] = [v(F, f"{c}277") for c in cols_anio]
fc["financiamiento"] = dict(prestamo=v(F, "J284"), aporte=v(F, "J285"), part_prestamo=v(F, "I284"), part_aporte=v(F, "I285"),
                            tea=v(F, "I286"), plazo=v(F, "I287"), cuota=v(F, "I288"),
                            cronograma=[dict(anio=v(F, f"H{r}"), saldo_ini=v(F, f"I{r}"), interes=v(F, f"J{r}"), amortizacion=v(F, f"K{r}"),
                                             cuota=v(F, f"L{r}"), saldo_fin=v(F, f"M{r}")) for r in range(291, 302)],
                            total=dict(interes=v(F, "J302"), amortizacion=v(F, "K302"), cuota=v(F, "L302")))
fc["cok"] = dict(componentes=[dict(concepto=v(F, f"H{r}"), valor=v(F, f"I{r}")) for r in range(310, 318)],
                 rf=v(F, "I310"), beta_u=v(F, "I311"), de=v(F, "I312"), beta_l=v(F, "I313"), prima_mercado=v(F, "I314"),
                 riesgo_pais=v(F, "I315"), prima_tamano=v(F, "I316"), cok=v(F, "I317"),
                 aportes=[dict(concepto=v(F, f"H{r}"), valor=v(F, f"I{r}")) for r in range(320, 324)])
fc["cppc"] = dict(valor=v(F, "I358"), pasivos=dict(monto=v(F, "K359"), part=v(F, "L359"), costo=v(F, "M359"), pond=v(F, "N359")),
                  patrimonio=dict(monto=v(F, "K360"), part=v(F, "L360"), costo=v(F, "M360"), pond=v(F, "N360")),
                  total=dict(monto=v(F, "K361"), part=v(F, "L361"), pond=v(F, "N361")))
fc["detalle_flujo"] = {k: [v(F, f"{c}{r}") for c in cols_anio] for k, r in dict(
    adopcion=334, ingresos_ot=335, ahorros=336, total_ingresos=337, costo_ot=338, sostenimiento=339, total_gastos=340,
    ingresos=345, gastos=346, depreciacion=347, interes=348, utilidad_neta=349, mas_dep=350, pago_principal=351,
    impuesto=352, reinversion=353, residual=354, flujo=355, valor_presente=369).items()}
fc["detalle_flujo"]["flujo_anio0"] = v(F, "I368")
fc["van"] = dict(valor=v(F, "I371"), manual=v(F, "I374"), suma_vp=v(F, "T369"), inversion=v(F, "I329"), financiado=v(F, "I330"),
                 txt=v(F, "J371"), txt_manual=v(F, "J374"))
fc["tir"] = dict(valor=v(F, "I377"), txt=v(F, "J377"))
fc["bc"] = dict(valor=v(F, "I379"), txt=v(F, "J379"))
cols_a0 = [L(CI("I") + k) for k in range(11)]  # I..S (año 0..10)
fc["recuperacion"] = dict(
    acum_simple=[v(F, f"{c}384") for c in cols_a0], acum_desc=[v(F, f"{c}385") for c in cols_a0],
    simple_anios=v(F, "I390"), simple_txt=v(F, "J390"), desc_anios=v(F, "I391"), desc_txt=v(F, "J391"))
fc["ref"] = "Flujo Caja (hoja completa; valores guardados en las celdas del Excel)"
D["flujo"] = fc

# ============================== CONCLUSIONES (interpretación por objetivo) ==============================
C = "Conclusiones"
D["conclusiones"] = dict(
    general=dict(texto=v(C, "I11"), as_is=v(C, "M10"), to_be=v(C, "N10"), estandar=v(C, "O10"), resultado=v(C, "P10"), variacion=v(C, "T10")),
    especificas={cod: dict(texto=v(C, f"I{r + 1}"), as_is=v(C, f"M{r}"), to_be=v(C, f"N{r}"), estandar=v(C, f"O{r}"),
                           resultado=v(C, f"P{r}"), variacion=v(C, f"T{r}"))
                 for cod, r in zip(CODIGOS, (39, 79, 59, 99, 119))})
# Orden en Conclusiones: 1 ADDIE (39), 2 TPM (59), 3 SMED (79), 4 ISOS (99), 5 DMAIC (119)
D["conclusiones"]["especificas"] = {cod: dict(texto=v(C, f"I{r + 1}"), as_is=v(C, f"M{r}"), to_be=v(C, f"N{r}"), estandar=v(C, f"O{r}"),
                                              resultado=v(C, f"P{r}"), variacion=v(C, f"T{r}"))
                                    for cod, r in zip(CODIGOS, (39, 59, 79, 99, 119))}

# ============================== PROBLEMAS (Tabla 1 / 2) ==============================
Pb = "5 PROBLEMAS"
D["problema_matriz"] = dict(problema=v(Pb, "M11"), observable=v(Pb, "L24"), definible=v(Pb, "M24"), medible=v(Pb, "N24"),
                            registrable=v(Pb, "O24"), impactante=v(Pb, "P24"), tratable=v(Pb, "Q24"), controlable=v(Pb, "R24"))

# ============================== METODOLOGÍAS (hojas Desarrollo de Objetivo 1-5) ==============================
def seccion_qse_hace(h, fila_ini, n=14):
    """Busca 'QUÉ SE HACE' / 'PREGUNTA...' / 'ENTREGABLES...' a partir de una fila de título de etapa."""
    d = hoja(h)
    etiquetas = {}
    for r in range(fila_ini + 1, fila_ini + n):
        for col in range(1, 130):
            val = d.get(f"{L(col)}{r}")
            if isinstance(val, str) and val.strip() in ("QUÉ SE HACE",) or (isinstance(val, str) and val.startswith(("PREGUNTA QUE RESPONDE", "ENTREGABLES DE LA ETAPA"))):
                etiquetas[val.strip()] = (r, col)
    res = {}
    for etq, (r, col) in etiquetas.items():
        for rr in range(r + 1, r + 6):
            t = d.get(f"{L(col)}{rr}")
            if t:
                res[etq] = limpia(t)
                break
    return dict(pregunta=res.get("PREGUNTA QUE RESPONDE ESTA ETAPA"), que_se_hace=res.get("QUÉ SE HACE"), entregables=res.get("ENTREGABLES DE LA ETAPA"))


met = {}
# ADDIE (Obj 1): 5 etapas
O1 = "Desarrollo de Objetivo 1"
met["ADDIE"] = dict(hoja=O1, etapas=[dict(nombre=limpia(v(O1, f"I{r}")), **seccion_qse_hace(O1, r)) for r in (188, 270, 421, 576, 695)])
# TPM (Obj 2): 6 pasos con fechas y responsables
O2 = "Desarrollo de Objetivo 2"
met["TPM"] = dict(hoja=O2, etapas=[dict(nombre=v(O2, f"AS{r}"), inicio=v(O2, f"BE{r}"), fin=v(O2, f"BK{r}"), responsable=v(O2, f"BP{r}")) for r in range(581, 592, 2)],
                  pilares=[dict(pilar=v(O2, f"K{r}"), aplica=v(O2, f"W{r}"), motivo=v(O2, f"AA{r}")) for r in range(192, 207, 2) if v(O2, f"K{r}")],
                  nota=v(O2, "AS593"))
# SMED (Obj 3): 8 etapas
O3 = "Desarrollo de Objetivo 3"
met["SMED"] = dict(hoja=O3, etapas=[dict(nombre=limpia(v(O3, f"K{r}")), descripcion=limpia(v(O3, f"K{r + 4}")))
                                    for r in (261, 372, 489, 554, 691, 781, 853, 923)])
# ISOS (Obj 4): componentes A-I
O4 = "Desarrollo de Objetivo 4"
isos = []
for ref in ("AS179", "K240", "AB240", "K292", "AB292", "K361", "AB361", "K441", "AL441"):
    t = v(O4, ref)
    if t:
        isos.append(dict(nombre=limpia(t)))
met["ISOS"] = dict(hoja=O4, etapas=isos)
# DMAIC (Obj 5): 5 fases
O5 = "Desarrollo de Objetivo 5"
met["DMAIC"] = dict(hoja=O5, etapas=[dict(nombre=v(O5, f"K{r}"), objetivo=v(O5, f"P{r}"), herramientas=v(O5, f"Z{r}"), fecha=v(O5, f"AK{r}"))
                                     for r in range(198, 208, 2)])
# Herramientas de diagnóstico (títulos de la Fase 1 de cada hoja) y de proyección (Fase 3)
CLAVES = {"ÁRBOL DE PROBLEMAS": "Árbol de problemas", "ISHIKAWA": "Diagrama de Ishikawa", "PARETO": "Diagrama de Pareto",
          "DIAGRAMA DE FLUJO DE PROCESO AS-IS": "Diagrama de flujo AS-IS (BPMN)", "CURSOGRAMA": "Cursograma analítico",
          "5 PORQUÉS": "5 porqués", "DIAGRAMA DE FLUJO DEL PROCESO AS-IS": "Diagrama de flujo AS-IS (BPMN)"}
for i, cod in enumerate(CODIGOS, start=1):
    h = f"Desarrollo de Objetivo {i}"
    d = hoja(h)
    enc = set()
    for ref, val in d.items():
        if isinstance(val, str) and len(val) < 120 and val.upper() == val:
            for k, nom in CLAVES.items():
                if k in val:
                    enc.add(nom)
    ruta = [f"{h}!{ref}" for ref, val in d.items() if isinstance(val, str) and "SIMULACIÓN DE RESULTADOS" in val]
    met[cod]["diagnostico"] = sorted(enc)
    met[cod]["simulacion"] = limpia(next((val for val in d.values() if isinstance(val, str) and "SIMULACIÓN DE RESULTADOS" in val), None))
D["metodologias"] = met

# Problema central: relación causa → herramienta → barrera (DESARROLLO DEL PROBLEMA CENTRAL AC56:AL60)
Dp = "DESARROLLO DEL PROBLEMA CENTRAL"
D["barreras"] = [dict(causa=v(Dp, f"AC{r}"), herramienta=v(Dp, f"AG{r}"), barrera=v(Dp, f"AL{r}")) for r in range(56, 61)]
D["problema_central_texto"] = dict(titulo=v(Dp, "G2"), objetivo=v(Dp, "L9"), indicador_def=limpia(v(Dp, "R43")), valor_ref=limpia(v(Dp, "R53")),
                                   relacion=limpia(v(Dp, "AC39")))

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(D, f, ensure_ascii=False, indent=1, default=str)
print("OK ->", OUT)
