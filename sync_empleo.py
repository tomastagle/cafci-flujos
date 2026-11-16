# -*- coding: utf-8 -*-
"""
sync_empleo.py — datos de la solapa "Empleo y Social" del Monitor Argentina (Argento, bloques/EMP).

Tres fuentes públicas, ninguna carga manual:
  1. SIPA (Secretaría de Trabajo, "Situación y evolución del trabajo registrado"): Excel mensual
     trabajoregistrado_AAMM_estadisticas.xlsx de argentina.gob.ar. Hojas T.2.1/T.2.2 (personas con trabajo
     registrado por modalidad, con y sin estacionalidad), A.2.1/A.2.2 (asalariados privados por rama) y A.4
     (remuneración mediana y promedio). Respaldo si no se llega al Excel: datos.gob.ar 151.1, 152.1 y 153.1
     (mismas series, con uno o dos meses más de rezago).
  2. INDEC, Cuenta de generación del ingreso e insumo de mano de obra (datos.gob.ar 324.1, trimestral): puestos
     de trabajo totales de la economía por categoría (asalariados registrados, no registrados, no asalariados)
     y por sector.
  3. INDEC, API oficial: lo que ya dejó sync_indec_api.py en macro/indec_api.json (tasas de la EPH, pobreza e
     indigencia, índice de salarios, IPC nivel general y canastas). Este script corre después de aquel.

Deja macro/monitor_emp.json = {"meta": {...}, "patch": {...}}; "patch" tiene la forma exacta de la constante
EMP de la página (Argento lo escribe tal cual en bloques/EMP):
  fetched, src
  q: trimestres "AAAA-Tn";  eph: {act, emp, des, sub, dem, form, inf}  (tasas EPH, 31 aglomerados, alineadas a q)
  s: semestres "AAAA-Sn";   pob: {pp, ip, ph, ih}
  m: meses "AAAA-MM" desde 2012-01; todo lo mensual va alineado a m (null donde no hay dato):
     sipa: {ult, prov, via, archivo, t: {priv,pub,cp,aut,mono,ms,tot: {sa, o}}, r: {A..O,X,T: {sa, o}},
            rem: {med, med_sa, prom, prom_sa}}       (miles de personas; remuneración en $ corrientes)
     sal: {tot, reg, priv, pub, noreg}  (índice de salarios INDEC),  ipc (nivel general, dic-16 = 100),
     can: {cbt_h2, cba_h2}  ($ por mes, hogar 2, GBA)
  cq: trimestres;  cgi: {TOT, PRIV, PUB, A..P: {t, reg, nr, na}}  (miles de puestos)

Uso:  python sync_empleo.py            # produce macro/monitor_emp.json
      python sync_empleo.py --check    # además imprime los últimos datos
      python sync_empleo.py --xlsx F   # usa un Excel de trabajo registrado ya bajado (pruebas)
"""
from __future__ import annotations
import os, io, re, json, time, argparse, warnings, urllib.request, urllib.parse
from datetime import datetime, timezone

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "macro")
os.makedirs(OUT, exist_ok=True)
HDR = {"User-Agent": "Mozilla/5.0 (monitor-argentina-sync)"}
PAGINAS = ["https://www.argentina.gob.ar/trabajo/estadisticas/situacion-y-evolucion-del-trabajo-registrado",
           "https://www.argentina.gob.ar/trabajo/estadisticas"]
SERIES_API = "https://apis.datos.gob.ar/series/api/series/"
DESDE = "2012-01"
LOG, FALLOS = [], []


def log(m):
    LOG.append(str(m)); print(m)


def http(url, tries=4, binario=False):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=120) as r:
                b = r.read()
                return b if binario else b.decode("utf-8", "replace")
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


# ---------------------------------------------------------------- SIPA (Excel de la Secretaría de Trabajo)
TIPOS = ["priv", "pub", "cp", "aut", "mono", "ms", "tot"]            # orden de columnas de T.2.1 / T.2.2
RAMAS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N", "O", "X", "T"]   # orden de A.2.1 / A.2.2
MES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "set": 9, "sept": 9,
       "oct": 10, "nov": 11, "dic": 12}


def periodo(c):
    """(AAAA-MM, provisorio) de la primera celda de una fila, o None."""
    if isinstance(c, datetime):
        return c.strftime("%Y-%m"), False
    if isinstance(c, str):
        m = re.match(r"^\s*([a-záé]{3,4})[\.\-\s/]+(\d{2,4})\s*(\*?)", c.strip().lower())
        if m and m.group(1) in MES:
            y = int(m.group(2)); y += 2000 if y < 100 else 0
            return "%04d-%02d" % (y, MES[m.group(1)]), bool(m.group(3))
    return None


def hoja(wb, nombre, ncols):
    """{mes: [valores]} y lista de meses provisorios de una hoja con el período en la columna A."""
    out, prov = {}, []
    for r in wb[nombre].iter_rows(max_col=ncols + 1, values_only=True):
        p = periodo(r[0]) if r else None
        if not p:
            continue
        v = list(r[1:ncols + 1])
        if not all(isinstance(x, (int, float)) for x in v):
            continue
        out[p[0]] = [float(x) for x in v]
        if p[1]:
            prov.append(p[0])
    return out, prov


def sipa_excel(ruta=None):
    import openpyxl
    archivo = os.path.basename(ruta) if ruta else None
    if not ruta:
        url = None
        for pg in PAGINAS:
            try:
                h = http(pg)
            except Exception as e:
                log("SIPA: no abre %s (%r)" % (pg, e)); continue
            fs = sorted(set(re.findall(r"https?://[^\"'#\s]*?/trabajoregistrado_(\d{4})_estadisticas\.xlsx", h)))
            if fs:
                archivo = "trabajoregistrado_%s_estadisticas.xlsx" % fs[-1]
                url = "https://www.argentina.gob.ar/sites/default/files/" + archivo
                break
        if not url:
            raise RuntimeError("no se encontró el link al Excel de trabajo registrado")
        ruta = os.path.join(OUT, "_trabajoregistrado.xlsx")
        io.open(ruta, "wb").write(http(url, binario=True))
        log("SIPA: bajado %s (%.1f MB)" % (archivo, os.path.getsize(ruta) / 1e6))
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    t_o, prov = hoja(wb, "T.2.1", 7); t_sa, _ = hoja(wb, "T.2.2", 7)
    r_o, _ = hoja(wb, "A.2.1", 16); r_sa, _ = hoja(wb, "A.2.2", 16)
    rem, _ = hoja(wb, "A.4", 4)
    wb.close()
    if os.path.basename(ruta) == "_trabajoregistrado.xlsx":
        os.remove(ruta)
    # controles: el total de cada hoja es la suma de sus columnas
    for nom, d, n in (("T.2.1", t_o, 6), ("T.2.2", t_sa, 6), ("A.2.1", r_o, 15)):
        mal = [k for k, v in d.items() if abs(sum(v[:n]) - v[n]) > 0.6]
        if mal:
            raise RuntimeError("%s: el total no cierra con las columnas en %s" % (nom, mal[:3]))
    return {"via": "Secretaría de Trabajo, " + archivo, "archivo": archivo, "prov": sorted(prov),
            "t": {"o": t_o, "sa": t_sa, "cols": TIPOS}, "r": {"o": r_o, "sa": r_sa, "cols": RAMAS},
            "rem": {k: {"med": v[0], "med_sa": v[1], "prom": v[2], "prom_sa": v[3]} for k, v in rem.items()}}


# respaldo: las mismas series en datos.gob.ar (sspm 151.1, 152.1, 153.1)
DG_T = {"priv": ("151.1_AARIADOTAC_2012_M_26", "151.1_AARIADODAD_2012_M_31"), "pub": ("151.1_AARIADOTAC_2012_M_25", "151.1_AARIADODAD_2012_M_30"),
        "cp": ("151.1_AARIADOTAC_2012_M_40", "151.1_AARIADODAD_2012_M_45"), "aut": ("151.1_IPENDIETAC_2012_M_34", "151.1_IPENDIEDAD_2012_M_39"),
        "mono": ("151.1_IPENDIETAC_2012_M_36", "151.1_IPENDIEDAD_2012_M_41"), "ms": ("151.1_IPENDIETAC_2012_M_43", "151.1_IPENDIEDAD_2012_M_48"),
        "tot": ("151.1_TL_SIN_TAC_2012_M_15", "151.1_TL_ESTADAD_2012_M_20")}
DG_R = {"A": ("152.1_ACULTUREST_2009_M_47", "152.1_ACULTUREST_2009_M_43"), "B": ("152.1_PA_SIN_EST_2009_M_13", "152.1_PA_ESTEST_2009_M_9"),
        "C": ("152.1_EOTACIOEST_2009_M_34", "152.1_EOTACIOEST_2009_M_30"), "D": ("152.1_ISTRIASEST_2009_M_33", "152.1_ISTRIASEST_2009_M_29"),
        "E": ("152.1_SNISTROEST_2009_M_40", "152.1_SNISTROEST_2009_M_36"), "F": ("152.1_CTRUCCIEST_2009_M_20", "152.1_CTRUCCIEST_2009_M_16"),
        "G": ("152.1_CRCIO_REST_2009_M_29", "152.1_CRCIO_REST_2009_M_25"), "H": ("152.1_HLES_REEST_2009_M_28", "152.1_HLES_REEST_2009_M_24"),
        "I": ("152.1_TSPORTEEST_2009_M_39", "152.1_TSPORTEEST_2009_M_35"), "J": ("152.1_IRMEDIAEST_2009_M_33", "152.1_IRMEDIAEST_2009_M_29"),
        "K": ("152.1_AVIDADEEST_2009_M_56", "152.1_AVIDADEEST_2009_M_52"), "M": ("152.1_ENANZA_EST_2009_M_17", "152.1_ENANZA_EST_2009_M_13"),
        "N": ("152.1_SICIOS_EST_2009_M_32", "152.1_SICIOS_EST_2009_M_28"), "O": ("152.1_SICIOS_EST_2009_M_50", "152.1_SICIOS_EST_2009_M_46"),
        "X": ("152.1_SESPECIEST_2009_M_23", "152.1_SESPECIEST_2009_M_19"), "T": ("152.1_TL_SIN_EST_2009_M_13", "152.1_TL_ESTEST_2009_M_9")}
DG_REM = {"med": "153.1_RNERACIANA_2009_M_20", "prom": "153.1_RNERACIDIO_2009_M_21"}


def datosgob(ids, desde):
    """{id: {período: valor}} con la API de series de datos.gob.ar (de a 20 ids)."""
    out = {i: {} for i in ids}
    for k in range(0, len(ids), 20):
        sub = ids[k:k + 20]
        j = json.loads(http(SERIES_API + "?" + urllib.parse.urlencode({"ids": ",".join(sub), "start_date": desde, "limit": 5000, "format": "json"})))
        for row in j.get("data", []):
            for i, v in zip(sub, row[1:]):
                if v is not None:
                    out[i][row[0]] = float(v)
        time.sleep(0.4)
    return out


def sipa_datosgob():
    ids = [x for par in list(DG_T.values()) + list(DG_R.values()) for x in par] + list(DG_REM.values())
    d = datosgob(ids, DESDE + "-01")
    meses = sorted({f[:7] for i in ids for f in d[i]})
    def arma(mapa, cols):
        o, sa = {}, {}
        for m in meses:
            f = m + "-01"
            vs = [d[mapa[c][0]].get(f) for c in cols]; vo = [d[mapa[c][1]].get(f) for c in cols]
            if all(v is not None for v in vs): sa[m] = vs
            if all(v is not None for v in vo): o[m] = vo
        return {"o": o, "sa": sa, "cols": cols}
    rem = {}
    for m in meses:
        a, b = d[DG_REM["med"]].get(m + "-01"), d[DG_REM["prom"]].get(m + "-01")
        if a is not None and b is not None:
            rem[m] = {"med": a, "med_sa": None, "prom": b, "prom_sa": None}
    return {"via": "datos.gob.ar 151.1, 152.1 y 153.1 (respaldo: no se llegó al Excel de Trabajo)", "archivo": None, "prov": [],
            "t": arma(DG_T, TIPOS), "r": arma(DG_R, RAMAS), "rem": rem}


# ---------------------------------------------------------------- INDEC, puestos de trabajo (datos.gob.ar 324.1)
P = "324.1_"
CGI = {
    "TOT": {"t": "TOTAL_GENEAJO__29", "reg": "TOTAL_GENEADO__51", "nr": "TOTAL_GENEREG__32", "na": "TOTAL_GENESAL__34"},
    "PRIV": {"t": "TOTAL_EXCLAJO__45", "reg": "TOTAL_EXCLREG__52", "nr": "TOTAL_EXCLREG__48", "na": "TOTAL_EXCLSAL__50"},
    "PUB": {"t": "TOTAL_SECTAJO__36", "reg": "TOTAL_SECTADO__58"},
    "A": {"t": "A_AGRICULTAJO__49", "reg": "A_AGRICULTREG__58", "nr": "A_AGRICULTREG__52", "na": "A_AGRICULTSAL__54"},
    "B": {"t": "B_PESCA_PUAJO__23", "reg": "B_PESCA_PUADO__45", "nr": "B_PESCA_ASREG__26", "na": "B_PESCA_PUSAL__28"},
    "C": {"t": "C_EXPLOTACAJO__44", "reg": "C_EXPLOTACREG__59", "nr": "C_EXPLOTACREG__47", "na": "C_EXPLOTACSAL__49"},
    "D": {"t": "D_INDUSTRIAJO__41", "reg": "D_INDUSTRIREG__56", "nr": "D_INDUSTRIREG__44", "na": "D_INDUSTRISAL__46"},
    "E": {"t": "E_ELECT_GAAJO__32", "reg": "E_ELECT_GAADO__54", "nr": "E_ELECT_GAREG__35", "na": "E_ELECT_GASAL__37"},
    "F": {"t": "F_CONSTRUCAJO__30", "reg": "F_CONSTRUCADO__52", "nr": "F_CONSTRUCREG__33", "na": "F_CONSTRUCSAL__35"},
    "G": {"t": "G_COMERCIOAJO__26", "reg": "G_COMERCIOADO__48", "nr": "G_COMERCIOREG__29", "na": "G_COMERCIOSAL__31"},
    "H": {"t": "H_HOTELES_AJO__38", "reg": "H_HOTELES_ADO__60", "nr": "H_HOTELES_REG__41", "na": "H_HOTELES_SAL__43"},
    "I": {"t": "I_TRANSPORAJO__58", "reg": "I_TRANSPORREG__58", "nr": "I_TRANS_ALREG__56", "na": "I_TRANS_ALSAL__58"},
    "J": {"t": "J_INTERMEDAJO__43", "reg": "J_INTERMEDADO__65", "nr": "J_INTERMEDREG__46", "na": "J_INTERMEDSAL__48"},
    "K": {"t": "K_ACT_INMOAJO__58", "reg": "K_ACT_INMOADO__57", "nr": "K_ACT_INMOREG__53", "na": "K_ACT_INMOSAL__55"},
    "L": {"t": "L_ADM_PUB_AJO__58", "reg": "L_ADM_PUB_REG__59"},
    "M": {"t": "M_ENSENIANZAJO__27", "reg": "M_ENSENIANZADO__49", "nr": "M_ENSENIANZREG__30", "na": "M_ENSENIANZSAL__32"},
    "N": {"t": "N_SERVICIOAJO__42", "reg": "N_SERVICIOREG__57", "nr": "N_SERVICIOREG__45", "na": "N_SERVICIOSAL__47"},
    "O": {"t": "O_OTRAS_ACAJO__60", "reg": "O_OTRAS_ACREG__51", "nr": "O_OTRAS_ACREG__47", "na": "O_OTRAS_ACSAL__56"},
    "P": {"t": "P_HOGARES_AJO__47", "reg": "P_HOGARES_REG__56", "nr": "P_HOGARES_REG__50"},
}


def cgi():
    ids = [P + v for s in CGI.values() for v in s.values()]
    d = datosgob(ids, "2016-01-01")
    fechas = sorted({f for i in ids for f in d[i]})
    cq = ["%s-T%d" % (f[:4], (int(f[5:7]) - 1) // 3 + 1) for f in fechas]
    out = {}
    for s, cols in CGI.items():
        out[s] = {c: [round(d[P + i][f], 3) if f in d[P + i] else None for f in fechas] for c, i in cols.items()}
        partes = [c for c in ("reg", "nr", "na") if c in cols]
        for k in range(len(fechas)):                    # control del mapeo: total = suma de categorías
            v = [out[s][c][k] for c in partes]
            if out[s]["t"][k] is not None and all(x is not None for x in v) and abs(sum(v) - out[s]["t"][k]) > 1.5:
                raise RuntimeError("CGI %s %s: total %.1f distinto de la suma de categorías %.1f" % (s, cq[k], out[s]["t"][k], sum(v)))
    return cq, out


# ---------------------------------------------------------------- armado del parche
def meses_entre(a, b):
    out, (y, m) = [], (int(a[:4]), int(a[5:]))
    while "%04d-%02d" % (y, m) <= b:
        out.append("%04d-%02d" % (y, m)); m += 1
        if m == 13: y, m = y + 1, 1
    return out


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); ap.add_argument("--xlsx"); a = ap.parse_args()
    p_out = os.path.join(OUT, "monitor_emp.json")
    prev = {}
    try:
        prev = json.load(io.open(p_out, encoding="utf-8")).get("patch", {})
    except Exception:
        pass
    patch = {}

    # 1. SIPA
    sp = None
    try:
        sp = sipa_excel(a.xlsx)
    except Exception as e:
        log("SIPA Excel: fallo (%r); voy a datos.gob.ar" % e)
        try:
            sp = sipa_datosgob()
        except Exception as e2:
            log("SIPA datos.gob.ar: fallo (%r)" % e2); FALLOS.append("sipa")
    # 2. INDEC API (ya en disco)
    ia = None
    try:
        ia = json.load(io.open(os.path.join(OUT, "indec_api.json"), encoding="utf-8"))
    except Exception as e:
        log("indec_api.json: fallo (%r)" % e); FALLOS.append("indec_api")
    S = (ia or {}).get("series", {})
    ser = lambda k: dict(S[k]["d"]) if k in S else {}

    sal = {"tot": ser("salarios|0|Indice"), "reg": ser("salarios|100|Indice"), "priv": ser("salarios|110|Indice"),
           "pub": ser("salarios|120|Indice"), "noreg": ser("salarios|200|Indice")}
    ipc = ser("ipc|Nacional|0|Indice")
    if ipc:
        ipc["2016-12"] = 100.0
    can = {"cbt_h2": ser("canastas|CBT H2|Hogar 2"), "cba_h2": ser("canastas|CBA H2|Hogar 2")}

    fin = max([DESDE] + [max(d) for d in list(sal.values()) + [ipc] + list(can.values()) if d]
              + ([max(sp["t"]["sa"])] if sp else []))
    M = meses_entre(DESDE, fin)
    al = lambda d, nd=2: [round(d[m], nd) if d.get(m) is not None else None for m in M]
    patch["m"] = M
    if sp:
        ult = max(sp["t"]["sa"])
        t = {c: {"sa": al({m: v[j] for m, v in sp["t"]["sa"].items()}, 3), "o": al({m: v[j] for m, v in sp["t"]["o"].items()}, 3)} for j, c in enumerate(TIPOS)}
        r = {c: {"sa": al({m: v[j] for m, v in sp["r"]["sa"].items()}, 3), "o": al({m: v[j] for m, v in sp["r"]["o"].items()}, 3)} for j, c in enumerate(RAMAS)}
        rem = {k: al({m: v[k] for m, v in sp["rem"].items()}, 0) for k in ("med", "med_sa", "prom", "prom_sa")}
        patch["sipa"] = {"ult": ult, "prov": sp["prov"], "via": sp["via"], "archivo": sp["archivo"], "t": t, "r": r, "rem": rem}
        log("SIPA: último %s (%s), %d meses provisorios" % (ult, sp["via"], len(sp["prov"])))
    elif prev.get("sipa"):
        log("SIPA: se conserva el bloque anterior (%s)" % prev["sipa"].get("ult"))
        pm = {m: i for i, m in enumerate(prev.get("m", []))}
        re_al = lambda arr: [arr[pm[m]] if m in pm and pm[m] < len(arr) else None for m in M]
        ps = prev["sipa"]
        patch["sipa"] = dict(ps, t={c: {k: re_al(v) for k, v in d.items()} for c, d in ps["t"].items()},
                             r={c: {k: re_al(v) for k, v in d.items()} for c, d in ps["r"].items()},
                             rem={k: re_al(v) for k, v in ps["rem"].items()})
    if ia:
        patch["sal"] = {k: al(v) for k, v in sal.items()}
        patch["ipc"] = al(ipc, 4)
        patch["can"] = {k: al(v) for k, v in can.items()}
        E = {"act": "tasa_de_actividad", "emp": "tasa_de_empleo", "des": "tasa_de_desocupacion", "sub": "tasa_de_subocupacion",
             "dem": "tasa_de_ocupacion_demandante", "form": "tasa_de_empleo_formal", "inf": "tasa_de_empleo_informal"}
        e = {k: ser("eph|T_31_A|" + v) for k, v in E.items()}
        Q = sorted(set().union(*[set(d) for d in e.values()]))
        patch["q"] = Q; patch["eph"] = {k: [d.get(x) for x in Q] for k, d in e.items()}
        B = {"pp": "tasa_pobreza_personas", "ip": "tasa_indigencia_personas", "ph": "tasa_pobreza_hogares", "ih": "tasa_indigencia_hogares"}
        b = {k: ser("pobreza|T_31_A|" + v) for k, v in B.items()}
        Sm = sorted(set().union(*[set(d) for d in b.values()]))
        patch["s"] = Sm; patch["pob"] = {k: [d.get(x) for x in Sm] for k, d in b.items()}
        if not Q or not Sm or not ipc or not sal["tot"]:
            FALLOS.append("indec_api incompleto (EPH, pobreza, IPC o salarios)")
        log("INDEC API: EPH %s, pobreza %s, salarios %s, IPC %s, canastas %s" % (Q[-1] if Q else "s/d", Sm[-1] if Sm else "s/d",
            max(sal["tot"]) if sal["tot"] else "s/d", max(ipc) if ipc else "s/d", max(can["cbt_h2"]) if can["cbt_h2"] else "s/d"))
    else:
        for k in ("sal", "ipc", "can", "q", "eph", "s", "pob"):
            if k in prev: patch[k] = prev[k]
    # 3. puestos de trabajo (INDEC, cuenta de generación del ingreso)
    try:
        cq, c = cgi()
        patch["cq"] = cq; patch["cgi"] = c
        log("Puestos de trabajo (INDEC 324.1): último %s" % cq[-1])
    except Exception as e:
        log("Puestos de trabajo 324.1: fallo (%r)" % e); FALLOS.append("cgi")
        for k in ("cq", "cgi"):
            if k in prev: patch[k] = prev[k]

    falta = [k for k in ("sipa", "eph", "pob", "sal", "ipc", "cgi") if k not in patch]
    if falta:
        log("faltan bloques %s y no hay versión anterior: no se escribe monitor_emp.json" % falta); return
    hoy = datetime.now(timezone.utc)
    patch["fetched"] = hoy.strftime("%Y-%m-%d")
    patch["src"] = "Secretaría de Trabajo (SIPA) e INDEC (EPH, cuenta de generación del ingreso, salarios, canastas), vía cafci-flujos/macro/monitor_emp.json"
    obj = {"meta": {"updatedAt": hoy.strftime("%Y-%m-%dT%H:%MZ"), "by": "sync_empleo.py (GitHub Actions)",
                    "fecha": {"sipa": patch["sipa"]["ult"], "eph": patch["q"][-1], "pobreza": patch["s"][-1], "cgi": patch["cq"][-1],
                              "salarios": max(m for m, v in zip(M, patch["sal"]["tot"]) if v is not None)},
                    "sipa_via": patch["sipa"]["via"], "indec_api": (ia or {}).get("meta", {}).get("updatedAt"),
                    "fallos": FALLOS, "log": LOG},
           "patch": patch}
    json.dump(obj, io.open(p_out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print("monitor_emp.json ->", p_out, "| %.0f KB" % (os.path.getsize(p_out) / 1e3), "| fallos:", FALLOS or "ninguno")
    if a.check:
        i = M.index(patch["sipa"]["ult"])
        for c in TIPOS:
            v = patch["sipa"]["t"][c]["sa"]
            print("  SIPA %-5s %s: %10.1f mil s.e. | m/m %+6.1f | i.a. %+7.1f" % (c, M[i], v[i], v[i] - v[i - 1], v[i] - v[i - 12]))
        print("  EPH %s: desocupación %s, empleo %s, actividad %s" % (patch["q"][-1], patch["eph"]["des"][-1], patch["eph"]["emp"][-1], patch["eph"]["act"][-1]))
        print("  Pobreza %s: %s personas, indigencia %s" % (patch["s"][-1], patch["pob"]["pp"][-1], patch["pob"]["ip"][-1]))
        g = patch["cgi"]["TOT"]
        print("  Puestos %s: total %.0f mil | reg %.0f | no reg %.0f | no asal %.0f" % (patch["cq"][-1], g["t"][-1], g["reg"][-1], g["nr"][-1], g["na"][-1]))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FALLO GENERAL sync_empleo: %r" % e)  # nunca tirar abajo el job
