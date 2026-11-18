# -*- coding: utf-8 -*-
"""
sync_nafta.py — nafta en surtidor contra el Brent (Monitor Argentina, bloque NAF). PROTOTIPO 18-nov-2026.

Fuentes, todas públicas y sin carga manual:
  1. Secretaría de Energía, Res. 1104/2004 "Precios y volúmenes EESS" (datos.energia.gob.ar): una fila por
     estación, producto y mes con precio con impuestos, precio sin impuestos y volumen; desde dic-2024 trae
     además ICL, IDC e IVA por litro. Sale con ~1 mes de rezago (el mes t se publica a fines de t+1).
     Se agrega a promedio país ponderado por volumen (y YPF, resto, YPF en CABA), canal "Al público".
  2. Secretaría de Energía, Res. 314/2016 "Precios en surtidor": precio vigente por estación, se actualiza
     todos los días. Sirve para estirar el último mes de la 1104 hasta hoy (mediana de estaciones). Ojo: YPF
     casi no informa a este registro desde fines de 2025, así que la extensión sale del resto de las banderas.
  3. Brent: EIA, "Europe Brent Spot Price FOB" diario (RBRTE, xls sin clave; rezago de 5 a 7 días).
     Respaldo: FRED DCOILBRENTEU (mismo dato). Opcional en la PC: --bbg agrega los últimos días con
     CO1 Comdty de Bloomberg (primer futuro; no es la fuente primaria porque la nube no ve la terminal).
  4. Tipo de cambio: A3500, API BCRA v4 id 5.
  5. Peso de los combustibles en el IPC: constante PESO_IPC (INDEC, ver abajo).

Metodología (resumen; el detalle va en el JSON, clave "metodo"):
  nafta sin impuestos en USD/l = precio sin impuestos / A3500 (promedios del mes)
  Brent en USD/l = Brent / 158,987
  diferencial = nafta sin impuestos USD/l − Brent USD/l (refinación + biocombustible + flete + márgenes)
  paridad = Brent USD/l + diferencial promedio del período de referencia (precios libres, REF)
  atraso = paridad / precio observado − 1 (positivo: la nafta está por debajo de lo que pide el Brent)
  a surtidor: precio = sin impuestos × k + ICL + IDC, con k = (con impuestos − ICL − IDC) / sin impuestos

Uso:  python sync_nafta.py --backfill          # primera vez: baja toda la historia (≈1,4 GB) y arma la base
      python sync_nafta.py                     # diario: baja solo el CSV 2025+ si cambió, el vigente, EIA y BCRA
      python sync_nafta.py --cache DIR         # usa archivos ya bajados en DIR (pruebas)
      python sync_nafta.py --bbg               # en la PC con terminal: completa el Brent hasta hoy
"""
from __future__ import annotations
import os, sys, io, csv, json, time, bisect, argparse, collections, statistics as st, urllib.request
import datetime as dt

AQUI = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(AQUI, "macro")
os.makedirs(OUT, exist_ok=True)
HDR = {"User-Agent": "Mozilla/5.0 (monitor-argentina-sync)"}
CKAN = "http://datos.energia.gob.ar/api/3/action/package_show?id="
EIA_BRENT = "https://www.eia.gov/dnav/pet/hist_xls/RBRTEd.xls"
FRED_BRENT = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DCOILBRENTEU"
BCRA = "https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/5?desde=%s&hasta=%s&limit=3000&offset=%d"
PROD = {"Nafta (súper) entre 92 y 95 Ron": "sup", "Nafta (premium) de más de 95 Ron": "pre",
        "Gas Oil Grado 2": "go2", "Gas Oil Grado 3": "go3"}
PID314 = {"2": "sup", "3": "pre", "19": "go2", "21": "go3"}
TAXC = ("impuesto_combustible_liquidos", "impuesto_dioxido_carbono", "iva")
BBL = 158.987                       # litros por barril
REF = ("2024-07", "2026-02")        # período de referencia: precios libres, antes del shock de 2026
REF2 = ("2017-01", "2018-04")       # segundo período sin regulación fuerte (solo para el traslado histórico)
BUFFER_DESDE = "2026-03"            # desde cuándo se acumula la cuenta del amortiguador
ESC_BRENT = (-20, -10, -5, 5, 10, 20)
ESC_FX = (-10, -5, 5, 10, 20)
EIA_GULF = "https://www.eia.gov/dnav/pet/hist_xls/EER_EPMRU_PF4_RGC_DPGd.xls"
GALON = 3.78541
IPC_APERTURAS = "https://www.indec.gob.ar/ftp/cuadros/economia/sh_ipc_aperturas.xls"
CLASE = "Combustibles y lubricantes para veh"
# INDEC, Metodología 32, Anexo II, cuadro 7: ponderación de la clase 07.2.2 "Combustibles y lubricantes para
# vehículos de uso del hogar" por región (base dic-2016, %). El INDEC no publica la nacional: se arma con el
# peso de cada región en el IPC nacional (el mismo de skills/argento/ipc_pesos.json). Orden = el del Excel.
REGIONES = (("GBA", 3.78, 45.196), ("Pampeana", 4.59, 33.8297), ("Noroeste", 3.80, 6.8034),
            ("Noreste", 4.26, 4.5207), ("Cuyo", 5.63, 5.1378), ("Patagonia", 5.12, 4.5124))
LOG, FALLOS = [], []


def log(m):
    LOG.append(str(m)); print(m)


def bajar(url, destino=None, tries=4):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=600) as r:
                if destino:
                    with open(destino, "wb") as fh:
                        while True:
                            b = r.read(1 << 20)
                            if not b: break
                            fh.write(b)
                    return destino
                return r.read()
        except Exception:
            if i == tries - 1: raise
            time.sleep(4 * (i + 1))


def f(x):
    try: return float(x)
    except Exception: return None


def media(xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) / len(xs) if xs else None


# ---------------------------------------------------------------- 1. Res. 1104 (mensual, por estación)
def agregar_1104(path):
    """CSV de la Res. 1104 → {(mes, prod): {pais, ypf, resto, caba_ypf, tax}}. Robusto a filas basura:
    descarta precios fuera de 0,6x a 1,6x la mediana del mes y pasa a m3 los volúmenes informados en litros."""
    rows = collections.defaultdict(list)
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            p = PROD.get(r["producto"])
            if not p or r["canal_de_comercializacion"] != "Al público": continue
            per = r["periodo"].replace("/", "-") if "periodo" in r else "%s-%02d" % (r["anio"], int(r["mes"]))
            v, pci, psi = f(r["volumen"]), f(r["precio_con_impuestos"]), f(r["precio_sin_impuestos"])
            if not v or v <= 0 or not pci or pci <= 0 or not psi or psi <= 0 or psi >= pci: continue
            tx = tuple(f(r[c]) for c in TAXC) if "iva" in r else None
            rows[(per, p)].append((pci, psi, v / 1000.0 if v > 3000 else v, r["bandera"].strip().upper() == "YPF",
                                   r["provincia"].strip().upper() == "CAPITAL FEDERAL", tx))
    out = {}
    for (per, p), L in rows.items():
        m = st.median(x[0] for x in L)
        L2 = [x for x in L if 0.6 * m <= x[0] <= 1.6 * m]

        def ag(S):
            if len(S) < 5: return None
            V = sum(x[2] for x in S)
            return {"n": len(S), "vol": round(V), "pci": round(sum(x[2] * x[0] for x in S) / V, 2),
                    "psi": round(sum(x[2] * x[1] for x in S) / V, 2)}
        o = {"pais": ag(L2), "ypf": ag([x for x in L2 if x[3]]), "resto": ag([x for x in L2 if not x[3]]),
             "caba_ypf": ag([x for x in L2 if x[3] and x[4]])}
        T = [x[5] for x in L2 if x[5] and x[5][0] and 20 < x[5][0] < 2000 and (x[5][2] or 0) > 30]   # filas en $/l
        if len(T) > 500:
            o["icl"], o["idc"], o["iva"] = [round(st.median(t[i] or 0 for t in T), 2) for i in range(3)]
        out["%s|%s" % (per, p)] = o
    return out


def recursos(dataset):
    d = json.loads(bajar(CKAN + dataset).decode("utf-8"))
    return [(r.get("name") or "", r["url"], (r.get("last_modified") or "")[:19]) for r in d["result"]["resources"]]


def nafta_1104(cache, backfill, previo):
    """Devuelve ({mes|prod: agregados}, meta). Los años cerrados quedan en macro/nafta_hist.json."""
    hist_path = os.path.join(OUT, "nafta_hist.json")
    hist = json.load(open(hist_path, encoding="utf-8")) if os.path.exists(hist_path) else {"m1104": {}, "s314": {}}
    try:
        rec = recursos("precios-eess---resolucion-1104-04")
        actual = [r for r in rec if "desde" in r[0].lower()][0]
    except Exception as e:
        if backfill or not (previo and previo.get("m1104_actual")): raise
        FALLOS.append("Energía 1104 no respondió (%s): se reusa la última bajada" % e)
        m = dict(hist["m1104"]); m.update(previo["m1104_actual"])
        return m, {"csv_1104_modificado": previo["meta"].get("csv_1104_modificado")}, previo["m1104_actual"], hist
    if backfill:
        for nom, url, _ in rec:
            anio = nom.strip()[-4:]
            if "desde" in nom.lower() or not anio.isdigit() or int(anio) < 2017: continue
            p = os.path.join(cache, "eess_%s.csv" % anio)
            if not os.path.exists(p): log("bajando 1104 %s" % anio); bajar(url, p)
            hist["m1104"].update(agregar_1104(p))
        json.dump(hist, open(hist_path, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    meta = {"csv_1104_modificado": actual[2]}
    prev_meta = (previo or {}).get("meta", {})
    if not backfill and prev_meta.get("csv_1104_modificado") == actual[2] and previo.get("m1104_actual"):
        log("1104: el CSV 2025+ no cambió (%s), se reusa" % actual[2])
        act = previo["m1104_actual"]
    else:
        p = os.path.join(cache, "eess_2025.csv")
        if not (os.path.exists(p) and cache != OUT and not backfill and os.environ.get("NAFTA_REUSAR")):
            log("bajando 1104 2025+ (%s)" % actual[2]); bajar(actual[1], p)
        act = agregar_1104(p)
    m = dict(hist["m1104"]); m.update(act)
    return m, meta, act, hist


# ---------------------------------------------------------------- 2. Res. 314 (precio vigente por estación)
def _fecha(s):
    s = s.strip()
    try:
        return dt.datetime.strptime(s[:10], "%d/%m/%Y").date() if "/" in s else dt.date.fromisoformat(s[:10])
    except Exception:
        return None


def leer_314(path, hoy):
    ev = collections.defaultdict(list)
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            p = PID314.get(r["idproducto"])
            if not p or r["idtipohorario"] != "2": continue
            d, px = _fecha(r["fecha_vigencia"]), f(r["precio"])
            if not d or not px or px <= 0 or d > hoy or d.year < 2016: continue
            ev[(p, r["idempresa"] + "|" + r["direccion"])].append((d, px, r["empresabandera"].strip().upper() == "YPF"))
    return ev


def mediana_314(ev, dias, prod="sup", tope=365):
    """Mediana del último precio vigente de cada estación (con cambio en los últimos `tope` días)."""
    S = [(sorted(set(L)),) for (p, _), L in ev.items() if p == prod]
    S = [(L, [x[0] for x in L]) for (L,) in S]
    out = []
    for d in dias:
        v = {"pais": [], "ypf": [], "resto": []}
        for L, F in S:
            i = bisect.bisect_right(F, d) - 1
            if i < 0 or (d - F[i]).days > tope: continue
            v["pais"].append(L[i][1]); v["ypf" if L[i][2] else "resto"].append(L[i][1])
        out.append([d.isoformat()] + [round(st.median(v[k]), 1) if len(v[k]) >= 30 else None for k in ("pais", "ypf", "resto")]
                   + [len(v["pais"]), len(v["ypf"])])
    return out


def nafta_314(cache, backfill, hist, hoy):
    """Serie semanal (miércoles) + hoy de la mediana de precio vigente. La historia sale una vez del CSV
    histórico (779 MB, --backfill) y después se estira con el CSV de vigentes de cada corrida."""
    try:
        rec = recursos("precios-en-surtidor")
        vig = [r for r in rec if "vigentes" in r[0].lower()][0]
        pv = os.path.join(cache, "vigentes.csv")
        if not (os.path.exists(pv) and os.environ.get("NAFTA_REUSAR")): bajar(vig[1], pv)
        ev = leer_314(pv, hoy)
    except Exception as e:
        if backfill: raise
        FALLOS.append("Energía 314 no respondió (%s): el surtidor de hoy queda con el último punto guardado" % e)
        return {k: list(v) for k, v in hist["s314"].items()}, {"csv_314_modificado": None}
    if backfill:
        ph = os.path.join(cache, "hist314.csv")
        if not os.path.exists(ph):
            log("bajando histórico 314 (779 MB)"); bajar([r for r in rec if "hist" in r[0].lower()][0][1], ph)
        for k, L in leer_314(ph, hoy).items(): ev[k].extend(L)
        d, dias = dt.date(2024, 1, 3), []
        while d < hoy: dias.append(d); d += dt.timedelta(days=7)
        for prod in ("sup", "pre", "go2", "go3"):
            hist["s314"][prod] = [r for r in mediana_314(ev, dias, prod) if r[1]]
        json.dump(hist, open(os.path.join(OUT, "nafta_hist.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    # corrida diaria: el CSV de vigentes solo permite calcular el punto de hoy; se agrega a la base
    out = {}
    for prod in ("sup", "pre", "go2", "go3"):
        base = [r for r in hist["s314"].get(prod, []) if r[0] < hoy.isoformat()]
        out[prod] = base + [r for r in mediana_314(ev, [hoy], prod) if r[1]]
        hist["s314"][prod] = out[prod]
    json.dump(hist, open(os.path.join(OUT, "nafta_hist.json"), "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    return out, {"csv_314_modificado": vig[2]}


# ---------------------------------------------------------------- 3. Brent y 4. A3500
def brent(cache, usar_bbg):
    d, via = {}, None
    try:
        d = eia_xls(EIA_BRENT)
        via = "EIA RBRTE (Brent spot FOB)"
    except Exception as e:
        log("EIA falló (%s), pruebo FRED" % e)
        try:
            for ln in bajar(FRED_BRENT).decode().splitlines()[1:]:
                a, b = ln.split(",")[:2]
                if f(b): d[a] = float(b)
            via = "FRED DCOILBRENTEU (Brent spot FOB)"
        except Exception as e2:
            FALLOS.append("brent: EIA y FRED fallaron (%s)" % e2)
    ult_pub = max(d) if d else None
    if usar_bbg:
        try:
            sys.path.insert(0, os.path.join(AQUI, "..", "reporte_rf"))
            import bbg
            h = bbg.bdh(["CO1 Comdty"], ["PX_LAST"], ult_pub or "2017-01-01", bbg.hoy()).get("CO1 Comdty", [])
            n = 0
            for fe, v in h:
                if fe > (ult_pub or "") and v.get("PX_LAST"): d[fe] = v["PX_LAST"]; n += 1
            if n: via += " + Bloomberg CO1 (%d ruedas finales)" % n
        except Exception as e:
            log("Bloomberg no disponible (%s): queda el Brent público" % e)
    return d, via, ult_pub


def eia_xls(url):
    import xlrd
    wb = xlrd.open_workbook(file_contents=bajar(url)); s = wb.sheet_by_index(1); d = {}
    for i in range(3, s.nrows):
        if s.cell_value(i, 1) != "":
            d[xlrd.xldate_as_datetime(s.cell_value(i, 0), wb.datemode).date().isoformat()] = float(s.cell_value(i, 1))
    return d


def ipc_combustibles(cache):
    """Peso de la clase combustibles en el IPC nacional (base y a precios de hoy) y su variación mensual,
    armados con el Excel de aperturas del INDEC (la clase solo se publica por región)."""
    import xlrd
    p = os.path.join(cache, "sh_ipc_aperturas.xls")
    if not (os.path.exists(p) and os.environ.get("NAFTA_REUSAR")): bajar(IPC_APERTURAS, p)
    wb = xlrd.open_workbook(p); s = [x for x in wb.sheets() if x.name.lower().startswith("índices") or x.name.lower().startswith("indices")][0]
    ini = [i for i in range(s.nrows) if str(s.cell_value(i, 0)).startswith("Región")]
    assert len(ini) == 6, "el Excel de aperturas cambió de formato"
    meses = [xlrd.xldate_as_datetime(v, wb.datemode).strftime("%Y-%m") for v in s.row_values(ini[0])[1:] if isinstance(v, float)]
    n = len(meses); gen, cla = [], []
    for j, i0 in enumerate(ini):
        fin = ini[j + 1] if j < 5 else s.nrows
        g = [i for i in range(i0, fin) if str(s.cell_value(i, 0)).startswith("Nivel general")][0]
        c = [i for i in range(i0, fin) if str(s.cell_value(i, 0)).startswith(CLASE)][0]
        gen.append(s.row_values(g)[1:1 + n]); cla.append(s.row_values(c)[1:1 + n])
    W = [r[2] / 100 for r in REGIONES]; w0 = [r[1] for r in REGIONES]
    nac_g = [sum(W[r] * gen[r][t] for r in range(6)) for t in range(n)]
    nac_c = [sum(W[r] * w0[r] * cla[r][t] for r in range(6)) / sum(W[r] * w0[r] for r in range(6)) for t in range(n)]
    base = sum(W[r] * w0[r] for r in range(6))
    hoy = base * nac_c[-1] / nac_g[-1]
    return {"valor": round(hoy, 2), "base_dic16": round(base, 2), "mes": meses[-1], "provisorio": False,
            "gba_base": 3.78, "gba_hoy": round(3.78 * cla[0][-1] / gen[0][-1], 2),
            "fuente": "INDEC, Metodología 32 (cuadro 7) y sh_ipc_aperturas.xls; nacional = regiones ponderadas",
            "m": meses, "idx_clase": [round(v, 2) for v in nac_c], "idx_general": [round(v, 2) for v in nac_g]}


def a3500(hasta):
    d, off = {}, 0
    while True:
        j = json.loads(bajar(BCRA % ("2017-01-01", hasta, off)).decode("utf-8"))
        det = j["results"][0]["detalle"]
        for r in det: d[r["fecha"]] = r["valor"]
        off += 3000
        if off >= j["metadata"]["resultset"]["count"]: break
    return d


def mensual(d):
    m = collections.defaultdict(list)
    for k, v in d.items(): m[k[:7]].append(v)
    return {k: sum(v) / len(v) for k, v in m.items()}


# ---------------------------------------------------------------- 5. modelo
def ols(x, y):
    mx, my = media(x), media(y)
    b = sum((a - mx) * (c - my) for a, c in zip(x, y)) / sum((a - mx) ** 2 for a in x)
    return my - b * mx, b


def armar(m1104, s314, br, fx, gu, ipc, brent_via, brent_pub, hoy):
    brm, fxm, gum = mensual(br), mensual(fx), mensual(gu)
    meses = sorted({k.split("|")[0] for k in m1104 if k.endswith("|sup")})
    meses = [x for x in meses if x in brm and x in fxm and m1104[x + "|sup"]["pais"]]
    S = {k: [] for k in ("pci", "psi", "pci_ypf", "pci_resto", "pci_caba_ypf", "vol", "fx", "brent", "icl", "idc", "k",
                         "psi_usd", "brent_l", "dif", "par_psi", "par_pci", "atraso_pct", "atraso_ars", "pci_pre", "pci_go2", "pci_go3",
                         "gulf_l", "dif_g", "par_pci_g", "atraso_pct_g")}
    kprev = None
    for x in meses:
        o = m1104[x + "|sup"]; a = o["pais"]
        S["pci"].append(a["pci"]); S["psi"].append(a["psi"]); S["vol"].append(a["vol"])
        for g in ("ypf", "resto", "caba_ypf"): S["pci_" + g].append(o[g]["pci"] if o.get(g) else None)
        for q in ("pre", "go2", "go3"):
            oq = m1104.get(x + "|" + q); S["pci_" + q].append(oq["pais"]["pci"] if oq and oq["pais"] else None)
        S["fx"].append(round(fxm[x], 2)); S["brent"].append(round(brm[x], 2))
        S["icl"].append(o.get("icl")); S["idc"].append(o.get("idc"))
        S["k"].append(round((a["pci"] - o["icl"] - o["idc"]) / a["psi"], 4) if o.get("icl") else None)
        S["psi_usd"].append(a["psi"] / fxm[x]); S["brent_l"].append(brm[x] / BBL)
        S["dif"].append(S["psi_usd"][-1] - S["brent_l"][-1])
        S["gulf_l"].append(gum[x] / GALON if x in gum else None)
        S["dif_g"].append(S["psi_usd"][-1] - S["gulf_l"][-1] if x in gum else None)
    ix = lambda v: [i for i, x in enumerate(meses) if v[0] <= x <= v[1]]
    R1, R2 = ix(REF), ix(REF2)
    dif_ref = media([S["dif"][i] for i in R1]); dif_sd = st.pstdev([S["dif"][i] for i in R1])
    dif_ref_g = media([S["dif_g"][i] for i in R1]); dif_sd_g = st.pstdev([S["dif_g"][i] for i in R1 if S["dif_g"][i] is not None])
    # traslado histórico: pendiente de nafta sin impuestos (USD/l) contra Brent (USD/l), con ordenada propia por período
    xs, ys = [], []
    for R in (R1, R2):
        mx, my = media([S["brent_l"][i] for i in R]), media([S["psi_usd"][i] for i in R])
        xs += [S["brent_l"][i] - mx for i in R]; ys += [S["psi_usd"][i] - my for i in R]
    beta = sum(a * b for a, b in zip(xs, ys)) / sum(a * a for a in xs)
    betas = {"ref": ols([S["brent_l"][i] for i in R1], [S["psi_usd"][i] for i in R1])[1],
             "ref2": ols([S["brent_l"][i] for i in R2], [S["psi_usd"][i] for i in R2])[1]}
    kk = [k for k in S["k"] if k]; k_ult = S["k"][-1] or media(kk); k_med = media(kk)
    for i, x in enumerate(meses):
        par = S["brent_l"][i] + dif_ref
        k = S["k"][i] or k_med
        gap = (par - S["psi_usd"][i]) * S["fx"][i] * k              # $/l en surtidor
        S["par_psi"].append(round(par * S["fx"][i], 2)); S["par_pci"].append(round(S["pci"][i] + gap, 2))
        S["atraso_ars"].append(round(gap, 1)); S["atraso_pct"].append(round(100 * gap / S["pci"][i], 2))
        gg = (S["gulf_l"][i] + dif_ref_g - S["psi_usd"][i]) * S["fx"][i] * k if S["gulf_l"][i] is not None else None
        S["par_pci_g"].append(round(S["pci"][i] + gg, 2) if gg is not None else None); S["atraso_pct_g"].append(round(100 * gg / S["pci"][i], 2) if gg is not None else None)
    # cuenta del amortiguador: (precio sin impuestos − paridad) × volumen de súper, acumulado en USD MM
    cta, acum, cta_g, acum_g = [], 0.0, [], 0.0
    volref = st.median(S["vol"][-18:])
    for i, x in enumerate(meses):
        if x < BUFFER_DESDE: cta.append(None); cta_g.append(None); continue
        v = min(S["vol"][i], 1.3 * volref)                           # m3; se acota un mes con volumen anómalo
        acum += (S["psi_usd"][i] - (S["brent_l"][i] + dif_ref)) * v * 1000 / 1e6
        cta.append(round(acum, 1))
        if S["gulf_l"][i] is not None: acum_g += (S["psi_usd"][i] - (S["gulf_l"][i] + dif_ref_g)) * v * 1000 / 1e6
        cta_g.append(round(acum_g, 1))
    # ---- foto de hoy: último mes 1104 estirado con la Res. 314, Brent y A3500 del último día
    ult = meses[-1]; s = s314.get("sup", [])
    base314 = media([r[1] for r in s if r[0][:7] == ult]); hoy314 = s[-1][1] if s else None
    fac = hoy314 / base314 if base314 and hoy314 else 1.0
    icl, idc = S["icl"][-1] or 0.0, S["idc"][-1] or 0.0
    pci_h = S["pci"][-1] * fac; psi_h = (pci_h - icl - idc) / k_ult
    fb, ffx = max(br), max(fx); b_h, fx_h = br[fb], fx[ffx]
    fg = max(gu); g_h = gu[fg] / GALON; bg_h = br.get(fg, b_h)      # nafta del Golfo y Brent del mismo día

    def foto(b, tc):
        par_psi = (b / BBL + dif_ref) * tc; par_pci = par_psi * k_ult + icl + idc
        return par_psi, par_pci
    par_psi_h, par_pci_h = foto(b_h, fx_h)
    # elasticidades en el punto de hoy
    e_brent = (b_h / BBL) * fx_h * k_ult / par_pci_h                 # paridad en surtidor ante 1% de Brent
    e_fx = par_psi_h * k_ult / par_pci_h                              # ante 1% de tipo de cambio
    e_hist = beta * e_brent                                           # con el traslado histórico
    w = ipc["valor"]
    par_pci_g_h = (g_h + (b_h - bg_h) / BBL + dif_ref_g) * fx_h * k_ult + icl + idc   # Golfo llevado al Brent de hoy, crack constante

    def fila(sb, sfx):
        b, tc = b_h * (1 + sb / 100), fx_h * (1 + sfx / 100)
        pp, pc = foto(b, tc)
        d_par = 100 * (pc / par_pci_h - 1)
        pcg = (g_h + (b - bg_h) / BBL + dif_ref_g) * tc * k_ult + icl + idc
        d_hist = 100 * (beta * (b - b_h) / BBL * tc * k_ult + (psi_h / fx_h) * (tc - fx_h) * k_ult) / pci_h
        return {"brent": round(b, 1), "fx": round(tc, 1), "paridad": round(pc), "d_paridad_pct": round(d_par, 1),
                "atraso_pct": round(100 * (pc / pci_h - 1), 1), "atraso_ars": round(pc - pci_h),
                "paridad_g": round(pcg), "atraso_g_pct": round(100 * (pcg / pci_h - 1), 1), "atraso_g_ars": round(pcg - pci_h),
                "ipc_cierre_g_pp": round(w * (pcg / pci_h - 1), 2),
                "suba_hist_pct": round(d_hist, 1), "surtidor_hist": round(pci_h * (1 + d_hist / 100)),
                "ipc_pleno_pp": round(w * d_par / 100, 2) if w else None, "ipc_hist_pp": round(w * d_hist / 100, 2) if w else None,
                "ipc_cierre_pp": round(w * (pc / pci_h - 1), 2) if w else None}
    r3 = lambda v: round(v, 3)
    return {
        "fetched": hoy.isoformat(),
        "m": meses,
        "s": {k: [round(v, 4) if isinstance(v, float) else v for v in S[k]] for k in S},
        "cuenta_usd_mm": cta, "cuenta_g_usd_mm": cta_g,
        "s314": {"sup": s314.get("sup", [])[-60:]},
        "modelo": {"ref": REF, "ref2": REF2, "dif_ref_usd_l": r3(dif_ref), "dif_sd_usd_l": r3(dif_sd), "beta_traslado": r3(beta),
                   "beta_ref": r3(betas["ref"]), "beta_ref2": r3(betas["ref2"]), "k": round(k_ult, 4), "icl": icl, "idc": idc,
                   "elast_brent_paridad": r3(e_brent), "elast_fx_paridad": r3(e_fx), "elast_brent_hist": r3(e_hist),
                   "dif_ref_g_usd_l": r3(dif_ref_g), "dif_sd_g_usd_l": r3(dif_sd_g), "peso_ipc": {k: v for k, v in ipc.items() if k not in ("m", "idx_clase", "idx_general")}},
        "ipc": {k: ipc[k] for k in ("m", "idx_clase", "idx_general")},
        "hoy": {"fecha": hoy.isoformat(), "ult_mes_1104": ult, "factor_314": round(fac, 4), "fecha_314": s[-1][0] if s else None,
                "surtidor": round(pci_h), "sin_imp": round(psi_h), "brent": round(b_h, 2), "fecha_brent": fb, "brent_publico_hasta": brent_pub,
                "fx": round(fx_h, 2), "fecha_fx": ffx, "paridad": round(par_pci_h), "paridad_sin_imp": round(par_psi_h),
                "atraso_pct": round(100 * (par_pci_h / pci_h - 1), 1), "atraso_ars": round(par_pci_h - pci_h),
                "dif_usd_l": r3(psi_h / fx_h - b_h / BBL), "surtidor_usd": r3(pci_h / fx_h),
                "gulf_usd_l": r3(g_h), "fecha_gulf": fg, "crack_usd_l": r3(g_h - bg_h / BBL), "paridad_g": round(par_pci_g_h),
                "atraso_g_pct": round(100 * (par_pci_g_h / pci_h - 1), 1), "atraso_g_ars": round(par_pci_g_h - pci_h)},
        "esc_brent": [dict(shock=x, **fila(x, 0)) for x in ESC_BRENT],
        "esc_fx": [dict(shock=x, **fila(0, x)) for x in ESC_FX],
        "src": {"nafta": "Secretaría de Energía, Res. 1104/2004 (mensual) y Res. 314/2016 (vigentes)", "brent": brent_via, "nafta_intl": "EIA, nafta regular convencional spot Costa del Golfo (EER_EPMRU_PF4_RGC_DPG)", "fx": "BCRA, Com. A 3500 (API v4 id 5)"},
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--backfill", action="store_true"); ap.add_argument("--cache", default=None)
    ap.add_argument("--bbg", action="store_true"); ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    cache = a.cache or os.path.join(OUT, "_nafta_cache"); os.makedirs(cache, exist_ok=True)
    hoy = dt.date.today()
    salida = os.path.join(OUT, "nafta.json")
    previo = json.load(open(salida, encoding="utf-8")) if os.path.exists(salida) else None
    m1104, meta, act, hist = nafta_1104(cache, a.backfill, previo)
    s314, meta2 = nafta_314(cache, a.backfill, hist, hoy); meta.update(meta2)
    br, via, pub = brent(cache, a.bbg)
    fx = a3500(hoy.isoformat())
    if not br or not fx: raise SystemExit("sin Brent o sin A3500: no se escribe")
    try: gu = eia_xls(EIA_GULF)
    except Exception as e: gu = {}; FALLOS.append("nafta del Golfo (EIA): %s" % e)
    try: ipc = ipc_combustibles(cache)
    except Exception as e:
        pp = (previo or {}).get("patch", {})
        if not pp.get("ipc"): raise SystemExit("sin el Excel de aperturas del INDEC y sin corrida previa: no se escribe (%s)" % e)
        FALLOS.append("INDEC aperturas no respondió (%s): se reusa la ponderación previa" % e)
        ipc = dict(pp["modelo"]["peso_ipc"], **pp["ipc"])
    patch = armar(m1104, s314, br, fx, gu, ipc, via, pub, hoy)
    meta.update({"updatedAt": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "fallos": FALLOS,
                 "fecha": {"nafta_1104": patch["m"][-1], "nafta_314": patch["hoy"]["fecha_314"], "brent": patch["hoy"]["fecha_brent"], "fx": patch["hoy"]["fecha_fx"]}})
    json.dump({"meta": meta, "patch": patch, "m1104_actual": act}, open(salida, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    log("ok %s (%d KB)" % (salida, os.path.getsize(salida) // 1024))
    if a.check:
        print(json.dumps(patch["hoy"], ensure_ascii=False, indent=1)); print(json.dumps(patch["modelo"], ensure_ascii=False, indent=1))
        for r in patch["esc_brent"] + patch["esc_fx"]: print(r)
        print("mes     surtidor  sin_imp  brent   fx    dif   paridad atraso% cuenta")
        for i, x in enumerate(patch["m"]):
            if x >= "2025-06" or x[5:] == "01":
                s = patch["s"]; print(x, s["pci"][i], s["psi"][i], s["brent"][i], s["fx"][i], round(s["dif"][i], 3), s["par_pci"][i], s["atraso_pct"][i], patch["cuenta_usd_mm"][i], "| golfo", round(s["gulf_l"][i], 3), round(s["dif_g"][i], 3), s["atraso_pct_g"][i], patch["cuenta_g_usd_mm"][i])


if __name__ == "__main__":
    main()
