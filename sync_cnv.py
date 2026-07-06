#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sincronizador de FLUJOS de FCI desde CNV (fuente oficial, publica, sin token).

Circuito (todo server-side, verificado 06-jul-2026):
  1) GET  https://www.cnv.gov.ar/SitioWeb/FondosComunesInversion/CuotaPartes
        -> tabla con (fecha_documento, fecha_recepcion, GUID de presentacion en el link al AIF)
  2) GET  https://aif2.cnv.gov.ar/Presentations/publicview/{presentacionGUID}
        -> HTML con un JSON embebido: {"nombreArchivo":"AAAAMMDD_Planilla_Diaria_A.xlsx",...,"guid":"<fileGUID>"}
  3) POST https://blob.cnv.gov.ar/BlobWebService.svc/DownloadBlob/{fileGUID}
        -> binario xlsx (Planilla Diaria de Cuotapartes: 1 fila por fondo)

El xlsx trae, por fondo: Clasificacion, Sociedad Gerente, Plazo Liq. (T+0/T+1), Moneda,
Codigo CAFCI, y VCP + cantidad de cuotapartes de HOY y AYER -> flujo diario autosuficiente.

FLUJO (por fondo, dia) = (cuotapartes_hoy - cuotapartes_ayer) * VCP_hoy/1000
  (patrimonio = cuotapartes * VCP/1000; el cambio de cuotapartes es suscripciones netas).

Agregados por bucket (tipo de fondo) y por gestora, en ventanas 1D/1W/1M/YTD/1Y, mostrando
INGRESOS brutos (suma de flujos>0), EGRESOS brutos (suma de flujos<0) y NETO.
Las ventanas >1D requieren historia: usar --backfill N para bajar los ultimos N documentos.

Uso:
  python sync_cnv.py                 # baja el ultimo dia y actualiza history + flujos_latest.json
  python sync_cnv.py --backfill 260  # siembra ~1 anio de historia (una vez)
  python sync_cnv.py --date 2026-07-03
"""
import os
import re
import io
import sys
import csv
import json
import time
import argparse
import datetime as dt

try:
    from curl_cffi import requests as _rq
    _IMPERSONATE = "chrome"
except Exception:
    import requests as _rq
    _IMPERSONATE = None

import openpyxl

CNV_LIST = "https://www.cnv.gov.ar/SitioWeb/FondosComunesInversion/CuotaPartes"
AIF_VIEW = "https://aif2.cnv.gov.ar/Presentations/publicview/"
BLOB_DL = "https://blob.cnv.gov.ar/BlobWebService.svc/DownloadBlob/"

HEADERS = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
           "Accept-Language": "es-AR,es;q=0.9"}

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
HISTORY = os.path.join(OUT_DIR, "cnv_history.csv")
HIST_COLS = ["fecha", "key", "fondo", "clasif", "gerente", "plazo", "moneda",
             "cuotapartes", "vcp", "patrimonio"]

_MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
          "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12}


# ------------------------------- HTTP ------------------------------------------
def _req(method, url, tries=5, backoff=2.5, **kw):
    kw.setdefault("timeout", 90)
    kw.setdefault("headers", HEADERS)
    if _IMPERSONATE:
        kw["impersonate"] = _IMPERSONATE
    for i in range(tries):
        try:
            r = _rq.request(method, url, **kw)
            if r.status_code == 200:
                return r
            print("   ! HTTP %s %s" % (r.status_code, url), file=sys.stderr)
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(backoff * (i + 1)); continue
            return None
        except Exception as e:
            print("   ! %s (%s)" % (str(e)[:120], url), file=sys.stderr)
            time.sleep(backoff * (i + 1))
    return None


# ------------------------------- 1) lista de documentos ------------------------
def listar_documentos():
    """[(fecha_iso_documento, presentacion_guid)] ordenado como en la pagina (recepcion desc)."""
    r = _req("GET", CNV_LIST)
    if not r:
        raise SystemExit("No se pudo leer la lista de CuotaPartes de CNV.")
    html = r.text
    # <a href=".../publicview/GUID">3 jul. 2026</a>
    pat = re.compile(r'publicview/([0-9A-Fa-f-]{30,})"\s*>\s*(\d{1,2})\s+([a-z]{3})\.?\s+(\d{4})', re.I)
    out = []
    for m in pat.finditer(html):
        guid = m.group(1)
        d, mon, y = int(m.group(2)), _MESES.get(m.group(3).lower()), int(m.group(4))
        if not mon:
            continue
        out.append(("%04d-%02d-%02d" % (y, mon, d), guid))
    return out


# ------------------------------- 2) guid del archivo ---------------------------
def file_guid(presentacion_guid):
    r = _req("GET", AIF_VIEW + presentacion_guid)
    if not r:
        return None
    m = re.search(r'"nombreArchivo":"([^"]*Planilla_Diaria[^"]*\.xlsx)"\s*,\s*"tamano":"[^"]*"\s*,\s*"guid":"([0-9a-f-]+)"',
                  r.text, re.I)
    if not m:
        # fallback: cualquier nombreArchivo .xlsx seguido de guid
        m = re.search(r'"nombreArchivo":"([^"]*\.xlsx)"[^{}]*?"guid":"([0-9a-f-]+)"', r.text, re.I)
    return (m.group(1), m.group(2)) if m else None


# ------------------------------- 3) descargar xlsx -----------------------------
def descargar_xlsx(fguid):
    r = _req("POST", BLOB_DL + fguid, data="")
    if not r:
        return None
    return r.content


# ------------------------------- parseo xlsx -----------------------------------
def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def parse_planilla(content):
    """Devuelve (fecha_iso, [filas por fondo])."""
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    ws = wb["Sheet 1"] if "Sheet 1" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(min_row=9, values_only=True))
    wb.close()
    clasif = None
    fondos = []
    fecha_iso = None
    for r in rows:
        c0 = r[0]
        vcp_now = _num(r[5]); patr_now = _num(r[14]); cuot_now = _num(r[12])
        if c0 and vcp_now is None and patr_now is None:
            clasif = str(c0).strip()
            continue
        if not c0 or vcp_now is None:
            continue
        # fecha del dato (col 4, dd/mm/yy)
        if fecha_iso is None and r[4]:
            try:
                fecha_iso = dt.datetime.strptime(str(r[4]).strip(), "%d/%m/%y").strftime("%Y-%m-%d")
            except Exception:
                try:
                    fecha_iso = dt.datetime.strptime(str(r[4]).strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
                except Exception:
                    pass
        cod_cafci = r[20]
        key = str(cod_cafci).strip() if cod_cafci not in (None, "") else str(c0).strip()
        fondos.append(dict(
            key=key, fondo=str(c0).strip(), clasif=clasif,
            gerente=(str(r[23]).strip() if r[23] else None),
            moneda=(str(r[1]).strip() if r[1] else None),
            plazo=r[37],
            cuotapartes=cuot_now, vcp=vcp_now, patrimonio=patr_now,
        ))
    return fecha_iso, fondos


# ------------------------------- history ---------------------------------------
def append_history(fecha_iso, fondos):
    os.makedirs(OUT_DIR, exist_ok=True)
    existing = set()
    if os.path.exists(HISTORY):
        with open(HISTORY, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                existing.add((r["fecha"], r["key"]))
    write_header = not os.path.exists(HISTORY)
    n = 0
    with open(HISTORY, "a", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=HIST_COLS)
        if write_header:
            w.writeheader()
        for f in fondos:
            if (fecha_iso, f["key"]) in existing:
                continue
            w.writerow({"fecha": fecha_iso, "key": f["key"], "fondo": f["fondo"],
                        "clasif": f["clasif"], "gerente": f["gerente"], "plazo": f["plazo"],
                        "moneda": f["moneda"], "cuotapartes": f["cuotapartes"],
                        "vcp": f["vcp"], "patrimonio": f["patrimonio"]})
            n += 1
    print("   history += %d filas (%s)" % (n, fecha_iso))


def load_history():
    if not os.path.exists(HISTORY):
        return []
    with open(HISTORY, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    for r in rows:
        r["cuotapartes"] = _num(r["cuotapartes"])
        r["vcp"] = _num(r["vcp"])
        r["patrimonio"] = _num(r["patrimonio"])
    return rows


# ------------------------------- buckets + flujos ------------------------------
def bucket(clasif, plazo, moneda):
    c = (clasif or "").upper()
    usd = "DOLAR" in c or (moneda or "").upper() in ("USD", "U$S", "DOL")
    try:
        p = int(plazo)
    except (TypeError, ValueError):
        p = None
    if "MERCADO DE DINERO" in c:
        return "Money Market USD" if usd else "Money Market ARS (T+0)"
    if "RENTA FIJA" in c:
        if usd:
            return "Renta Fija USD"
        return "Renta Fija ARS T+1" if p == 1 else "Renta Fija ARS"
    if "RENTA MIXTA" in c:
        return "Renta Mixta"
    if "RENTA VARIABLE" in c:
        return "Renta Variable"
    if "PYME" in c:
        return "PyMEs"
    if "INFRAEST" in c:
        return "Infraestructura"
    if "RETORNO TOTAL" in c:
        return "Retorno Total"
    return (clasif or "Otros").replace(" Peso Argentina", "").replace(" Dolar Estadounidense", " USD")


WINDOWS = ["1D", "1W", "1M", "YTD", "1Y"]


def _window_start(fechas, cierre):
    """Devuelve {ventana: fecha_inicio disponible en el history}."""
    c = dt.date.fromisoformat(cierre)
    objetivos = {
        "1D": None,  # dia habil anterior (se resuelve como la fecha inmediatamente previa)
        "1W": c - dt.timedelta(days=7),
        "1M": c - dt.timedelta(days=30),
        "YTD": dt.date(c.year, 1, 1),
        "1Y": c - dt.timedelta(days=365),
    }
    fset = sorted(dt.date.fromisoformat(f) for f in fechas)
    res = {}
    idx = fset.index(c) if c in fset else len(fset) - 1
    res["1D"] = fset[idx - 1].isoformat() if idx > 0 else None
    for w in ("1W", "1M", "YTD", "1Y"):
        tgt = objetivos[w]
        prev = [f for f in fset if f <= tgt]
        res[w] = (prev[-1] if prev else fset[0]).isoformat()
    return res


def flujos(cierre):
    rows = load_history()
    if not rows:
        return None
    fechas = sorted({r["fecha"] for r in rows})
    if cierre not in fechas:
        cierre = fechas[-1]
    starts = _window_start(fechas, cierre)

    by_date = {}
    for r in rows:
        by_date.setdefault(r["fecha"], {})[r["key"]] = r
    cur = by_date[cierre]

    # atributos (bucket/gestora) desde el dia de cierre
    attr = {k: dict(bucket=bucket(r["clasif"], r["plazo"], r["moneda"]),
                    gerente=r["gerente"] or "?") for k, r in cur.items()}

    def flujo_entre(f0, f1):
        """flujo por fondo entre 2 fechas (Delta cuotapartes * vcp_f1/1000)."""
        m0, m1 = by_date.get(f0, {}), by_date.get(f1, {})
        out = {}
        for k, r1 in m1.items():
            r0 = m0.get(k)
            if not r0:
                continue
            c0, c1, v1 = r0["cuotapartes"], r1["cuotapartes"], r1["vcp"]
            if None not in (c0, c1, v1):
                out[k] = (c1 - c0) * v1 / 1000.0
        return out

    def agg_window(dim, win):
        f0 = starts.get(win)
        if not f0:
            return {}
        fl = flujo_entre(f0, cierre)
        m = {}
        for k, v in fl.items():
            key = attr.get(k, {}).get(dim, "?")
            d = m.setdefault(key, {"in": 0.0, "out": 0.0, "net": 0.0})
            if v >= 0:
                d["in"] += v
            else:
                d["out"] += v
            d["net"] += v
        return m

    res = {"cierre": cierre, "ventanas_inicio": starts, "por_bucket": {}, "por_gestora": {}}
    for w in WINDOWS:
        res["por_bucket"][w] = dict(sorted(agg_window("bucket", w).items(), key=lambda kv: kv[1]["net"], reverse=True))
        res["por_gestora"][w] = dict(sorted(agg_window("gerente", w).items(), key=lambda kv: kv[1]["net"], reverse=True))
    # AUM actual por bucket/gestora
    aum_b, aum_g = {}, {}
    for k, r in cur.items():
        aum_b[attr[k]["bucket"]] = aum_b.get(attr[k]["bucket"], 0) + (r["patrimonio"] or 0)
        aum_g[attr[k]["gerente"]] = aum_g.get(attr[k]["gerente"], 0) + (r["patrimonio"] or 0)
    res["aum_por_bucket"] = dict(sorted(aum_b.items(), key=lambda kv: -kv[1]))
    res["aum_por_gestora"] = dict(sorted(aum_g.items(), key=lambda kv: -kv[1]))
    return res


# ------------------------------- orquestacion ----------------------------------
def bajar_un_dia(fecha_target=None):
    docs = listar_documentos()
    if not docs:
        raise SystemExit("Lista de documentos vacia.")
    if fecha_target:
        cand = [g for (f, g) in docs if f == fecha_target]
        if not cand:
            raise SystemExit("No hay documento para %s en la primera pagina." % fecha_target)
        pres_guid = cand[0]
        fecha_doc = fecha_target
    else:
        fecha_doc, pres_guid = docs[0]
    print("Documento %s (presentacion %s)" % (fecha_doc, pres_guid))
    fg = file_guid(pres_guid)
    if not fg:
        raise SystemExit("No se pudo extraer el guid del archivo.")
    print("  archivo: %s" % fg[0])
    content = descargar_xlsx(fg[1])
    if not content:
        raise SystemExit("No se pudo descargar el xlsx (blob CNV).")
    fecha_iso, fondos = parse_planilla(content)
    print("  parseados %d fondos (fecha dato %s)" % (len(fondos), fecha_iso))
    append_history(fecha_iso or fecha_doc, fondos)
    return fecha_iso or fecha_doc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="AAAA-MM-DD especifico (default: ultimo)")
    ap.add_argument("--backfill", type=int, default=0, help="baja los ultimos N documentos para sembrar historia")
    args = ap.parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)

    if args.backfill:
        docs = listar_documentos()
        vistas = set()
        for (fecha_doc, pres_guid) in docs[:args.backfill]:
            if fecha_doc in vistas:
                continue
            vistas.add(fecha_doc)
            try:
                fg = file_guid(pres_guid)
                if not fg:
                    continue
                content = descargar_xlsx(fg[1])
                if not content:
                    continue
                fecha_iso, fondos = parse_planilla(content)
                append_history(fecha_iso or fecha_doc, fondos)
                time.sleep(1.0)
            except Exception as e:
                print("  ! backfill %s: %s" % (fecha_doc, str(e)[:80]), file=sys.stderr)
        cierre = sorted(vistas)[-1]
    else:
        cierre = bajar_un_dia(args.date)

    res = flujos(cierre)
    payload = {"generado": dt.datetime.utcnow().isoformat() + "Z", "flujos": res}
    with open(os.path.join(OUT_DIR, "flujos_latest.json"), "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=1)
    print("OK -> data/flujos_latest.json (cierre %s)" % (res["cierre"] if res else "?"))
    if res:
        b = res["por_bucket"]["1D"]
        tot = sum(v["net"] for v in b.values())
        print("   Flujo neto 1D (industria): %+.0f mill ARS" % (tot / 1e6))


if __name__ == "__main__":
    main()
