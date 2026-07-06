#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sincronizador diario de FLUJOS de FCI desde la API publica de CAFCI.

Corre SERVER-SIDE (GitHub Action / cron). El sandbox de Cowork y el navegador NO
pueden pegarle a la API (sin internet / CORS); un servidor si.

CLOUDFLARE: CAFCI suele bloquear IPs de nube (403) por 'huella' TLS. Por eso usamos
curl_cffi con impersonate=chrome (imita a un navegador de verdad). Si no esta
instalado, cae a requests normal (y probablemente el 403 vuelva -> ver el log).

Pipeline:
  1) universo(): fondos + clases (categoria=tipoRenta, gestora=gerente,
     plazo=diasLiquidacion -> 0 = Money Market/T+0, 1 = T+1).
  2) fichas():  por cada clase, ficha diaria (patrimonio + VCP + fecha).
  3) history.csv: se anexa el snapshot del dia.
  4) flujos():  flujo neto por clase = patrimonio_t - patrimonio_{t-1}*(vcp_t/vcp_{t-1}).
  5) data/flujos_latest.json (lo que consume el reporte) + snapshot del dia.
"""
import os
import sys
import csv
import json
import time
import argparse
import datetime as dt
from concurrent.futures import ThreadPoolExecutor, as_completed

# --- cliente HTTP: preferimos curl_cffi (evita el bloqueo de Cloudflare) ---
try:
    from curl_cffi import requests as _rq
    _IMPERSONATE = "chrome"
    _CLIENT = "curl_cffi (impersonate=chrome)"
except Exception:
    import requests as _rq
    _IMPERSONATE = None
    _CLIENT = "requests (sin impersonate; si CAFCI da 403, instalar curl_cffi)"

HOSTS = ["https://api.cafci.org.ar", "https://api.pub.cafci.org.ar"]
WORKING_HOST = None  # se fija en universo()

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es-AR,es;q=0.9,en;q=0.8",
    "Referer": "https://www.cafci.org.ar/",
    "Origin": "https://www.cafci.org.ar",
}

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
HISTORY = os.path.join(OUT_DIR, "history.csv")
HIST_COLS = ["fecha", "fondoId", "claseId", "categoria", "gestora", "plazo",
             "moneda", "patrimonio", "vcp"]


# ----------------------------- HTTP --------------------------------------------
def _raw_get(url):
    kw = {"timeout": 40, "headers": HEADERS}
    if _IMPERSONATE:
        kw["impersonate"] = _IMPERSONATE
    return _rq.get(url, **kw)


def _get(url, tries=4, backoff=2.0):
    for i in range(tries):
        try:
            r = _raw_get(url)
            if r.status_code == 200:
                try:
                    return r.json()
                except Exception:
                    print("   ! 200 pero body no-JSON: %s" % (r.text or "")[:200], file=sys.stderr)
                    return None
            body = (r.text or "").replace("\n", " ")[:200]
            print("   ! HTTP %s :: %s :: %s" % (r.status_code, url, body), file=sys.stderr)
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(backoff * (i + 1)); continue
            return None
        except Exception as e:
            print("   ! excepcion %s (%s)" % (str(e)[:140], url), file=sys.stderr)
            time.sleep(backoff * (i + 1))
    return None


def _num(x):
    if x in (None, "", "-"):
        return None
    try:
        s = str(x)
        return float(s.replace(".", "").replace(",", ".")) if "," in s else float(s)
    except (ValueError, TypeError):
        return None


# ----------------------------- 1) universo -------------------------------------
def universo():
    global WORKING_HOST
    print("Cliente HTTP: %s" % _CLIENT)
    print("Bajando universo de fondos/clases...")
    inc_attr = ("/fondo?estado=1&include=entidad;depositaria,entidad;gerente,tipoRenta,"
                "moneda,horizonte,duration,tipo_fondo&limit=0&order=fondo.nombre")
    inc_cls = "/fondo?estado=1&include=clase_fondo,entidad;gerente&limit=0"

    dattr = dcls = None
    for host in HOSTS:
        print("  probando host %s ..." % host)
        dattr = _get(host + inc_attr)
        dcls = _get(host + inc_cls) if dattr else None
        if dattr and dcls:
            WORKING_HOST = host
            break
    if not dattr or not dcls:
        raise SystemExit("No se pudo bajar el universo (ver los HTTP de arriba: 403=bloqueo, 503=caida).")
    print("  host OK: %s" % WORKING_HOST)

    attr = {}
    for f in dattr.get("data", []):
        attr[f["id"]] = dict(
            fondo=f.get("nombre"),
            gestora=(f.get("gerente") or {}).get("nombreCorto"),
            categoria=(f.get("tipoRenta") or {}).get("nombre"),
            horizonte=(f.get("horizonte") or {}).get("nombre"),
            duration=(f.get("duration") or {}).get("nombre"),
            tipoFondo=(f.get("tipoFondo") or {}).get("nombre"),
            moneda=(f.get("moneda") or {}).get("codigoCafci"),
            plazo=f.get("diasLiquidacion"),
        )
    uni = {}
    for f in dcls.get("data", []):
        fid = f["id"]
        b = attr.get(fid, {})
        for c in f.get("clase_fondos", []):
            uni[(fid, c["id"])] = dict(b, fondoId=fid, claseId=c["id"],
                                       clase=c.get("nombre"), ticker=c.get("tickerBloomberg"))
    print("  -> %d clases en %d fondos" % (len(uni), len(attr)))
    return uni


# ----------------------------- 2) fichas ---------------------------------------
def _ficha(fid, cid):
    d = _get("%s/fondo/%s/clase/%s/ficha" % (WORKING_HOST, fid, cid))
    if not d or "data" not in d:
        return None
    diaria = ((d["data"].get("info") or {}).get("diaria")) or {}
    act = diaria.get("actual") or {}
    return dict(fecha=diaria.get("referenceDay"),
                patrimonio=_num(act.get("patrimonio")),
                vcp=_num(act.get("vcpUnitario")))


def fichas(uni, workers=8):
    print("Bajando fichas diarias (%d clases)..." % len(uni))
    out = {}
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_ficha, k[0], k[1]): k for k in uni}
        done = 0
        for fu in as_completed(futs):
            k = futs[fu]
            v = fu.result()
            if v and v.get("patrimonio") is not None and v.get("vcp"):
                out[k] = v
            done += 1
            if done % 300 == 0:
                print("   %d/%d" % (done, len(uni)))
    print("  -> %d fichas con dato" % len(out))
    return out


# ----------------------------- 3) history --------------------------------------
def _iso(fecha_ddmmyyyy):
    return dt.datetime.strptime(fecha_ddmmyyyy, "%d/%m/%Y").strftime("%Y-%m-%d")


def append_history(uni, fis):
    os.makedirs(OUT_DIR, exist_ok=True)
    fecha_iso = None
    rows = []
    for k, fi in fis.items():
        u = uni[k]
        try:
            fecha_iso = _iso(fi["fecha"])
        except Exception:
            continue
        rows.append({"fecha": fecha_iso, "fondoId": k[0], "claseId": k[1],
                     "categoria": u.get("categoria"), "gestora": u.get("gestora"),
                     "plazo": u.get("plazo"), "moneda": u.get("moneda"),
                     "patrimonio": fi["patrimonio"], "vcp": fi["vcp"]})
    if not rows:
        raise SystemExit("Sin filas para guardar (fichas vacias).")
    write_header = not os.path.exists(HISTORY)
    with open(HISTORY, "a", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=HIST_COLS)
        if write_header:
            w.writeheader()
        for r in rows:
            w.writerow(r)
    print("  -> history += %d filas (fecha %s)" % (len(rows), fecha_iso))
    return fecha_iso


def load_history():
    if not os.path.exists(HISTORY):
        return []
    with open(HISTORY, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    for r in rows:
        r["patrimonio"] = _num(r["patrimonio"])
        r["vcp"] = _num(r["vcp"])
    return rows


# ----------------------------- 4) buckets + flujos -----------------------------
def bucket(cat, plazo, moneda):
    cat = (cat or "").upper()
    usd = (moneda or "").upper() in ("USD", "DOL", "U$S")
    try:
        p = int(plazo)
    except (TypeError, ValueError):
        p = None
    if "MERCADO DE DINERO" in cat or "MONEY" in cat:
        return "Money Market USD" if usd else "Money Market ARS (T+0)"
    if "RENTA FIJA" in cat:
        if usd:
            return "Renta Fija USD"
        return "Renta Fija ARS T+1" if p == 1 else "Renta Fija ARS T+%s" % (p if p is not None else "?")
    if "RENTA VARIABLE" in cat:
        return "Renta Variable"
    if "RENTA MIXTA" in cat:
        return "Renta Mixta"
    return cat.title() or "Otros"


def flujos(fecha_iso):
    rows = load_history()
    if not rows:
        return None
    fechas = sorted({r["fecha"] for r in rows})
    idx = fechas.index(fecha_iso) if fecha_iso in fechas else len(fechas) - 1
    prev = fechas[idx - 1] if idx > 0 else None
    cur = fechas[idx]
    if not prev:
        print("  (solo hay 1 dia de history: los flujos arrancan manana)")

    def by_key(f):
        return {(r["fondoId"], r["claseId"]): r for r in rows if r["fecha"] == f}
    cur_m, prev_m = by_key(cur), (by_key(prev) if prev else {})

    per = []
    for k, r in cur_m.items():
        flujo = None
        if prev and k in prev_m:
            p0, v0 = prev_m[k]["patrimonio"], prev_m[k]["vcp"]
            pt, vt = r["patrimonio"], r["vcp"]
            if None not in (p0, pt, v0, vt) and v0:
                flujo = pt - p0 * (vt / v0)
        per.append(dict(gestora=r["gestora"], categoria=r["categoria"],
                        bucket=bucket(r["categoria"], r["plazo"], r["moneda"]),
                        patrimonio=r["patrimonio"], flujo=flujo))

    def agg(dim):
        m = {}
        for c in per:
            d = m.setdefault(c[dim] or "?", {"flujo": 0.0, "aum": 0.0, "n": 0})
            if c["flujo"] is not None:
                d["flujo"] += c["flujo"]
            if c["patrimonio"] is not None:
                d["aum"] += c["patrimonio"]
            d["n"] += 1
        return dict(sorted(m.items(), key=lambda kv: -kv[1]["aum"]))

    return {"fecha": cur, "fecha_previa": prev,
            "total_aum": sum((c["patrimonio"] or 0) for c in per),
            "total_flujo_1d": sum((c["flujo"] or 0) for c in per),
            "por_bucket": agg("bucket"), "por_gestora": agg("gestora"),
            "por_categoria": agg("categoria")}


# ----------------------------- main --------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--discover", action="store_true",
                    help="Baja universo + 3 fichas crudas y las muestra (para validar).")
    args = ap.parse_args()
    os.makedirs(OUT_DIR, exist_ok=True)

    uni = universo()
    if args.discover:
        for k in list(uni)[:3]:
            d = _get("%s/fondo/%s/clase/%s/ficha" % (WORKING_HOST, k[0], k[1]))
            print("\n== ficha %s ==\n%s" % (k, json.dumps(d, ensure_ascii=False)[:1200] if d else "None"))
        return

    fis = fichas(uni)
    fecha_iso = append_history(uni, fis)
    res = flujos(fecha_iso)
    payload = {"generado": dt.datetime.utcnow().isoformat() + "Z", "flujos": res}
    for p in (os.path.join(OUT_DIR, "flujos_latest.json"),
              os.path.join(OUT_DIR, "flujos_%s.json" % fecha_iso.replace("-", ""))):
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=1)
    print("OK -> data/flujos_latest.json")
    if res:
        print("   AUM total: %.0f  |  Flujo neto 1D: %+.0f" % (res["total_aum"], res["total_flujo_1d"]))


if __name__ == "__main__":
    main()
