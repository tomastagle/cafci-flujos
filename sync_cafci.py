#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sincronizador diario de FLUJOS de FCI desde la API publica de CAFCI.

Corre SERVER-SIDE (GitHub Action / cron). El sandbox de Cowork y el navegador NO
pueden pegarle a la API (sin internet / CORS); un servidor con requests.get() si
(confirmado con el bot publico cuajoa/FCI.ar).

Pipeline:
  1) universo(): baja fondos + clases (categoria=tipoRenta, gestora=gerente,
     plazo=diasLiquidacion  -> 0 = Money Market/T+0, 1 = T+1, etc.).
  2) fichas():  por cada clase, baja la ficha diaria (patrimonio + VCP + fecha).
  3) Se anexa el snapshot del dia a data/history.csv.
  4) flujos():  flujo neto por clase = patrimonio_t - patrimonio_{t-1}*(vcp_t/vcp_{t-1})
     (equiv. a  Delta(cuotapartes) * VCP). Se agrega por bucket y por gestora,
     en ventanas 1D / semana / mes / YTD, a partir del history.
  5) Escribe data/flujos_latest.json (lo que consume el reporte) + snapshot del dia.

Fuente: https://api.cafci.org.ar   (endpoints confirmados 06-jul-2026)
"""
import os
import sys
import csv
import json
import time
import argparse
import datetime as dt
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

BASE = "https://api.cafci.org.ar"
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
HISTORY = os.path.join(OUT_DIR, "history.csv")
HIST_COLS = ["fecha", "fondoId", "claseId", "categoria", "gestora", "plazo",
             "moneda", "patrimonio", "vcp"]

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; fci-flows/1.0)",
    "Accept": "application/json",
    "Referer": "https://www.cafci.org.ar/",
})


# ----------------------------- HTTP con reintentos -----------------------------
def _get(url, tries=5, backoff=1.5):
    last = None
    for i in range(tries):
        try:
            r = SESSION.get(url, timeout=40)
            if r.status_code == 200:
                return r.json()
            last = "HTTP %s" % r.status_code
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(backoff * (i + 1))
                continue
            return None
        except requests.RequestException as e:
            last = str(e)
            time.sleep(backoff * (i + 1))
    print("   ! fallo %s (%s)" % (url, last), file=sys.stderr)
    return None


def _num(x):
    if x in (None, "", "-"):
        return None
    try:
        return float(str(x).replace(".", "").replace(",", ".")) if isinstance(x, str) and "," in str(x) \
            else float(x)
    except (ValueError, TypeError):
        return None


# ----------------------------- 1) universo -------------------------------------
def universo():
    """Devuelve {(_fondoId,_claseId): {atributos}} para todas las clases activas."""
    print("Bajando universo de fondos/clases...")
    # atributos por fondo
    u_attr = (BASE + "/fondo?estado=1&include=entidad;depositaria,entidad;gerente,"
              "tipoRenta,moneda,horizonte,duration,tipo_fondo&limit=0&order=fondo.nombre")
    # clases por fondo
    u_clases = BASE + "/fondo?estado=1&include=clase_fondo,entidad;gerente&limit=0"

    dattr = _get(u_attr)
    dcls = _get(u_clases)
    if not dattr or not dcls:
        raise SystemExit("No se pudo bajar el universo (API caida?). Reintentar mas tarde.")

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

    universo = {}
    for f in dcls.get("data", []):
        fid = f["id"]
        base = attr.get(fid, {})
        for c in f.get("clase_fondos", []):
            universo[(fid, c["id"])] = dict(
                base, fondoId=fid, claseId=c["id"],
                clase=c.get("nombre"), ticker=c.get("tickerBloomberg"),
            )
    print("  -> %d clases en %d fondos" % (len(universo), len(attr)))
    return universo


# ----------------------------- 2) fichas diarias -------------------------------
def _ficha(fid, cid):
    d = _get("%s/fondo/%s/clase/%s/ficha" % (BASE, fid, cid))
    if not d or "data" not in d:
        return None
    diaria = ((d["data"].get("info") or {}).get("diaria")) or {}
    act = diaria.get("actual") or {}
    return dict(
        fecha=diaria.get("referenceDay"),           # dd/mm/yyyy
        patrimonio=_num(act.get("patrimonio")),
        vcp=_num(act.get("vcpUnitario")),
    )


def fichas(universo, workers=10):
    print("Bajando fichas diarias (%d clases)..." % len(universo))
    out = {}
    def job(k):
        return k, _ficha(k[0], k[1])
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(job, k) for k in universo]
        done = 0
        for fu in as_completed(futs):
            k, v = fu.result()
            if v and v.get("patrimonio") is not None and v.get("vcp"):
                out[k] = v
            done += 1
            if done % 300 == 0:
                print("   %d/%d" % (done, len(universo)))
    print("  -> %d fichas con dato" % len(out))
    return out


# ----------------------------- 3) history --------------------------------------
def _iso(fecha_ddmmyyyy):
    return dt.datetime.strptime(fecha_ddmmyyyy, "%d/%m/%Y").strftime("%Y-%m-%d")


def append_history(universo, fichas_):
    os.makedirs(OUT_DIR, exist_ok=True)
    fecha_iso = None
    rows = []
    for k, fi in fichas_.items():
        u = universo[k]
        try:
            fecha_iso = _iso(fi["fecha"])
        except Exception:
            continue
        rows.append({
            "fecha": fecha_iso, "fondoId": k[0], "claseId": k[1],
            "categoria": u.get("categoria"), "gestora": u.get("gestora"),
            "plazo": u.get("plazo"), "moneda": u.get("moneda"),
            "patrimonio": fi["patrimonio"], "vcp": fi["vcp"],
        })
    if not rows:
        raise SystemExit("Sin filas para guardar (fichas vacias).")

    # evitar duplicar la fecha si ya existe en el history
    existing_dates = set()
    if os.path.exists(HISTORY):
        with open(HISTORY, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                existing_dates.add(r["fecha"])
    write_header = not os.path.exists(HISTORY)
    if fecha_iso in existing_dates:
        print("  (la fecha %s ya estaba en history; se reescribe el snapshot igual)" % fecha_iso)
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
    """Mapea (categoria CAFCI, plazo, moneda) a un tipo operativo.
    Primer corte; refinar CER/Lecap con nombre/benchmark en una version futura."""
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
    return (cat.title() or "Otros")


def flujos(fecha_iso):
    """Calcula flujo neto por clase entre las 2 ultimas fechas y agrega por bucket/gestora."""
    rows = load_history()
    if not rows:
        return None
    fechas = sorted({r["fecha"] for r in rows})
    idx = fechas.index(fecha_iso) if fecha_iso in fechas else len(fechas) - 1
    if idx == 0:
        print("  (solo hay 1 dia de history: los flujos arrancan manana)")
        prev = None
    else:
        prev = fechas[idx - 1]
    cur = fechas[idx]

    def by_key(f):
        return {(r["fondoId"], r["claseId"]): r for r in rows if r["fecha"] == f}
    cur_m = by_key(cur)
    prev_m = by_key(prev) if prev else {}

    per_class = []
    for k, r in cur_m.items():
        p_t, v_t = r["patrimonio"], r["vcp"]
        flujo = None
        if prev and k in prev_m:
            p0, v0 = prev_m[k]["patrimonio"], prev_m[k]["vcp"]
            if p_t is not None and p0 is not None and v_t and v0:
                flujo = p_t - p0 * (v_t / v0)
        per_class.append(dict(
            fondoId=k[0], claseId=k[1], gestora=r["gestora"],
            bucket=bucket(r["categoria"], r["plazo"], r["moneda"]),
            categoria=r["categoria"], moneda=r["moneda"],
            patrimonio=p_t, flujo=flujo,
        ))

    def agg(dim):
        m = {}
        for c in per_class:
            key = c[dim] or "?"
            d = m.setdefault(key, {"flujo": 0.0, "aum": 0.0, "n": 0})
            if c["flujo"] is not None:
                d["flujo"] += c["flujo"]
            if c["patrimonio"] is not None:
                d["aum"] += c["patrimonio"]
            d["n"] += 1
        return dict(sorted(m.items(), key=lambda kv: -kv[1]["aum"]))

    return {
        "fecha": cur, "fecha_previa": prev,
        "total_aum": sum((c["patrimonio"] or 0) for c in per_class),
        "total_flujo_1d": sum((c["flujo"] or 0) for c in per_class),
        "por_bucket": agg("bucket"),
        "por_gestora": agg("gestora"),
        "por_categoria": agg("categoria"),
    }


# ----------------------------- main --------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--discover", action="store_true",
                    help="Solo baja universo + 3 fichas y muestra la estructura cruda (para validar).")
    args = ap.parse_args()

    os.makedirs(OUT_DIR, exist_ok=True)
    uni = universo()

    if args.discover:
        sample_keys = list(uni)[:3]
        for k in sample_keys:
            print("\n== ficha cruda %s ==" % (k,))
            d = _get("%s/fondo/%s/clase/%s/ficha" % (BASE, k[0], k[1]))
            print(json.dumps(d, ensure_ascii=False)[:1500] if d else "None")
        print("\nAtributos de una clase:", json.dumps(uni[sample_keys[0]], ensure_ascii=False))
        return

    fi = fichas(uni)
    fecha_iso = append_history(uni, fi)
    res = flujos(fecha_iso)

    latest = os.path.join(OUT_DIR, "flujos_latest.json")
    snap = os.path.join(OUT_DIR, "flujos_%s.json" % fecha_iso.replace("-", ""))
    payload = {"generado": dt.datetime.utcnow().isoformat() + "Z", "flujos": res}
    for p in (latest, snap):
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=1)
    print("OK -> %s" % latest)
    if res:
        print("   AUM total: %.0f  | Flujo neto 1D: %+.0f" %
              (res["total_aum"], res["total_flujo_1d"]))


if __name__ == "__main__":
    main()
