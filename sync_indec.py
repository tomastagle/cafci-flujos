# -*- coding: utf-8 -*-
"""
sync_indec.py — series mensuales/trimestrales de datos.gob.ar para el Monitor Argentina (agente Argento).

Enumera, con el catálogo oficial del portal (data.json de sspm), TODAS las series de las distribuciones que usa
el Monitor (recaudación ARCA 142.3, SPN base caja 379.9, ICA 74.3 y 76.3, balance de pagos 160.2, 182.1,
IPC 148.3, EMAE 143.3) y baja sus valores con la API de series. Deja macro/monitor_macro.json:
  {"meta": {"updatedAt": ISO, "distribuciones": {...}},
   "series": {"142.3_...": {"t": título, "u": unidades, "f": frecuencia, "dist": "142.3", "d": [["2026-09-01", v], ...]}}}
Argento lee ese archivo desde raw.githubusercontent.com y actualiza los bloques del Monitor sin tocar el HTML.

Uso:  python sync_indec.py            # produce macro/monitor_macro.json
      python sync_indec.py --check    # imprime último dato por distribución
"""
from __future__ import annotations
import os, io, json, time, argparse, urllib.request, urllib.parse
from datetime import datetime, timezone

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "macro")
os.makedirs(OUT, exist_ok=True)
HDR = {"User-Agent": "Mozilla/5.0 (monitor-argentina-sync)"}
CATALOGO = "https://datos.gob.ar/catalog/sspm/data.json"
API = "https://apis.datos.gob.ar/series/api/series/"
# distribución -> (desde, nota)
DISTS = {
    "142.3": ("2023-01-01", "Recaudación tributaria mensual (ARCA / Hacienda DNIAF)"),
    "379.9": ("2023-01-01", "Sector Público Nacional base caja, metodología 2017 (Hacienda)"),
    "74.3": ("2023-01-01", "ICA: exportaciones e importaciones mensuales (INDEC)"),
    "76.3": ("2023-01-01", "ICA: índices de valor, precio y cantidad (INDEC)"),
    "160.2": ("2016-01-01", "Balance de pagos trimestral MBP6 (INDEC)"),
    "182.1": ("2023-01-01", "Cuenta capital y financiera (INDEC)"),
    "148.3": ("2017-01-01", "IPC nacional (INDEC)"),
    "143.3": ("2023-01-01", "EMAE (INDEC)"),
}
CHUNK = 25


def get_json(url, tries=4):
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=HDR)
            with urllib.request.urlopen(req, timeout=90) as r:
                return json.loads(r.read().decode("utf-8"))
        except Exception as e:
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def enumerar():
    """{dist: [{"id","title","units","description"}]} a partir del data.json de sspm."""
    cat = get_json(CATALOGO)
    out = {d: [] for d in DISTS}
    for ds in cat.get("dataset", []):
        for dist in ds.get("distribution", []):
            did = str(dist.get("identifier", ""))
            if did in out:
                for f in dist.get("field", []):
                    fid = f.get("id")
                    if not fid or fid.startswith("indice_tiempo"):
                        continue
                    out[did].append({"id": fid, "title": f.get("title", ""), "units": f.get("units", ""),
                                     "description": f.get("description", ""),
                                     "freq": dist.get("accrualPeriodicity", ds.get("accrualPeriodicity", ""))})
    return out


def bajar(ids, desde):
    """{id: [[fecha, valor], ...]} con la API de series (de a CHUNK ids, limit 5000)."""
    vals = {i: [] for i in ids}
    for k in range(0, len(ids), CHUNK):
        sub = ids[k:k + CHUNK]
        url = API + "?" + urllib.parse.urlencode({"ids": ",".join(sub), "start_date": desde, "limit": 5000, "format": "json"})
        j = get_json(url)
        for row in j.get("data", []):
            f = row[0]
            for i, v in zip(sub, row[1:]):
                if v is not None:
                    vals[i].append([f, float(v)])
        time.sleep(0.5)
    return vals


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    enum = enumerar()
    series, meta_d = {}, {}
    for dist, (desde, nota) in DISTS.items():
        campos = enum.get(dist, [])
        if not campos:
            print("%s: sin series en el catálogo (¿cambió el identificador?)" % dist); meta_d[dist] = {"n": 0, "nota": nota}; continue
        ids = [c["id"] for c in campos]
        try:
            vals = bajar(ids, desde)
        except Exception as e:
            print("%s: fallo bajando (%s)" % (dist, e)); meta_d[dist] = {"n": 0, "nota": nota, "error": str(e)}; continue
        ult = ""
        for c in campos:
            d = vals.get(c["id"], [])
            if not d:
                continue
            series[c["id"]] = {"t": c["title"], "u": c["units"], "f": c["freq"], "dist": dist, "d": d}
            ult = max(ult, d[-1][0])
        meta_d[dist] = {"n": sum(1 for c in campos if vals.get(c["id"])), "ult": ult, "nota": nota}
        print("%s: %d series, último %s" % (dist, meta_d[dist]["n"], ult))
    obj = {"meta": {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "fuente": "datos.gob.ar (sspm) series API",
                    "by": "sync_indec.py (GitHub Actions)", "distribuciones": meta_d}, "series": series}
    p = os.path.join(OUT, "monitor_macro.json")
    json.dump(obj, io.open(p, "w", encoding="utf-8"), ensure_ascii=False)
    print("monitor_macro.json ->", p, "|", len(series), "series, %.1f MB" % (os.path.getsize(p) / 1e6))


if __name__ == "__main__":
    main()
