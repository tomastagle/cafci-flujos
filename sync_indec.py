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
import os, io, json, time, argparse, urllib.request, urllib.parse, urllib.error
from datetime import datetime, timezone

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "macro")
os.makedirs(OUT, exist_ok=True)
HDR = {"User-Agent": "Mozilla/5.0 (monitor-argentina-sync)"}
CATALOGOS = ["https://infra.datos.gob.ar/catalog/sspm/data.json", "https://datos.gob.ar/catalog/sspm/data.json"]
SEARCH = "https://apis.datos.gob.ar/series/api/search/"
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


def enumerar_busqueda():
    """Plan B: API de búsqueda, filtrando por prefijo de distribución (q = texto de la nota)."""
    out = {d: [] for d in DISTS}
    for dist, (_, nota) in DISTS.items():
        q = nota.split("(")[0].strip()
        for offset in (0, 1000, 2000):
            try:
                j = get_json(SEARCH + "?" + urllib.parse.urlencode({"q": q, "limit": 1000, "offset": offset, "catalog_id": "sspm"}))
            except Exception as e:
                log("busqueda %s: fallo (%r)" % (dist, e)); break
            rows = j.get("data", [])
            for r in rows:
                f = r.get("field", {}); fid = f.get("id", "")
                if fid.startswith(dist + "_"):
                    out[dist].append({"id": fid, "title": f.get("title", "") or f.get("description", ""), "units": f.get("units", ""),
                                      "description": f.get("description", ""), "freq": f.get("frequency", "")})
            if len(rows) < 1000:
                break
        # sin duplicados
        seen = set(); out[dist] = [c for c in out[dist] if not (c["id"] in seen or seen.add(c["id"]))]
        log("busqueda %s: %d series" % (dist, len(out[dist])))
    return out


def enumerar():
    """{dist: [{"id","title","units","description"}]} a partir del data.json de sspm (o la busqueda, plan B)."""
    cat = None
    for u in CATALOGOS:
        try:
            cat = get_json(u); log("catalogo: " + u); break
        except Exception as e:
            log("catalogo %s: fallo (%r)" % (u, e))
    if cat is None:
        return enumerar_busqueda()
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
    if any(not v for v in out.values()):
        alt = enumerar_busqueda()
        for d in out:
            if not out[d]:
                out[d] = alt.get(d, [])
    return out


MALOS = []
def _bajar_lote(sub, desde, vals):
    """Baja un lote; ante 400 lo parte en mitades hasta aislar el id que la API rechaza (y lo saltea)."""
    url = API + "?" + urllib.parse.urlencode({"ids": ",".join(sub), "start_date": desde, "limit": 5000, "format": "json"})
    try:
        j = get_json(url, tries=2)
    except urllib.error.HTTPError as e:
        if e.code == 400 and len(sub) > 1:
            m = len(sub) // 2
            _bajar_lote(sub[:m], desde, vals); _bajar_lote(sub[m:], desde, vals); return
        MALOS.append(sub[0] if len(sub) == 1 else ",".join(sub)); return
    for row in j.get("data", []):
        f = row[0]
        for i, v in zip(sub, row[1:]):
            if v is not None:
                vals[i].append([f, float(v)])
    time.sleep(0.4)


def bajar(ids, desde):
    """{id: [[fecha, valor], ...]} con la API de series (de a CHUNK ids, limit 5000)."""
    vals = {i: [] for i in ids}
    for k in range(0, len(ids), CHUNK):
        _bajar_lote(ids[k:k + CHUNK], desde, vals)
    return vals


LOG = []
def log(m):
    LOG.append(str(m)); print(m)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    try:
        enum = enumerar()
        log("catalogo OK: " + ", ".join("%s=%d" % (d, len(v)) for d, v in enum.items()))
    except Exception as e:
        log("catalogo: fallo (%r)" % e); enum = {}
    series, meta_d = {}, {}
    for dist, (desde, nota) in DISTS.items():
        campos = enum.get(dist, [])
        if not campos:
            log("%s: sin series en el catálogo (¿cambió el identificador?)" % dist); meta_d[dist] = {"n": 0, "nota": nota}; continue
        ids = [c["id"] for c in campos]
        try:
            vals = bajar(ids, desde)
        except Exception as e:
            log("%s: fallo bajando (%r)" % (dist, e)); meta_d[dist] = {"n": 0, "nota": nota, "error": repr(e)}; continue
        ult = ""
        for c in campos:
            d = vals.get(c["id"], [])
            if not d:
                continue
            series[c["id"]] = {"t": c["title"], "u": c["units"], "f": c["freq"], "dist": dist, "d": d}
            ult = max(ult, d[-1][0])
        meta_d[dist] = {"n": sum(1 for c in campos if vals.get(c["id"])), "ult": ult, "nota": nota}
        log("%s: %d series, último %s" % (dist, meta_d[dist]["n"], ult))
    if MALOS:
        log("ids rechazados por la API (salteados): " + ", ".join(MALOS[:40]) + (" ..." if len(MALOS) > 40 else ""))
    obj = {"meta": {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "fuente": "datos.gob.ar (sspm) series API",
                    "by": "sync_indec.py (GitHub Actions)", "distribuciones": meta_d, "log": LOG}, "series": series}
    p = os.path.join(OUT, "monitor_macro.json")
    json.dump(obj, io.open(p, "w", encoding="utf-8"), ensure_ascii=False)
    print("monitor_macro.json ->", p, "|", len(series), "series, %.1f MB" % (os.path.getsize(p) / 1e6))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # nunca tirar abajo el job: dejar el diagnostico en el JSON
        log("FALLO GENERAL: %r" % e)
        p = os.path.join(OUT, "monitor_macro.json")
        json.dump({"meta": {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "log": LOG, "error": repr(e)}, "series": {}},
                  io.open(p, "w", encoding="utf-8"), ensure_ascii=False)
