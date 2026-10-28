# -*- coding: utf-8 -*-
"""
sync_monitor.py — series CRUDAS para el artifact "Monitor Argentina" (agente Argento).

Baja de la API v4 del BCRA las 25 series que el Monitor embebe (ids de SEED) y el índice IPC
nacional de datos.gob.ar, y deja macro/monitor_series.json con la forma que la página lee de su
base de datos (colección "series", un doc por id):
  {"meta": {"updatedAt": ISO, "desde": "2025-09-01", "fuente": ...},
   "series": {"15": {"n": "Base monetaria", "d": [["2025-09-01", 42598268], ...]}, ..., "ipcidx": {...}}}

Corre en GitHub Actions (tiene red). Argento lo baja de raw.githubusercontent.com y lo vuelca
en la base del artifact con ArtifactData, sin que ningún número pase por el modelo.

Uso:  python sync_monitor.py            # produce macro/monitor_series.json
      python sync_monitor.py --check    # imprime último dato por serie
"""
from __future__ import annotations
import os, json, argparse, urllib.request, time
from datetime import date, datetime, timezone
from sync_bcra import serie, OUT, HDR

DESDE = "2025-09-01"
DESDE_MENSUAL = "2024-01-01"
# id -> nombre tal como lo muestra el Monitor (mismo texto que SEED)
IDS = {
    "15": "Base monetaria",
    "16": "Circulación monetaria",
    "19": "Depósitos de las entidades financieras en cuenta corriente en el BCRA",
    "91": "Depósitos en pesos de los sectores público y privados no financieros (incluye cedros)",
    "109": "M2",
    "25": "Variación interanual del promedio móvil de 30 días del M2 privado.",
    "44": "Tasa de interes TAMAR de bancos privados",
    "7": "Tasa de interés BADLAR de bancos privados",
    "8": "Tasa de interés TM20 de bancos privados",
    "150": "Tasa de interés por operaciones de pases entre terceros a 1 día",
    "156": "LEBAC y NOBAC , LEGAR y LEMIN",
    "1221": "Liquidez en pesos de las entidades financieras",
    "30": "Coeficiente de estabilización de referencia (base 2.2.02=1)",
    "1": "Reservas internacionales",
    "5": "Tipo de cambio mayorista de referencia",
    "4": "Tipo de cambio minorista (promedio vendedor)",
    "78": "Variación de reservas internacionales por compra de divisas",
    "27": "Inflación mensual.",
    "28": "Inflación interanual.",
    "85": "Cuentas corrientes de los sectores público y privado no financieros",
    "86": "Cajas de ahorro de los sectores público y privado no financieros",
    "87": "Depósitos a plazo no ajustables por CER/UVAS de los sectores público y privado no financieros",
    "88": "Depósitos a plazo ajustables por CER/UVAS de los sectores público y privado no financieros",
    "89": "Otros depósitos de los sectores público y privado no financieros",
    "47": "Efecto monetario de las compras netas de divisas al sector privado y otros",
}
MENSUALES = {"27", "28"}
IPC_URL = ("https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26"
           "&start_date=2024-01-01&limit=1000&format=json")
IPC_NOMBRE = "IPC nivel general nacional, índice base dic-2016 (INDEC vía datos.gob.ar 148.3_INIVELNAL_DICI_M_26)"


def _ipcidx():
    req = urllib.request.Request(IPC_URL, headers=HDR)
    with urllib.request.urlopen(req, timeout=60) as r:
        j = json.loads(r.read().decode("utf-8"))
    out = []
    for f, v in j.get("data", []):
        if v is None:
            continue
        # datos.gob.ar da el primer día del mes; el Monitor guarda fin de mes (como la API BCRA)
        y, m = int(f[:4]), int(f[5:7])
        import calendar
        out.append(["%04d-%02d-%02d" % (y, m, calendar.monthrange(y, m)[1]), float(v)])
    return out


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    series = {}
    for idv, nombre in IDS.items():
        desde = DESDE_MENSUAL if idv in MENSUALES else DESDE
        d = serie(int(idv), desde=desde)
        pts = [[f, (float(v) if v is not None else None)] for f, v in sorted(d.items()) if v is not None]
        series[idv] = {"n": nombre, "d": pts}
        time.sleep(0.3)
    try:
        series["ipcidx"] = {"n": IPC_NOMBRE, "d": _ipcidx()}
    except Exception as e:
        print("ipcidx: fallo datos.gob.ar (%s); se omite" % e)
    obj = {"meta": {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"),
                    "desde": DESDE, "fuente": "BCRA API v4 (monetarias) + datos.gob.ar 148.3 (IPC índice)",
                    "by": "sync_monitor.py (GitHub Actions)"},
           "series": series}
    p = os.path.join(OUT, "monitor_series.json")
    json.dump(obj, open(p, "w", encoding="utf-8"), ensure_ascii=False)
    print("monitor_series.json ->", p, "|", len(series), "series")
    if a.check:
        for k, s in series.items():
            print("  %-7s %-40s n=%-4d ult=%s" % (k, s["n"][:40], len(s["d"]), s["d"][-1] if s["d"] else None))


if __name__ == "__main__":
    main()
