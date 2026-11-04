# -*- coding: utf-8 -*-
"""
sync_monitor_mon.py — promedios MENSUALES para la solapa Monetaria del Monitor Argentina (Argento).

La solapa Monetaria embebe `const MON = {months, M, d150, d44, fecha, src, ...}`: M son promedios
mensuales de series diarias del BCRA desde 2019-01 (los factores, sumas mensuales). Hasta el
04-nov-2026 esos promedios se refrescaban a mano y quedaban viejos. Este script los recalcula con
la API v4 y deja macro/monitor_mon.json con la forma EXACTA del parche que Argento escribe en
`bloques/MON` de la base del artifact:
  {"meta": {...}, "patch": {"months": [...], "M": {"15": [...], ...}, "d150": [[f, v]...],
                            "d44": [[f, v]...], "fecha": "AAAA-MM-DD", "src": "..."}}

Incluye las series del bloque F (crédito al sector privado, 04-nov-2026): préstamos en pesos por
línea (110-117), depósitos privados en pesos (100), depósitos y préstamos en dólares (107, 108,
118-125) y la tasa de préstamos personales (14).

Uso:  python sync_monitor_mon.py            # produce macro/monitor_mon.json
      python sync_monitor_mon.py --check    # además imprime el último mes de cada serie
"""
from __future__ import annotations
import os, json, argparse, time
from datetime import date, datetime, timezone
from sync_bcra import serie, OUT

DESDE = "2019-01-01"
# ids con PROMEDIO mensual de saldos/tasas diarios
PROM = ["1", "7", "8", "12", "13", "14", "15", "16", "19", "44", "85", "86", "87", "91", "94", "95", "96",
        "100", "104", "106", "107", "108", "109", "110", "111", "112", "113", "114", "115", "116", "117",
        "118", "119", "120", "121", "122", "123", "124", "125", "150", "152", "156", "158", "1221"]
# ids con SUMA mensual (flujos diarios: factores de la base monetaria y compras de divisas)
SUMA = ["46", "47", "48", "49", "50", "51", "52", "53", "55", "58", "59", "78"]
SRC = ("BCRA, API de estadísticas v4.0 (promedios mensuales de series diarias; factores y compras: suma "
       "mensual). Refrescado por el Action bcra-daily de cafci-flujos (sync_monitor_mon.py) el {hoy} con "
       "datos al {ult}. Colchón: series.xlsm, hoja INSTRUMENTOS DEL BCRA, columna Otros (nota 13).")


def meses(desde, hasta):
    y, m = int(desde[:4]), int(desde[5:7]); out = []
    while (y, m) <= (int(hasta[:4]), int(hasta[5:7])):
        out.append("%04d-%02d" % (y, m)); m += 1
        if m == 13: y, m = y + 1, 1
    return out


def mensual(d, ms, modo):
    acc = {}
    for f, v in d.items():
        if v is None: continue
        acc.setdefault(f[:7], []).append(float(v))
    out = []
    for m in ms:
        xs = acc.get(m)
        if not xs: out.append(None); continue
        s = sum(xs)
        out.append(round(s if modo == "suma" else s / len(xs), 4))
    return out


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    hoy = date.today().isoformat()
    raw, fallos = {}, []
    for idv in PROM + SUMA:
        try:
            raw[idv] = serie(int(idv), desde=DESDE, hasta=hoy)
        except Exception as e:
            fallos.append(idv); print("serie %s: fallo (%r)" % (idv, e))
        time.sleep(0.2)
    # último mes = el del último dato de la base monetaria (15); fecha = último día común de 15/44/150
    ult15 = max(raw["15"]) if raw.get("15") else hoy
    ms = meses(DESDE, ult15)
    M = {}
    for idv in PROM:
        if idv in raw: M[idv] = mensual(raw[idv], ms, "prom")
    for idv in SUMA:
        if idv in raw: M[idv] = mensual(raw[idv], ms, "suma")
    d150 = [[f, raw["150"][f]] for f in sorted(raw.get("150", {})) if f >= "2025-09-01" and raw["150"][f] is not None]
    d44 = [[f, raw["44"][f]] for f in sorted(raw.get("44", {})) if f >= "2025-01-01" and raw["44"][f] is not None]
    comunes = set(raw.get("15", {})) & set(raw.get("44", {})) & set(raw.get("150", {}))
    fecha = max(comunes) if comunes else ult15
    patch = {"months": ms, "M": M, "d150": d150, "d44": d44, "fecha": fecha, "src": SRC.format(hoy=hoy, ult=fecha)}
    obj = {"meta": {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "desde": DESDE,
                    "ult_mes": ms[-1], "fecha": fecha, "fallos": fallos, "by": "sync_monitor_mon.py (GitHub Actions)",
                    "suma": SUMA, "nota": "patch = bloques/MON del Monitor Argentina; arrays alineados con months"},
           "patch": patch}
    p = os.path.join(OUT, "monitor_mon.json")
    json.dump(obj, open(p, "w", encoding="utf-8"), ensure_ascii=False)
    print("monitor_mon.json ->", p, "|", len(M), "series,", len(ms), "meses, fecha", fecha, "| fallos:", fallos or "ninguno")
    if a.check:
        for k in sorted(M, key=int):
            v = M[k]; i = max((j for j, x in enumerate(v) if x is not None), default=-1)
            print("  %-5s ult=%s %s" % (k, ms[i] if i >= 0 else None, v[i] if i >= 0 else None))


if __name__ == "__main__":
    main()
