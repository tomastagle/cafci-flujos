# -*- coding: utf-8 -*-
"""
sync_ica.py — Intercambio comercial argentino (ICA) directo del INDEC, sin esperar a datos.gob.ar.

Por qué: sync_indec.py toma el ICA de datos.gob.ar (distribuciones 74.3 y 76.3), que lo republica con unas semanas de
rezago: el INDEC publicó octubre el 19-nov-2026 y datos.gob.ar seguía en septiembre. Este script corre DESPUÉS de
sync_indec.py y pisa en macro/monitor_macro.json los meses que el INDEC ya tiene, con dos cuadros de la página del ICA
(www.indec.gob.ar/Nivel4/Tema/3/2/40, sin token):
  expo_grandes_rubros_<aaaa>_<aaaa>.xls   exportaciones mensuales por grandes rubros (total, PP, MOA, MOI, CyE)
  impo_uso_economico_<aaaa>_<aaaa>.xls    importaciones mensuales por usos económicos (total, BK, BI, CyL, PyA, BC, VA, resto)
Cada archivo trae el año en curso y el anterior, en millones de dólares, con las revisiones del INDEC. El nombre cambia
con el año: siempre se llega por el link de la página (si la página falla, se prueba el nombre armado con el año).

No se usa la API del INDEC para esto: su microdato de exportaciones manda a un rubro "confidencial" (secreto
estadístico) unos USD 1.000 MM por mes que el cuadro oficial sí reparte entre PP, MOA, MOI y CyE.

Qué escribe (mismas series y títulos que datos.gob.ar, así Argento y la ficha macro no cambian nada):
  - agrega los meses nuevos y corrige los ya cargados cuando el INDEC los revisó (quedan anotados en meta.ica_directo);
  - calcula las derivadas: saldo comercial, BK + piezas, BI + combustibles, consumo + vehículos + resto.

Control (si falla, NO se escribe ninguna serie y queda el aviso en meta.ica_directo.fallos):
  - en cada mes, la suma de los rubros y de los usos da el total del cuadro;
  - los meses que están en las dos fuentes tienen que coincidir en su gran mayoría (las diferencias son revisiones
    recientes del INDEC; si difiere mucho, cambió el formato);
  - no puede faltar un mes entre lo cargado y lo nuevo.

Uso:  python sync_ica.py [--check]
"""
from __future__ import annotations
import os, io, re, json, time, argparse, unicodedata, urllib.request
from datetime import datetime, timezone

BASE = os.path.dirname(os.path.abspath(__file__))
ARCH = os.path.join(BASE, "macro", "monitor_macro.json")
HDR = {"User-Agent": "Mozilla/5.0 (monitor-argentina-sync)"}
SITIO = "https://www.indec.gob.ar"
PAGINA = SITIO + "/Nivel4/Tema/3/2/40"
FTP = SITIO + "/ftp/cuadros/economia/"
MES3 = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
TOL_SUMA = 0.15      # millones de USD: "los totales por suma pueden no coincidir por redondeos"
TOL_IGUAL = 0.02     # millones de USD: por debajo de esto el dato es el mismo que el de datos.gob.ar
MIN_COMUNES = 100    # pares (serie, mes) en común para dar por bueno el control
MIN_IGUALES = 0.80   # proporción de esos pares que tiene que coincidir

# (regex sobre el encabezado normalizado, títulos de las series 74.3 / 76.3 que alimenta)
EXPO = [
    (r"^total", ["ica_expo_totales"]),
    (r"^productos primarios", ["ica_exportacion_productos_primarios"]),
    (r"origen agropecuario", ["ica_exportacion_manufacturas_origen_agropecuario"]),
    (r"origen industrial", ["ica_exportacion_manufacturas_origen_industrial"]),
    (r"^combustibles? y energia", ["ica_exportacion_combustible_energia"]),
]
IMPO = [
    (r"^total", ["ica_importaciones_totales", "ica_total_general"]),
    (r"^bienes de capital", ["ica_importaciones_bienes_capital", "ica_bienes_capital"]),
    (r"^bienes intermedios", ["ica_importaciones_bienes_intermedios", "ica_bienes_intermedios"]),
    (r"^combustibles y lubricantes", ["ica_importaciones_combustibles_lubricantes", "ica_combustible"]),
    (r"^piezas y accesorios", ["ica_importaciones_piezas_accesorios_bienes_capital", "ica_piezas_accesorios_bienes_capital"]),
    (r"^bienes de consumo", ["ica_importaciones_bienes_consumo", "ica_bienes_consumo"]),
    (r"^vehiculos", ["ica_importaciones_vehiculos_automotores_pasajeros", "ica_vehiculos_automotores_pasajeros"]),
    (r"^resto", ["ica_importaciones_resto", "ica_resto"]),
]
# derivadas: título -> [(signo, título base), ...]
DERIVADAS = {
    "ica_saldo_comercial": [(1, "ica_expo_totales"), (-1, "ica_importaciones_totales")],
    "ica_bienes_capital_partes_piezas": [(1, "ica_importaciones_bienes_capital"), (1, "ica_importaciones_piezas_accesorios_bienes_capital")],
    "ica_bienes_intermedios_combustibles_lubricantes": [(1, "ica_importaciones_bienes_intermedios"), (1, "ica_importaciones_combustibles_lubricantes")],
    "ica_importaciones_bs_consumo_vehiculos_automotor_pasajeros": [(1, "ica_importaciones_bienes_consumo"),
                                                                   (1, "ica_importaciones_vehiculos_automotores_pasajeros"), (1, "ica_importaciones_resto")],
}
DISTS = ("74.3", "76.3")
LOG = []


def log(m):
    LOG.append(str(m)); print(m)


def get(url, tries=3):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=90) as r:
                return r.read()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(4 * (i + 1))


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"\s+", " ", s).strip()


def links():
    """URL de los dos cuadros: por el link de la página del ICA; si no, por el nombre armado con el año."""
    out = {}
    try:
        html = get(PAGINA).decode("utf-8", "replace")
        for clave, pat in (("expo", r"expo_grandes_rubros_\d{4}_\d{4}\.xlsx?"), ("impo", r"impo_uso_economico_\d{4}_\d{4}\.xlsx?")):
            m = sorted(set(re.findall(pat, html)))
            if m:
                out[clave] = FTP + m[-1]
    except Exception as e:
        log("pagina del ICA: fallo (%r), pruebo con el nombre por año" % e)
    a = datetime.now(timezone.utc).year
    for clave, nom in (("expo", "expo_grandes_rubros_%d_%d.xls"), ("impo", "impo_uso_economico_%d_%d.xls")):
        if clave not in out:
            for y in (a, a - 1):
                try:
                    u = FTP + nom % (y - 1, y); get(u, tries=1); out[clave] = u; break
                except Exception:
                    continue
    return out


def leer(url, grupos):
    """{título: {"AAAA-MM-01": valor}} de un cuadro mensual del INDEC (grupos en columnas, año en curso y anterior)."""
    import xlrd
    sh = xlrd.open_workbook(file_contents=get(url)).sheet_by_index(0)
    filas = [sh.row_values(i) for i in range(sh.nrows)]
    # fila de encabezados: la que tiene "Total" y al menos otros dos grupos
    cab = None
    for i, r in enumerate(filas[:15]):
        n = [norm(c) for c in r]
        if sum(1 for c in n if c and any(re.search(p, c) for p, _ in grupos)) >= 3:
            cab = i; break
    if cab is None:
        raise RuntimeError("no encontre la fila de encabezados")
    anios = None
    for r in filas[cab:cab + 4]:
        if sum(1 for c in r if re.search(r"a.o \d{4}", norm(c))) >= 4:
            anios = r; break
    if anios is None:
        raise RuntimeError("no encontre la fila de años")
    cols = {}   # título -> [(columna, año)]
    for c, v in enumerate(filas[cab]):
        n = norm(v)
        if not n:
            continue
        for pat, tits in grupos:
            if re.search(pat, n):
                ys = []
                for cc in (c, c + 1):
                    m = re.search(r"a.o (\d{4})", norm(anios[cc])) if cc < len(anios) else None
                    if m:
                        ys.append((cc, int(m.group(1))))
                if len(ys) != 2:
                    raise RuntimeError("grupo '%s': no encontre sus dos años" % n)
                for t in tits:
                    cols[t] = ys
                break
    falt = [t for _, tits in grupos for t in tits if t not in cols]
    if falt:
        raise RuntimeError("faltan columnas: %s" % ", ".join(falt))
    out = {t: {} for t in cols}
    for r in filas[cab + 1:]:
        et = next((norm(c) for c in r[:3] if isinstance(c, str) and c.strip()), "")
        if et not in MES3:
            continue
        mes = MES3.index(et) + 1
        for t, ys in cols.items():
            for c, y in ys:
                v = r[c]
                if isinstance(v, (int, float)):
                    out[t]["%04d-%02d-01" % (y, mes)] = float(v)
    return out


def _sig(p, n):
    a, m = int(p[:4]), int(p[5:7]) + n
    return "%04d-%02d-01" % (a + (m - 1) // 12, (m - 1) % 12 + 1)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    obj = json.load(io.open(ARCH, encoding="utf-8"))
    ser = {s["t"]: s for s in obj.get("series", {}).values() if s.get("dist") in DISTS}
    info = {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "fuente": "INDEC, cuadros mensuales del ICA (expo por grandes rubros, impo por usos economicos)",
            "pagina": PAGINA, "archivos": {}, "agregado": [], "revisado": [], "validacion": None, "fallos": []}
    try:
        if "ica_expo_totales" not in ser:
            raise RuntimeError("monitor_macro.json no trae la distribucion 74.3: no hay contra que controlar")
        ult = ser["ica_expo_totales"]["d"][-1][0]                               # último mes de datos.gob.ar
        lk = links()
        if set(lk) != {"expo", "impo"}:
            raise RuntimeError("no encontre los cuadros del ICA en el INDEC (%s)" % (", ".join(sorted(lk)) or "ninguno"))
        info["archivos"] = lk
        vals = leer(lk["expo"], EXPO); vals.update(leer(lk["impo"], IMPO))
        meses = sorted(set(vals["ica_expo_totales"]) & set(vals["ica_importaciones_totales"]))
        if not meses:
            raise RuntimeError("los cuadros no traen ningun mes con exportaciones e importaciones")
        # 1) la suma de los rubros y de los usos da el total
        for p in meses:
            for tot, grupos in (("ica_expo_totales", EXPO), ("ica_importaciones_totales", IMPO)):
                partes = [vals[tits[0]].get(p) for _, tits in grupos[1:]]
                if None in partes or abs(sum(partes) - vals[tot][p]) > TOL_SUMA:
                    raise RuntimeError("%s: las partes no suman %s (cambio el formato del cuadro)" % (p, tot))
        for t, comp in DERIVADAS.items():
            vals[t] = {p: sum(sg * vals[b][p] for sg, b in comp) for p in meses}
        # 2) contra datos.gob.ar en los meses comunes
        n = ig = 0; peor = 0.0; revis = {}
        for t, dv in vals.items():
            if t not in ser:
                continue
            ref = dict(ser[t]["d"])
            for p in meses:
                if p in ref and p in dv:
                    n += 1; dif = abs(dv[p] - ref[p]); peor = max(peor, dif)
                    if dif <= TOL_IGUAL:
                        ig += 1
                    else:
                        revis.setdefault(p, []).append(t)
        ok = n >= MIN_COMUNES and ig >= MIN_IGUALES * n
        info["validacion"] = {"comunes": n, "iguales": ig, "max_dif": round(peor, 3), "ok": ok, "datos_gob_ar_hasta": ult, "indec_hasta": meses[-1]}
        if not ok:
            raise RuntimeError("control contra datos.gob.ar: %d de %d datos en comun coinciden (minimo %d y %d%%)" % (ig, n, MIN_COMUNES, MIN_IGUALES * 100))
        log("INDEC hasta %s | datos.gob.ar hasta %s | control: %d de %d datos en comun iguales (dif max %.3f)" % (meses[-1], ult, ig, n, peor))
        nuevos = [p for p in meses if p > ult]
        if nuevos and nuevos != [_sig(ult, i + 1) for i in range(len(nuevos))]:
            raise RuntimeError("falta un mes entre datos.gob.ar (%s) y el INDEC (%s)" % (ult, ", ".join(nuevos)))
        # 3) escribir: meses nuevos y meses revisados
        for t, dv in vals.items():
            if t not in ser:
                continue
            d = dict(ser[t]["d"])
            for p in meses:
                if p in dv and (p not in d or abs(dv[p] - d[p]) > TOL_IGUAL):
                    d[p] = round(dv[p], 8)
            ser[t]["d"] = [[p, d[p]] for p in sorted(d)]
        for p in nuevos:
            info["agregado"].append({"periodo": p, "expo": round(vals["ica_expo_totales"][p], 2), "impo": round(vals["ica_importaciones_totales"][p], 2),
                                     "saldo": round(vals["ica_saldo_comercial"][p], 2)})
            log("%s agregado: expo %.1f, impo %.1f, saldo %.1f" % (p, vals["ica_expo_totales"][p], vals["ica_importaciones_totales"][p], vals["ica_saldo_comercial"][p]))
        for p in sorted(revis):
            info["revisado"].append({"periodo": p, "series": sorted(revis[p])})
            log("%s revisado por el INDEC: %s" % (p, ", ".join(sorted(revis[p]))))
        if not nuevos:
            log("sin meses nuevos: datos.gob.ar ya tiene lo que publico el INDEC")
        for dist in DISTS:
            d = obj["meta"].setdefault("distribuciones", {}).setdefault(dist, {})
            if nuevos:
                d["ult"] = nuevos[-1]; d["ult_datos_gob_ar"] = ult
                d["nota"] = (d.get("nota", "") + " | ultimo mes directo del INDEC").strip(" |")
    except Exception as e:
        info["fallos"].append(repr(e)); info["agregado"] = []; info["revisado"] = []; log("ICA DIRECTO: %r" % e)
        obj = json.load(io.open(ARCH, encoding="utf-8"))                         # ante cualquier falla, el archivo queda como estaba
    obj.setdefault("meta", {})["ica_directo"] = info
    obj["meta"].setdefault("log", []).extend("ica_directo: " + l for l in LOG)
    json.dump(obj, io.open(ARCH, "w", encoding="utf-8"), ensure_ascii=False)
    if a.check:
        for t in ("ica_expo_totales", "ica_importaciones_totales", "ica_saldo_comercial"):
            s = {x["t"]: x for x in obj["series"].values() if x.get("dist") in DISTS}.get(t)
            print("  %-28s %s" % (t, s["d"][-2:] if s else "FALTA"))


if __name__ == "__main__":
    main()
