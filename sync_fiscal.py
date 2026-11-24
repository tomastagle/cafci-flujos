# -*- coding: utf-8 -*-
"""
sync_fiscal.py — resultado fiscal del SPN base caja directo de la Secretaría de Hacienda, sin el rezago de datos.gob.ar.

datos.gob.ar (distribución 379.9) republica el esquema Ahorro - Inversión - Financiamiento de Hacienda con unas semanas
de demora. Hacienda cuelga el Excel "Base Caja" en el comunicado del día del dato (alrededor del 16 de cada mes):
  https://www.argentina.gob.ar/economia/sechacienda  ->  /noticias/en-<mes>-el-sector-publico-nacional-...  ->  <mes>_<aa>.xlsx
Este script corre DESPUÉS de sync_indec.py y completa en macro/monitor_macro.json las series 379.9 con los meses que
Hacienda ya publicó y datos.gob.ar todavía no cargó. Nada cambia para quien lee el archivo (Argento, ficha macro).

Control en cada corrida: el último mes que tienen las dos fuentes se baja también de Hacienda y se compara serie por
serie contra datos.gob.ar. Si no coincide (cambió el formato del Excel), no se escribe nada y queda el aviso en
meta.fiscal_directo.fallos.

Uso:  python sync_fiscal.py            # completa macro/monitor_macro.json
      python sync_fiscal.py --check    # además imprime lo que hizo
"""
from __future__ import annotations
import os, io, re, json, time, argparse, unicodedata, urllib.request
from datetime import datetime, timezone

BASE = os.path.dirname(os.path.abspath(__file__))
ARCH = os.path.join(BASE, "macro", "monitor_macro.json")
HDR = {"User-Agent": "Mozilla/5.0 (monitor-argentina-sync)"}
SITIO = "https://www.argentina.gob.ar"
PORTADAS = [SITIO + "/economia/sechacienda", SITIO + "/economia/sechacienda/infoestadistica"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
TOL = 0.6          # millones de pesos: los Excel vienen a un decimal
DIST = "379.9"

# Filas del Excel, EN ORDEN de aparición, y el título de la serie 379.9 que les corresponde (None = no hay serie).
# El orden resuelve las etiquetas repetidas ("Otras", "Provincias y CABA", "Otros gastos").
FILAS = [
    (r"^ingresos corrientes", "ing_corr_2017"),
    (r"^ingresos impositivos", "ing_corr_ing_tributarios_2017"),
    (r"^aportes y contrib", "ing_corr_aportes_contrib_seg_soc_2017"),
    (r"^ingresos no impositivos", "ing_corr_ing_no_tributarios_2017"),
    (r"^ventas de bs", "ing_corr_ventas_bs_ser_adm_pub_2017"),
    (r"^ingresos de operacion", "ing_corr_ing_operacion_2017"),
    (r"^rentas de la propiedad netas", "ing_corr_ren_prop_ren_prop_netas_2017"),
    (r"^transferencias corrientes", "ing_corr_transf_corr_2017"),
    (r"^otros ingresos", "ing_corr_ots_ing_2017"),
    (r"^gastos corrientes", "gtos_corr_2017"),
    (r"^remuneraciones", "gtos_corr_gtos_cons_operacion_remuneraciones_2017"),
    (r"^bienes y servicios", "gtos_corr_gtos_cons_operacion_bienes_servicios_2017"),
    (r"^otros gastos$", "gtos_corr_gtos_cons_operacion_ots_gtos_2017"),
    (r"^intereses y otras rentas", "_int_y_rentas"),
    (r"^intereses netos", "gtos_corr_int_ot_ren_prop_int_netos_2017"),
    (r"^otras rentas", "gtos_corr_int_ot_ren_prop_ot_ren_2017"),
    (r"^prestaciones de la seguridad social", "gtos_corr_prestaciones_seg_soc_2017"),
    (r"^otros gastos corrientes", "gtos_corr_ots_gtos_corr_2017"),
    (r"^al sector privado", "gtos_corr_transf_corr_sec_priv_total_sec_priv_2017"),
    (r"^provincias y caba", "gtos_corr_transf_corr_sect_pub_caba_total_caba_2017"),
    (r"^universidades", "gtos_corr_transf_corr_sec_pub_uni_2017"),
    (r"^otras$", "gtos_corr_transf_corr_sec_pub_ot_2017"),
    (r"^al sector externo", "gtos_corr_transf_corr_sec_ext_2017"),
    (r"^otros gastos$", "gtos_corr_ots_gtos_2017"),
    (r"^deficit operativo empresas", "gtos_corr_deficit_operativo_empresas_pub_2017"),
    (r"^result\.? ?econ", "resultado_economico_ahorro_desahorro_2017"),
    (r"^recursos de capital", "recursos_cap_total_recursos_cap_2017"),
    (r"^gastos de capital", "gtos_cap_2017"),
    (r"^inversion real directa", "gtos_cap_inversion_real_directa_2017"),
    (r"^a provincias y caba", "gtos_cap_transf_cap_prov_caba_total_prov_caba_2017"),
    (r"^otras$", "gtos_cap_transf_cap_ot_2017"),
    (r"^inversion financiera", "gtos_cap_resto_2017"),
    (r"^ingresos antes de figurat", "ing_antes_figurativos_2017"),
    (r"^gastos antes de figurat", "gtos_antes_figurativos_2017"),
    (r"^result\.? ?financ\.? ?antes de figurat", "resultado_fin_antes_figurativos_2017"),
    (r"^del tesoro nacional", "contrib_figurat_tesoro_nacional_2017"),
    (r"^de recursos afectados", "contrib_figurat_recursos_afectados_2017"),
    (r"^de organismos descentralizados", "contrib_figurat_org_descentralizados_2017"),
    (r"^de instituciones de seguridad social", "contrib_figurat_inst_seg_soc_2017"),
    (r"^de pami", "contrib_figurat_pami_fdos_fidu_ots_2017"),
    (r"^gastos figurativos", "gtos_figurativos_2017"),
    (r"^ingresos despues de figurat", "ing_despues_figurativos_2017"),
    (r"^gastos primarios despues de figurat", "gtos_primarios_despues_figurativos_2017"),
    (r"^gastos despues de figurat", "gtos_despues_figurativos_2017"),
    (r"^resultado primario", "superavit_primario_2017"),
    (r"^resultado financiero", "resultado_fin_2017"),
    (r"^rentas percibidas del bcra", "ing_corr_ren_prop_ren_percidas_bcra_2017"),
    (r"^intereses pagados intra", "gtos_corr_int_ot_ren_prop_int_pag_intra_stor_pub_2017"),
]
# Series que el esquema no trae como fila y se arman con otras (misma cuenta que hace datos.gob.ar)
DERIVADAS = {
    "ing_prario_antes_figurativos_2017": lambda v: v["ing_antes_figurativos_2017"] - v["ing_corr_ren_prop_ren_prop_netas_2017"],
    "gtos_primario_antes_figurativos_2017": lambda v: v["gtos_antes_figurativos_2017"] - v["_int_y_rentas"],
    "resultados_primarios_sin_rentas": lambda v: (v["ing_antes_figurativos_2017"] - v["ing_corr_ren_prop_ren_prop_netas_2017"])
                                                 - (v["gtos_antes_figurativos_2017"] - v["_int_y_rentas"]),
}
MINIMAS = ["superavit_primario_2017", "resultado_fin_2017", "ing_antes_figurativos_2017", "gtos_antes_figurativos_2017",
           "ing_despues_figurativos_2017", "gtos_primarios_despues_figurativos_2017"]

LOG = []
def log(m):
    LOG.append(str(m)); print(m)


def get(url, tries=3):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=60) as r:
                return r.read()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"\(\d+\)", " ", s)               # llamadas a notas: (1), (4)
    s = re.sub(r"^[\s.\-]+", "", s)               # viñetas: "- ", ". ", ".. "
    return re.sub(r"\s+", " ", s).strip()


def comunicados():
    """{'2026-10-01': url del comunicado} con lo que haya en las portadas de Hacienda."""
    out = {}
    for p in PORTADAS:
        try:
            html = get(p).decode("utf-8", "ignore")
        except Exception as e:
            log("portada %s: fallo (%r)" % (p, e)); continue
        for path in set(re.findall(r'href="(?:https://www\.argentina\.gob\.ar)?(/noticias/en-[a-z]+-el-sector-publico-nacional[^"#?]*)"', html)):
            mes = path.split("/noticias/en-")[1].split("-")[0]
            if mes in MESES:
                out.setdefault(mes, SITIO + path)
    return out


def excel_base_caja(url_comunicado):
    """Baja los .xlsx del comunicado y devuelve (periodo 'AAAA-MM-01', {titulo: valor}, url del Excel) del esquema base caja."""
    import openpyxl
    html = get(url_comunicado).decode("utf-8", "ignore")
    links = []
    for u in re.findall(r'href="([^"]+\.xlsx)"', html, flags=re.I):
        u = u if u.startswith("http") else SITIO + u
        if u.startswith(SITIO + "/sites/default/files/") and u not in links:
            links.append(u)
    links.sort(key=lambda u: "imig" in u.lower())          # el de base caja primero; el IMIG tiene otro formato
    for u in links:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(get(u)), data_only=True, read_only=True)
        except Exception as e:
            log("excel %s: fallo (%r)" % (u, e)); continue
        ws = wb.worksheets[0]
        filas = [list(r) for r in ws.iter_rows(values_only=True)]
        cab = " ".join(norm(c) for r in filas[:6] for c in r if isinstance(c, str))
        m = re.search(r"sector publico base caja - (%s) (\d{4})" % "|".join(MESES), cab)
        if not m or "acumulado" in cab:
            continue
        periodo = "%s-%02d-01" % (m.group(2), MESES.index(m.group(1)) + 1)
        # columna TOTAL: la que dice "T O T A L" en el encabezado
        col = None
        for r in filas[:12]:
            for j, c in enumerate(r):
                if isinstance(c, str) and re.sub(r"\s+", "", c).upper() == "TOTAL":
                    col = j                                  # queda la más a la derecha (total del SPN)
        if col is None:
            log("excel %s: sin columna TOTAL" % u); continue
        vals, k = {}, 0
        for r in filas:
            if k >= len(FILAS):
                break
            et = max((c for c in r[:col] if isinstance(c, str)), key=len, default=None)   # la etiqueta, no el "III)"
            if et is None:
                continue
            if re.search(FILAS[k][0], norm(et)):
                v = r[col] if col < len(r) else None
                if isinstance(v, (int, float)):
                    vals[FILAS[k][1]] = float(v)
                k += 1
        if k < len(FILAS):
            log("excel %s: no encontre la fila %r (formato distinto)" % (u, FILAS[k][0])); continue
        for t, f in DERIVADAS.items():
            try:
                vals[t] = f(vals)
            except KeyError:
                pass
        return periodo, {t: v for t, v in vals.items() if not t.startswith("_")}, u
    return None, {}, None


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); ap.parse_args()
    obj = json.load(io.open(ARCH, encoding="utf-8"))
    ser = {s["t"]: s for s in obj.get("series", {}).values() if s.get("dist") == DIST}
    info = {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "fuente": "Secretaria de Hacienda, Excel base caja del comunicado mensual",
            "agregado": [], "validacion": None, "fallos": []}
    try:
        if "superavit_primario_2017" not in ser:
            raise RuntimeError("monitor_macro.json no trae la distribucion 379.9: no hay contra que controlar")
        ult = ser["superavit_primario_2017"]["d"][-1][0]                     # último mes de datos.gob.ar
        com = comunicados()
        if not com:
            raise RuntimeError("no encontre comunicados en las portadas de Hacienda")
        # período de cada comunicado: el mes del título, del año que corresponda (se confirma con el Excel)
        leidos = {}
        for mes, url in com.items():
            try:
                p, vals, xl = excel_base_caja(url)
            except Exception as e:
                log("comunicado %s: fallo (%r)" % (url, e)); continue
            if p:
                leidos[p] = (vals, xl, url)
        log("Hacienda: %s | datos.gob.ar 379.9 hasta %s" % (", ".join(sorted(leidos)) or "nada", ult))
        if not leidos:
            raise RuntimeError("ningun Excel de base caja se pudo leer (cambio el formato o el sitio)")
        nuevos = sorted(p for p in leidos if p > ult)
        if nuevos:
            # control: el último mes que está en las dos fuentes tiene que coincidir serie por serie
            comunes = sorted(p for p in leidos if p <= ult)
            if not comunes:
                raise RuntimeError("no hay un mes en comun para controlar el formato del Excel (datos.gob.ar hasta %s)" % ult)
            pc = comunes[-1]; vc = leidos[pc][0]; n = 0; peor = 0.0; malas = []
            for t, v in vc.items():
                ref = dict(ser[t]["d"]).get(pc) if t in ser else None
                if ref is None:
                    continue
                n += 1; dif = abs(v - ref); peor = max(peor, dif)
                if dif > TOL:
                    malas.append("%s: Hacienda %.1f, datos.gob.ar %.1f" % (t, v, ref))
            info["validacion"] = {"periodo": pc, "series": n, "max_dif": round(peor, 3), "ok": not malas and n >= 40}
            if malas or n < 40:
                raise RuntimeError("control contra datos.gob.ar en %s: %d series, no coinciden %s" % (pc, n, "; ".join(malas[:6]) or "(pocas series)"))
            log("control %s: %d series iguales a datos.gob.ar (dif max %.3f)" % (pc, n, peor))
            if nuevos != [_sig(ult, i + 1) for i in range(len(nuevos))]:
                raise RuntimeError("falta un mes entre datos.gob.ar (%s) y Hacienda (%s)" % (ult, ", ".join(nuevos)))
            for p in nuevos:
                vals, xl, url = leidos[p]
                falt = [t for t in MINIMAS if t not in vals]
                if falt:
                    raise RuntimeError("%s: faltan %s" % (p, ", ".join(falt)))
                # identidades del esquema, por si una fila quedó corrida
                if abs(vals["ing_despues_figurativos_2017"] - vals["gtos_primarios_despues_figurativos_2017"] - vals["superavit_primario_2017"]) > TOL \
                        or abs(vals["ing_antes_figurativos_2017"] - vals["gtos_antes_figurativos_2017"] - vals["resultado_fin_2017"]) > TOL:
                    raise RuntimeError("%s: no cierran las identidades del esquema" % p)
                k = 0
                for t, v in vals.items():
                    if t in ser and ser[t]["d"] and ser[t]["d"][-1][0] < p:
                        ser[t]["d"].append([p, round(v, 1)]); k += 1
                info["agregado"].append({"periodo": p, "series": k, "excel": xl, "comunicado": url,
                                         "primario": vals["superavit_primario_2017"], "financiero": vals["resultado_fin_2017"]})
                log("%s agregado: %d series | primario %.1f, financiero %.1f | %s" % (p, k, vals["superavit_primario_2017"], vals["resultado_fin_2017"], xl))
            d = obj["meta"].setdefault("distribuciones", {}).setdefault(DIST, {})
            d["ult"] = nuevos[-1]; d["ult_datos_gob_ar"] = ult
            d["nota"] = (d.get("nota", "") + " | ultimo mes directo de Hacienda").strip(" |")
        else:
            log("sin meses nuevos: datos.gob.ar ya tiene lo que publico Hacienda")
    except Exception as e:
        info["fallos"].append(repr(e)); info["agregado"] = []; log("FISCAL DIRECTO: %r" % e)
        obj = json.load(io.open(ARCH, encoding="utf-8"))                       # ante cualquier falla, el archivo queda como estaba
    obj.setdefault("meta", {})["fiscal_directo"] = info
    obj["meta"].setdefault("log", []).extend("fiscal_directo: " + l for l in LOG)
    json.dump(obj, io.open(ARCH, "w", encoding="utf-8"), ensure_ascii=False)


def _sig(p, n):
    a, m = int(p[:4]), int(p[5:7]) + n
    return "%04d-%02d-01" % (a + (m - 1) // 12, (m - 1) % 12 + 1)


if __name__ == "__main__":
    main()
