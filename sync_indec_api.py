# -*- coding: utf-8 -*-
"""
sync_indec_api.py — series del INDEC desde su API oficial (api.indec.gob.ar, beta, PostgREST) para el
Monitor Argentina (Argento) y la ficha macro del Ministro.

Por qué: datos.gob.ar carga el IPC, el EMAE y el ICA con ~1 mes de rezago y sin los componentes del IPC
(núcleo, regulados, estacionales). La API del INDEC los tiene el día del informe (validado 04-nov-2026:
IPC sep-26 con núcleo 1,7 / regulados 2,3 / estacionales 2,9; EMAE e IPI ago-26; salarios ago-26;
EPH 2T-26; pobreza 1S-26). datos.gob.ar (sync_indec.py) queda como respaldo y para lo que la API no trae
(recaudación, SPN caja, ICA precios/cantidades, balance de pagos, PIB, ISAC).

Token: secret INDEC_TOKEN del repo (JWT personal de Tomás, VENCE A LOS 12 MESES: ver meta.token_vence).
Sin token, el script no toca el archivo anterior y sale sin error (el job sigue).

Deja macro/indec_api.json:
  {"meta": {"updatedAt", "fuente", "token_vence", "datasets": {slug: {"filas", "ult"}}, "log": [...]},
   "series": {"ipc|Nacional|N|v_m": {"t": "IPC Núcleo, Nacional, var. % mensual", "u": "%", "ds": "ipc_nacional",
                                      "d": [["2017-01", 1.6], ...]}, ...}}
Períodos: mensual "AAAA-MM"; trimestral "AAAA-T1..T4"; semestral "AAAA-S1|S2".

Uso:  python sync_indec_api.py            # produce macro/indec_api.json
      python sync_indec_api.py --check    # además imprime el último dato de las series clave
"""
from __future__ import annotations
import os, io, json, time, base64, argparse, urllib.request, urllib.parse, urllib.error
from datetime import datetime, timezone

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "macro")
os.makedirs(OUT, exist_ok=True)
API = "https://api.indec.gob.ar/v1/"
PAGE = 1000  # la API topea en 1000 filas por respuesta
LOG = []


def log(m):
    LOG.append(str(m)); print(m)


def token():
    t = os.environ.get("INDEC_TOKEN", "").strip()
    if not t and os.environ.get("INDEC_TOKEN_FILE"):
        t = json.load(open(os.environ["INDEC_TOKEN_FILE"], encoding="utf-8"))["token"].strip()
    return t


def vence(t):
    try:
        p = t.split(".")[1]; p += "=" * (-len(p) % 4)
        return datetime.fromtimestamp(json.loads(base64.urlsafe_b64decode(p))["exp"], timezone.utc).strftime("%Y-%m-%d")
    except Exception:
        return None


def get(path, tok, tries=4):
    url = API + urllib.parse.quote(path, safe="/?=&.,()*_:-")
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"Authorization": "Bearer " + tok, "User-Agent": "monitor-argentina-sync"})
            with urllib.request.urlopen(req, timeout=90) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (400, 401, 403) or i == tries - 1:
                raise
            time.sleep(3 * (i + 1))
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def todo(slug, filtro, order, tok):
    """Todas las filas (pagina de a 1000 con orden estable por la clave de la tabla)."""
    out, off = [], 0
    while True:
        q = "%s?%s%sorder=%s&limit=%d&offset=%d" % (slug, filtro, "&" if filtro else "", order, PAGE, off)
        d = get(q, tok); out += d
        if len(d) < PAGE:
            return out
        off += PAGE
        time.sleep(0.2)


def per(p):
    p = str(p)
    if len(p) == 6 and p.isdigit():
        return p[:4] + "-" + p[4:]
    if len(p) == 6 and p[4] in "12345" and p[5] in "TS":  # 20262T / 20261S
        return p[:4] + "-" + p[5] + p[4]
    return p


TIPO = {"v_m": "var. % mensual", "v_i_a": "var. % interanual", "v_ia": "var. % interanual", "v_acum": "var. % acumulada del año",
        "Indice": "índice", "Índice": "índice"}

# (slug, filtro, orden, función fila -> (clave, título, unidad) o None, nota)
def _ipc(r):
    t = r["tipo_de_indicador"]
    return ("ipc|%s|%s|%s" % (r["region"], r["codigo"], t),
            "IPC %s, %s, %s" % (r["descripcion"], r["region"], TIPO.get(t, t)), "%" if t != "Indice" else "índice dic-16=100")

def _emae(r):
    t = r["tipo_de_indicador_descripcion"]
    return ("emae|0|" + t, "EMAE nivel general, " + t, "índice 2004=100" if "Índice" in t else "%")

def _emae_sec(r):
    t = r["tipo_de_indicador"]
    return ("emae_sec|%s|%s" % (r["codigo"], "idx" if t.startswith("Índice") else "v_ia"),
            "EMAE %s, %s" % (r["descripcion"], "índice 2004=100" if t.startswith("Índice") else "var. % interanual"), "índice" if t.startswith("Índice") else "%")

def _ipi(r):
    return ("ipi|%s|%s|%s" % (r["codigo_CLaNAE_2004"], r["serie"], r["tipo_de_indicador"]),
            "IPI manufacturero %s, %s, %s" % (r["descripcion"].replace("IPI Manufacturero - ", ""), r["serie"].lower(), r["tipo_de_indicador"].lower()),
            "índice 2016=100" if r["tipo_de_indicador"] == "Índice" else "%")

def _sal(r):
    t = r["tipo_de_indicador"]
    return ("salarios|%s|%s" % (r["codigo"], t), "Índice de salarios %s, %s" % (r["descripcion"], TIPO.get(t, t)), "%" if t != "Indice" else "índice oct-16=100")

def _sipm(r):
    t = r["tipo_de_indicador"]
    return ("sipm|%s|%s" % (r["indicador"], t), "Precios mayoristas %s nivel general, %s" % (r["indicador"], TIPO.get(t, t)), "%" if t != "Indice" else "índice 1993=100")

def _cba(r):
    t = r["tipo_de_indicador"]
    return ("canastas|%s|%s" % (r["codigo"], t), "%s (%s), %s" % (r["descripcion"], r["region"], TIPO.get(t, t if t not in ("CBA", "CBT", "ICE") and not t.startswith("Hogar") else "valor")),
            "%" if t in ("v_m", "v_ia") else ("coeficiente" if t == "ICE" else "$"))

DATASETS = [
    ("ipc_nacional", "region=in.(Nacional,GBA)&nivel=in.(Categoría,División)&periodo=gte.201701",
     "periodo,region,clasificador,nivel,codigo,tipo_de_indicador", _ipc, "IPC por categoría y división (Nacional, GBA)"),
    ("ipc_nacional", "nivel=eq.Nivel General&periodo=gte.201701",
     "periodo,region,tipo_de_indicador", _ipc, "IPC nivel general por región"),
    ("estimador-mensual-de-actividad-economica-nivel-general", "", "periodo,tipo_de_indicador_descripcion", _emae, "EMAE nivel general"),
    ("estimador-mensual-de-actividad-economica-sectores", "periodo=gte.201701", "periodo,codigo,tipo_de_indicador", _emae_sec, "EMAE por sector"),
    ("ipi_manufacturero", "periodo=gte.202101", "periodo,cuadro_web,codigo_CLaNAE_2004,serie,tipo_de_indicador", _ipi, "IPI manufacturero"),
    ("indice_salarios", "", "periodo,codigo,tipo_de_indicador", _sal, "Índice de salarios"),
    ("sipm_precios_mayoristas", "nivel=eq.Nivel general&periodo=gte.201701", "periodo,indicador,tipo_de_indicador", _sipm, "SIPM nivel general"),
    ("valorizacion_canastas_cba_cbt", "", "periodo,codigo,tipo_de_indicador", _cba, "CBA y CBT"),
]
# trimestrales/semestrales: columnas anchas -> una serie por (área, columna)
ANCHAS = [
    ("mercado_trabajo_eph", "periodo", ["tasa_de_actividad", "tasa_de_empleo", "tasa_de_desocupacion", "tasa_de_subocupacion",
                                         "tasa_de_empleo_formal", "tasa_de_empleo_informal", "tasa_de_ocupacion_demandante"], "EPH mercado de trabajo"),
    ("pobreza_eph", "periodo", ["tasa_pobreza_personas", "tasa_indigencia_personas", "tasa_pobreza_hogares", "tasa_indigencia_hogares"], "Pobreza e indigencia"),
]
CLAVE_CHECK = ["ipc|Nacional|0|v_m", "ipc|Nacional|N|v_m", "ipc|Nacional|R|v_m", "ipc|Nacional|E|v_m", "ipc|Nacional|0|v_ia",
               "emae|0|Var % mensual serie desestacionalizada", "emae|0|Var % interanual serie original",
               "ipi|D|Serie original|Variación interanual", "salarios|0|v_m", "sipm|IPIM|v_m", "canastas|CBT|v_m",
               "eph|T_31_A|tasa_de_desocupacion", "pobreza|T_31_A|tasa_pobreza_personas"]


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--check", action="store_true"); a = ap.parse_args()
    tok = token()
    p = os.path.join(OUT, "indec_api.json")
    if not tok:
        log("sin INDEC_TOKEN: no se toca indec_api.json (cargar el secret en el repo)"); return
    tv = vence(tok)
    if tv:
        dias = (datetime.strptime(tv, "%Y-%m-%d").replace(tzinfo=timezone.utc) - datetime.now(timezone.utc)).days
        log("token vence %s (%d días)%s" % (tv, dias, " ** RENOVAR **" if dias < 30 else ""))
    series, meta_ds, fallos = {}, {}, []
    for slug, filtro, order, fn, nota in DATASETS:
        try:
            rows = todo(slug, filtro, order, tok)
        except Exception as e:
            log("%s: fallo (%r)" % (nota, e)); fallos.append(nota); continue
        ult = ""
        for r in rows:
            v = r.get("valor")
            if v is None:
                continue
            k, t, u = fn(r)
            s = series.setdefault(k, {"t": t, "u": u, "ds": slug, "d": []})
            s["d"].append([per(r["periodo"]), float(v)]); ult = max(ult, per(r["periodo"]))
        meta_ds[nota] = {"slug": slug, "filas": len(rows), "ult": ult}
        log("%s: %d filas, último %s" % (nota, len(rows), ult))
    for slug, pc, cols, nota in ANCHAS:
        try:
            rows = todo(slug, "", "%s,area_geografica_cod" % pc, tok)
        except Exception as e:
            log("%s: fallo (%r)" % (nota, e)); fallos.append(nota); continue
        ult = ""; pref = "eph" if slug.startswith("mercado") else "pobreza"
        for r in rows:
            for c in cols:
                v = r.get(c)
                if v is None:
                    continue
                area = r.get("area_geografica_cod") or r.get("area_geografica_nom")  # hoy el código viene en _nom ("T_31_A")
                k = "%s|%s|%s" % (pref, area, c)
                s = series.setdefault(k, {"t": "%s, %s" % (c.replace("_", " ").capitalize(), "31 aglomerados" if area == "T_31_A" else area), "u": "%", "ds": slug, "d": []})
                s["d"].append([per(r[pc]), float(v)]); ult = max(ult, per(r[pc]))
        meta_ds[nota] = {"slug": slug, "filas": len(rows), "ult": ult}
        log("%s: %d filas, último %s" % (nota, len(rows), ult))
    for s in series.values():
        s["d"].sort()
    if not series:
        log("ninguna serie: no se pisa indec_api.json"); return
    obj = {"meta": {"updatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ"), "fuente": "INDEC, API oficial v1 (beta)",
                    "by": "sync_indec_api.py (GitHub Actions)", "token_vence": tv, "datasets": meta_ds, "fallos": fallos, "log": LOG},
           "series": series}
    json.dump(obj, io.open(p, "w", encoding="utf-8"), ensure_ascii=False)
    print("indec_api.json ->", p, "|", len(series), "series, %.1f MB" % (os.path.getsize(p) / 1e6))
    if a.check:
        for k in CLAVE_CHECK:
            s = series.get(k)
            print("  %-48s %s" % (k, s["d"][-1] if s else "FALTA"))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("FALLO GENERAL sync_indec_api: %r" % e)  # nunca tirar abajo el job
