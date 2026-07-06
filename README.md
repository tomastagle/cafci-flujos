# CAFCI flujos — sync diario (fuente propia)

Baja todos los días la data de FCI de la **API pública de CAFCI** (`api.cafci.org.ar`),
calcula el **flujo neto por fondo** (suscripciones − rescates) y lo agrega por **tipo de fondo**
(Money Market, Renta Fija T+1, USD…) y por **gestora**. Corre solo en **GitHub Actions** y deja
el resultado en `data/flujos_latest.json`, que después consume el Monitor de Renta Fija.

## Por qué GitHub Actions (server-side)
La API de CAFCI se puede llamar con un `GET` simple **desde un servidor** (confirmado), pero
**no** desde un navegador (CORS) ni desde el sandbox de Cowork (sin salida a internet). Por eso
la ingesta vive acá, corre sola, y publica un JSON que el reporte lee.

## Cómo se calcula el flujo
`flujo_neto(t) = patrimonio(t) − patrimonio(t−1) × ( vcp(t) / vcp(t−1) )`
Equivale a Δ(cuotapartes) × VCP: el cambio de patrimonio que **no** se explica por el rendimiento
del día es dinero que entró o salió. Solo usa patrimonio + valor de cuotaparte, ambos de la ficha diaria.

## Endpoints usados (CAFCI)
- Universo:  `GET /fondo?estado=1&include=…gerente,tipoRenta,moneda,horizonte,duration,tipo_fondo&limit=0`
- Clases:    `GET /fondo?estado=1&include=clase_fondo,entidad;gerente&limit=0`
- Ficha:     `GET /fondo/{fondoId}/clase/{claseId}/ficha`  → `data.info.diaria.actual.{patrimonio,vcpUnitario}`
`tipoRenta` = categoría · `gerente` = gestora · `diasLiquidacion` = plazo (0 = MM/T+0, 1 = T+1).

## Setup (una vez)
1. Crear un repo nuevo en GitHub (privado está bien) y subir **el contenido de esta carpeta** a la raíz
   (`sync_cafci.py`, `requirements.txt`, `.github/workflows/cafci-daily.yml`, este README).
2. En el repo → **Settings → Actions → General → Workflow permissions** = *Read and write*.
3. Ir a la pestaña **Actions**, elegir "CAFCI flujos diarios" y correrlo a mano (**Run workflow**)
   para validar. Después corre solo cada noche de día hábil (cron `30 1 * * 2-6`, ~22:30 ART; ajustable).
4. Primera corrida: se crea `data/history.csv` con el snapshot del día. **Los flujos aparecen desde el
   segundo día** (necesitan el día anterior). El histórico se va acumulando en el repo.

### Validar la estructura antes (opcional)
`python sync_cafci.py --discover` baja el universo + 3 fichas crudas y las imprime, para chequear
que los campos siguen igual. Útil si CAFCI cambia algo.

## Salidas (carpeta `data/`)
- `history.csv` — una fila por clase y día (fecha, fondoId, claseId, categoría, gestora, plazo, moneda, patrimonio, vcp). Es la base para calcular flujos y ventanas.
- `flujos_latest.json` — agregado del último día: flujo neto y AUM **por bucket**, **por gestora** y **por categoría**.
- `flujos_YYYYMMDD.json` — snapshot del día.

## Cómo lo consume el reporte
El Monitor leerá `flujos_latest.json` (o el `history.csv` para ventanas semana/mes/YTD) y dibujará
la sección "Flujos de FCI". Se sincroniza a la carpeta del proyecto con un `git pull` del repo, o
apuntando el generador al repo clonado.

## Pendientes / a refinar
- **Buckets finos**: hoy el corte es categoría + plazo + moneda. Para separar **CER / Lecap / Dollar-Linked**
  dentro de Renta Fija hay que sumar heurística por nombre/benchmark (o un mapa manual). Ver `bucket()`.
- **Backfill histórico**: para arrancar con ventanas largas sin esperar, se puede sembrar `history.csv`
  con series de patrimonio históricas (endpoint de rendimiento/histórico de CAFCI). Queda para una v2.
- **Ventanas semana/mes/YTD**: el cálculo por bucket ya está para 1D; extender a más ventanas leyendo `history.csv`.
- **Validación**: contrastar contra CNV (valores diarios de cuotaparte / patrimonio, fuente oficial).
