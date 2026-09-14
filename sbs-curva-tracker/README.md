# SBS Curva Soberana — Tracker diario

Descarga automáticamente, todos los días hábiles, las curvas cupón cero
publicadas por la SBS (https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/CurvaSoberana/ConsultaHistorica)
y guarda el histórico en `data/*.json`. Un proceso separado en Claude lee este
repo cada día para armar el reporte y las alertas (ver sección final).

Curvas configuradas: `CCPSS` (Soberana Soles), `CCPEDS` (Dólares Globales),
`CBCRS` (BCRP CDBCRP).

## 1. Crear el repositorio

1. En GitHub, crea un repo nuevo (público o privado, da igual), por ejemplo
   `sbs-curva-tracker`.
2. Descomprime el zip que te envié y sube **toda la carpeta** (arrastrar y
   soltar en "Add file → Upload files" conserva la estructura de subcarpetas,
   incluida `.github/workflows/`).
3. Confirma el commit inicial.

## 2. Activar el workflow

GitHub Actions se activa solo al detectar `.github/workflows/daily.yml`. Ve a
la pestaña **Actions** del repo y, si aparece un aviso, dale a "I understand
my workflows, enable them".

## 3. Backfill inicial (sembrar histórico)

Por defecto el job diario solo pide los últimos 5 días. Para tener un
histórico más largo desde el día 1:

1. Ve a **Actions → Actualizar curvas cupón cero SBS → Run workflow**.
2. Deja `tipo_curva` vacío (corre las 3), pon `fecha_inicio` en algo como
   `2026-06-01` y `fecha_fin` vacío.
3. Ejecuta. Revisa que al terminar existan `data/CCPSS.json`,
   `data/CCPEDS.json`, `data/CBCRS.json` con varios registros.

## 4. Verificación

- Cada corrida escribe/actualiza `data/latest_run.json` con el estado
  (`ok`/`error`) de cada curva — esto es lo que el reporte de Claude usa para
  detectar fallas de actualización.
- Si un run fallа, GitHub te manda un correo automático (viene activado por
  defecto en tu cuenta) y además queda un artifact `debug-<curva>` con
  screenshot + HTML de la página en el momento del error, útil para
  diagnosticar sin tener que reproducirlo.

## Nota sobre el WAF de SBS

El portal está protegido por Imperva/Incapsula (cookies `incap_ses_*`). El
scraper usa un navegador real (Playwright) para pasar ese filtro igual que lo
haría una persona, pero **Imperva a veces bloquea rangos de IP de
datacenters** (incluyendo los runners de GitHub Actions). Si ves fallos
repetidos con un error de tipo "Request unsuccessful" o timeout en todos los
runs (revisa el artifact de debug), probablemente sea esto — la alternativa
en ese caso es correr el mismo script desde una máquina con IP residencial
(por ejemplo tu propia laptop, hay un modo listo para eso: avísame y lo
armamos).

## Cómo lo consume el reporte de Claude

Claude lee directamente estos archivos vía:
```
https://raw.githubusercontent.com/<tu-usuario>/<tu-repo>/main/data/CCPSS.json
https://raw.githubusercontent.com/<tu-usuario>/<tu-repo>/main/data/latest_run.json
```
No necesitas hacer nada más aquí — el reporte y las alertas se generan y
actualizan desde el lado de Claude (scheduled task) usando estos datos.
