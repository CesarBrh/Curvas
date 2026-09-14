#!/usr/bin/env python3
"""
fetch_curva.py — Descarga histórica/diaria de curvas cupón cero desde el
Portal Curva Soberana de la SBS (Perú) y actualiza el histórico local en /data.

Fuente: https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/CurvaSoberana/ConsultaHistorica

El portal es una SPA .NET protegida por Imperva/Incapsula (WAF con cookies de
sesión) — por eso este script usa un navegador real (Playwright/Chromium) en
vez de peticiones HTTP crudas: replica exactamente la interacción manual
(seleccionar curva, fechas, click en "Consultar") y captura el archivo que el
portal descarga como respuesta al POST interno a
ExportarListadoHistoricoCurvaSoberana.

Uso:
    python fetch_curva.py --tipo-curva CCPSS --fecha-inicio 2026-08-14 --fecha-fin 2026-09-13
    python fetch_curva.py --tipo-curva CCPEDS --dias-atras 5   (modo diario normal)

Salida:
    data/<TIPO_CURVA>.json   histórico acumulado, deduplicado por fecha
    data/latest_run.json     metadata de la última corrida (para alertas)
"""
import argparse
import json
import sys
import traceback
from datetime import datetime, timedelta, date
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE_URL = "https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/CurvaSoberana/ConsultaHistorica"
DATA_DIR = Path(__file__).resolve().parent.parent / "data"
DEBUG_DIR = Path(__file__).resolve().parent.parent / "debug"

CURVAS_VALIDAS = {
    "CCPSS": "Curva Soberana Soles",
    "CCPVS": "Curva Soberana Soles VAC",
    "CCPEDS": "Curva Cupón Cero Dólares Globales",
    "CCINFS": "Curva Cupón Cero de Inflación en Soles",
    "CCCLD": "Curva Cupón Cero Libor",
    "CBCRS": "Curva Banco Central de Reserva CDBCRP",
    "CBCRPS": "Curva Banco Central de Reserva CDBCRP NR",
    "CSBCRD": "Curva Cupón Cero Dólares Sintética",
    "CCSDF": "Curva Dólares Corto Plazo",
}


def log(msg: str) -> None:
    print(f"[{datetime.utcnow().isoformat()}Z] {msg}", flush=True)


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--tipo-curva", required=True, choices=sorted(CURVAS_VALIDAS))
    p.add_argument("--fecha-inicio", help="YYYY-MM-DD (default: hoy - dias-atras)")
    p.add_argument("--fecha-fin", help="YYYY-MM-DD (default: hoy)")
    p.add_argument("--dias-atras", type=int, default=5,
                    help="Si no se pasa fecha-inicio, cuántos días naturales atrás consultar (default 5, cubre fines de semana/feriados)")
    p.add_argument("--headful", action="store_true", help="Correr con navegador visible (debug local)")
    return p.parse_args()


def ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def descargar_excel(tipo_curva: str, fecha_inicio: date, fecha_fin: date, headful: bool = False) -> Path:
    """Abre el portal SBS, hace la consulta y devuelve el path al archivo descargado."""
    DEBUG_DIR.mkdir(exist_ok=True)
    download_path = None

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not headful)
        context = browser.new_context(
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"),
            locale="es-PE",
        )
        page = context.new_page()
        try:
            log(f"Navegando a {BASE_URL}")
            page.goto(BASE_URL, wait_until="networkidle", timeout=45000)

            page.wait_for_selector("#cboFiltroTipoCurva", timeout=20000)
            page.select_option("#cboFiltroTipoCurva", value=tipo_curva)

            page.fill("#txtFiltroFechaInicio", "")
            page.fill("#txtFiltroFechaInicio", ddmmyyyy(fecha_inicio))
            page.fill("#txtFiltroFechaFin", "")
            page.fill("#txtFiltroFechaFin", ddmmyyyy(fecha_fin))
            # Cerrar cualquier datepicker abierto haciendo click fuera
            page.click("h5:has-text('Curvas Cupón Cero')", timeout=5000)

            log(f"Consultando {tipo_curva} {fecha_inicio}→{fecha_fin}")
            with page.expect_download(timeout=30000) as download_info:
                page.click("#btnBuscarInformacionHistorica")
            download = download_info.value
            download_path = DEBUG_DIR / f"{tipo_curva}_{fecha_inicio}_{fecha_fin}{Path(download.suggested_filename).suffix or '.xlsx'}"
            download.save_as(download_path)
            log(f"Descargado: {download_path} ({download_path.stat().st_size} bytes)")
        except PWTimeout as e:
            page.screenshot(path=str(DEBUG_DIR / f"timeout_{tipo_curva}.png"))
            (DEBUG_DIR / f"timeout_{tipo_curva}.html").write_text(page.content())
            raise RuntimeError(f"Timeout esperando respuesta de SBS para {tipo_curva}: {e}")
        finally:
            browser.close()

    return download_path


def _dias_a_años(dias: float) -> str:
    """Convierte un plazo en días a una clave de años legible (convención 360:
    90d→0.25, 360d→1, 450d→1.25, etc. — confirmado por los plazos que exporta
    SBS, todos múltiplos exactos de 90 días)."""
    años = dias / 360
    if años == int(años):
        return str(int(años))
    return f"{años:.2f}".rstrip("0").rstrip(".")


def parsear_excel(path: Path) -> list[dict]:
    """
    Convierte el archivo exportado por SBS a una lista de registros:
    [{"fecha": "YYYY-MM-DD", "tasas": {"<plazo_años>": tasa_pct, ...}}, ...]

    El export de SBS viene en formato LARGO (una fila por combinación
    fecha × plazo), con una fila de título antes del encabezado real:

        Fila 0: "Rango de fechas del | 01/06/2026 | al | 14/09/2026"  (metadata)
        Fila 1: "Fecha de Proceso | Tipo de Curva | Plazo (DIAS) | Tasas (%)"  (encabezado real)
        Fila 2+: datos, ej. "01/06/2026 | CCPSS | 90 | 4.03528"

    OJO: la fila de metadata también contiene la palabra "fecha" (en "Rango
    de fechas del"), así que buscar solo esa palabra detecta la fila
    equivocada como encabezado. Por eso exigimos una coincidencia EXACTA con
    "Fecha de Proceso" (no solo "contiene fecha").
    """
    def _cargar(header_row):
        try:
            return pd.read_excel(path, header=header_row)
        except Exception:
            return pd.read_csv(path, sep=None, engine="python", header=header_row)

    crudo = _cargar(None)
    log("Vista cruda del archivo (primeras 15 filas):\n" + crudo.head(15).to_string())
    header_row = None
    for i in range(min(10, len(crudo))):
        valores = [str(v).strip().lower() for v in crudo.iloc[i].tolist()]
        if "fecha de proceso" in valores and any("plazo" in v for v in valores):
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            "No se encontró la fila de encabezado ('Fecha de Proceso' + 'Plazo') "
            f"en las primeras 10 filas del archivo. Primeras filas:\n{crudo.head(10).to_string()}"
        )

    df = _cargar(header_row)
    df.columns = [str(c).strip() for c in df.columns]
    log("Columnas detectadas (fila de encabezado " + str(header_row) + "): " + str(list(df.columns)))

    col_fecha = next((c for c in df.columns if c.lower() == "fecha de proceso"), None)
    col_plazo = next((c for c in df.columns if "plazo" in c.lower()), None)
    col_tasa = next((c for c in df.columns if "tasa" in c.lower()), None)
    if not (col_fecha and col_plazo and col_tasa):
        raise ValueError(
            "No se encontraron las columnas esperadas (fecha/plazo/tasa) tras fijar "
            f"encabezado en fila {header_row}. Columnas: {list(df.columns)}"
        )

    por_fecha: dict[str, dict[str, float]] = {}
    filas_omitidas = 0
    for idx, row in df.iterrows():
        fecha_raw = row[col_fecha]
        if isinstance(fecha_raw, pd.Series):
            fecha_raw = fecha_raw.iloc[0]
        if pd.isna(fecha_raw):
            continue
        try:
            fecha = pd.to_datetime(fecha_raw, dayfirst=True).date().isoformat()
        except Exception as e:
            filas_omitidas += 1
            log(f"Fila {idx} omitida: valor de fecha no parseable ({fecha_raw!r}): {e}")
            continue

        plazo_raw = row[col_plazo]
        tasa_raw = row[col_tasa]
        if isinstance(plazo_raw, pd.Series):
            plazo_raw = plazo_raw.iloc[0]
        if isinstance(tasa_raw, pd.Series):
            tasa_raw = tasa_raw.iloc[0]
        if pd.isna(plazo_raw) or pd.isna(tasa_raw):
            continue
        try:
            dias = float(plazo_raw)
            tasa = float(tasa_raw)
        except (TypeError, ValueError):
            continue

        plazo_key = _dias_a_años(dias)
        por_fecha.setdefault(fecha, {})[plazo_key] = tasa

    if filas_omitidas:
        log(f"Total filas omitidas por fecha no parseable: {filas_omitidas}")

    registros = [{"fecha": f, "tasas": t} for f, t in sorted(por_fecha.items())]
    return registros


def actualizar_historico(tipo_curva: str, nuevos_registros: list[dict]) -> dict:
    DATA_DIR.mkdir(exist_ok=True)
    hist_path = DATA_DIR / f"{tipo_curva}.json"
    historico = json.loads(hist_path.read_text()) if hist_path.exists() else []
    por_fecha = {r["fecha"]: r for r in historico}
    agregados = 0
    for r in nuevos_registros:
        if r["fecha"] not in por_fecha:
            agregados += 1
        por_fecha[r["fecha"]] = r  # nuevos sobrescriben (por si SBS corrige un dato)
    historico_final = sorted(por_fecha.values(), key=lambda r: r["fecha"])
    hist_path.write_text(json.dumps(historico_final, ensure_ascii=False, indent=2))
    return {
        "tipo_curva": tipo_curva,
        "nombre": CURVAS_VALIDAS[tipo_curva],
        "registros_totales": len(historico_final),
        "registros_agregados_hoy": agregados,
        "ultima_fecha_dato": historico_final[-1]["fecha"] if historico_final else None,
    }


def main():
    args = parse_args()
    hoy = date.today()
    fecha_fin = datetime.strptime(args.fecha_fin, "%Y-%m-%d").date() if args.fecha_fin else hoy
    fecha_inicio = (datetime.strptime(args.fecha_inicio, "%Y-%m-%d").date() if args.fecha_inicio
                     else fecha_fin - timedelta(days=args.dias_atras))

    resultado_run = {
        "tipo_curva": args.tipo_curva,
        "ejecutado_utc": datetime.utcnow().isoformat() + "Z",
        "fecha_inicio_consultada": fecha_inicio.isoformat(),
        "fecha_fin_consultada": fecha_fin.isoformat(),
        "status": "error",
    }

    try:
        excel_path = descargar_excel(args.tipo_curva, fecha_inicio, fecha_fin, headful=args.headful)
        registros = parsear_excel(excel_path)
        if not registros:
            raise ValueError("El archivo se descargó pero no se pudo extraer ningún registro (revisar layout en debug/).")
        resumen = actualizar_historico(args.tipo_curva, registros)
        resultado_run.update(status="ok", **resumen)
        log(f"OK: {resumen}")
    except Exception as e:
        resultado_run["error"] = str(e)
        resultado_run["traceback"] = traceback.format_exc()
        log(f"ERROR: {e}")

    DATA_DIR.mkdir(exist_ok=True)
    runs_path = DATA_DIR / "latest_run.json"
    runs = json.loads(runs_path.read_text()) if runs_path.exists() else {}
    runs[args.tipo_curva] = resultado_run
    runs_path.write_text(json.dumps(runs, ensure_ascii=False, indent=2))

    if resultado_run["status"] != "ok":
        sys.exit(1)


if __name__ == "__main__":
    main()
