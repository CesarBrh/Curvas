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


def parsear_excel(path: Path) -> list[dict]:
    """
    Convierte el archivo exportado por SBS a una lista de registros:
    [{"fecha": "YYYY-MM-DD", "tasas": {"<plazo_dias>": tasa_pct, ...}}, ...]

    El layout exacto del archivo no está documentado públicamente, así que el
    parser es defensivo: busca una columna de fecha y trata el resto de
    columnas numéricas como plazos (en días). Si el layout viene distinto,
    lanza un error claro con las columnas encontradas para poder ajustar
    rápido (revisar debug/*.xlsx en los artifacts del run que falló).
    """
    def _cargar(header_row):
        try:
            return pd.read_excel(path, header=header_row)
        except Exception:
            return pd.read_csv(path, sep=None, engine="python", header=header_row)

    # El export de SBS trae 1-2 filas de título/metadata antes del encabezado
    # real, así que buscamos la fila que contiene "fecha" en las primeras 10
    # filas en vez de asumir que el encabezado está en la fila 0.
    crudo = _cargar(None)
    log("Vista cruda del archivo (primeras 15 filas):\n" + crudo.head(15).to_string())
    header_row = None
    for i in range(min(10, len(crudo))):
        valores = [str(v) for v in crudo.iloc[i].tolist()]
        if any("fecha" in v.lower() for v in valores):
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            "No se encontró ninguna fila con 'fecha' en las primeras 10 filas del archivo. "
            f"Primeras filas:\n{crudo.head(10).to_string()}"
        )

    df = _cargar(header_row)
    df.columns = [str(c).strip() for c in df.columns]
    log("Columnas detectadas (fila de encabezado " + str(header_row) + "): " + str(list(df.columns)))
    col_fecha = next((c for c in df.columns if "fecha" in c.lower()), None)
    if col_fecha is None:
        raise ValueError(f"No se encontró columna de fecha tras fijar encabezado en fila {header_row}. Columnas: {list(df.columns)}")

    plazo_cols = [c for c in df.columns if c != col_fecha]
    registros = []
    filas_omitidas = 0
    for idx, row in df.iterrows():
        fecha_raw = row[col_fecha]
        # Si hay columnas con nombre duplicado, row[col_fecha] puede venir
        # como Series en vez de escalar — nos quedamos con el primer valor.
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
        tasas = {}
        for c in plazo_cols:
            val = row[c]
            if isinstance(val, pd.Series):
                val = val.iloc[0]
            if pd.isna(val):
                continue
            try:
                tasas[str(c)] = float(val)
            except (TypeError, ValueError):
                continue
        if tasas:
            registros.append({"fecha": fecha, "tasas": tasas})
    if filas_omitidas:
        log(f"Total filas omitidas por fecha no parseable: {filas_omitidas}")
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
