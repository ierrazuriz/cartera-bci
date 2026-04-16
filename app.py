"""
Aplicación web — Cartera BCI: EL LTDA y EMF SPA
Ejecutar localmente: python app.py
En Railway: gunicorn app:app

=== FIXES APLICADOS ===
Fix 1: Scheduler con retry cada 15 min entre 10:00 y 14:00 (antes: ventana fija 09:30-09:59)
Fix 3: Query Gmail usa email directo (antes: display name que podía no matchear)
Fix 4: Nuevo endpoint /api/actualizar_facturas + sync automático de facturas
"""

import os
import json
import logging
import threading
import requests as _requests
from datetime import date
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify

from cartera_calc import (
    calcular_el, calcular_emf,
    cargar_datos_cartola,
    PRECIOS_DEFAULT, INSTRUMENTOS_META,
)

app = Flask(__name__)

# Precios se guardan en este archivo JSON (persiste entre reinicios)
DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
DATA_FILE = os.path.join(DATA_DIR, "precios.json")


# ── Helpers ──────────────────────────────────────────────────────────────────

def load_precios():
    try:
        with open(DATA_FILE) as f:
            data = json.load(f)
        for k, v in PRECIOS_DEFAULT.items():
            data.setdefault(k, v)
        return data
    except (FileNotFoundError, json.JSONDecodeError):
        return PRECIOS_DEFAULT.copy()


def save_precios(precios):
    with open(DATA_FILE, "w") as f:
        json.dump(precios, f, indent=2)


# ── Jinja2 filters ───────────────────────────────────────────────────────────

@app.template_filter("clp")
def clp_filter(n):
    try:
        return f"$ {int(round(float(n))):,}".replace(",", ".")
    except Exception:
        return "—"


@app.template_filter("miles")
def miles_filter(n):
    try:
        return f"{int(round(float(n))):,}".replace(",", ".")
    except Exception:
        return "—"


@app.template_filter("num")
def num_filter(n, dec=2):
    try:
        return f"{float(n):,.{dec}f}"
    except Exception:
        return "—"


@app.template_filter("pct")
def pct_filter(n):
    try:
        v = float(n) * 100
        sign = "+" if v >= 0 else ""
        return f"{sign}{v:.2f}%"
    except Exception:
        return "—"


@app.template_filter("fdate")
def fdate_filter(d):
    try:
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(d)


@app.template_filter("signo")
def signo_filter(n):
    """Clase CSS según signo del número."""
    try:
        return "pos" if float(n) >= 0 else "neg"
    except Exception:
        return ""


# ── Rutas ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    precios = load_precios()
    hoy = date.today()

    # FIX 2: cargar datos dinámicos desde cartola_data.json
    datos_cartola = cargar_datos_cartola()
    el = calcular_el(precios, hoy, datos_cartola)
    emf = calcular_emf(precios, hoy, datos_cartola)
    pat_total = el["patrimonio_clp"] + emf["patrimonio_clp"]

    # Fecha base de la cartola (para mostrar en el header)
    fecha_base = datos_cartola.get("fecha", "sin datos") if datos_cartola else "sin datos"

    return render_template(
        "index.html",
        precios=precios,
        el=el,
        emf=emf,
        hoy=hoy,
        pat_total=pat_total,
        instrumentos_meta=INSTRUMENTOS_META,
        precios_default=PRECIOS_DEFAULT,
        fecha_base=fecha_base,
    )


@app.route("/precios", methods=["POST"])
def update_precios():
    precios = load_precios()
    for key in PRECIOS_DEFAULT:
        val = request.form.get(key, "").strip().replace(",", ".")
        if val:
            try:
                precios[key] = float(val)
            except ValueError:
                pass
    save_precios(precios)
    return redirect(url_for("index"))


@app.route("/reset", methods=["POST"])
def reset_precios():
    save_precios(PRECIOS_DEFAULT.copy())
    return redirect(url_for("index"))


@app.route("/excel")
def download_excel():
    """Regenera y descarga el Excel."""
    try:
        import generar_cartera as gc
        gc.main()
        return send_file(
            gc.ARCHIVO,
            as_attachment=True,
            download_name="Cartera BCI - EL y EMF SPA.xlsx",
        )
    except Exception as e:
        return f"Error generando Excel: {e}", 500


def _parse_clp(text):
    """Convierte precio en formato chileno '6.527,90' -> 6527.9"""
    return float(text.strip().replace(".", "").replace(",", "."))


@app.route("/api/precios_auto")
def precios_auto():
    """Obtiene precios actuales desde mercadosenconsorcio.cl (scraping HTML)."""
    from bs4 import BeautifulSoup

    precios = load_precios()
    errores = []
    actualizados = []
    headers = {"User-Agent": "Mozilla/5.0"}
    ACCIONES_BUSCADAS = {"ABC", "AGUAS-A", "CENCOSUD", "CHILE", "COPEC", "ENELAM", "ITAUCL", "LTM"}

    # ── 1. Página principal: monedas (UF, USD obs., EUR/USD) ──────────────
    try:
        r = _requests.get("https://mercadosenconsorcio.cl/mercado/acciones", timeout=12, headers=headers)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        usd_obs = None
        eur_usd = None

        for table in soup.find_all("table"):
            th = table.find("th")
            if not th or "Moneda" not in th.get_text():
                continue
            for row in table.find_all("tr"):
                nemo_s = row.find("span", class_="instrument")
                price_s = row.find("span", attrs={"data-bind": "text: price"})
                if not (nemo_s and price_s):
                    continue
                name = nemo_s.get_text(strip=True).upper()
                try:
                    val = _parse_clp(price_s.get_text(strip=True))
                except ValueError:
                    continue
                if "OBS" in name:
                    usd_obs = val
                elif "EUR" in name and "USD" in name:
                    eur_usd = val
                elif name == "UF":
                    precios["UF"] = round(val, 2)
                    actualizados.append("UF")

        if usd_obs:
            precios["USD"] = round(usd_obs, 2)
            actualizados.append("USD")
        if eur_usd and usd_obs:
            precios["EUR"] = round(eur_usd * usd_obs, 2)
            actualizados.append("EUR")

        # Recoger acciones de la página
        for row in soup.find_all("tr"):
            nemo_s = row.find("span", class_="instrument")
            price_s = row.find("span", attrs={"data-bind": "text: price"})
            if not (nemo_s and price_s):
                continue
            nemo = nemo_s.get_text(strip=True)
            if nemo not in ACCIONES_BUSCADAS:
                continue
            try:
                val = _parse_clp(price_s.get_text(strip=True))
            except ValueError:
                continue
            if val > 0:
                precios[nemo] = round(val, 4)
                if nemo not in actualizados:
                    actualizados.append(nemo)

    except Exception as e:
        errores.append(f"mercadosenconsorcio (página principal): {e}")

    # ── 2. Constituyentes IPSA ─────────────────────────────────────────────
    faltantes = ACCIONES_BUSCADAS - set(actualizados)
    if faltantes:
        try:
            r = _requests.post(
                "https://mercadosenconsorcio.cl/www/global/constituyentes.html",
                data={"ORDER": "VOLUME", "SORT": "desc", "HASH": "x"},
                headers={**headers, "X-Requested-With": "XMLHttpRequest",
                         "Referer": "https://mercadosenconsorcio.cl/www/detalle.html"},
                timeout=12,
            )
            r.raise_for_status()
            soup2 = BeautifulSoup(r.text, "html.parser")
            for row in soup2.find_all("tr"):
                nemo_s = row.find("span", class_="instrumentList")
                price_s = row.find("span", attrs={"data-bind": "text: price"})
                if not (nemo_s and price_s):
                    continue
                nemo = nemo_s.get_text(strip=True)
                if nemo not in faltantes:
                    continue
                try:
                    val = _parse_clp(price_s.get_text(strip=True))
                except ValueError:
                    continue
                if val > 0:
                    precios[nemo] = round(val, 4)
                    actualizados.append(nemo)
                    faltantes.discard(nemo)
        except Exception as e:
            errores.append(f"mercadosenconsorcio (constituyentes): {e}")

    for nem in faltantes:
        errores.append(f"{nem}: no encontrado (se mantiene precio anterior)")

    save_precios(precios)
    return jsonify({
        "precios": precios,
        "actualizados": actualizados,
        "errores": errores,
        "manuales": ["CFIARRAA-E", "CFIMRCLP", "CFITRIPT-E", "ABC"],
    })


@app.route("/api/estado")
def api_estado():
    """Endpoint JSON con el estado actual de la cartera."""
    precios = load_precios()
    hoy = date.today()
    datos_cartola = cargar_datos_cartola()
    el = calcular_el(precios, hoy, datos_cartola)
    emf = calcular_emf(precios, hoy, datos_cartola)
    return jsonify({
        "fecha": hoy.isoformat(),
        "precios": precios,
        "el_ltda": {
            "patrimonio_clp": el["patrimonio_clp"],
            "patrimonio_uf": el["patrimonio_uf"],
        },
        "emf_spa": {
            "patrimonio_clp": emf["patrimonio_clp"],
            "patrimonio_uf": emf["patrimonio_uf"],
        },
        "total_clp": el["patrimonio_clp"] + emf["patrimonio_clp"],
    })


@app.route("/facturas")
def facturas():
    json_path = os.path.join(DATA_DIR, "facturas_data.json")
    try:
        with open(json_path, encoding='utf-8') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {'movimientos': [], 'dividendos': [], 'transferencias': [],
                'forwards': [], 'comparacion': [], 'sync_at': ''}
    return render_template("facturas.html", data=data, hoy=date.today())


# ── FIX 4: Endpoint para actualizar facturas desde Gmail ─────────────────────

@app.route("/api/actualizar_facturas", methods=["POST"])
def actualizar_facturas():
    """Descarga y parsea facturas desde Gmail."""
    token = request.headers.get("X-Actualizar-Token", "")
    secret = os.environ.get("ACTUALIZAR_SECRET", "")
    if secret and token != secret:
        return jsonify({"error": "no autorizado"}), 403

    def _run():
        try:
            import gmail_facturas as gf
            logging.info("Actualizando facturas desde Gmail...")
            gf.sync_facturas()
            logging.info("Facturas actualizadas correctamente")
        except Exception as e:
            logging.error("Error actualizando facturas: %s", e, exc_info=True)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"status": "actualizando facturas en segundo plano"})


# ── FIX 1: Endpoint manual para actualizar cartola ───────────────────────────

@app.route("/api/actualizar_cartola", methods=["POST"])
def actualizar_cartola():
    """
    Descarga la cartola de hoy desde Gmail y actualiza cartola_data.json.
    Llamado por el scheduler diario o manualmente.
    """
    token = request.headers.get("X-Actualizar-Token", "")
    secret = os.environ.get("ACTUALIZAR_SECRET", "")
    if secret and token != secret:
        return jsonify({"error": "no autorizado"}), 403

    def _run():
        try:
            _ejecutar_sync_cartola()
        except Exception as e:
            logging.error("Error actualizando cartola (manual): %s", e, exc_info=True)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"status": "actualizando en segundo plano"})


# ── Lógica de sync compartida ────────────────────────────────────────────────

def _ejecutar_sync_cartola(fecha=None):
    """
    Descarga, parsea y guarda la cartola.
    También actualiza precios.json con los precios de la cartola.
    """
    import gmail_bci as gb
    import parsear_cartola as pc

    logging.info("Sync cartola: autenticando Gmail...")
    creds = gb.autenticar()
    service = gb.build("gmail", "v1", credentials=creds)

    logging.info("Sync cartola: buscando email BCI...")
    msg_id = gb.buscar_email_bci(service, fecha)

    logging.info("Sync cartola: descargando ZIP...")
    zip_data, zip_name, fecha_email = gb.descargar_zip(service, msg_id)
    logging.info("Sync cartola: ZIP descargado: %s (%d bytes)", zip_name, len(zip_data))

    pdfs = gb.extraer_pdfs(zip_data)
    logging.info("Sync cartola: %d PDFs extraídos", len(pdfs))

    data = pc.parsear(pdfs)
    pc.guardar(data)

    # Actualizar precios.json con los precios de la cartola
    precios_cartola = data.get("precios", {})
    if precios_cartola:
        precios = load_precios()
        precios.update(precios_cartola)
        save_precios(precios)
        logging.info("Sync cartola: precios actualizados: %s", list(precios_cartola.keys()))

    logging.info("Sync cartola: completado — fecha cartola: %s", data.get("fecha"))
    return data


# ══════════════════════════════════════════════════════════════════════════════
# FIX 1: Scheduler con retry cada 15 min entre 10:00 y 14:00
# ══════════════════════════════════════════════════════════════════════════════

def _iniciar_scheduler():
    """
    Scheduler mejorado:
    - Intenta cada 15 minutos entre 10:00 y 14:00 hora Chile (días hábiles)
    - Si encuentra cartola de HOY, la procesa y para hasta mañana
    - Si falla, reintenta en 15 min (máximo hasta las 14:00)
    - También sincroniza facturas una vez al día a las 14:30
    """
    import time
    from datetime import datetime
    import zoneinfo

    TZ_CHILE = zoneinfo.ZoneInfo("America/Santiago")
    INTERVALO_RETRY = 15 * 60  # 15 minutos en segundos

    def _loop():
        cartola_ultimo_dia = None
        facturas_ultimo_dia = None

        while True:
            try:
                ahora = datetime.now(TZ_CHILE)
                hoy = ahora.date()
                es_habil = ahora.weekday() < 5
                hora = ahora.hour
                minuto = ahora.minute

                # ── Sync cartola: 10:00 a 14:00, cada 15 min ──────────
                if (es_habil
                    and 10 <= hora < 14
                    and hoy != cartola_ultimo_dia):

                    logging.info("Scheduler: intentando sync cartola (intento %02d:%02d)",
                                 hora, minuto)
                    try:
                        _ejecutar_sync_cartola()
                        cartola_ultimo_dia = hoy
                        logging.info("Scheduler: cartola actualizada exitosamente (%s)", hoy)
                    except ValueError as e:
                        # "No se encontró email" — reintentará en 15 min
                        logging.warning("Scheduler: cartola no disponible aún: %s", e)
                    except Exception as e:
                        logging.error("Scheduler: error en sync cartola: %s", e, exc_info=True)

                # ── Sync facturas: una vez al día a las 14:30+ ────────
                if (es_habil
                    and hora == 14 and minuto >= 30
                    and hoy != facturas_ultimo_dia):

                    logging.info("Scheduler: sincronizando facturas...")
                    try:
                        import gmail_facturas as gf
                        gf.sync_facturas()
                        facturas_ultimo_dia = hoy
                        logging.info("Scheduler: facturas sincronizadas (%s)", hoy)
                    except Exception as e:
                        logging.error("Scheduler: error sync facturas: %s", e, exc_info=True)

            except Exception as e:
                logging.error("Scheduler loop error: %s", e, exc_info=True)

            time.sleep(INTERVALO_RETRY)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    logging.info("Scheduler iniciado: cartola 10:00-14:00 cada 15min, facturas 14:30")


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# Solo iniciar el scheduler en producción (Railway)
if os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("INICIAR_SCHEDULER"):
    _iniciar_scheduler()


if __name__ == "__main__":
    _iniciar_scheduler()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
