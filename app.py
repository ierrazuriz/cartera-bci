"""
Aplicación web — Cartera BCI: EL LTDA y EMF SPA
Ejecutar localmente: python app.py
En Railway: gunicorn app:app
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
    PRECIOS_DEFAULT, INSTRUMENTOS_META,
)

app = Flask(__name__)

# Precios se guardan en este archivo JSON (persiste entre reinicios)
DATA_DIR  = os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
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
    el  = calcular_el(precios, hoy)
    emf = calcular_emf(precios, hoy)
    pat_total = el["patrimonio_clp"] + emf["patrimonio_clp"]
    return render_template(
        "index.html",
        precios=precios,
        el=el,
        emf=emf,
        hoy=hoy,
        pat_total=pat_total,
        instrumentos_meta=INSTRUMENTOS_META,
        precios_default=PRECIOS_DEFAULT,
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

        # Recoger acciones que aparezcan en las secciones top de la página
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

    # ── 2. Constituyentes IPSA: cubre las acciones no en las secciones top ─
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

    # Stocks no encontrados (ej. ABC, fuera del IPSA): conservan precio anterior
    for nem in faltantes:
        errores.append(f"{nem}: no encontrado en mercadosenconsorcio.cl (se mantiene precio anterior)")

    save_precios(precios)
    return jsonify({
        "precios":    precios,
        "actualizados": actualizados,
        "errores":    errores,
        "manuales":   ["CFIARRAA-E", "CFIMRCLP", "CFITRIPT-E", "ABC"],
    })


@app.route("/api/estado")
def api_estado():
    """Endpoint JSON con el estado actual de la cartera."""
    precios = load_precios()
    hoy = date.today()
    el  = calcular_el(precios, hoy)
    emf = calcular_emf(precios, hoy)
    return jsonify({
        "fecha": hoy.isoformat(),
        "precios": precios,
        "el_ltda": {
            "patrimonio_clp": el["patrimonio_clp"],
            "patrimonio_uf":  el["patrimonio_uf"],
        },
        "emf_spa": {
            "patrimonio_clp": emf["patrimonio_clp"],
            "patrimonio_uf":  emf["patrimonio_uf"],
        },
        "total_clp": el["patrimonio_clp"] + emf["patrimonio_clp"],
    })


@app.route("/facturas")
def facturas():
    import json, os
    json_path = os.path.join(DATA_DIR, "facturas_data.json")
    try:
        with open(json_path, encoding='utf-8') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {'movimientos': [], 'dividendos': [], 'transferencias': [], 'forwards': [], 'comparacion': [], 'sync_at': ''}
    return render_template("facturas.html", data=data, hoy=date.today())


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
            import gmail_bci as gb
            import parsear_cartola as pc
            logging.info("Actualizando cartola desde Gmail...")
            creds   = gb.autenticar()
            service = gb.build("gmail", "v1", credentials=creds)
            msg_id  = gb.buscar_email_bci(service)
            zip_data, zip_name, fecha_email = gb.descargar_zip(service, msg_id)
            pdfs    = gb.extraer_pdfs(zip_data)
            data    = pc.parsear(pdfs)
            pc.guardar(data)
            logging.info("Cartola actualizada: %s", data.get("fecha"))
        except Exception as e:
            logging.error("Error actualizando cartola: %s", e)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"status": "actualizando en segundo plano"})


# ── Scheduler diario ─────────────────────────────────────────────────────────

def _iniciar_scheduler():
    """Corre la actualización de cartola cada día hábil a las 09:30 hora Chile."""
    import time
    from datetime import datetime
    import zoneinfo

    TZ_CHILE = zoneinfo.ZoneInfo("America/Santiago")

    def _loop():
        ultimo_dia = None
        while True:
            try:
                ahora = datetime.now(TZ_CHILE)
                hoy   = ahora.date()
                es_habil = ahora.weekday() < 5
                hora_ok  = ahora.hour == 9 and ahora.minute >= 30
                if es_habil and hora_ok and hoy != ultimo_dia:
                    logging.info("Scheduler: iniciando actualización de cartola")
                    try:
                        import gmail_bci as gb
                        import parsear_cartola as pc
                        creds   = gb.autenticar()
                        service = gb.build("gmail", "v1", credentials=creds)
                        msg_id  = gb.buscar_email_bci(service)
                        zip_data, _, _ = gb.descargar_zip(service, msg_id)
                        pdfs    = gb.extraer_pdfs(zip_data)
                        data    = pc.parsear(pdfs)
                        pc.guardar(data)
                        ultimo_dia = hoy
                        logging.info("Scheduler: cartola actualizada (%s)", hoy)
                    except Exception as e:
                        logging.error("Scheduler: error en actualización: %s", e)
            except Exception as e:
                logging.error("Scheduler loop error: %s", e)
            time.sleep(60)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    logging.info("Scheduler de cartola iniciado (09:30 hora Chile, días hábiles)")


logging.basicConfig(level=logging.INFO)

# Solo iniciar el scheduler en producción (Railway) para evitar dobles en dev
if os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("INICIAR_SCHEDULER"):
    _iniciar_scheduler()


if __name__ == "__main__":
    _iniciar_scheduler()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
