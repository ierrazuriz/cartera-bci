"""
Aplicación web — Cartera BCI: EL LTDA y EMF SPA
Ejecutar localmente: python app.py
En Railway: gunicorn app:app
"""
import os
import json
import requests as _requests
from datetime import date
from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify
from cartera_calc import (
    calcular_el, calcular_emf,
    PRECIOS_DEFAULT, INSTRUMENTOS_META,
    EL_ACCIONES, EL_CFI, EL_SIM, EMF_CFI, EMF_FWD,
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


@app.route("/api/precios_auto")
def precios_auto():
    """Obtiene precios actuales via yfinance (Yahoo Finance) + mindicador.cl para UF."""
    import yfinance as yf

    precios = load_precios()
    errores = []
    actualizados = []

    # ── Acciones y FX — yfinance ──────────────────────────────────────────
    TICKERS_YF = {
        "ABC":      "ABC.SN",
        "AGUAS-A":  "AGUAS-A.SN",
        "CENCOSUD": "CENCOSUD.SN",
        "CHILE":    "CHILE.SN",
        "COPEC":    "COPEC.SN",
        "ENELAM":   "ENELAM.SN",
        "LTM":      "LTM.SN",
            "ITAUCL":   "ITAUCL.SN",
        "USD":      "USDCLP=X",
        "EUR":      "EURCLP=X",
    }
    try:
        symbols = list(TICKERS_YF.values())
        data = yf.download(symbols, period="1d", auto_adjust=True, progress=False)
        closes = data["Close"].iloc[-1] if not data.empty else {}

        sym_to_nem = {v: k for k, v in TICKERS_YF.items()}
        for sym, precio in closes.items():
            nem = sym_to_nem.get(sym)
            if nem and precio and float(precio) > 0:
                precios[nem] = round(float(precio), 4)
                actualizados.append(nem)
            elif nem:
                errores.append(f"{nem}: sin precio en Yahoo Finance")
    except Exception as e:
        errores.append(f"yfinance: {e}")

    # ── UF — mindicador.cl ────────────────────────────────────────────────
    try:
        r = _requests.get("https://mindicador.cl/api/uf", timeout=8)
        r.raise_for_status()
        uf_val = r.json()["serie"][0]["valor"]
        precios["UF"] = round(float(uf_val), 2)
        actualizados.append("UF")
    except Exception as e:
        errores.append(f"UF (mindicador.cl): {e}")

    return jsonify({
        "precios":      precios,
        "actualizados": actualizados,
        "errores":      errores,
                    "manuales":    ["CFIARRAA-E"],
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
