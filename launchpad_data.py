"""
Launchpad — datos de mercado desde fuentes 100% gratuitas.

Fuentes:
  - Yahoo Finance (via yfinance)   -> índices, FX, cripto, futuros, commodities, tasas UST, charts
  - mindicador.cl                   -> UF, dólar observado, euro, IPC, UTM, TPM (Chile, API pública)
  - ForexFactory (nfs.faireconomy)  -> calendario económico semanal (JSON público, sin key)
  - Google News RSS                 -> feed de noticias (sin key)

Nota: varios instrumentos del Launchpad original de Bloomberg (harina de pescado,
litio FOB, Baltic Dry Index, bonos soberanos España/Portugal/Italia/Grecia, tasas
chilenas PESO2B/BCU) no tienen una fuente gratuita confiable y se omiten o marcan
como no disponibles ("N/D") en vez de simular datos falsos.
"""

import time
import logging
import threading
from urllib.parse import quote
import xml.etree.ElementTree as ET

import requests
import yfinance as yf

log = logging.getLogger(__name__)

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; CarteraBCI-Launchpad/1.0)"}

# ── Mapas de instrumentos (etiqueta -> ticker Yahoo Finance) ─────────────────

INDICES = {
    "INDU": "^DJI",
    "SPX": "^GSPC",
    "NASDAQ": "^IXIC",
    "MEXBOL": "^MXX",
    "RUSSELL 2000": "^RUT",
    "IPSA": "^IPSA",
    "IBOV": "^BVSP",
    "MERVAL": "^MERV",
    "DAX": "^GDAXI",
    "IBEX": "^IBEX",
    "ITALIA": "FTSEMIB.MI",
    "CAC": "^FCHI",
    "INGLATERRA": "^FTSE",
    "JAPON": "^N225",
    "HONG KONG": "^HSI",
}

FX_CRYPTO = {
    "CLP": "CLP=X",
    "REAL": "BRL=X",
    "PESO MEX": "MXN=X",
    "COLOMBIANO": "COP=X",
    "PESO ARG": "ARS=X",
    "SUDAFRICANO": "ZAR=X",
    "TRY": "TRY=X",
    "PEN": "PEN=X",
    "ETHEREUM": "ETH-USD",
    "BITCOIN": "BTC-USD",
    "AUSTRALIANO": "AUDUSD=X",
    "INDICE DOLAR": "DX-Y.NYB",
    "GBP": "GBPUSD=X",
    "EURO": "EURUSD=X",
    "YUAN": "CNY=X",
    "VIX": "^VIX",
    "YEN": "JPY=X",
    "EEM": "EEM",
    "PYG": "PYG=X",
}

FUTURES = {
    "DOW JONES": "YM=F",
    "S&P 500 MINI": "ES=F",
    "NASDAQ 100": "NQ=F",
}

COMMODITIES = {
    "COBRE (COMEX)": "HG=F",
    "WTI": "CL=F",
    "PLATA": "SI=F",
    "GAS NATURAL": "NG=F",
    "ORO": "GC=F",
    "ALUMINIO": "ALI=F",
}

RATES_US = {
    "USGG5YR": "^FVX",
    "USGG10YR": "^TNX",
    "USGG30YR": "^TYX",
    "USGG13W": "^IRX",
}
# Yahoo cotiza ^TNX/^TYX/^FVX como rendimiento x 10 (ej. 42.5 = 4.25%)
RATES_SCALE = {"^TNX": 0.1, "^TYX": 0.1, "^FVX": 0.1}

CHARTS = {
    "g4": ("IPSA INDEX", "^IPSA"),
    "g5": ("INDU INDEX", "^DJI"),
    "g6": ("USGG10YR INDEX", "^TNX"),
    "g7": ("USGG30YR INDEX", "^TYX"),
}

MINDICADOR_KEYS = ["uf", "dolar", "euro", "ipc", "utm", "tpm", "imacec", "bitcoin"]

# ── Cache en memoria con TTL ─────────────────────────────────────────────────

_cache = {}
_cache_lock = threading.Lock()


def cached(key, ttl, fn, *args, **kwargs):
    now = time.time()
    with _cache_lock:
        entry = _cache.get(key)
        if entry and now - entry[0] < ttl:
            return entry[1]
    try:
        value = fn(*args, **kwargs)
    except Exception as e:
        log.warning("launchpad: error refrescando %s: %s", key, e)
        with _cache_lock:
            entry = _cache.get(key)
        return entry[1] if entry else None
    with _cache_lock:
        _cache[key] = (now, value)
    return value


# ── Cotizaciones (Yahoo Finance) ─────────────────────────────────────────────

def _build_group(df, symbols, group_map):
    out = {}
    single = len(symbols) == 1
    for label, sym in group_map.items():
        scale = RATES_SCALE.get(sym, 1)
        try:
            if df is None:
                raise ValueError("sin datos")
            sub = df if single else df[sym]
            closes = sub["Close"].dropna()
            if len(closes) < 2:
                raise ValueError("historial insuficiente")
            price = float(closes.iloc[-1]) * scale
            prev = float(closes.iloc[-2]) * scale
            chg = price - prev
            chg_pct = (chg / prev) if prev else None
            out[label] = {
                "symbol": sym, "price": price, "change": chg,
                "change_pct": chg_pct, "ok": True,
            }
        except Exception:
            out[label] = {
                "symbol": sym, "price": None, "change": None,
                "change_pct": None, "ok": False,
            }
    return out


def fetch_all_quotes():
    """Una sola descarga batch para todos los grupos de instrumentos."""
    all_map = {}
    all_map.update(INDICES)
    all_map.update(FX_CRYPTO)
    all_map.update(FUTURES)
    all_map.update(COMMODITIES)
    all_map.update(RATES_US)
    symbols = sorted(set(all_map.values()))

    try:
        df = yf.download(
            symbols, period="5d", interval="1d", group_by="ticker",
            progress=False, threads=True, auto_adjust=False,
        )
    except Exception as e:
        log.warning("launchpad: yfinance download error: %s", e)
        df = None

    return {
        "indices": _build_group(df, symbols, INDICES),
        "fx": _build_group(df, symbols, FX_CRYPTO),
        "futures": _build_group(df, symbols, FUTURES),
        "commodities": _build_group(df, symbols, COMMODITIES),
        "rates": _build_group(df, symbols, RATES_US),
        "ts": time.time(),
    }


def fetch_chart(symbol, period="1d", interval="5m"):
    try:
        df = yf.download(
            symbol, period=period, interval=interval,
            progress=False, auto_adjust=False,
        )
        if df is None or df.empty:
            return {"symbol": symbol, "points": [], "ok": False}
        df = df.dropna(subset=["Close"])
        scale = RATES_SCALE.get(symbol, 1)
        points = [
            {"t": idx.isoformat(), "c": float(row["Close"]) * scale}
            for idx, row in df.iterrows()
        ]
        return {"symbol": symbol, "points": points, "ok": True}
    except Exception as e:
        log.warning("launchpad: chart error %s: %s", symbol, e)
        return {"symbol": symbol, "points": [], "ok": False}


# ── Indicadores Chile (mindicador.cl) ────────────────────────────────────────

def fetch_mindicador():
    try:
        r = requests.get("https://mindicador.cl/api", timeout=10, headers=HEADERS)
        r.raise_for_status()
        data = r.json()
        out = {}
        for k in MINDICADOR_KEYS:
            if k in data and isinstance(data[k], dict):
                out[k] = {
                    "nombre": data[k].get("nombre"),
                    "valor": data[k].get("valor"),
                    "fecha": data[k].get("fecha"),
                    "unidad": data[k].get("unidad_medida"),
                }
        return out
    except Exception as e:
        log.warning("launchpad: mindicador.cl error: %s", e)
        return {}


# ── Noticias (Google News RSS) ───────────────────────────────────────────────

def fetch_news(query, hl="es-419", gl="CL", ceid="CL:es", limit=14):
    url = (
        "https://news.google.com/rss/search?q="
        f"{quote(query)}&hl={hl}&gl={gl}&ceid={ceid}"
    )
    try:
        r = requests.get(url, timeout=10, headers=HEADERS)
        r.raise_for_status()
        root = ET.fromstring(r.content)
        items = []
        for item in root.findall(".//item")[:limit]:
            title = (item.findtext("title") or "").strip()
            link = (item.findtext("link") or "").strip()
            pub = (item.findtext("pubDate") or "").strip()
            source_el = item.find("source")
            source = source_el.text.strip() if source_el is not None and source_el.text else ""
            items.append({"title": title, "link": link, "pub": pub, "source": source})
        return items
    except Exception as e:
        log.warning("launchpad: news error (%s): %s", query, e)
        return []


def fetch_news_bundle():
    return {
        "chile": fetch_news(
            "mercado bolsa Chile IPSA economia peso dolar",
            hl="es-419", gl="CL", ceid="CL:es",
        ),
        "global": fetch_news(
            "stock market economy federal reserve markets",
            hl="en-US", gl="US", ceid="US:en",
        ),
    }


# ── Calendario económico (ForexFactory JSON público) ─────────────────────────

def fetch_calendar():
    url = "https://nfs.faireconomy.media/ff_calendar_thisweek.json"
    try:
        r = requests.get(url, timeout=10, headers=HEADERS)
        r.raise_for_status()
        data = r.json()
        if not isinstance(data, list):
            return []
        out = []
        for ev in data:
            out.append({
                "title": ev.get("title"),
                "country": ev.get("country"),
                "date": ev.get("date"),
                "impact": ev.get("impact"),
                "forecast": ev.get("forecast"),
                "previous": ev.get("previous"),
                "actual": ev.get("actual"),
            })
        return out
    except Exception as e:
        log.warning("launchpad: calendar error: %s", e)
        return []
