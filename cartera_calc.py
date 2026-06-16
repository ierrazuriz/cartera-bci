"""
Lógica de cálculo de cartera — BCI: EL LTDA y EMF SPA.

Las posiciones se leen dinámicamente de cartola_data.json (actualizado
diariamente por el sync de Gmail/BCI). Si cartola_data.json no existe
o está vacío, se usan los valores de la cartola 12/06/2026 como fallback.
"""
from datetime import date, datetime
import json
from pathlib import Path

DATA_DIR = Path(__file__).parent
CARTOLA_FILE = DATA_DIR / "cartola_data.json"


# ── Precios default (fallback cartola 12/06/2026) ──────────────────────────

PRECIOS_DEFAULT = {
    "UF": 40771.41,
    "USD": 909.02,
    "EUR": 1050.53,
    "ABC": 10.48,
    "AGUAS-A": 331.0,
    "CENCOSUD": 2180.1,
    "CFIARRAA-E": 52867.97,
    "CFITRIPT-E": 14000.0,
    "CHILE": 178.25,
    "COPEC": 6159.0,
    "ENELAM": 77.21,
    "ITAUCL": 18152.0,
    "LTM": 23.15,
}

INSTRUMENTOS_META = {
    "ABC":        {"nombre": "Abc S.A.",                 "tipo": "accion", "fmt": ".4f"},
    "AGUAS-A":    {"nombre": "Aguas Andinas S.A.",        "tipo": "accion", "fmt": ".4f"},
    "CENCOSUD":   {"nombre": "Cencosud S.A.",             "tipo": "accion", "fmt": ".4f"},
    "CHILE":      {"nombre": "Banco De Chile",            "tipo": "accion", "fmt": ".4f"},
    "COPEC":      {"nombre": "Empresas Copec S.A.",       "tipo": "accion", "fmt": ".4f"},
    "ENELAM":     {"nombre": "Enel Americas S.A.",        "tipo": "accion", "fmt": ".4f"},
    "ITAUCL":     {"nombre": "Banco Itau Chile",          "tipo": "accion", "fmt": ".4f"},
    "LTM":        {"nombre": "Latam Airlines Group S.A.", "tipo": "accion", "fmt": ".4f"},
    "CFIARRAA-E": {"nombre": "Cfiarraa-E",               "tipo": "cfi",    "fmt": ".4f"},
    "CFITRIPT-E": {"nombre": "Cfitript-E",               "tipo": "cfi",    "fmt": ".4f"},
}


# ── Datos fallback (cartola 12/06/2026) ─────────────────────────────────────
# Se usan si cartola_data.json no existe o no tiene datos válidos.

_FALLBACK = {
    "fecha": "2026-06-12",
    "el": {
        "caja":         -9_914_264,
        "ops_liquidar": 202_497_822,
        "acciones": [
            {"nem": "ABC",      "cant_activo": 23_210_430, "cant_pasivo": 0, "precio_compra": 10.48,    "precio_cartola": 10.48},
            {"nem": "AGUAS-A",  "cant_activo":  1_819_069, "cant_pasivo": 0, "precio_compra": 331.0,   "precio_cartola": 331.0},
            {"nem": "CENCOSUD", "cant_activo":    136_229, "cant_pasivo": 0, "precio_compra": 2180.1,  "precio_cartola": 2180.1},
            {"nem": "CHILE",    "cant_activo":  1_126_593, "cant_pasivo": 0, "precio_compra": 178.25,  "precio_cartola": 178.25},
            {"nem": "COPEC",    "cant_activo":     21_055, "cant_pasivo": 0, "precio_compra": 6159.0,  "precio_cartola": 6159.0},
            {"nem": "ENELAM",   "cant_activo": 10_158_102, "cant_pasivo": 0, "precio_compra": 77.21,   "precio_cartola": 77.21},
            {"nem": "ITAUCL",   "cant_activo":      3_801, "cant_pasivo": 0, "precio_compra": 18152.0, "precio_cartola": 18152.0},
            {"nem": "LTM",      "cant_activo": 74_285_174, "cant_pasivo": 0, "precio_compra": 23.15,   "precio_cartola": 23.15},
        ],
        "cfis": [
            {"nem": "CFIARRAA-E", "cantidad": 4_187, "precio_compra": 48138.424, "precio_cartola": 52867.97},
            {"nem": "CFITRIPT-E", "cantidad": 1_471, "precio_compra": 13280.761, "precio_cartola": 14000.0},
        ],
        "sims": [
            {"instrumento": "ENELAM", "cantidad": 3_806_521,  "f_venta": "2026-05-29", "monto_venta": 300_677_094, "f_compra": "2026-06-30", "monto_compra": 302_216_451},
            {"instrumento": "ENELAM", "cantidad":    26_216,  "f_venta": "2026-06-05", "monto_venta":   2_000_019, "f_compra": "2026-07-03", "monto_compra":   2_008_979},
            {"instrumento": "ENELAM", "cantidad": 1_209_472,  "f_venta": "2026-06-10", "monto_venta":  90_988_579, "f_compra": "2026-07-10", "monto_compra":  91_425_319},
            {"instrumento": "LTM",    "cantidad": 41_799_270, "f_venta": "2026-06-10", "monto_venta": 920_837_918, "f_compra": "2026-07-10", "monto_compra": 925_256_101},
        ],
    },
    "emf": {
        "caja": 106_996_120,
        "cfis": [
            {"nem": "CFIARRAA-E", "cantidad": 500, "precio_compra": 47154.0, "precio_cartola": 52867.97},
        ],
        "fwds": [
            {"folio": 1845333, "tipo": "V", "usd": 500_000, "tc_fwd": 912.27, "f_inicio": "2026-06-08", "f_termino": "2026-07-17"},
        ],
    },
}


# ── Helpers ──────────────────────────────────────────────────────────────────

def _to_date(val) -> date:
    """Convierte ISO string '2026-06-10' o date object a date."""
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        return datetime.strptime(val[:10], "%Y-%m-%d").date()
    raise ValueError(f"No se puede convertir a date: {val!r}")


# ── Carga de datos de cartola ─────────────────────────────────────────────────

def cargar_datos_cartola() -> dict:
    """
    Lee cartola_data.json (generado por parsear_cartola.py tras cada sync).
    Si no existe o está vacío, retorna el fallback hardcodeado.
    """
    try:
        with open(CARTOLA_FILE, encoding="utf-8") as f:
            data = json.load(f)
        # Validar que tiene datos reales
        if data and data.get("el") and data["el"].get("acciones"):
            return data
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    return _FALLBACK


# ── Cálculo EL LTDA ───────────────────────────────────────────────────────────

def calcular_el(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    datos = cargar_datos_cartola()
    el = datos.get("el", _FALLBACK["el"])

    caja_el      = el.get("caja",         _FALLBACK["el"]["caja"])
    ops_liquidar = el.get("ops_liquidar", _FALLBACK["el"]["ops_liquidar"])

    # ── Acciones ──────────────────────────────────────────────────────────────
    acciones = []
    for a in el.get("acciones", []):
        nem    = a["nem"]
        nombre = INSTRUMENTOS_META.get(nem, {}).get("nombre", nem)
        p_c    = a.get("precio_cartola", 0)
        p      = precios.get(nem, p_c)
        cant_a = a.get("cant_activo", 0)
        cant_p = a.get("cant_pasivo", 0)
        va     = cant_a * p
        vp     = cant_p * p
        acciones.append({
            "nem": nem, "nombre": nombre,
            "cant_activo": cant_a, "cant_pasivo": cant_p,
            "precio_cartola": p_c, "precio_hoy": p,
            "valor_activo": va, "valor_pasivo": vp,
            "valor_neto": va + vp,
            "var_pct": (p - p_c) / p_c if p_c else 0,
        })

    # ── CFIs ──────────────────────────────────────────────────────────────────
    cfis = []
    for c in el.get("cfis", []):
        nem    = c["nem"]
        nombre = INSTRUMENTOS_META.get(nem, {}).get("nombre", nem)
        p_c    = c.get("precio_cartola", 0)
        p      = precios.get(nem, p_c)
        cant   = c.get("cantidad", 0)
        p_comp = c.get("precio_compra", p_c)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_c,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_c) / p_c if p_c else 0,
        })

    # ── Simultáneas ───────────────────────────────────────────────────────────
    sims = []
    for s in el.get("sims", []):
        inst   = s["instrumento"]
        cant   = s["cantidad"]
        f_vta  = _to_date(s["f_venta"])
        m_vta  = s["monto_venta"]
        f_cpra = _to_date(s["f_compra"])
        m_cpra = s["monto_compra"]

        total_days = (f_cpra - f_vta).days
        elapsed    = max(0, min((hoy - f_vta).days, total_days))
        amort      = m_vta + (m_cpra - m_vta) * elapsed / total_days if total_days else m_cpra
        p          = precios.get(inst, 0)
        vm         = cant * p

        sims.append({
            "instrumento": inst, "cantidad": cant,
            "f_venta": f_vta, "monto_venta": m_vta,
            "f_compra": f_cpra, "monto_compra": m_cpra,
            "monto_amortizado": amort, "valor_mercado": vm,
            "resultado": vm - amort,
            "dias_restantes": max(0, (f_cpra - hoy).days),
            "vencida": hoy >= f_cpra,
        })

    tot_acc_neto  = sum(a["valor_neto"]       for a in acciones)
    tot_cfi       = sum(c["valor_mercado"]    for c in cfis)
    tot_sim_amort = sum(s["monto_amortizado"] for s in sims)

    patrimonio = caja_el + ops_liquidar + tot_acc_neto + tot_cfi - tot_sim_amort
    return {
        "acciones": acciones, "cfis": cfis, "sims": sims,
        "tot_acc_activo":    sum(a["valor_activo"]  for a in acciones),
        "tot_acc_pasivo":    sum(a["valor_pasivo"]  for a in acciones),
        "tot_acc_neto":      tot_acc_neto,
        "tot_cfi":           tot_cfi,
        "tot_sim_amort":     tot_sim_amort,
        "tot_sim_vm":        sum(s["valor_mercado"] for s in sims),
        "tot_sim_resultado": sum(s["resultado"]     for s in sims),
        "caja":              caja_el,
        "ops_liquidar":      ops_liquidar,
        "patrimonio_clp":    patrimonio,
        "patrimonio_uf":     patrimonio / precios.get("UF",  39_841.72),
        "patrimonio_usd":    patrimonio / precios.get("USD",    927.46),
    }


# ── Cálculo EMF SPA ───────────────────────────────────────────────────────────

def calcular_emf(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    datos = cargar_datos_cartola()
    emf = datos.get("emf", _FALLBACK["emf"])

    caja_emf = emf.get("caja", _FALLBACK["emf"]["caja"])

    # ── CFIs ──────────────────────────────────────────────────────────────────
    cfis = []
    for c in emf.get("cfis", []):
        nem    = c["nem"]
        nombre = INSTRUMENTOS_META.get(nem, {}).get("nombre", nem)
        p_c    = c.get("precio_cartola", 0)
        p      = precios.get(nem, p_c)
        cant   = c.get("cantidad", 0)
        p_comp = c.get("precio_compra", p_c)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_c,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_c) / p_c if p_c else 0,
        })

    # ── Forwards ──────────────────────────────────────────────────────────────
    spot = precios.get("USD", 927.46)
    fwds = []
    compra_usd = venta_usd = 0
    for f in emf.get("fwds", []):
        folio  = f["folio"]
        tipo   = f["tipo"]
        usd    = f["usd"]
        tc_fwd = f["tc_fwd"]
        f_ini  = _to_date(f["f_inicio"])
        f_term = _to_date(f["f_termino"])

        resultado = (spot - tc_fwd) * usd if tipo == "C" else (tc_fwd - spot) * usd
        if tipo == "C":
            compra_usd += usd
        else:
            venta_usd += usd

        fwds.append({
            "folio": folio, "tipo": tipo, "usd": usd,
            "tc_fwd": tc_fwd, "f_inicio": f_ini, "f_termino": f_term,
            "tc_spot": spot, "resultado": resultado,
            "dias_restantes": max(0, (f_term - hoy).days),
            "vencido": hoy >= f_term,
        })

    tot_cfi = sum(c["valor_mercado"] for c in cfis)
    tot_fwd = sum(f["resultado"]     for f in fwds)

    patrimonio = caja_emf + tot_cfi
    return {
        "cfis": cfis, "fwds": fwds,
        "tot_cfi": tot_cfi, "tot_fwd": tot_fwd,
        "compra_usd": compra_usd, "venta_usd": venta_usd,
        "descalce_usd": compra_usd - venta_usd,
        "caja": caja_emf,
        "patrimonio_clp":  patrimonio,
        "patrimonio_uf":   patrimonio / precios.get("UF",  39_841.72),
        "patrimonio_usd":  patrimonio / precios.get("USD",    927.46),
    }
