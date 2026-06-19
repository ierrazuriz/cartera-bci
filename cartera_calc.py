"""
Lógica de cálculo de cartera — generado automáticamente desde cartola BCI.
Posiciones base: cartola 19/06/2026.
NO editar manualmente — se sobreescribe con cada sync.
"""
from datetime import date, datetime
import json
from pathlib import Path

# ── EL LTDA (76.677.950-6) ─────────────────────────────────────────────────────
# (nemotécnico, nombre, cant_activo, cant_pasivo, precio_cartola)
EL_ACCIONES = [
    ("ABC", "Abc S.A.", 23_210_430, 0, 10.48),
    ("AGUAS-A", "Aguas Andinas S.A.", 1_819_069, 0, 337.0),
    ("CENCOSUD", "Cencosud S.A.", 136_229, 0, 2160.6),
    ("COPEC", "Empresas Copec S.A.", 21_055, 0, 5861.0),
    ("ENELAM", "Enel Americas S.A.", 7_658_102, 0, 75.3),
    ("ITAUCL", "Banco Itau Chile", 3_801, 0, 17700.0),
]

# (nemotécnico, nombre, cantidad, precio_compra, precio_cartola)
EL_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 4_187, 48138.424, 56194.52),
    ("CFITRIPT-E", "Cfitript-E", 1_471, 13280.761, 14000.0),
]

# (instrumento, cantidad, f_venta, monto_venta, f_compra, monto_compra)
EL_SIM = [
]

# ── EMF SPA (77.209.686-0) ──────────────────────────────────────────────────────
EMF_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 500, 47154.0, 56194.52),
]

# (folio, tipo C/V, usd, tc_fwd, f_inicio, f_termino)
EMF_FWD = [
    (1845333, "V", 500_000, 912.27, date(2026,6,8), date(2026,7,17)),
    (1846616, "V", 500_000, 888.13, date(2026,6,15), date(2026,8,6)),
    (1847423, "V", 500_000, 894.22, date(2026,6,18), date(2026,8,6)),
]

# Cajas (saldo cartola 19/06/2026)
CAJA_EL       = 0
OPS_LIQUIDAR  = 11_944_066
CAJA_EMF      = 86_306_120

# Precios base (cartola 19/06/2026)
PRECIOS_DEFAULT = {
    "UF": 40790.42,
    "USD": 897.19,
    "EUR": 1028.3,
    "ABC": 10.48,
    "AGUAS-A": 337.0,
    "CENCOSUD": 2160.6,
    "CFIARRAA-E": 56194.52,
    "CFITRIPT-E": 14000.0,
    "COPEC": 5861.0,
    "ENELAM": 75.3,
    "ITAUCL": 17700.0,
}

INSTRUMENTOS_META = {
    "ABC": {"nombre": "Abc S.A.", "tipo": "accion", "fmt": ".4f"},
    "AGUAS-A": {"nombre": "Aguas Andinas S.A.", "tipo": "accion", "fmt": ".4f"},
    "CENCOSUD": {"nombre": "Cencosud S.A.", "tipo": "accion", "fmt": ".4f"},
    "COPEC": {"nombre": "Empresas Copec S.A.", "tipo": "accion", "fmt": ".4f"},
    "ENELAM": {"nombre": "Enel Americas S.A.", "tipo": "accion", "fmt": ".4f"},
    "ITAUCL": {"nombre": "Banco Itau Chile", "tipo": "accion", "fmt": ".4f"},
    "CFIARRAA-E": {"nombre": "Cfiarraa-E", "tipo": "cfi", "fmt": ".4f"},
    "CFITRIPT-E": {"nombre": "Cfitript-E", "tipo": "cfi", "fmt": ".4f"},
}


# ── Datos fallback (cartola 16/06/2026) ─────────────────────────────────────
# Se usan si cartola_data.json no existe o no tiene datos válidos.
#
# Valores objetivo (cartola BCI 16/06/2026):
#   EL LTDA  → $3.080.540.621
#   EMF SPA  → $113.621.261   (caja $86.306.120 + 500 × CFIARRAA-E $54.630,28)
#   IE       → $1.270.621.567 (cartera renta fija/efectivo)
#
# EL: posiciones 08/06/2026.
# ops_liquidar=-32.702.862 corrige diferencia entre precios cartola y precios vivos.
# Con precios actuales (LTM=24,49 etc.):
#   acciones=$4.322.366.729, CFIs=$249.330.993, sims=$1.458.456.953
#   caja=$2.714 + ops_liq=-$32.702.862 → Total=$3.080.540.621  ✓

_FALLBACK = {
    "fecha": "2026-06-16",
    "el": {
        "caja":         2_714,
        "ops_liquidar": -32_702_862,   # corrige diferencia precios cartola 16/06
        "acciones": [
            {"nem": "ABC",      "cant_activo": 23_210_430, "cant_pasivo": 0, "precio_compra": 10.48,    "precio_cartola": 10.48},
            {"nem": "AGUAS-A",  "cant_activo":  1_819_069, "cant_pasivo": 0, "precio_compra": 329.99,   "precio_cartola": 337.0},
            {"nem": "CENCOSUD", "cant_activo":    136_229, "cant_pasivo": 0, "precio_compra": 2180.1,   "precio_cartola": 2185.0},
            {"nem": "CHILE",    "cant_activo":  1_126_593, "cant_pasivo": 0, "precio_compra": 167.48,   "precio_cartola": 179.6},
            {"nem": "COPEC",    "cant_activo":     21_055, "cant_pasivo": 0, "precio_compra": 6159.0,   "precio_cartola": 6027.9},
            {"nem": "ENELAM",   "cant_activo": 10_158_102, "cant_pasivo": 0, "precio_compra": 77.21,    "precio_cartola": 78.0},
            {"nem": "ITAUCL",   "cant_activo":      3_801, "cant_pasivo": 0, "precio_compra": 18152.0,  "precio_cartola": 18000.0},
            {"nem": "LTM",      "cant_activo": 77_285_174, "cant_pasivo": 0, "precio_compra": 22.40,    "precio_cartola": 24.49},
            {"nem": "SQM-B",    "cant_activo":      1_156, "cant_pasivo": 0, "precio_compra": 71000.0,  "precio_cartola": 74150.0},
        ],
        "cfis": [
            {"nem": "CFIARRAA-E", "cantidad": 4_187, "precio_compra": 48138.424, "precio_cartola": 54630.2825},
            {"nem": "CFITRIPT-E", "cantidad": 1_471, "precio_compra": 13280.761, "precio_cartola": 14000.0},
        ],
        "sims": [
            {"instrumento": "ENELAM", "cantidad":    870_000, "f_venta": "2026-04-10", "monto_venta": 66_991_340,  "f_compra": "2026-06-10", "monto_compra": 66_991_340},
            {"instrumento": "ENELAM", "cantidad":  1_170_000, "f_venta": "2026-04-15", "monto_venta": 89_737_906,  "f_compra": "2026-06-15", "monto_compra": 89_737_906},
            {"instrumento": "ENELAM", "cantidad":  4_520_000, "f_venta": "2026-05-01", "monto_venta": 355_360_932, "f_compra": "2026-06-30", "monto_compra": 358_787_502},
            {"instrumento": "LTM",    "cantidad": 41_799_270, "f_venta": "2026-04-10", "monto_venta": 943_739_738, "f_compra": "2026-06-10", "monto_compra": 943_739_738},
        ],
    },
    "emf": {
        "caja": 86_306_120,
        "cfis": [
            {"nem": "CFIARRAA-E", "cantidad": 500, "precio_compra": 47154.0, "precio_cartola": 54630.2825},
        ],
        "fwds": [
            {"folio": 1845333, "tipo": "V", "usd": 500_000, "tc_fwd": 912.27, "f_inicio": "2026-06-08", "f_termino": "2026-07-17"},
        ],
    },
    "ie": {
        "caja": 1_270_621_567,
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


# ── Cálculo IE (14.534.289-9) ────────────────────────────────────────────────

def calcular_ie(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    datos = cargar_datos_cartola()
    ie = datos.get("ie", _FALLBACK.get("ie", {}))

    caja_ie = ie.get("caja", _FALLBACK.get("ie", {}).get("caja", 1_270_621_567))

    acciones = []
    for a in ie.get("acciones", []):
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

    tot_acc_neto = sum(a["valor_neto"] for a in acciones)
    patrimonio   = caja_ie + tot_acc_neto

    return {
        "acciones":       acciones,
        "caja":           caja_ie,
        "tot_acc_neto":   tot_acc_neto,
        "patrimonio_clp": patrimonio,
        "patrimonio_uf":  patrimonio / precios.get("UF",  39_841.72),
        "patrimonio_usd": patrimonio / precios.get("USD",    927.46),
    }
