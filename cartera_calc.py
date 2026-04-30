"""
Lógica de cálculo de cartera — generado automáticamente desde cartola BCI.
Lee posiciones desde cartola_data.json (actualizado por el sync diario).
NO editar manualmente — se sobreescribe con cada sync.
"""
import os
import json
from datetime import date, datetime


def _cargar_json():
    """Lee cartola_data.json y retorna el dict, o None si no existe."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cartola_data.json")
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _to_date(s):
    """Convierte string ISO '2026-04-30' a date object."""
    if isinstance(s, date):
        return s
    return datetime.strptime(s, "%Y-%m-%d").date()


# ── Constantes fallback (cartola 29/04/2026) ────────────────────────────────
# (nemotécnico, nombre, cant_activo, cant_pasivo, precio_cartola)
EL_ACCIONES = [
    ("ABC", "Abc S.A.", 23_210_430, 0, 12.26),
    ("AGUAS-A", "Aguas Andinas S.A.", 1_819_069, 0, 347.0),
    ("CENCOSUD", "Cencosud S.A.", 86_229, 0, 2278.0),
    ("CHILE", "Banco De Chile", 5_000_000, 0, 168.25),
    ("COPEC", "Empresas Copec S.A.", 21_055, 0, 6515.0),
    ("ENELAM", "Enel Americas S.A.", 10_158_102, 0, 83.92),
    ("ITAUCL", "Banco Itau Chile", 3_801, 0, 19000.0),
    ("LTM", "Latam Airlines Group S.A.", 77_285_174, 0, 21.62),
    ("SQM-B", "Sociedad Quimica Y Minera De Chile S.A.", 5_553, 0, 79302.0),
]

# (nemotécnico, nombre, cantidad, precio_compra, precio_cartola)
EL_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 4_187, 48138.424, 56498.5811),
    ("CFIMRCLP", "Moneda Renta Clp Fi, Serie A", 11_172, 19592.0, 20019.16),
    ("CFITRIPT-E", "Cfitript-E", 1_471, 13280.761, 12000.0),
]

# (instrumento, cantidad, f_venta, monto_venta, f_compra, monto_compra)
EL_SIM = [
    ("AGUAS-A", 438_600, date(2026,4,23), 156_615_288, date(2026,5,25), 157_417_137),
    ("COPEC", 2_286, date(2026,4,23), 15_233_218, date(2026,5,25), 15_311_212),
    ("ENELAM", 816_050, date(2026,4,9), 68_540_040, date(2026,5,11), 68_883_678),
    ("ENELAM", 1_136_332, date(2026,4,17), 99_429_050, date(2026,5,15), 99_874_492),
    ("ENELAM", 7_905_720, date(2026,4,1), 640_284_263, date(2026,5,29), 646_226_202),
    ("LTM", 45_034_136, date(2026,4,9), 1_051_547_076, date(2026,5,11), 1_056_820_573),
    ("LTM", 8_391_584, date(2026,4,17), 207_356_041, date(2026,5,15), 208_284_989),
    ("LTM", 5_296_810, date(2026,4,1), 125_004_716, date(2026,5,29), 126_164_717),
    ("SQM-B", 5_553, date(2026,4,23), 425_903_994, date(2026,5,25), 428_084_622),
]

# ── EMF SPA (77.209.686-0) ──────────────────────────────────────────────────────
EMF_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 500, 47154.0, 56498.5811),
]

# (folio, tipo C/V, usd, tc_fwd, f_inicio, f_termino)
EMF_FWD = [
    (1835288, "C", 500_000, 891.8, date(2026,4,8), date(2026,5,5)),
    (1835290, "C", 500_000, 892.15, date(2026,4,8), date(2026,5,5)),
    (1836514, "C", 250_000, 887.06, date(2026,4,15), date(2026,5,5)),
    (1837496, "V", 250_000, 889.01, date(2026,4,21), date(2026,5,5)),
    (1836070, "V", 500_000, 897.82, date(2026,4,13), date(2026,5,5)),
    (1834324, "V", 500_000, 922.32, date(2026,4,2), date(2026,5,5)),
]

# Cajas fallback
CAJA_EL = 1_131_573
OPS_LIQUIDAR = 640_324_462
CAJA_EMF = 73_351_120

# Precios base fallback
PRECIOS_DEFAULT = {
    "UF": 40106.89,
    "USD": 896.03,
    "EUR": 1049.58,
    "ABC": 12.26,
    "AGUAS-A": 347.0,
    "CENCOSUD": 2278.0,
    "CFIARRAA-E": 56498.5811,
    "CFIMRCLP": 20019.16,
    "CFITRIPT-E": 12000.0,
    "CHILE": 168.25,
    "COPEC": 6515.0,
    "ENELAM": 83.92,
    "ITAUCL": 19000.0,
    "LTM": 21.62,
    "SQM-B": 79302.0,
}

INSTRUMENTOS_META = {
    "ABC": {"nombre": "Abc S.A.", "tipo": "accion", "fmt": ".4f"},
    "AGUAS-A": {"nombre": "Aguas Andinas S.A.", "tipo": "accion", "fmt": ".4f"},
    "CENCOSUD": {"nombre": "Cencosud S.A.", "tipo": "accion", "fmt": ".4f"},
    "CHILE": {"nombre": "Banco De Chile", "tipo": "accion", "fmt": ".4f"},
    "COPEC": {"nombre": "Empresas Copec S.A.", "tipo": "accion", "fmt": ".4f"},
    "ENELAM": {"nombre": "Enel Americas S.A.", "tipo": "accion", "fmt": ".4f"},
    "ITAUCL": {"nombre": "Banco Itau Chile", "tipo": "accion", "fmt": ".4f"},
    "LTM": {"nombre": "Latam Airlines Group S.A.", "tipo": "accion", "fmt": ".4f"},
    "SQM-B": {"nombre": "Sociedad Quimica Y Minera De Chile S.A.", "tipo": "accion", "fmt": ".4f"},
    "CFIARRAA-E": {"nombre": "Cfiarraa-E", "tipo": "cfi", "fmt": ".4f"},
    "CFIMRCLP": {"nombre": "Moneda Renta Clp Fi, Serie A", "tipo": "cfi", "fmt": ".4f"},
    "CFITRIPT-E": {"nombre": "Cfitript-E", "tipo": "cfi", "fmt": ".4f"},
}


def calcular_el(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    # ── Leer posiciones desde cartola_data.json si está disponible ──────────
    data = _cargar_json()
    el = data.get("el", {}) if data else {}

    # Acciones
    if el.get("acciones"):
        acciones_raw = [
            (a["nem"], INSTRUMENTOS_META.get(a["nem"], {}).get("nombre", a["nem"]),
             a["cant_activo"], a["cant_pasivo"], a["precio_cartola"])
            for a in el["acciones"]
        ]
    else:
        acciones_raw = EL_ACCIONES

    # CFIs EL
    if el.get("cfis"):
        cfis_raw = [
            (c["nem"], INSTRUMENTOS_META.get(c["nem"], {}).get("nombre", c["nem"]),
             c["cantidad"], c["precio_compra"], c["precio_cartola"])
            for c in el["cfis"]
        ]
    else:
        cfis_raw = EL_CFI

    # SIMs
    if el.get("sims"):
        sims_raw = [
            (s["instrumento"], s["cantidad"],
             _to_date(s["f_venta"]), s["monto_venta"],
             _to_date(s["f_compra"]), s["monto_compra"])
            for s in el["sims"]
        ]
    else:
        sims_raw = EL_SIM

    # Caja y ops_liquidar
    caja_el = el.get("caja", CAJA_EL)
    ops_liquidar = el.get("ops_liquidar", OPS_LIQUIDAR)

    # ── Calcular ─────────────────────────────────────────────────────────────
    acciones = []
    for nem, nombre, cant_a, cant_p, p_c in acciones_raw:
        p = precios.get(nem, p_c)
        va = cant_a * p
        vp = cant_p * p
        acciones.append({
            "nem": nem, "nombre": nombre,
            "cant_activo": cant_a, "cant_pasivo": cant_p,
            "precio_cartola": p_c, "precio_hoy": p,
            "valor_activo": va, "valor_pasivo": vp,
            "valor_neto": va + vp,
            "var_pct": (p - p_c) / p_c if p_c else 0,
        })

    cfis = []
    for nem, nombre, cant, p_comp, p_cart in cfis_raw:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    sims = []
    for inst, cant, f_vta, m_vta, f_cpra, m_cpra in sims_raw:
        total_days = (f_cpra - f_vta).days
        elapsed = max(0, min((hoy - f_vta).days, total_days))
        amort = m_vta + (m_cpra - m_vta) * elapsed / total_days if total_days else m_cpra
        p = precios.get(inst, 0)
        vm = cant * p
        sims.append({
            "instrumento": inst, "cantidad": cant,
            "f_venta": f_vta, "monto_venta": m_vta,
            "f_compra": f_cpra, "monto_compra": m_cpra,
            "monto_amortizado": amort, "valor_mercado": vm,
            "resultado": vm - amort,
            "dias_restantes": max(0, (f_cpra - hoy).days),
            "vencida": hoy >= f_cpra,
        })

    tot_acc_neto = sum(a["valor_neto"] for a in acciones)
    tot_cfi = sum(c["valor_mercado"] for c in cfis)
    tot_sim_amort = sum(s["monto_amortizado"] for s in sims)

    patrimonio = caja_el + ops_liquidar + tot_acc_neto + tot_cfi - tot_sim_amort
    return {
        "acciones": acciones, "cfis": cfis, "sims": sims,
        "tot_acc_activo": sum(a["valor_activo"] for a in acciones),
        "tot_acc_pasivo": sum(a["valor_pasivo"] for a in acciones),
        "tot_acc_neto": tot_acc_neto,
        "tot_cfi": tot_cfi,
        "tot_sim_amort": tot_sim_amort,
        "tot_sim_vm": sum(s["valor_mercado"] for s in sims),
        "tot_sim_resultado": sum(s["resultado"] for s in sims),
        "caja": caja_el,
        "ops_liquidar": ops_liquidar,
        "patrimonio_clp": patrimonio,
        "patrimonio_uf": patrimonio / precios.get("UF", 39_841.72),
        "patrimonio_usd": patrimonio / precios.get("USD", 927.46),
    }


def calcular_emf(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    # ── Leer posiciones desde cartola_data.json si está disponible ──────────
    data = _cargar_json()
    emf = data.get("emf", {}) if data else {}

    # CFIs EMF
    if emf.get("cfis"):
        cfis_raw = [
            (c["nem"], INSTRUMENTOS_META.get(c["nem"], {}).get("nombre", c["nem"]),
             c["cantidad"], c["precio_compra"], c["precio_cartola"])
            for c in emf["cfis"]
        ]
    else:
        cfis_raw = EMF_CFI

    # Forwards
    if emf.get("fwds"):
        fwds_raw = [
            (f["folio"], f["tipo"], f["usd"], f["tc_fwd"],
             _to_date(f["f_inicio"]), _to_date(f["f_termino"]))
            for f in emf["fwds"]
        ]
    else:
        fwds_raw = EMF_FWD

    # Caja
    caja_emf = emf.get("caja", CAJA_EMF)

    # ── Calcular ─────────────────────────────────────────────────────────────
    cfis = []
    for nem, nombre, cant, p_comp, p_cart in cfis_raw:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    spot = precios.get("USD", 927.46)
    fwds = []
    compra_usd = venta_usd = 0
    for folio, tipo, usd, tc_fwd, f_ini, f_term in fwds_raw:
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
    tot_fwd = sum(f["resultado"] for f in fwds)

    patrimonio = caja_emf + tot_cfi
    return {
        "cfis": cfis, "fwds": fwds,
        "tot_cfi": tot_cfi, "tot_fwd": tot_fwd,
        "compra_usd": compra_usd, "venta_usd": venta_usd,
        "descalce_usd": compra_usd - venta_usd,
        "caja": caja_emf,
        "patrimonio_clp": patrimonio,
        "patrimonio_uf": patrimonio / precios.get("UF", 39_841.72),
        "patrimonio_usd": patrimonio / precios.get("USD", 927.46),
    }


def cargar_datos_cartola(path=None):
    """Lee cartola_data.json y retorna el dict, o None si no existe."""
    import os, json
    if path is None:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cartola_data.json")
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None
