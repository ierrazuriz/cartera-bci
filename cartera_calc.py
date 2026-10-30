"""
Lógica de cálculo de cartera — generado automáticamente desde cartola BCI.
Posiciones base: cartola 30/10/2026.
NO editar manualmente — se sobreescribe con cada sync.
"""
from datetime import date

# ── EL LTDA (76.677.950-6) ─────────────────────────────────────────────────────
# (nemotécnico, nombre, cant_activo, cant_pasivo, precio_cartola)
EL_ACCIONES = [
    ("ABC", "Abc S.A.", 23_210_430, 0, 8.54),
    ("AGUAS-A", "Aguas Andinas S.A.", 1_819_069, 0, 325.82),
    ("CENCOMALLS", "Cencosud Shopping S.A.", 201_126, 0, 2190.0),
    ("CENCOSUD", "Cencosud S.A.", 336_229, 0, 1807.9),
    ("CHILE", "Banco De Chile", 7_009_086, 0, 189.01),
    ("CMPC", "Empresas Cmpc S.A.", 150_000, 0, 945.0),
    ("COPEC", "Empresas Copec S.A.", 15_136, 0, 6530.0),
    ("SOCOVESA", "Socovesa S.A.", 10_734, 0, 157.45),
]

# (nemotécnico, nombre, cantidad, precio_compra, precio_cartola)
EL_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 4_132, 48151.1358, 52330.0),
    ("CFITRIPT-E", "Cfitript-E", 1_471, 13280.7614, 14000.0),
]

# (instrumento, cantidad, f_venta, monto_venta, f_compra, monto_compra)
EL_SIM = [
    ("CHILE", 1_141_018, date(2026,10,5), 216_565_216, date(2026,11,4), 217_604_684),
    ("CHILE", 11_561, date(2026,10,13), 2_200_174, date(2026,11,12), 2_210_735),
    ("CHILE", 9_043, date(2026,10,16), 1_733_181, date(2026,11,16), 1_741_778),
]

# ── EMF SPA (77.209.686-0) ──────────────────────────────────────────────────────
EMF_CFI = [
]

# (folio, tipo C/V, usd, tc_fwd, f_inicio, f_termino)
EMF_FWD = [
    (1868899, "C", 500_000, 979.8, date(2026,10,27), date(2026,11,6)),
    (1869060, "C", 500_000, 979.16, date(2026,10,28), date(2026,11,6)),
    (1869063, "C", 500_000, 980.91, date(2026,10,28), date(2026,11,6)),
    (1869065, "C", 500_000, 980.56, date(2026,10,28), date(2026,11,6)),
    (1869085, "C", 500_000, 980.99, date(2026,10,28), date(2026,11,6)),
    (1869374, "V", 500_000, 995.3, date(2026,10,29), date(2026,11,6)),
    (1868123, "V", 500_000, 986.35, date(2026,10,22), date(2026,11,6)),
    (1868124, "V", 500_000, 985.45, date(2026,10,22), date(2026,11,6)),
    (1868408, "V", 500_000, 988.9, date(2026,10,23), date(2026,11,6)),
    (1868409, "V", 500_000, 991.75, date(2026,10,23), date(2026,11,6)),
]

# Cajas (saldo cartola 30/10/2026)
CAJA_EL       = -486_333_377
OPS_LIQUIDAR  = 0
CAJA_EMF      = 97_593_444

# Precios base (cartola 30/10/2026)
PRECIOS_DEFAULT = {
    "UF": 41242.32,
    "USD": 993.22,
    "EUR": 1109.37,
    "ABC": 8.54,
    "AGUAS-A": 325.82,
    "CENCOMALLS": 2190.0,
    "CENCOSUD": 1807.9,
    "CFIARRAA-E": 52330.0,
    "CFITRIPT-E": 14000.0,
    "CHILE": 189.01,
    "CMPC": 945.0,
    "COPEC": 6530.0,
    "SOCOVESA": 157.45,
}

INSTRUMENTOS_META = {
    "ABC": {"nombre": "Abc S.A.", "tipo": "accion", "fmt": ".4f"},
    "AGUAS-A": {"nombre": "Aguas Andinas S.A.", "tipo": "accion", "fmt": ".4f"},
    "CENCOMALLS": {"nombre": "Cencosud Shopping S.A.", "tipo": "accion", "fmt": ".4f"},
    "CENCOSUD": {"nombre": "Cencosud S.A.", "tipo": "accion", "fmt": ".4f"},
    "CHILE": {"nombre": "Banco De Chile", "tipo": "accion", "fmt": ".4f"},
    "CMPC": {"nombre": "Empresas Cmpc S.A.", "tipo": "accion", "fmt": ".4f"},
    "COPEC": {"nombre": "Empresas Copec S.A.", "tipo": "accion", "fmt": ".4f"},
    "SOCOVESA": {"nombre": "Socovesa S.A.", "tipo": "accion", "fmt": ".4f"},
    "CFIARRAA-E": {"nombre": "Cfiarraa-E", "tipo": "cfi", "fmt": ".4f"},
    "CFITRIPT-E": {"nombre": "Cfitript-E", "tipo": "cfi", "fmt": ".4f"},
}


def calcular_el(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    acciones = []
    for nem, nombre, cant_a, cant_p, p_c in EL_ACCIONES:
        p  = precios.get(nem, p_c)
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
    for nem, nombre, cant, p_comp, p_cart in EL_CFI:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    sims = []
    for inst, cant, f_vta, m_vta, f_cpra, m_cpra in EL_SIM:
        total_days = (f_cpra - f_vta).days
        elapsed    = max(0, min((hoy - f_vta).days, total_days))
        amort      = m_vta + (m_cpra - m_vta) * elapsed / total_days if total_days else m_cpra
        p  = precios.get(inst, 0)
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

    tot_acc_neto  = sum(a["valor_neto"]    for a in acciones)
    tot_cfi       = sum(c["valor_mercado"] for c in cfis)
    tot_sim_amort = sum(s["monto_amortizado"] for s in sims)

    patrimonio = CAJA_EL + OPS_LIQUIDAR + tot_acc_neto + tot_cfi - tot_sim_amort
    return {
        "acciones": acciones, "cfis": cfis, "sims": sims,
        "tot_acc_activo":  sum(a["valor_activo"]  for a in acciones),
        "tot_acc_pasivo":  sum(a["valor_pasivo"]  for a in acciones),
        "tot_acc_neto":    tot_acc_neto,
        "tot_cfi":         tot_cfi,
        "tot_sim_amort":   tot_sim_amort,
        "tot_sim_vm":      sum(s["valor_mercado"] for s in sims),
        "tot_sim_resultado": sum(s["resultado"]   for s in sims),
        "caja":            CAJA_EL,
        "ops_liquidar":    OPS_LIQUIDAR,
        "patrimonio_clp":  patrimonio,
        "patrimonio_uf":   patrimonio / precios.get("UF",  39_841.72),
        "patrimonio_usd":  patrimonio / precios.get("USD",    927.46),
    }


def calcular_emf(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    cfis = []
    for nem, nombre, cant, p_comp, p_cart in EMF_CFI:
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
    for folio, tipo, usd, tc_fwd, f_ini, f_term in EMF_FWD:
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

    patrimonio = CAJA_EMF + tot_cfi
    return {
        "cfis": cfis, "fwds": fwds,
        "tot_cfi": tot_cfi, "tot_fwd": tot_fwd,
        "compra_usd": compra_usd, "venta_usd": venta_usd,
        "descalce_usd": compra_usd - venta_usd,
        "caja": CAJA_EMF,
        "patrimonio_clp":  patrimonio,
        "patrimonio_uf":   patrimonio / precios.get("UF",  39_841.72),
        "patrimonio_usd":  patrimonio / precios.get("USD",    927.46),
    }
