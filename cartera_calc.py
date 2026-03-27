"""
Lógica de cálculo de cartera — independiente de Excel.
Posiciones base: cartola 27/03/2026.
"""
from datetime import date

# ── EL LTDA (76.677.950-6) ──────────────────────────────────────────────────
# (nemotécnico, nombre, cant_activo, cant_pasivo, precio_cartola)
EL_ACCIONES = [
    ("ABC",      "ABC S.A.",                  23_210_430,         0,   12.09),
    ("AGUAS-A",  "Aguas Andinas S.A.",          1_939_069,  -120_000,  348.00),
    ("CENCOSUD", "Cencosud S.A.",                  86_229,         0, 2_340.00),
    ("CHILE",    "Banco de Chile",             6_000_000, -1_000_000,  168.00),
    ("COPEC",    "Empresas Copec S.A.",            21_055,         0, 6_290.00),
    ("ENELAM",   "Enel Americas S.A.",         13_158_102,         0,   77.25),
    ("LTM",      "Latam Airlines Group S.A.", 104_595_390, -5_000_000,  22.83),
]

# (nemotécnico, nombre, cantidad, precio_compra, precio_cartola)
EL_CFI = [
    ("CFIARRAA-E", "CFI Arraa Serie E",             3_801, 47_777.8556, 53_279.4282),
    ("CFIMRCLP",   "Moneda Renta CLP FI, Serie A", 11_172, 19_592.0000, 20_090.7100),
    ("CFITRIPT-E", "CFI Tript Serie E",              1_471, 13_280.7610, 12_000.0000),
]

# (instrumento, cantidad, f_venta, monto_venta, f_compra, monto_compra)
EL_SIM = [
    ("AGUAS-A",   558_600, date(2026,3,26), 193_834_200, date(2026,4,23), 194_684_501),
    ("CHILE",     238_830, date(2026,3,20),  42_370_830, date(2026,4,17),  42_556_688),
    ("COPEC",       2_286, date(2026,3,26),  14_401_800, date(2026,4,23),  14_464_976),
    ("ENELAM", 7_905_720,  date(2026,3, 6), 639_335_576, date(2026,4, 1), 641_939_721),
    ("ENELAM", 1_816_050,  date(2026,3,12), 139_091_270, date(2026,4, 9), 139_701_462),
    ("ENELAM",    75_000,  date(2026,3,16),   6_000_000, date(2026,4,15),   6_028_200),
    ("ENELAM", 1_061_332,  date(2026,3,20),  85_437_226, date(2026,4,17),  85_811_982),
    ("LTM",   15_310_300,  date(2026,3, 6), 335_601_776, date(2026,4, 1), 336_968_986),
    ("LTM",   69_285_090,  date(2026,3,12),1_525_657_682,date(2026,4, 9),1_532_350_621),
    ("LTM",   20_000_000,  date(2026,3,20), 446_200_000, date(2026,4,17), 447_878_000),
]

# ── EMF SPA (77.209.686-0) ───────────────────────────────────────────────────
EMF_CFI = [
    ("CFIARRAA-E", "CFI Arraa Serie E", 500, 47_154.0000, 53_279.4282),
]

# (folio, tipo C/V, usd, tc_fwd, f_inicio, f_termino)
EMF_FWD = [
    (1832537, "C", 500_000, 908.20, date(2026,3,23), date(2026,4, 2)),
    (1833221, "V", 500_000, 920.90, date(2026,3,26), date(2026,4, 8)),
    (1833037, "V", 500_000, 916.40, date(2026,3,25), date(2026,4, 9)),
    (1832341, "V", 500_000, 915.55, date(2026,3,20), date(2026,4,10)),
    (1832343, "V", 500_000, 921.35, date(2026,3,20), date(2026,4,10)),
    (1832345, "V", 500_000, 926.25, date(2026,3,20), date(2026,4,10)),
    (1831846, "V", 500_000, 912.98, date(2026,3,18), date(2026,4, 2)),
]

# Cajas fijas (saldo cartola 27/03/2026)
CAJA_BRUTA_EL = 555_976_435   # saldo final período
OPS_LIQUIDAR  = 205_558_872   # compras pendientes de liquidación (activo)
CAJA_EL       = 350_417_563   # = CAJA_BRUTA_EL - OPS_LIQUIDAR (caja disponible)
CAJA_EMF      =  16_826_120

# Precios base (cartola 27/03/2026) — usados como valor inicial
PRECIOS_DEFAULT = {
    "UF":          39_841.72,
    "USD":            923.25,
    "EUR":          1_064.63,
    "ABC":             12.09,
    "AGUAS-A":        348.00,
    "CENCOSUD":     2_340.00,
    "CHILE":          168.00,
    "COPEC":        6_290.00,
    "ENELAM":          77.25,
    "LTM":             22.83,
    "CFIARRAA-E":  53_279.4282,
    "CFIMRCLP":    20_090.71,
    "CFITRIPT-E":  12_000.00,
}

INSTRUMENTOS_META = {
    "ABC":         {"nombre": "ABC S.A.",                    "tipo": "accion",  "fmt": ".4f"},
    "AGUAS-A":     {"nombre": "Aguas Andinas S.A.",           "tipo": "accion",  "fmt": ".2f"},
    "CENCOSUD":    {"nombre": "Cencosud S.A.",                "tipo": "accion",  "fmt": ".2f"},
    "CHILE":       {"nombre": "Banco de Chile",               "tipo": "accion",  "fmt": ".2f"},
    "COPEC":       {"nombre": "Empresas Copec S.A.",          "tipo": "accion",  "fmt": ".2f"},
    "ENELAM":      {"nombre": "Enel Americas S.A.",           "tipo": "accion",  "fmt": ".2f"},
    "LTM":         {"nombre": "Latam Airlines Group S.A.",    "tipo": "accion",  "fmt": ".4f"},
    "CFIARRAA-E":  {"nombre": "CFI Arraa Serie E",            "tipo": "cfi",     "fmt": ".4f"},
    "CFIMRCLP":    {"nombre": "Moneda Renta CLP FI, S.A",    "tipo": "cfi",     "fmt": ".4f"},
    "CFITRIPT-E":  {"nombre": "CFI Tript Serie E",            "tipo": "cfi",     "fmt": ".4f"},
}


def calcular_el(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    acciones = []
    for nem, nombre, cant_a, cant_p, p_c in EL_ACCIONES:
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
    for nem, nombre, cant, p_comp, p_cart in EL_CFI:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p,
            "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    sims = []
    for inst, cant, f_vta, m_vta, f_cpra, m_cpra in EL_SIM:
        total_days = (f_cpra - f_vta).days
        elapsed = max(0, min((hoy - f_vta).days, total_days))
        amort = m_vta + (m_cpra - m_vta) * elapsed / total_days if total_days else m_cpra
        p = precios.get(inst, 0)
        vm = cant * p
        sims.append({
            "instrumento": inst, "cantidad": cant,
            "f_venta": f_vta, "monto_venta": m_vta,
            "f_compra": f_cpra, "monto_compra": m_cpra,
            "monto_amortizado": amort,
            "valor_mercado": vm,
            "resultado": vm - amort,
            "dias_restantes": max(0, (f_cpra - hoy).days),
            "vencida": hoy >= f_cpra,
        })

    tot_acc_neto = sum(a["valor_neto"] for a in acciones)
    tot_cfi      = sum(c["valor_mercado"] for c in cfis)
    tot_sim_amort = sum(s["monto_amortizado"] for s in sims)
    # Patrimonio = Caja_disp + OPS_LIQUIDAR + Acciones_neto + CFI - Sim_amort
    # equivale a: CAJA_BRUTA + Acciones_neto + CFI - Sim_amort
    patrimonio = CAJA_EL + OPS_LIQUIDAR + tot_acc_neto + tot_cfi - tot_sim_amort

    return {
        "acciones": acciones,
        "cfis": cfis,
        "sims": sims,
        "tot_acc_activo":   sum(a["valor_activo"] for a in acciones),
        "tot_acc_pasivo":   sum(a["valor_pasivo"] for a in acciones),
        "tot_acc_neto":     tot_acc_neto,
        "tot_cfi":          tot_cfi,
        "tot_sim_amort":    tot_sim_amort,
        "tot_sim_vm":       sum(s["valor_mercado"] for s in sims),
        "tot_sim_resultado":sum(s["resultado"] for s in sims),
        "caja":             CAJA_EL,
        "ops_liquidar":     OPS_LIQUIDAR,
        "patrimonio_clp":   patrimonio,
        "patrimonio_uf":    patrimonio / precios.get("UF", 39_841.72),
        "patrimonio_usd":   patrimonio / precios.get("USD", 923.25),
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
            "precio_hoy": p,
            "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    spot = precios.get("USD", 923.25)
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
    tot_fwd = sum(f["resultado"] for f in fwds)
    # Patrimonio BCI no incluye resultado forwards (son informativo/contingente)
    patrimonio = CAJA_EMF + tot_cfi

    return {
        "cfis": cfis,
        "fwds": fwds,
        "tot_cfi":        tot_cfi,
        "tot_fwd":        tot_fwd,
        "compra_usd":     compra_usd,
        "venta_usd":      venta_usd,
        "descalce_usd":   compra_usd - venta_usd,
        "caja":           CAJA_EMF,
        "patrimonio_clp": patrimonio,
        "patrimonio_uf":  patrimonio / precios.get("UF", 39_841.72),
        "patrimonio_usd": patrimonio / precios.get("USD", 923.25),
    }
