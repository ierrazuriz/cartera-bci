from datetime import date
import json
import os
from datetime import datetime
"""
Lógica de cálculo de cartera — generado automáticamente desde cartola BCI.
Posiciones base: cartola 17/04/2026.
NO editar manualmente — se sobreescribe con cada sync.
"""

# ── EL LTDA (76.677.950-6) ─────────────────────────────────────────────────────
# (nemotécnico, nombre, cant_activo, cant_pasivo, precio_cartola)
EL_ACCIONES = [
    ("ABC", "Abc S.A.", 23_210_430, 0, 12.86),
    ("AGUAS-A", "Aguas Andinas S.A.", 1_819_069, 0, 362.0),
    ("CENCOSUD", "Cencosud S.A.", 86_229, 0, 2473.0),
    ("CHILE", "Banco De Chile", 5_000_000, 0, 180.5),
    ("COPEC", "Empresas Copec S.A.", 21_055, 0, 6817.0),
    ("ENELAM", "Enel Americas S.A.", 12_158_102, 0, 86.97),
    ("ITAUCL", "Banco Itau Chile", 3_801, 0, 19969.0),
    ("LTM", "Latam Airlines Group S.A.", 77_285_174, 0, 23.58),
]

# (nemotécnico, nombre, cantidad, precio_compra, precio_cartola)
EL_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 4_187, 48138.424, 57622.0),
    ("CFIMRCLP", "Moneda Renta Clp Fi, Serie A", 11_172, 19592.0, 20122.91),
    ("CFITRIPT-E", "Cfitript-E", 1_471, 13280.761, 12000.0),
]

# (instrumento, cantidad, f_venta, monto_venta, f_compra, monto_compra)
EL_SIM = [
    ("AGUAS-A", 438_600, date(2026,3,26), 152_194_200, date(2026,4,23), 152_861_837),
    ("COPEC", 2_286, date(2026,3,26), 14_401_800, date(2026,4,23), 14_464_976),
    ("ENELAM", 816_050, date(2026,4,9), 68_540_040, date(2026,5,11), 68_883_678),
    ("ENELAM", 7_905_720, date(2026,4,1), 640_284_263, date(2026,5,29), 646_226_202),
    ("LTM", 45_034_136, date(2026,4,9), 1_051_547_076, date(2026,5,11), 1_056_820_573),
    ("LTM", 5_296_810, date(2026,4,1), 125_004_716, date(2026,5,29), 126_164_717),
]

# ── EMF SPA (77.209.686-0) ──────────────────────────────────────────────────────
EMF_CFI = [
    ("CFIARRAA-E", "Cfiarraa-E", 500, 47154.0, 57622.0),
]

# (folio, tipo C/V, usd, tc_fwd, f_inicio, f_termino)
EMF_FWD = [
    (1835288, "C", 500_000, 891.8, date(2026,4,8), date(2026,5,5)),
    (1835290, "C", 500_000, 892.15, date(2026,4,8), date(2026,5,5)),
    (1836514, "C", 250_000, 887.06, date(2026,4,15), date(2026,5,5)),
    (1834324, "V", 500_000, 922.32, date(2026,4,2), date(2026,5,5)),
    (1836070, "V", 500_000, 897.82, date(2026,4,13), date(2026,5,5)),
]

# Cajas (saldo cartola 17/04/2026)
CAJA_EL       = -306_785_070
OPS_LIQUIDAR  = 4_512_816
CAJA_EMF      = 73_351_120

# Precios base (cartola 17/04/2026)
PRECIOS_DEFAULT = {
    "UF": 39947.58,
    "USD": 886.32,
    "EUR": 1043.96,
    "ABC": 12.86,
    "AGUAS-A": 362.0,
    "CENCOSUD": 2473.0,
    "CFIARRAA-E": 57622.0,
    "CFIMRCLP": 20122.91,
    "CFITRIPT-E": 12000.0,
    "CHILE": 180.5,
    "COPEC": 6817.0,
    "ENELAM": 86.97,
    "ITAUCL": 19969.0,
    "LTM": 23.58,
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
    "CFIARRAA-E": {"nombre": "Cfiarraa-E", "tipo": "cfi", "fmt": ".4f"},
    "CFIMRCLP": {"nombre": "Moneda Renta Clp Fi, Serie A", "tipo": "cfi", "fmt": ".4f"},
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

# -- Carga dinamica desde JSON --
CARTOLA_FILE = os.path.join(os.path.dirname(__file__), "cartola_data.json")
_NOMBRES = {m: v["nombre"] for m, v in INSTRUMENTOS_META.items()}
_EL_ACCIONES_DEFAULT = EL_ACCIONES
_EL_CFI_DEFAULT = EL_CFI
_EL_SIM_DEFAULT = EL_SIM
_EMF_CFI_DEFAULT = EMF_CFI
_EMF_FWD_DEFAULT = EMF_FWD
_CAJA_EL_DEFAULT = CAJA_EL
_OPS_LIQUIDAR_DEFAULT = OPS_LIQUIDAR
_CAJA_EMF_DEFAULT = CAJA_EMF


def cargar_datos_cartola(path=None):
    p = path or CARTOLA_FILE
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
                    return None
