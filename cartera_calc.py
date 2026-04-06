"""
Lógica de cálculo de cartera.

Lee posiciones de cartola_data.json (generado por parsear_cartola.py).
Si no existe, usa los valores base de la cartola 06/04/2026 como fallback.
"""
from datetime import date
from pathlib import Path
import json

# ── Fallback: valores base cartola 06/04/2026 ────────────────────────────────

_EL_ACCIONES_BASE = [
    ("ABC",     "ABC S.A.",                  23_210_430,          0,  12.15),
    ("AGUAS-A", "Aguas Andinas S.A.",          1_380_469,  -438_600, 352.30),
    ("CENCOSUD","Cencosud S.A.",                  86_229,          0, 2_500.00),
    ("CHILE",   "Banco de Chile",             5_000_000,           0,  167.00),
    ("COPEC",   "Empresas Copec S.A.",           18_769,     -2_286, 6_527.90),
    ("ENELAM",  "Enel Americas S.A.",          2_300_000,-10_858_102,   81.98),
    ("ITAUCL",  "Banco Itau Chile",             3_801,           0, 20_200.00),
    ("LTM",     "Latam Airlines Group S.A.", 14_770_737,-82_514_437,   22.80),
]

_EL_CFI_BASE = [
    ("CFIARRAA-E", "CFI Arraa Serie E",          4_187, 48_138.4240, 51_702.0000),
    ("CFIMRCLP",   "Moneda Renta CLP FI Serie A",11_172, 19_592.0000, 20_102.1800),
    ("CFITRIPT-E", "CFI Trip-T Serie E",          1_471, 13_280.7610, 12_000.0000),
]

_EL_SIM_BASE = [
    ("AGUAS-A",   438_600, date(2026,3,26),  152_194_200, date(2026,4,23),  152_861_837),
    ("COPEC",       2_286, date(2026,3,26),   14_401_800, date(2026,4,23),   14_464_976),
    ("ENELAM",  1_816_050, date(2026,3,12),  139_091_270, date(2026,4, 9),  139_701_462),
    ("ENELAM",     75_000, date(2026,3,16),    6_000_000, date(2026,4,15),    6_028_200),
    ("ENELAM",  1_061_332, date(2026,3,20),   85_437_226, date(2026,4,17),   85_811_982),
    ("ENELAM",  7_905_720, date(2026,4, 1),  640_284_263, date(2026,5,29),  646_226_202),
    ("LTM",    57_217_627, date(2026,3,12),1_259_932_147, date(2026,4, 9),1_265_459_369),
    ("LTM",    20_000_000, date(2026,3,20),  446_200_000, date(2026,4,17),  447_878_000),
    ("LTM",     5_296_810, date(2026,4, 1),  125_004_716, date(2026,5,29),  126_164_717),
]

_EMF_CFI_BASE = [
    ("CFIARRAA-E", "CFI Arraa Serie E", 500, 47_154.0000, 51_702.0000),
]

_EMF_FWD_BASE = [
    (1834140, "C", 500_000, 916.86, date(2026,4, 1), date(2026,4, 8)),
    (1834324, "V", 500_000, 922.32, date(2026,4, 2), date(2026,5, 5)),
    (1832341, "V", 500_000, 915.55, date(2026,3,20), date(2026,4,10)),
    (1832343, "V", 500_000, 921.35, date(2026,3,20), date(2026,4,10)),
    (1832345, "V", 500_000, 926.25, date(2026,3,20), date(2026,4,10)),
    (1833037, "V", 500_000, 916.40, date(2026,3,25), date(2026,4, 9)),
    (1833221, "V", 500_000, 920.90, date(2026,3,26), date(2026,4, 8)),
]

_CAJA_EL_BASE  = -42_693_925
_CAJA_EMF_BASE =  19_216_120

PRECIOS_DEFAULT = {
    "UF":          39_841.72,
    "USD":            922.17,
    "EUR":          1_064.37,
    "ABC":             12.15,
    "AGUAS-A":        352.30,
    "CENCOSUD":     2_500.00,
    "CHILE":          167.00,
    "COPEC":        6_527.90,
    "ENELAM":          81.98,
    "ITAUCL":      20_200.00,
    "LTM":             22.80,
    "CFIARRAA-E":  51_702.0000,
    "CFIMRCLP":    20_102.18,
    "CFITRIPT-E":  12_000.00,
}

INSTRUMENTOS_META = {
    "ABC":        {"nombre": "ABC S.A.",                        "tipo": "accion", "fmt": ".4f"},
    "AGUAS-A":    {"nombre": "Aguas Andinas S.A.",              "tipo": "accion", "fmt": ".2f"},
    "CENCOSUD":   {"nombre": "Cencosud S.A.",                   "tipo": "accion", "fmt": ".2f"},
    "CHILE":      {"nombre": "Banco de Chile",                  "tipo": "accion", "fmt": ".2f"},
    "COPEC":      {"nombre": "Empresas Copec S.A.",             "tipo": "accion", "fmt": ".2f"},
    "ENELAM":     {"nombre": "Enel Americas S.A.",              "tipo": "accion", "fmt": ".2f"},
    "ITAUCL":     {"nombre": "Banco Itau Chile",                "tipo": "accion", "fmt": ".2f"},
    "LTM":        {"nombre": "Latam Airlines Group S.A.",       "tipo": "accion", "fmt": ".4f"},
    "CFIARRAA-E": {"nombre": "CFI Arraa Serie E",               "tipo": "cfi",    "fmt": ".4f"},
    "CFIMRCLP":   {"nombre": "Moneda Renta CLP FI Serie A",     "tipo": "cfi",    "fmt": ".2f"},
    "CFITRIPT-E": {"nombre": "CFI Trip-T Serie E",              "tipo": "cfi",    "fmt": ".4f"},
}

_NOMBRES_ACCION = {a[0]: a[1] for a in _EL_ACCIONES_BASE}
_NOMBRES_CFI    = {
    "CFIARRAA-E": "CFI Arraa Serie E",
    "CFIMRCLP":   "Moneda Renta CLP FI Serie A",
    "CFITRIPT-E": "CFI Trip-T Serie E",
}


# ── Carga de cartola_data.json ────────────────────────────────────────────────

def _cargar_cartola() -> dict | None:
    path = Path(__file__).parent / "cartola_data.json"
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _parse_date(s: str) -> date:
    return date.fromisoformat(s)


# ── Acceso a posiciones ───────────────────────────────────────────────────────

def _get_el_acciones(cartola: dict | None):
    if cartola and cartola.get("el", {}).get("acciones"):
        out = []
        for a in cartola["el"]["acciones"]:
            nem    = a["nem"]
            nombre = _NOMBRES_ACCION.get(nem, nem)
            out.append((nem, nombre, a["cant_activo"], a["cant_pasivo"], a["precio_cartola"]))
        return out
    return _EL_ACCIONES_BASE


def _get_el_cfis(cartola: dict | None):
    if cartola and cartola.get("el", {}).get("cfis"):
        out = []
        for c in cartola["el"]["cfis"]:
            nem    = c["nem"]
            nombre = _NOMBRES_CFI.get(nem, nem)
            out.append((nem, nombre, c["cantidad"], c["precio_compra"], c["precio_cartola"]))
        return out
    return _EL_CFI_BASE


def _get_el_sims(cartola: dict | None):
    if cartola and cartola.get("el", {}).get("sims"):
        out = []
        for s in cartola["el"]["sims"]:
            out.append((
                s["instrumento"],
                s["cantidad"],
                _parse_date(s["f_venta"]),
                s["monto_venta"],
                _parse_date(s["f_compra"]),
                s["monto_compra"],
            ))
        return out
    return _EL_SIM_BASE


def _get_emf_cfis(cartola: dict | None):
    if cartola and cartola.get("emf", {}).get("cfis"):
        out = []
        for c in cartola["emf"]["cfis"]:
            nem    = c["nem"]
            nombre = _NOMBRES_CFI.get(nem, nem)
            out.append((nem, nombre, c["cantidad"], c["precio_compra"], c["precio_cartola"]))
        return out
    return _EMF_CFI_BASE


def _get_emf_fwds(cartola: dict | None):
    if cartola and cartola.get("emf", {}).get("fwds"):
        out = []
        for f in cartola["emf"]["fwds"]:
            out.append((
                f["folio"],
                f["tipo"],
                f["usd"],
                f["tc_fwd"],
                _parse_date(f["f_inicio"]),
                _parse_date(f["f_termino"]),
            ))
        return out
    return _EMF_FWD_BASE


def _get_caja_el(cartola: dict | None) -> int:
    if cartola and cartola.get("el", {}).get("caja") is not None:
        return cartola["el"]["caja"]
    return _CAJA_EL_BASE


def _get_caja_emf(cartola: dict | None) -> int:
    if cartola and cartola.get("emf", {}).get("caja") is not None:
        return cartola["emf"]["caja"]
    return _CAJA_EMF_BASE


# ── Cálculos ──────────────────────────────────────────────────────────────────

def calcular_el(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    cartola = _cargar_cartola()
    el_acciones = _get_el_acciones(cartola)
    el_cfis     = _get_el_cfis(cartola)
    el_sims     = _get_el_sims(cartola)
    caja_el     = _get_caja_el(cartola)

    acciones = []
    for nem, nombre, cant_a, cant_p, p_c in el_acciones:
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
    for nem, nombre, cant, p_comp, p_cart in el_cfis:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    sims = []
    for inst, cant, f_vta, m_vta, f_cpra, m_cpra in el_sims:
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

    patrimonio = caja_el + tot_acc_neto + tot_cfi - tot_sim_amort
    return {
        "acciones": acciones, "cfis": cfis, "sims": sims,
        "tot_acc_activo":  sum(a["valor_activo"]  for a in acciones),
        "tot_acc_pasivo":  sum(a["valor_pasivo"]  for a in acciones),
        "tot_acc_neto":    tot_acc_neto,
        "tot_cfi":         tot_cfi,
        "tot_sim_amort":   tot_sim_amort,
        "tot_sim_vm":      sum(s["valor_mercado"] for s in sims),
        "tot_sim_resultado": sum(s["resultado"]   for s in sims),
        "caja":            caja_el,
        "ops_liquidar":    0,
        "patrimonio_clp":  patrimonio,
        "patrimonio_uf":   patrimonio / precios.get("UF",  39_841.72),
        "patrimonio_usd":  patrimonio / precios.get("USD",    922.17),
    }


def calcular_emf(precios, hoy=None):
    if hoy is None:
        hoy = date.today()

    cartola  = _cargar_cartola()
    emf_cfis = _get_emf_cfis(cartola)
    emf_fwds = _get_emf_fwds(cartola)
    caja_emf = _get_caja_emf(cartola)

    cfis = []
    for nem, nombre, cant, p_comp, p_cart in emf_cfis:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    spot = precios.get("USD", 922.17)
    fwds = []
    compra_usd = venta_usd = 0
    for folio, tipo, usd, tc_fwd, f_ini, f_term in emf_fwds:
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
        "patrimonio_usd":  patrimonio / precios.get("USD",    922.17),
    }
