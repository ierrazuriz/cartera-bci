"""
Lógica de cálculo de cartera — EL LTDA y EMF SPA.

=== FIX 2 APLICADO ===
Antes: posiciones hardcodeadas como constantes Python (se perdían en cada deploy).
Ahora: lee dinámicamente desde cartola_data.json. Si no existe, usa defaults.

Las funciones calcular_el() y calcular_emf() ahora aceptan un parámetro
opcional `datos_cartola` (dict cargado desde JSON).
"""

import json
import os
from datetime import date, datetime
from pathlib import Path

DATA_DIR = Path(os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__))))
CARTOLA_FILE = DATA_DIR / "cartola_data.json"

# ── Defaults (cartola 10/04/2026) — se usan si cartola_data.json no existe ──

_EL_ACCIONES_DEFAULT = [
    ("ABC",      "Abc S.A.",                  23_210_430, 0, 12.4),
    ("AGUAS-A",  "Aguas Andinas S.A.",         1_819_069, 0, 352.49),
    ("CENCOSUD", "Cencosud S.A.",                 86_229, 0, 2560.0),
    ("CHILE",    "Banco De Chile",             5_000_000, 0, 171.99),
    ("COPEC",    "Empresas Copec S.A.",           21_055, 0, 6565.9),
    ("ENELAM",   "Enel Americas S.A.",        12_158_102, 0, 84.1),
    ("ITAUCL",   "Banco Itau Chile",              3_801, 0, 20950.0),
    ("LTM",     "Latam Airlines Group S.A.", 85_101_683, 0, 23.79),
]

_EL_CFI_DEFAULT = [
    ("CFIARRAA-E", "Cfiarraa-E",                   4_187, 48138.424, 50500.0),
    ("CFIMRCLP",   "Moneda Renta Clp Fi, Serie A", 11_172, 19592.0,  19425.24),
    ("CFITRIPT-E", "Cfitript-E",                    1_471, 13280.761, 12000.0),
]

_EL_SIM_DEFAULT = [
    ("AGUAS-A",  438_600,   "2026-03-26", 152_194_200, "2026-04-23", 152_861_837),
    ("COPEC",      2_286,   "2026-03-26",  14_401_800, "2026-04-23",  14_464_976),
    ("ENELAM",    75_000,   "2026-03-16",   6_000_000, "2026-04-15",   6_028_200),
    ("ENELAM", 1_061_332,   "2026-03-20",  85_437_226, "2026-04-17",  85_811_982),
    ("ENELAM",   816_050,   "2026-04-09",  68_540_040, "2026-05-11",  68_883_678),
    ("ENELAM", 7_905_720,   "2026-04-01", 640_284_263, "2026-05-29", 646_226_202),
    ("LTM",   20_000_000,   "2026-03-20", 446_200_000, "2026-04-17", 447_878_000),
    ("LTM",   45_034_136,   "2026-04-09",1_051_547_076,"2026-05-11",1_056_820_573),
    ("LTM",    5_296_810,   "2026-04-01", 125_004_716, "2026-05-29", 126_164_717),
]

_EMF_CFI_DEFAULT = [
    ("CFIARRAA-E", "Cfiarraa-E", 500, 47154.0, 50500.0),
]

_EMF_FWD_DEFAULT = [
    (1835288, "C", 500_000, 891.8,  "2026-04-08", "2026-05-05"),
    (1835290, "C", 500_000, 892.15, "2026-04-08", "2026-05-05"),
    (1834324, "V", 500_000, 922.32, "2026-04-02", "2026-05-05"),
]

_CAJA_EL_DEFAULT = 19
_OPS_LIQUIDAR_DEFAULT = 232_916_818
_CAJA_EMF_DEFAULT = 33_031_120

PRECIOS_DEFAULT = {
    "UF": 39854.94, "USD": 894.17, "EUR": 1046.3,
    "ABC": 12.4, "AGUAS-A": 352.49, "CENCOSUD": 2560.0,
    "CFIARRAA-E": 50500.0, "CFIMRCLP": 19425.24, "CFITRIPT-E": 12000.0,
    "CHILE": 171.99, "COPEC": 6565.9, "ENELAM": 84.1,
    "ITAUCL": 20950.0, "LTM": 23.79,
}

INSTRUMENTOS_META = {
    "ABC":        {"nombre": "Abc S.A.",                  "tipo": "accion", "fmt": ".4f"},
    "AGUAS-A":    {"nombre": "Aguas Andinas S.A.",        "tipo": "accion", "fmt": ".4f"},
    "CENCOSUD":   {"nombre": "Cencosud S.A.",             "tipo": "accion", "fmt": ".4f"},
    "CHILE":      {"nombre": "Banco De Chile",            "tipo": "accion", "fmt": ".4f"},
    "COPEC":      {"nombre": "Empresas Copec S.A.",       "tipo": "accion", "fmt": ".4f"},
    "ENELAM":     {"nombre": "Enel Americas S.A.",        "tipo": "accion", "fmt": ".4f"},
    "ITAUCL":     {"nombre": "Banco Itau Chile",          "tipo": "accion", "fmt": ".4f"},
    "LTM":        {"nombre": "Latam Airlines Group S.A.", "tipo": "accion", "fmt": ".4f"},
    "CFIARRAA-E": {"nombre": "Cfiarraa-E",                "tipo": "cfi",    "fmt": ".4f"},
    "CFIMRCLP":   {"nombre": "Moneda Renta Clp Fi, Serie A", "tipo": "cfi", "fmt": ".4f"},
    "CFITRIPT-E": {"nombre": "Cfitript-E",                "tipo": "cfi",    "fmt": ".4f"},
}

# Mapa de nombres por nemotécnico (para cuando el JSON no trae nombre)
_NOMBRES = {m: v["nombre"] for m, v in INSTRUMENTOS_META.items()}


# ── Carga de datos desde JSON ────────────────────────────────────────────────

def cargar_datos_cartola(path=None):
    """
    Carga cartola_data.json. Retorna dict o None si no existe.
    """
    p = path or CARTOLA_FILE
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _parse_date(s):
    """Convierte string ISO 'YYYY-MM-DD' a date object."""
    if isinstance(s, date):
        return s
    return datetime.strptime(s, "%Y-%m-%d").date()


def _extraer_acciones(datos_cartola):
    """Extrae lista de acciones EL desde JSON o usa defaults."""
    if datos_cartola and "el" in datos_cartola:
        el_data = datos_cartola["el"]
        acciones_raw = el_data.get("acciones", [])
        if acciones_raw:
            result = []
            for a in acciones_raw:
                nem = a["nem"]
                result.append((
                    nem,
                    _NOMBRES.get(nem, nem),
                    int(a.get("cant_activo", 0)),
                    int(a.get("cant_pasivo", 0)),
                    float(a.get("precio_cartola", 0)),
                ))
            return result
    return _EL_ACCIONES_DEFAULT


def _extraer_cfis_el(datos_cartola):
    """Extrae CFIs de EL desde JSON o usa defaults."""
    if datos_cartola and "el" in datos_cartola:
        cfis_raw = datos_cartola["el"].get("cfis", [])
        if cfis_raw:
            result = []
            for c in cfis_raw:
                nem = c["nem"]
                result.append((
                    nem,
                    _NOMBRES.get(nem, nem),
                    int(c.get("cantidad", 0)),
                    float(c.get("precio_compra", 0)),
                    float(c.get("precio_cartola", 0)),
                ))
            return result
    return _EL_CFI_DEFAULT


def _extraer_sims(datos_cartola):
    """Extrae simultáneas de EL desde JSON o usa defaults."""
    if datos_cartola and "el" in datos_cartola:
        sims_raw = datos_cartola["el"].get("sims", [])
        if sims_raw:
            result = []
            for s in sims_raw:
                result.append((
                    s["instrumento"],
                    int(s["cantidad"]),
                    _parse_date(s["f_venta"]),
                    int(s["monto_venta"]),
                    _parse_date(s["f_compra"]),
                    int(s["monto_compra"]),
                ))
            return result
    # Defaults con strings → convertir a date
    return [
        (inst, cant, _parse_date(fv), mv, _parse_date(fc), mc)
        for inst, cant, fv, mv, fc, mc in _EL_SIM_DEFAULT
    ]


def _extraer_cfis_emf(datos_cartola):
    """Extrae CFIs de EMF desde JSON o usa defaults."""
    if datos_cartola and "emf" in datos_cartola:
        cfis_raw = datos_cartola["emf"].get("cfis", [])
        if cfis_raw:
            result = []
            for c in cfis_raw:
                nem = c["nem"]
                result.append((
                    nem,
                    _NOMBRES.get(nem, nem),
                    int(c.get("cantidad", 0)),
                    float(c.get("precio_compra", 0)),
                    float(c.get("precio_cartola", 0)),
                ))
            return result
    return _EMF_CFI_DEFAULT


def _extraer_forwards(datos_cartola):
    """Extrae forwards de EMF desde JSON o usa defaults."""
    if datos_cartola and "emf" in datos_cartola:
        fwds_raw = datos_cartola["emf"].get("fwds", [])
        if fwds_raw:
            result = []
            for f in fwds_raw:
                result.append((
                    int(f["folio"]),
                    f["tipo"],
                    int(f["usd"]),
                    float(f["tc_fwd"]),
                    _parse_date(f["f_inicio"]),
                    _parse_date(f["f_termino"]),
                ))
            return result
    return [
        (folio, tipo, usd, tc, _parse_date(fi), _parse_date(ft))
        for folio, tipo, usd, tc, fi, ft in _EMF_FWD_DEFAULT
    ]


def _extraer_caja_el(datos_cartola):
    if datos_cartola and "el" in datos_cartola:
        return datos_cartola["el"].get("caja", _CAJA_EL_DEFAULT)
    return _CAJA_EL_DEFAULT


def _extraer_ops_liquidar(datos_cartola):
    if datos_cartola and "el" in datos_cartola:
        return datos_cartola["el"].get("ops_liquidar", _OPS_LIQUIDAR_DEFAULT)
    return _OPS_LIQUIDAR_DEFAULT


def _extraer_caja_emf(datos_cartola):
    if datos_cartola and "emf" in datos_cartola:
        return datos_cartola["emf"].get("caja", _CAJA_EMF_DEFAULT)
    return _CAJA_EMF_DEFAULT


# ── Cálculos ─────────────────────────────────────────────────────────────────

def calcular_el(precios, hoy=None, datos_cartola=None):
    if hoy is None:
        hoy = date.today()

    # Cargar datos dinámicamente
    el_acciones = _extraer_acciones(datos_cartola)
    el_cfis_data = _extraer_cfis_el(datos_cartola)
    el_sims_data = _extraer_sims(datos_cartola)
    caja_el = _extraer_caja_el(datos_cartola)
    ops_liquidar = _extraer_ops_liquidar(datos_cartola)

    acciones = []
    for nem, nombre, cant_a, cant_p, p_c in el_acciones:
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
    for nem, nombre, cant, p_comp, p_cart in el_cfis_data:
        p = precios.get(nem, p_cart)
        cfis.append({
            "nem": nem, "nombre": nombre, "cantidad": cant,
            "precio_compra": p_comp, "precio_cartola": p_cart,
            "precio_hoy": p, "valor_mercado": cant * p,
            "var_pct": (p - p_cart) / p_cart if p_cart else 0,
        })

    sims = []
    for inst, cant, f_vta, m_vta, f_cpra, m_cpra in el_sims_data:
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


def calcular_emf(precios, hoy=None, datos_cartola=None):
    if hoy is None:
        hoy = date.today()

    emf_cfis_data = _extraer_cfis_emf(datos_cartola)
    emf_fwds_data = _extraer_forwards(datos_cartola)
    caja_emf = _extraer_caja_emf(datos_cartola)

    cfis = []
    for nem, nombre, cant, p_comp, p_cart in emf_cfis_data:
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
    for folio, tipo, usd, tc_fwd, f_ini, f_term in emf_fwds_data:
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
