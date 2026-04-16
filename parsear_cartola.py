"""
Parser de cartolas BCI.
Extrae de los PDFs de EL LTDA y EMF SPA:
 - precios (UF, USD, EUR, acciones, CFIs)
 - posiciones de acciones (cant_activo, cant_pasivo)
 - posiciones CFI
 - simultáneas EL
 - caja EL y EMF
 - forwards EMF
 - ops_liquidar

=== BUGS CORREGIDOS ===
1. cant_activo = Libre + En Garantía + Saldo a Plazo (antes solo usaba En Garantía)
   "Saldo a Plazo" son acciones propias en simultáneas, NO son pasivo
2. cant_pasivo viene de la línea "Pasivo:" (antes usaba -Saldo_a_Plazo, que era incorrecto)
3. Se extrae ops_liquidar del resumen de inversiones
4. Caja se lee correctamente (ya incluye ops_liquidar cuando es 0)
"""

import re
import json
import io
from datetime import datetime, date
from pathlib import Path
import pdfplumber

DATA_DIR = Path(__file__).parent
DATA_FILE = DATA_DIR / "cartola_data.json"

RUT_EL = "76677950"
RUT_EMF = "77209686"

# ── Helpers ───────────────────────────────────────────────────────────────────

def _num(s: str) -> float:
    """Convierte '1.234,56' → 1234.56 y '-856.291.313' → -856291313."""
    s = s.strip().replace("\xa0", "").replace(" ", "")
    negativo = s.startswith("-")
    s = s.lstrip("-")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(".", "")
    return -float(s) if negativo else float(s)


def _fecha(s: str) -> str:
    """Convierte '26-03-2026' o '26/03/2026' → '2026-03-26'."""
    s = s.strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError(f"Fecha no reconocida: {s!r}")


def _texto_pdf(pdf_bytes: bytes) -> str:
    lineas = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                lineas.append(t)
    return "\n".join(lineas)


# ── Extracción de precios de cabecera ─────────────────────────────────────────

def _extraer_precios_cabecera(texto: str) -> dict:
    precios = {}
    patrones = {
        "UF":  r"Valor UF[:\s]+\$\s*([\d\.]+,\d+)",
        "USD": r"Valor USD\s*\$?\s*([\d\.]+,\d+)",
        "EUR": r"Valor EUR[:\s]+\$\s*([\d\.]+,\d+)",
    }
    for nem, pat in patrones.items():
        m = re.search(pat, texto)
        if m:
            precios[nem] = _num(m.group(1))
    return precios


# ── Extracción de acciones (EL) ───────────────────────────────────────────────
# Formato del PDF (pdfplumber):
# ABC Activo: 0 0 23.210.430 0 11,7800 12,6200 292.915.627 0
# Rubro: Comerciales y Distribuidoras Pasivo: 0 0 0 0 12,6200 0 0
#
# Columnas Activo: Libre | Saldo Préstamo | En Garantía | Saldo a Plazo | Precio Compra | Precio Último | Valor Mercado | Dividendos
# Columnas Pasivo: Libre | Saldo Préstamo | En Garantía | Saldo a Plazo | Precio Último | Valor Mercado | Dividendos

# Regex para línea Activo
_RE_ACCION_ACTIVO = re.compile(
    r"^([A-Z][A-Z0-9\-]+)\s+Activo:\s+"
    r"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+"  # libre, prestamo, en_garantia, a_plazo
    r"([\d\.,]+)\s+([\d\.,]+)",                           # precio_compra, precio_ultimo
    re.MULTILINE,
)

# Regex para línea Pasivo (puede estar en la siguiente línea o dos después)
_RE_ACCION_PASIVO = re.compile(
    r"Pasivo:\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+(-?[\d\.]+)\s+"  # libre, prestamo, en_garantia, a_plazo
    r"([\d\.,]+)\s+(-?[\d\.,]+)",                                            # precio_ultimo, valor_mercado
    re.MULTILINE,
)


def _extraer_acciones(texto: str) -> list:
    """
    Retorna lista de dicts con posiciones de acciones.
    
    FIX: cant_activo = Libre + En Garantía + Saldo a Plazo (total acciones propias)
         cant_pasivo = viene de la línea Pasivo: columna Libre (acciones vendidas en contado de sim)
    """
    acciones = []
    lineas = texto.splitlines()
    
    for i, linea in enumerate(lineas):
        m_act = _RE_ACCION_ACTIVO.match(linea.strip())
        if not m_act:
            continue
            
        nem = m_act.group(1)
        if nem.startswith("CFI"):
            continue
        
        libre = int(_num(m_act.group(2)))
        # prestamo = int(_num(m_act.group(3)))  # no lo usamos
        en_garantia = int(_num(m_act.group(4)))
        a_plazo = int(_num(m_act.group(5)))
        p_compra = _num(m_act.group(6))
        p_ultimo = _num(m_act.group(7))
        
        # FIX 1: cant_activo es la suma de TODAS las posiciones activas
        cant_activo = libre + en_garantia + a_plazo
        
        # FIX 2: buscar la línea Pasivo: en las siguientes líneas
        cant_pasivo = 0
        for j in range(i + 1, min(i + 4, len(lineas))):
            m_pas = _RE_ACCION_PASIVO.search(lineas[j])
            if m_pas:
                pas_libre = int(_num(m_pas.group(1)))
                # Si hay valor negativo en Libre, es la cantidad pasiva
                if pas_libre != 0:
                    cant_pasivo = pas_libre  # ya viene negativo del PDF
                break
        
        acciones.append({
            "nem": nem,
            "cant_activo": cant_activo,
            "cant_pasivo": cant_pasivo,
            "precio_compra": p_compra,
            "precio_cartola": p_ultimo,
        })
    
    return acciones


# ── Extracción de CFIs ────────────────────────────────────────────────────────

_RE_CFI = re.compile(
    r"^(CFI[A-Z0-9\-]+)\s+Activo:\s+"
    r"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+"  # libre, prestamo, en_garantia, a_plazo
    r"([\d\.,]+)\s+([\d\.,]+)",                           # precio_compra, precio_ultimo
    re.MULTILINE,
)

def _extraer_cfis(texto: str) -> list:
    cfis = []
    vistas = set()
    for m in _RE_CFI.finditer(texto):
        nem = m.group(1)
        if nem in vistas:
            continue
        vistas.add(nem)
        # FIX: sum all quantity columns (Libre + En Garantía + Saldo a Plazo)
        libre = int(_num(m.group(2)))
        en_garantia = int(_num(m.group(4)))
        a_plazo = int(_num(m.group(5)))
        cantidad = libre + en_garantia + a_plazo
        cfis.append({
            "nem": nem,
            "cantidad": cantidad,
            "precio_compra": _num(m.group(6)),
            "precio_cartola": _num(m.group(7)),
        })
    return cfis


# ── Extracción de simultáneas (EL) ───────────────────────────────────────────

_RE_SIM1 = re.compile(
    r"^([A-Z][A-Z0-9\-]+)\s+([\d\.]+)\s+\d+días\s+[\d,]+%Venta Contado:\s*(\d{2}-\d{2}-\d{4})\s+"
    r"[\d.,]+\s+([\d\.]+)",
    re.MULTILINE,
)

_RE_SIM2 = re.compile(r"Compra Plazo:\s*(\d{2}-\d{2}-\d{4})\s+[\d\.,]+\s+([\d\.]+)")

def _extraer_sims(texto: str) -> list:
    sims = []
    lineas = texto.splitlines()
    i = 0
    while i < len(lineas):
        m1 = _RE_SIM1.match(lineas[i].strip())
        if m1:
            instrumento = m1.group(1)
            cantidad = int(_num(m1.group(2)))
            f_venta = _fecha(m1.group(3))
            monto_venta = int(_num(m1.group(4)))

            for j in range(i+1, min(i+4, len(lineas))):
                m2 = _RE_SIM2.search(lineas[j])
                if m2:
                    f_compra = _fecha(m2.group(1))
                    monto_compra = int(_num(m2.group(2)))
                    sims.append({
                        "instrumento": instrumento,
                        "cantidad": cantidad,
                        "f_venta": f_venta,
                        "monto_venta": monto_venta,
                        "f_compra": f_compra,
                        "monto_compra": monto_compra,
                    })
                    break
        i += 1
    return sims


# ── Extracción de caja ────────────────────────────────────────────────────────

_RE_CAJA = re.compile(r"Saldo Final del Periodo\s+(-?[\d\.]+)")

def _extraer_caja(texto: str) -> int:
    m = _RE_CAJA.search(texto)
    if m:
        return int(_num(m.group(1)))
    return 0


# ── Extracción de Operaciones por Liquidar ────────────────────────────────────

_RE_OPS_LIQ = re.compile(r"Operaciones por Liquidar\s+([\d\.]+)")

def _extraer_ops_liquidar(texto: str) -> int:
    m = _RE_OPS_LIQ.search(texto)
    if m:
        val = int(_num(m.group(1)))
        return val
    return 0


# ── Extracción de forwards (EMF) ──────────────────────────────────────────────

_RE_FWD1 = re.compile(
    r"^(\d{7})\s+(Compra|Venta)\s+Seguro de Cambio Nominal\s+"
    r"([\d\.,]+)\s+USD/CLP\s+Fecha Inicio(\d{2}-\d{2}-\d{4})\s+"
    r"Total\s+\d+\s+([\d\.,]+)",
    re.MULTILINE,
)

_RE_FWD2 = re.compile(r"Fecha Termino(\d{2}-\d{2}-\d{4})")

def _extraer_forwards(texto: str) -> list:
    fwds = []
    lineas = texto.splitlines()
    i = 0
    while i < len(lineas):
        m1 = _RE_FWD1.match(lineas[i].strip())
        if m1:
            folio = int(m1.group(1))
            tipo = "C" if m1.group(2) == "Compra" else "V"
            usd = int(_num(m1.group(3)))
            f_inicio = _fecha(m1.group(4))
            tc_fwd = _num(m1.group(5))

            for j in range(i+1, min(i+4, len(lineas))):
                m2 = _RE_FWD2.search(lineas[j])
                if m2:
                    f_termino = _fecha(m2.group(1))
                    fwds.append({
                        "folio": folio,
                        "tipo": tipo,
                        "usd": usd,
                        "tc_fwd": tc_fwd,
                        "f_inicio": f_inicio,
                        "f_termino": f_termino,
                    })
                    break
        i += 1
    return fwds


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear(pdfs: dict) -> dict:
    resultado = {
        "fecha": date.today().isoformat(),
        "precios": {},
        "el": {},
        "emf": {},
    }

    # ── EL ────────────────────────────────────────────────────────────────────
    if "EL" in pdfs:
        _, el_bytes = pdfs["EL"]
        texto_el = _texto_pdf(el_bytes)

        precios = _extraer_precios_cabecera(texto_el)
        resultado["precios"].update(precios)

        acciones = _extraer_acciones(texto_el)
        for a in acciones:
            resultado["precios"][a["nem"]] = a["precio_cartola"]

        cfis_el = _extraer_cfis(texto_el)
        for c in cfis_el:
            resultado["precios"][c["nem"]] = c["precio_cartola"]

        sims = _extraer_sims(texto_el)
        caja_el = _extraer_caja(texto_el)
        ops_liquidar = _extraer_ops_liquidar(texto_el)

        resultado["el"] = {
            "caja": caja_el,
            "ops_liquidar": ops_liquidar,
            "acciones": acciones,
            "cfis": cfis_el,
            "sims": sims,
        }

    # ── EMF ───────────────────────────────────────────────────────────────────
    if "EMF" in pdfs:
        _, emf_bytes = pdfs["EMF"]
        texto_emf = _texto_pdf(emf_bytes)

        if not resultado["precios"]:
            resultado["precios"].update(_extraer_precios_cabecera(texto_emf))

        cfis_emf = _extraer_cfis(texto_emf)
        for c in cfis_emf:
            resultado["precios"].setdefault(c["nem"], c["precio_cartola"])

        fwds = _extraer_forwards(texto_emf)
        caja_emf = _extraer_caja(texto_emf)

        resultado["emf"] = {
            "caja": caja_emf,
            "cfis": cfis_emf,
            "fwds": fwds,
        }

    return resultado


def guardar(data: dict, path: Path = DATA_FILE):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  Cartola guardada → {path}")


def cargar(path: Path = DATA_FILE) -> dict | None:
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


# ── CLI para pruebas ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import zipfile

    if len(sys.argv) < 2:
        print("Uso: python parsear_cartola.py <archivo.zip>")
        sys.exit(1)

    zip_path = sys.argv[1]
    with open(zip_path, "rb") as f:
        zip_data = f.read()

    RUT_MAP = {"76677950": "EL", "77209686": "EMF"}
    pdfs = {}
    with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
        for nombre in z.namelist():
            if not nombre.lower().endswith(".pdf"):
                continue
            nombre_limpio = nombre.replace(".", "").replace("-", "").replace(" ", "")
            for rut, key in RUT_MAP.items():
                if rut in nombre_limpio:
                    pdfs[key] = (nombre, z.read(nombre))
                    break

    data = parsear(pdfs)
    guardar(data)
    print(json.dumps(data, indent=2, ensure_ascii=False))
