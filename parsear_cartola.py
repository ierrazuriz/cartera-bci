"""
Parser de cartolas BCI.

Extrae de los PDFs de EL LTDA y EMF SPA:
  - precios (UF, USD, EUR, acciones, CFIs)
  - posiciones de acciones (cant_activo, cant_pasivo)
  - posiciones CFI
  - simultáneas EL
  - caja EL y EMF
  - forwards EMF

Genera cartola_data.json con toda la información.
"""
import re
import json
import io
from datetime import datetime, date
from pathlib import Path

import pdfplumber

DATA_DIR  = Path(__file__).parent
DATA_FILE = DATA_DIR / "cartola_data.json"

# RUTs conocidos
RUT_EL  = "76677950"
RUT_EMF = "77209686"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _num(s: str) -> float:
    """Convierte '1.234,56' → 1234.56  y  '-856.291.313' → -856291313."""
    s = s.strip().replace("\xa0", "").replace(" ", "")
    negativo = s.startswith("-")
    s = s.lstrip("-")
    # Si tiene coma → formato chileno (punto=miles, coma=decimal)
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
    """Extrae texto plano de todas las páginas."""
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

# Ejemplo de línea:
# "ABC Activo: 0 0 23.210.430 0 11,7800 12,1500 282.006.725 0"
_RE_ACCION = re.compile(
    r"^([A-Z][A-Z0-9\-]+)\s+Activo:\s+"
    r"([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+"
    r"([\d\.,]+)\s+([\d\.,]+)",
    re.MULTILINE,
)

def _extraer_acciones(texto: str) -> list:
    """
    Retorna lista de dicts:
      nem, cant_activo (en_garantia), cant_pasivo (-a_plazo),
      precio_compra, precio_cartola
    Excluye CFIs (que tienen estructura de columnas distinta).
    """
    acciones = []
    for m in _RE_ACCION.finditer(texto):
        nem = m.group(1)
        if nem.startswith("CFI"):
            continue
        en_garantia = int(_num(m.group(4)))
        a_plazo     = int(_num(m.group(5)))
        p_compra    = _num(m.group(6))
        p_ultimo    = _num(m.group(7))

        # cant_activo = acciones físicamente en la cuenta (en garantía de sim)
        # cant_pasivo = acciones vendidas en el contado de la sim (negativo)
        acciones.append({
            "nem":            nem,
            "cant_activo":    en_garantia,
            "cant_pasivo":    -a_plazo,
            "precio_compra":  p_compra,
            "precio_cartola": p_ultimo,
        })
    return acciones


# ── Extracción de CFIs ────────────────────────────────────────────────────────

# Ejemplo: "CFIARRAA-E Activo: 4.187 0 0 0 48.138,4240 51.702,0000 216.476.274 640.475"
_RE_CFI = re.compile(
    r"^(CFI[A-Z0-9\-]+)\s+Activo:\s+"
    r"([\d\.]+)\s+[\d\.]+\s+[\d\.]+\s+[\d\.]+\s+"
    r"([\d\.,]+)\s+([\d\.,]+)",
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
        cfis.append({
            "nem":            nem,
            "cantidad":       int(_num(m.group(2))),
            "precio_compra":  _num(m.group(3)),
            "precio_cartola": _num(m.group(4)),
        })
    return cfis


# ── Extracción de simultáneas (EL) ───────────────────────────────────────────

# Ejemplo de 2 líneas consecutivas:
# "AGUAS-A 438.600 28días 0,47%Venta Contado: 26-03-2026 347,000 152.194.200 152.456.481 ..."
# "Compra Plazo: 23-04-2026 348,522 152.861.837"
#
# Nota: el precio por acción usa coma decimal (ej. "347,000" = 347.000 CLP)
# mientras que los montos usan punto como separador de miles (ej. "152.194.200")
_RE_SIM1 = re.compile(
    r"^([A-Z][A-Z0-9\-]+)\s+([\d\.]+)\s+\d+días\s+[\d,]+%Venta Contado:\s*(\d{2}-\d{2}-\d{4})\s+"
    r"[\d.,]+\s+([\d\.]+)",         # precio_contado (puede tener punto y coma), luego monto_venta
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
            cantidad    = int(_num(m1.group(2)))
            f_venta     = _fecha(m1.group(3))
            monto_venta = int(_num(m1.group(4)))
            # Buscar la siguiente línea con "Compra Plazo:"
            for j in range(i+1, min(i+4, len(lineas))):
                m2 = _RE_SIM2.search(lineas[j])
                if m2:
                    f_compra     = _fecha(m2.group(1))
                    monto_compra = int(_num(m2.group(2)))
                    sims.append({
                        "instrumento": instrumento,
                        "cantidad":    cantidad,
                        "f_venta":     f_venta,
                        "monto_venta": monto_venta,
                        "f_compra":    f_compra,
                        "monto_compra":monto_compra,
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


# ── Extracción de forwards (EMF) ──────────────────────────────────────────────

# Ejemplo:
# "1834140 Compra Seguro de Cambio Nominal 500.000,00 USD/CLP Fecha Inicio01-04-2026 Total 7 916,760000 ..."
# "Fecha Termino08-04-2026 Residual 2"
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
            folio    = int(m1.group(1))
            tipo     = "C" if m1.group(2) == "Compra" else "V"
            usd      = int(_num(m1.group(3)))
            f_inicio = _fecha(m1.group(4))
            tc_fwd   = _num(m1.group(5))
            for j in range(i+1, min(i+4, len(lineas))):
                m2 = _RE_FWD2.search(lineas[j])
                if m2:
                    f_termino = _fecha(m2.group(1))
                    fwds.append({
                        "folio":     folio,
                        "tipo":      tipo,
                        "usd":       usd,
                        "tc_fwd":    tc_fwd,
                        "f_inicio":  f_inicio,
                        "f_termino": f_termino,
                    })
                    break
        i += 1
    return fwds


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear(pdfs: dict) -> dict:
    """
    pdfs: dict {'EL': ('nombre.pdf', bytes), 'EMF': ...}
    Retorna el dict de cartola_data.
    """
    resultado = {
        "fecha":   date.today().isoformat(),
        "precios": {},
        "el":      {},
        "emf":     {},
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

        resultado["el"] = {
            "caja":     caja_el,
            "acciones": acciones,
            "cfis":     cfis_el,
            "sims":     sims,
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

        fwds  = _extraer_forwards(texto_emf)
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
    import base64

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
