"""
Lector de facturas BCI desde Gmail.

=== FIX 4: NUEVO MÓDULO ===
Lee emails de facturas de BCI y actualiza facturas_data.json automáticamente.

Fuentes de email:
- clientesbcicorredordebolsa@bci.cl   → facturas con detalle de operaciones
- 965198008DTE@mail.bcs.cl            → documentos tributarios electrónicos (DTEs)

Los DTEs contienen PDFs con el detalle de cada operación (ticker, tipo, cantidad,
monto, comisión). Este módulo descarga esos PDFs, los parsea, y actualiza
facturas_data.json.
"""

import os
import io
import re
import json
import base64
import logging
from datetime import date, datetime
from pathlib import Path

import pdfplumber

# Reutilizamos la autenticación de gmail_bci
from gmail_bci import autenticar
from googleapiclient.discovery import build

DATA_DIR = Path(os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__))))
FACTURAS_FILE = DATA_DIR / "facturas_data.json"


# ── Carga/guardado ────────────────────────────────────────────────────────────

def cargar_facturas():
    try:
        with open(FACTURAS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "movimientos": [],
            "dividendos": [],
            "forwards": [],
            "transferencias": [],
            "comparacion": [],
            "sync_at": "",
        }


def guardar_facturas(data):
    data["sync_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(FACTURAS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    logging.info("Facturas guardadas → %s", FACTURAS_FILE)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _num(s):
    """'1.234.567' → 1234567, '1.234,56' → 1234.56"""
    s = s.strip().replace("\xa0", "").replace(" ", "")
    neg = s.startswith("-")
    s = s.lstrip("-")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(".", "")
    return -float(s) if neg else float(s)


def _fecha_iso(s):
    """'09/04/2026' o '09-04-2026' → '2026-04-09'"""
    s = s.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return s


# ── Parseo de DTE (factura electrónica en PDF) ───────────────────────────────

# Patrones comunes en facturas BCI
_RE_FOLIO = re.compile(r"(?:Folio|N[°º]\s*Factura)[:\s]*(\d+)", re.IGNORECASE)
_RE_FECHA = re.compile(r"Fecha[:\s]*(\d{2}[/-]\d{2}[/-]\d{4})")
_RE_RUT = re.compile(r"(?:RUT|Rut)[:\s]*([\d\.]+\-[\dkK])")

# Línea de detalle de operación:
# "COMPRA ACCIONES LTM 10.000.000 23,79 237.900.000"
_RE_OPERACION = re.compile(
    r"(COMPRA|VENTA)\s+"
    r"(?:ACCIONES|ACC\.?|CFI|CUOTAS?)\s+"
    r"([A-Z][A-Z0-9\-]+)\s+"
    r"([\d\.]+)\s+"          # cantidad
    r"([\d\.,]+)\s+"         # precio
    r"([\d\.]+)",            # monto total
    re.IGNORECASE,
)

# Fecha de pago
_RE_FPAGO = re.compile(r"(?:Fecha\s+Pago|Liquidaci[oó]n)[:\s]*(\d{2}[/-]\d{2}[/-]\d{4})", re.IGNORECASE)


def parsear_dte_pdf(pdf_bytes):
    """
    Parsea un PDF de DTE/factura electrónica de BCI.
    Retorna lista de operaciones encontradas.
    """
    operaciones = []
    texto = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    texto += t + "\n"
    except Exception as e:
        logging.warning("Error leyendo PDF de DTE: %s", e)
        return operaciones

    # Extraer metadata
    folio_m = _RE_FOLIO.search(texto)
    folio = folio_m.group(1) if folio_m else ""

    fecha_m = _RE_FECHA.search(texto)
    fecha = _fecha_iso(fecha_m.group(1)) if fecha_m else ""

    fpago_m = _RE_FPAGO.search(texto)
    fpago = _fecha_iso(fpago_m.group(1)) if fpago_m else ""

    # Determinar sociedad por RUT
    sociedad = "EL"
    if "77.209.686" in texto or "77209686" in texto:
        sociedad = "EMF"

    # Extraer operaciones
    for m in _RE_OPERACION.finditer(texto):
        tipo = m.group(1).upper()
        ticker = m.group(2).upper()
        try:
            cantidad = int(_num(m.group(3)))
            monto = int(_num(m.group(5)))
        except (ValueError, TypeError):
            continue

        operaciones.append({
            "fecha": fecha,
            "f_pago": fpago,
            "n_factura": folio,
            "sociedad": sociedad,
            "ticker": ticker,
            "tipo": tipo,
            "cantidad": cantidad,
            "monto_clp": monto,
        })

    return operaciones


# ── Búsqueda de emails de facturas ────────────────────────────────────────────

def buscar_facturas_gmail(service, dias_atras=30):
    """
    Busca emails de facturas BCI de los últimos N días.
    Retorna lista de message IDs.
    """
    desde = date.fromordinal(date.today().toordinal() - dias_atras)
    query = (
        f"(from:clientesbcicorredordebolsa@bci.cl OR from:965198008DTE@mail.bcs.cl) "
        f"has:attachment after:{desde.strftime('%Y/%m/%d')}"
    )

    msg_ids = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": query, "maxResults": 50}
        if page_token:
            kwargs["pageToken"] = page_token

        result = service.users().messages().list(**kwargs).execute()
        messages = result.get("messages", [])
        msg_ids.extend(m["id"] for m in messages)

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    logging.info("Encontrados %d emails de facturas BCI", len(msg_ids))
    return msg_ids


def descargar_adjuntos_pdf(service, msg_id):
    """Descarga todos los adjuntos PDF de un mensaje."""
    msg = service.users().messages().get(
        userId="me", id=msg_id, format="full"
    ).execute()

    pdfs = []

    def buscar(parts):
        for part in parts:
            fname = part.get("filename", "")
            mime = part.get("mimeType", "")
            if fname.lower().endswith(".pdf") or mime == "application/pdf":
                body = part.get("body", {})
                att_id = body.get("attachmentId")
                if att_id:
                    att = service.users().messages().attachments().get(
                        userId="me", messageId=msg_id, id=att_id
                    ).execute()
                    data = base64.urlsafe_b64decode(att["data"])
                    pdfs.append((fname, data))
            subparts = part.get("parts", [])
            if subparts:
                buscar(subparts)

    parts = msg["payload"].get("parts", [msg["payload"]])
    buscar(parts)
    return pdfs


# ── Sync principal ────────────────────────────────────────────────────────────

def sync_facturas(dias_atras=60):
    """
    Proceso completo: busca emails → descarga PDFs → parsea → guarda JSON.
    """
    logging.info("Iniciando sync de facturas (últimos %d días)...", dias_atras)

    creds = autenticar()
    service = build("gmail", "v1", credentials=creds)

    msg_ids = buscar_facturas_gmail(service, dias_atras)
    if not msg_ids:
        logging.info("No se encontraron emails de facturas nuevas")
        return

    # Cargar facturas existentes para evitar duplicados
    data = cargar_facturas()
    folios_existentes = {m.get("n_factura") for m in data.get("movimientos", []) if m.get("n_factura")}

    nuevas = 0
    for msg_id in msg_ids:
        try:
            pdfs = descargar_adjuntos_pdf(service, msg_id)
            for fname, pdf_bytes in pdfs:
                ops = parsear_dte_pdf(pdf_bytes)
                for op in ops:
                    # Evitar duplicados por folio
                    if op.get("n_factura") and op["n_factura"] in folios_existentes:
                        continue
                    data["movimientos"].append(op)
                    folios_existentes.add(op.get("n_factura"))
                    nuevas += 1
        except Exception as e:
            logging.warning("Error procesando email %s: %s", msg_id, e)

    # Ordenar por fecha descendente
    data["movimientos"].sort(key=lambda x: x.get("fecha", ""), reverse=True)

    guardar_facturas(data)
    logging.info("Sync facturas completado: %d operaciones nuevas", nuevas)


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    sync_facturas()
