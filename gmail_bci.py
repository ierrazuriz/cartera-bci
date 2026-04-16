"""
Descargador de cartolas BCI desde Gmail.

=== FIX 3 APLICADO ===
Antes: from:"Unidad Comercial Bci" (display name que podía no matchear)
Ahora: from:soportecomercialcb@bci.cl (email exacto del sender)

Uso:
    python gmail_bci.py                # descarga y procesa última cartola
    python gmail_bci.py --mostrar      # muestra texto crudo de los PDFs
    python gmail_bci.py --guardar      # guarda los PDFs en ./cartolas/
    python gmail_bci.py --fecha 2026-04-01  # busca cartola de fecha específica
"""

import os
import io
import json
import zipfile
import base64
import argparse
from datetime import date, datetime
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import pdfplumber

# ── Configuración ─────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

# RUTs que mapean cada PDF a una entidad
RUTS = {
    "76677950": "EL",   # EL LTDA (76.677.950-6)
    "77209686": "EMF",  # EMF SPA (77.209.686-0)
}

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(DATA_DIR, "precios.json")


# ── Autenticación Gmail ───────────────────────────────────────────────────────

def autenticar():
    """
    OAuth 2.0.
    En Railway: lee el token desde la variable de entorno GMAIL_TOKEN_JSON.
    Localmente: usa token.json.
    """
    creds = None

    # 1. Intentar desde variable de entorno (Railway)
    token_env = os.environ.get("GMAIL_TOKEN_JSON")
    if token_env:
        creds = Credentials.from_authorized_user_info(json.loads(token_env), SCOPES)

    # 2. Intentar desde archivo local
    if creds is None and os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # 3. Refrescar si expiró
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        # Guardar localmente si aplica
        if os.path.exists(TOKEN_FILE) or not token_env:
            with open(TOKEN_FILE, "w") as f:
                f.write(creds.to_json())

    # 4. Flujo interactivo (solo local)
    if not creds or not creds.valid:
        if not os.path.exists(CREDENTIALS_FILE):
            raise FileNotFoundError(
                f"Falta '{CREDENTIALS_FILE}'. "
                "Descárgalo desde Google Cloud Console > APIs & Services > Credentials."
            )
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return creds


# ── Búsqueda y descarga del email ─────────────────────────────────────────────

def buscar_email_bci(service, fecha: date = None):
    """
    Busca el email más reciente de BCI con adjunto ZIP.

    FIX 3: usa from:soportecomercialcb@bci.cl (email exacto)
    en vez de from:"Unidad Comercial Bci" (display name).
    """
    # ── Query corregida ──
    query = 'from:soportecomercialcb@bci.cl has:attachment filename:zip'

    if fecha:
        d_str = fecha.strftime("%Y/%m/%d")
        d_next = date.fromordinal(fecha.toordinal() + 1).strftime("%Y/%m/%d")
        query += f" after:{d_str} before:{d_next}"

    result = service.users().messages().list(
        userId="me", q=query, maxResults=1
    ).execute()

    messages = result.get("messages", [])
    if not messages:
        raise ValueError(
            f"No se encontró email de BCI con ZIP"
            + (f" para la fecha {fecha}" if fecha else "")
        )
    return messages[0]["id"]


def descargar_zip(service, msg_id: str):
    """Retorna (bytes del ZIP, nombre del archivo, fecha del email)."""
    msg = service.users().messages().get(
        userId="me", id=msg_id, format="full"
    ).execute()

    headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}
    fecha_hdr = headers.get("Date", "")

    def buscar_adjunto(parts):
        for part in parts:
            fname = part.get("filename", "")
            if fname.lower().endswith(".zip"):
                body = part.get("body", {})
                att_id = body.get("attachmentId")
                if att_id:
                    att = service.users().messages().attachments().get(
                        userId="me", messageId=msg_id, id=att_id
                    ).execute()
                    data = base64.urlsafe_b64decode(att["data"])
                    return data, fname
            subparts = part.get("parts", [])
            if subparts:
                result = buscar_adjunto(subparts)
                if result:
                    return result
        return None, None

    parts = msg["payload"].get("parts", [msg["payload"]])
    zip_data, zip_name = buscar_adjunto(parts)

    if zip_data is None:
        raise ValueError("No se encontró adjunto ZIP en el email de BCI")

    return zip_data, zip_name, fecha_hdr


# ── Extracción de PDFs del ZIP ────────────────────────────────────────────────

def extraer_pdfs(zip_data: bytes) -> dict:
    """
    Extrae los PDFs del ZIP.
    Retorna dict: {'EL': ('nombre.pdf', bytes), 'EMF': (...)}
    """
    pdfs = {}
    with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
        for nombre in z.namelist():
            if not nombre.lower().endswith(".pdf"):
                continue
            contenido = z.read(nombre)
            rut_key = None
            nombre_limpio = nombre.replace(".", "").replace("-", "").replace(" ", "")
            for rut, key in RUTS.items():
                if rut in nombre_limpio:
                    rut_key = key
                    break
            if rut_key is None:
                rut_key = Path(nombre).stem
            pdfs[rut_key] = (nombre, contenido)
    return pdfs


# ── Extracción de texto del PDF ───────────────────────────────────────────────

def extraer_texto_pdf(pdf_bytes: bytes) -> str:
    """Extrae texto y tablas de un PDF con pdfplumber."""
    lineas = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages):
            lineas.append(f"\n{'─'*60}")
            lineas.append(f" PÁGINA {i + 1}")
            lineas.append(f"{'─'*60}")
            texto = page.extract_text()
            if texto:
                lineas.append(texto)
            tablas = page.extract_tables()
            for j, tabla in enumerate(tablas):
                lineas.append(f"\n  [Tabla {j + 1}]")
                for fila in tabla:
                    celdas = [str(c).strip() if c else "" for c in fila]
                    lineas.append("  " + " | ".join(celdas))
    return "\n".join(lineas)


# ── Guardar PDFs localmente ───────────────────────────────────────────────────

def guardar_pdfs(pdfs: dict, directorio: str = "cartolas"):
    Path(directorio).mkdir(exist_ok=True)
    hoy = date.today().strftime("%Y%m%d")
    for key, (nombre, contenido) in pdfs.items():
        destino = os.path.join(directorio, f"{hoy}_{key}_{nombre}")
        with open(destino, "wb") as f:
            f.write(contenido)
        print(f"  Guardado: {destino}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Descargador de cartolas BCI desde Gmail")
    parser.add_argument("--mostrar", action="store_true")
    parser.add_argument("--guardar", action="store_true")
    parser.add_argument("--fecha", type=str, default=None)
    args = parser.parse_args()

    fecha = None
    if args.fecha:
        fecha = datetime.strptime(args.fecha, "%Y-%m-%d").date()

    print("Autenticando con Gmail...")
    creds = autenticar()
    service = build("gmail", "v1", credentials=creds)
    print("  OK")

    print(f"Buscando email de BCI{' del ' + str(fecha) if fecha else ''}...")
    msg_id = buscar_email_bci(service, fecha)
    print(f"  Encontrado (id: {msg_id})")

    print("Descargando adjunto ZIP...")
    zip_data, zip_name, fecha_email = descargar_zip(service, msg_id)
    print(f"  {zip_name} ({len(zip_data):,} bytes) — {fecha_email}")

    print("Extrayendo PDFs del ZIP...")
    pdfs = extraer_pdfs(zip_data)
    for key, (nombre, data) in pdfs.items():
        print(f"  [{key}] {nombre} ({len(data):,} bytes)")

    if not pdfs:
        print("ERROR: No se encontraron PDFs dentro del ZIP.")
        return

    if args.guardar:
        print("Guardando PDFs...")
        guardar_pdfs(pdfs)

    if args.mostrar:
        for key, (nombre, data) in pdfs.items():
            print(f"\n{'='*60}")
            print(f"  PDF: {key} ({nombre})")
            print(f"{'='*60}")
            print(extraer_texto_pdf(data))

    print("\nParsando cartola...")
    import parsear_cartola as pc
    cartola_data = pc.parsear(pdfs)
    pc.guardar(cartola_data)

    precios = cartola_data.get("precios", {})
    print(f"  Precios actualizados: {', '.join(precios.keys())}")

    el_caja = cartola_data.get("el", {}).get("caja")
    emf_caja = cartola_data.get("emf", {}).get("caja")
    if el_caja is not None:
        print(f"  Caja EL: {el_caja:,}")
    if emf_caja is not None:
        print(f"  Caja EMF: {emf_caja:,}")

    n_sims = len(cartola_data.get("el", {}).get("sims", []))
    n_fwds = len(cartola_data.get("emf", {}).get("fwds", []))
    print(f"  Simultáneas EL: {n_sims} | Forwards EMF: {n_fwds}")
    print("Listo.")


if __name__ == "__main__":
    main()
