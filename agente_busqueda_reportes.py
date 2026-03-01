"""
AGENTE 1: Busqueda de Correos de Reporte IA SSOMA
- Busca correos con asunto "REPORTE IA SSOMA-OBRA XXX" en Gmail
- Filtra por remitentes conocidos (obras registradas en config)
- Filtra por fecha (dia objetivo)
- Mapea cada correo a su obra correspondiente
"""
import re
from email.utils import parseaddr
from datetime import datetime, date

from config import OBRAS, GMAIL_FROM_QUERY, GMAIL_SUBJECT_QUERY, EXTENSIONES_VALIDAS


def buscar_reportes(service, fecha_objetivo):
    """
    Busca correos de Reporte IA SSOMA en Gmail para una fecha especifica.

    Args:
        service: Gmail API service
        fecha_objetivo: date object (el dia del reporte a buscar)

    Returns:
        Lista de diccionarios con info de cada correo encontrado
    """
    from datetime import timedelta

    # Construir query de fecha (after/before en formato YYYY/MM/DD)
    fecha_after = fecha_objetivo.strftime("%Y/%m/%d")

    # Si hoy es lunes (o sea fecha_objetivo es sabado), ampliar ventana
    hoy = date.today()
    if hoy.weekday() == 0 and fecha_objetivo.weekday() == 5:
        fecha_before = (hoy + timedelta(days=1)).strftime("%Y/%m/%d")
    else:
        fecha_before = (fecha_objetivo + timedelta(days=1)).strftime("%Y/%m/%d")

    # Query: de los remitentes conocidos + asunto de reporte SSOMA + rango de fecha
    query = f"({GMAIL_FROM_QUERY}) ({GMAIL_SUBJECT_QUERY}) after:{fecha_after} before:{fecha_before}"

    print(f"\n[AGENTE 1] Buscando reportes IA SSOMA del {fecha_objetivo.strftime('%d/%m/%Y')}")
    print(f"[AGENTE 1] Query: {query[:120]}...")

    resultados = []
    page_token = None

    while True:
        response = (
            service.users()
            .messages()
            .list(
                userId="me",
                q=query,
                maxResults=50,
                pageToken=page_token,
            )
            .execute()
        )

        messages = response.get("messages", [])
        if not messages:
            break

        for msg_ref in messages:
            msg_data = _procesar_mensaje(service, msg_ref["id"], fecha_objetivo)
            if msg_data:
                resultados.append(msg_data)

        page_token = response.get("nextPageToken")
        if not page_token:
            break

    print(f"[AGENTE 1] Se encontraron {len(resultados)} correos de reporte IA SSOMA.")
    return resultados


def _procesar_mensaje(service, message_id, fecha_objetivo):
    """Procesa un mensaje individual y extrae informacion relevante."""
    msg = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )

    headers = msg.get("payload", {}).get("headers", [])
    header_dict = {h["name"].lower(): h["value"] for h in headers}

    asunto = header_dict.get("subject", "(Sin asunto)")
    de = header_dict.get("from", "")
    fecha_raw = header_dict.get("date", "")

    # Extraer email del remitente
    _, de_email = parseaddr(de)
    de_email_lower = de_email.lower()

    # Mapear remitente a obra
    obra_key, obra_nombre = _mapear_email_a_obra(de_email_lower)
    if not obra_key:
        return None  # Remitente no reconocido

    # Parsear fecha de envio
    fecha_envio = _parsear_fecha(fecha_raw)

    # Buscar adjuntos Word/PDF
    payload = msg.get("payload", {})
    adjuntos = _buscar_adjuntos_documento(payload)

    # Link directo a Gmail
    gmail_link = f"https://mail.google.com/mail/u/0/#all/{msg['id']}"

    return {
        "id": msg["id"],
        "thread_id": msg.get("threadId", ""),
        "obra_key": obra_key,
        "obra_nombre": obra_nombre,
        "de": de,
        "de_email": de_email_lower,
        "asunto": asunto,
        "fecha_envio": fecha_envio,
        "fecha_raw": fecha_raw,
        "tiene_adjunto_documento": len(adjuntos) > 0,
        "adjuntos": adjuntos,
        "gmail_link": gmail_link,
    }


def _mapear_email_a_obra(email_lower):
    """Mapea un email de remitente a su obra correspondiente."""
    for key, obra in OBRAS.items():
        for email_obra in obra["emails"]:
            if email_lower == email_obra.lower():
                return key, obra["nombre"]
    return None, None


def _buscar_adjuntos_documento(payload, adjuntos=None):
    """Busca recursivamente adjuntos Word o PDF en el payload del mensaje."""
    if adjuntos is None:
        adjuntos = []

    filename = payload.get("filename", "")
    attachment_id = payload.get("body", {}).get("attachmentId")

    if filename and attachment_id:
        if any(filename.lower().endswith(ext) for ext in EXTENSIONES_VALIDAS):
            adjuntos.append({
                "filename": filename,
                "attachmentId": attachment_id,
                "mimeType": payload.get("mimeType", ""),
            })

    for part in payload.get("parts", []):
        _buscar_adjuntos_documento(part, adjuntos)

    return adjuntos


def _parsear_fecha(fecha_raw):
    """Parsea la fecha del header del correo a formato legible."""
    formatos = [
        "%a, %d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S %Z",
    ]
    fecha_limpia = re.sub(r"\s*\([^)]*\)\s*$", "", fecha_raw).strip()

    for fmt in formatos:
        try:
            dt = datetime.strptime(fecha_limpia, fmt)
            return dt.strftime("%d/%m/%Y %H:%M")
        except ValueError:
            continue
    return fecha_raw
