"""
AGENTE 2: Verificacion de Formato de Reporte IA SSOMA
- Descarga adjuntos Word (.docx) o PDF (.pdf) del correo
- Verifica que el documento inicie con datos de empresa y fecha
- Retorna resultado de validacion de formato
"""
import base64
import os
import re
import tempfile

from config import TEMP_DIR, COMPANY_NAME


def verificar_formato_reporte(service, message_id, adjuntos, fecha_objetivo):
    """
    Descarga y verifica el formato del reporte IA SSOMA adjunto.

    Args:
        service: Gmail API service
        message_id: ID del mensaje de Gmail
        adjuntos: Lista de adjuntos encontrados
        fecha_objetivo: date object (fecha esperada del reporte)

    Returns:
        dict con:
            formato_valido: bool
            tipo_archivo: str (word/pdf/desconocido)
            tiene_datos_empresa: bool
            tiene_fecha: bool
            fecha_documento: str o None
            fecha_correcta: bool
            detalle: str
    """
    if not adjuntos:
        return {
            "formato_valido": False,
            "tipo_archivo": "sin_adjunto",
            "tiene_datos_empresa": False,
            "tiene_fecha": False,
            "fecha_documento": None,
            "fecha_correcta": False,
            "detalle": "Sin adjunto Word/PDF",
        }

    os.makedirs(TEMP_DIR, exist_ok=True)

    # Tomar el primer adjunto valido
    adjunto = adjuntos[0]
    filename = adjunto["filename"]
    attachment_id = adjunto["attachmentId"]

    # Descargar adjunto
    try:
        att = (
            service.users()
            .messages()
            .attachments()
            .get(userId="me", messageId=message_id, id=attachment_id)
            .execute()
        )
        data = base64.urlsafe_b64decode(att["data"])
    except Exception as e:
        return {
            "formato_valido": False,
            "tipo_archivo": "error",
            "tiene_datos_empresa": False,
            "tiene_fecha": False,
            "fecha_documento": None,
            "fecha_correcta": False,
            "detalle": f"Error descargando adjunto: {e}",
        }

    # Guardar temporalmente
    filepath = os.path.join(TEMP_DIR, filename)
    with open(filepath, "wb") as f:
        f.write(data)

    try:
        # Extraer texto segun tipo
        if filename.lower().endswith(".pdf"):
            texto_inicio = _extraer_texto_pdf(filepath)
            tipo_archivo = "pdf"
        elif filename.lower().endswith((".docx", ".doc")):
            texto_inicio = _extraer_texto_word(filepath)
            tipo_archivo = "word"
        else:
            return {
                "formato_valido": False,
                "tipo_archivo": "desconocido",
                "tiene_datos_empresa": False,
                "tiene_fecha": False,
                "fecha_documento": None,
                "fecha_correcta": False,
                "detalle": f"Tipo de archivo no soportado: {filename}",
            }

        # Validar contenido inicial
        return _validar_contenido(texto_inicio, tipo_archivo, fecha_objetivo, filename)

    finally:
        # Limpiar archivo temporal
        try:
            os.remove(filepath)
        except OSError:
            pass


def _extraer_texto_pdf(filepath):
    """Extrae las primeras lineas de texto de un PDF."""
    try:
        import PyPDF2

        with open(filepath, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            if len(reader.pages) == 0:
                return ""
            # Solo leer la primera pagina
            page = reader.pages[0]
            texto = page.extract_text() or ""
            # Retornar los primeros 1000 caracteres
            return texto[:1000]
    except ImportError:
        # Fallback: intentar con pdfplumber
        try:
            import pdfplumber

            with pdfplumber.open(filepath) as pdf:
                if len(pdf.pages) == 0:
                    return ""
                texto = pdf.pages[0].extract_text() or ""
                return texto[:1000]
        except ImportError:
            return "[ERROR: Instalar PyPDF2 o pdfplumber para leer PDFs]"


def _extraer_texto_word(filepath):
    """Extrae las primeras lineas de texto de un documento Word."""
    try:
        from docx import Document

        doc = Document(filepath)
        lineas = []
        for para in doc.paragraphs[:20]:  # Primeros 20 parrafos
            texto = para.text.strip()
            if texto:
                lineas.append(texto)
            if len("\n".join(lineas)) > 1000:
                break
        return "\n".join(lineas)[:1000]
    except ImportError:
        return "[ERROR: Instalar python-docx para leer archivos Word]"
    except Exception as e:
        return f"[ERROR: {e}]"


def _validar_contenido(texto_inicio, tipo_archivo, fecha_objetivo, filename):
    """Valida que el texto del documento contenga datos de empresa y fecha."""
    texto_lower = texto_inicio.lower()

    # Verificar datos de empresa
    keywords_empresa = [
        "hergonsa", "hergon", "grupo hergonsa",
        "ssoma", "seguridad", "salud ocupacional",
        "reporte", "informe",
    ]
    tiene_datos_empresa = any(kw in texto_lower for kw in keywords_empresa)

    # Verificar fecha en el documento
    fecha_objetivo_str = fecha_objetivo.strftime("%d/%m/%Y")
    fecha_objetivo_str2 = fecha_objetivo.strftime("%d-%m-%Y")
    fecha_objetivo_str3 = fecha_objetivo.strftime("%d.%m.%Y")

    # Buscar la fecha exacta esperada
    tiene_fecha_exacta = any(
        f in texto_inicio
        for f in [fecha_objetivo_str, fecha_objetivo_str2, fecha_objetivo_str3]
    )

    # Buscar cualquier fecha en formato dd/mm/yyyy o similar
    patron_fecha = r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})'
    fechas_encontradas = re.findall(patron_fecha, texto_inicio)
    tiene_alguna_fecha = len(fechas_encontradas) > 0

    # Intentar extraer la primera fecha del documento
    fecha_documento = None
    fecha_correcta = False

    if fechas_encontradas:
        dia, mes, anio = fechas_encontradas[0]
        if len(anio) == 2:
            anio = f"20{anio}"
        fecha_documento = f"{dia.zfill(2)}/{mes.zfill(2)}/{anio}"

        # Verificar si la fecha coincide con la esperada
        try:
            from datetime import datetime
            fecha_doc_parsed = datetime.strptime(fecha_documento, "%d/%m/%Y").date()
            fecha_correcta = fecha_doc_parsed == fecha_objetivo
        except ValueError:
            pass

    # Determinar formato valido
    formato_valido = tiene_datos_empresa and tiene_alguna_fecha

    # Construir detalle
    detalles = []
    if not tiene_datos_empresa:
        detalles.append("Sin datos de empresa al inicio")
    if not tiene_alguna_fecha:
        detalles.append("Sin fecha en el documento")
    elif not fecha_correcta:
        detalles.append(f"Fecha documento: {fecha_documento} (esperado: {fecha_objetivo_str})")

    detalle = "; ".join(detalles) if detalles else "Formato correcto"

    return {
        "formato_valido": formato_valido,
        "tipo_archivo": tipo_archivo,
        "tiene_datos_empresa": tiene_datos_empresa,
        "tiene_fecha": tiene_alguna_fecha,
        "fecha_documento": fecha_documento,
        "fecha_correcta": fecha_correcta,
        "detalle": detalle,
        "nombre_archivo": filename,
    }
