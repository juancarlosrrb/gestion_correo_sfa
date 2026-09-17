# -*- coding: utf-8 -*-
"""
Utilidades para sacarle informacion limpia a un correo de Outlook.

Todo lo que toca win32com esta aca, para que los scripts 01 y 04
extraigan EXACTAMENTE los mismos campos. Si extraen distinto, el modelo
entrena con una cosa y predice con otra, y los resultados no sirven.
"""
import os
import re

from . import config

# Propiedades MAPI (esto es Outlook por dentro, se acceden por su codigo)
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"


# ----------------------------------------------------------------------
# 1. IDENTIFICADOR DEL CORREO
# ----------------------------------------------------------------------
def obtener_id_correo(m):
    """
    Devuelve un identificador unico y estable del correo.

    Por que NO uso ConversationID (que es lo que usa el bot actual):
    ConversationID identifica el HILO, no el correo. Si a tu hermano le
    responden tres veces el mismo asunto, los tres correos comparten
    ConversationID y el historial creeria que ya los leyo todos.

    Message-ID de internet es unico por correo y no cambia aunque el
    correo se mueva de carpeta. Si por alguna razon no existe, cae al
    EntryID.
    """
    try:
        mid = m.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
        if mid and str(mid).strip():
            return str(mid).strip()
    except Exception:
        pass
    try:
        return str(m.EntryID)
    except Exception:
        return ""


def obtener_entry_id(m):
    """
    EntryID: es el 'puntero' que Outlook usa para volver a abrir el correo.
    Es el que necesita el script 05 para mover el correo de carpeta.
    OJO: el EntryID CAMBIA cuando el correo se mueve de carpeta, por eso
    no sirve como llave del historial, pero si para mover en el momento.
    """
    try:
        return str(m.EntryID)
    except Exception:
        return ""


# ----------------------------------------------------------------------
# 2. DESTINATARIOS (el target del modelo)
# ----------------------------------------------------------------------
def _resolver_smtp(recipient):
    """Saca el correo real (SMTP) de un destinatario, no el nombre para mostrar."""
    try:
        direccion = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if direccion and "@" in str(direccion):
            return str(direccion).strip().lower()
    except Exception:
        pass
    try:
        direccion = recipient.Address
        if direccion and "@" in str(direccion):
            return str(direccion).strip().lower()
    except Exception:
        pass
    try:
        return str(recipient.Name).strip().lower()
    except Exception:
        return ""


def obtener_destinatarios(m):
    """
    Devuelve (destinatario_principal, todos_los_destinatarios).

    El destinatario principal es el primero del campo "Para" (Type == 1).
    Ese es el TARGET: a que area mando tu hermano el correo.
    Los de copia (CC, Type == 2) se guardan aparte pero no son el target.
    """
    principal = ""
    todos = []
    try:
        for r in m.Recipients:
            direccion = _resolver_smtp(r)
            if not direccion:
                continue
            tipo = 1
            try:
                tipo = int(r.Type)
            except Exception:
                pass
            todos.append(direccion)
            if tipo == 1 and not principal:
                principal = direccion
    except Exception:
        pass

    # Respaldo: si no se pudo resolver por Recipients, se parte el campo To
    if not principal:
        try:
            crudo = str(m.To or "")
            encontrados = re.findall(r"[\w\.\-\+]+@[\w\.\-]+", crudo)
            if encontrados:
                principal = encontrados[0].lower()
                todos = [e.lower() for e in encontrados]
        except Exception:
            pass

    return principal, "; ".join(todos)


# ----------------------------------------------------------------------
# 3. CUERPO Y ASUNTO
# ----------------------------------------------------------------------
def limpiar_texto(texto, max_chars=None):
    """Normaliza espacios, quita saltos de linea raros y recorta."""
    if texto is None:
        return ""
    texto = str(texto)
    texto = texto.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    if max_chars is not None and len(texto) > max_chars:
        texto = texto[:max_chars]
    return texto


def obtener_asunto(m):
    try:
        return limpiar_texto(m.Subject)
    except Exception:
        return ""


def obtener_cuerpo(m, max_chars=None):
    """Cuerpo en texto plano. Si falla, intenta el HTML sin etiquetas."""
    cuerpo = ""
    try:
        cuerpo = m.Body or ""
    except Exception:
        try:
            html = m.HTMLBody or ""
            cuerpo = re.sub(r"<[^>]+>", " ", html)
        except Exception:
            cuerpo = ""
    return limpiar_texto(cuerpo, max_chars)


# ----------------------------------------------------------------------
# 4. ADJUNTOS
# ----------------------------------------------------------------------
EXTENSIONES_LEIBLES = (".pdf", ".txt", ".csv", ".docx", ".xlsx", ".xls", ".html", ".htm")


def _leer_pdf(ruta, usar_ocr=False):
    """Texto nativo del PDF con PyMuPDF. Si viene escaneado y usar_ocr=True, hace OCR."""
    texto = ""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(ruta)
        total = min(len(doc), config.MAX_PAGINAS_PDF)
        for i in range(total):
            texto += doc[i].get_text() + " "
        doc.close()
    except Exception as e:
        texto = ""
        print(f"      ! No se pudo leer el PDF ({os.path.basename(ruta)}): {e}")

    # PDF escaneado: casi sin texto. OCR solo si se pidio explicitamente,
    # porque con 2.000 correos al dia el OCR se demora horas.
    if len(texto.strip()) < 50 and usar_ocr:
        texto = _ocr_pdf(ruta)

    return texto


_lector_ocr = None


def _ocr_pdf(ruta):
    """OCR con easyocr. Carga el lector una sola vez (es pesado)."""
    global _lector_ocr
    try:
        import numpy as np
        import easyocr
        from pdf2image import convert_from_path

        POPPLER_PATH = r"C:\poppler\Library\bin"

        if _lector_ocr is None:
            print("      * Cargando motor OCR por primera vez (tarda un poco)...")
            _lector_ocr = easyocr.Reader(["es"])

        texto = ""
        for i in range(1, config.MAX_PAGINAS_PDF + 1):
            paginas = convert_from_path(
                ruta, first_page=i, last_page=i, dpi=200, poppler_path=POPPLER_PATH
            )
            if not paginas:
                break
            resultado = _lector_ocr.readtext(np.array(paginas[0]), detail=0, paragraph=True)
            texto += " ".join(resultado) + " "
        return texto
    except Exception as e:
        print(f"      ! OCR fallo: {e}")
        return ""


def _leer_docx(ruta):
    try:
        import docx
        d = docx.Document(ruta)
        return " ".join(p.text for p in d.paragraphs)
    except Exception:
        return ""


def _leer_excel(ruta):
    try:
        import pandas as pd
        hojas = pd.read_excel(ruta, sheet_name=None, nrows=50)
        partes = []
        for _, df in hojas.items():
            partes.append(" ".join(df.astype(str).values.ravel().tolist()))
        return " ".join(partes)
    except Exception:
        return ""


def _leer_plano(ruta):
    try:
        with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(50000)
    except Exception:
        return ""


def procesar_adjuntos(m, usar_ocr=False, leer_contenido=True):
    """
    Devuelve (nombres_adjuntos, texto_adjuntos, cantidad).

    nombres_adjuntos: "demanda.pdf | anexo.docx"
    texto_adjuntos:   texto extraido de esos archivos, ya recortado

    Los archivos se bajan a una carpeta temporal y se borran de una vez.
    """
    nombres = []
    textos = []
    cantidad = 0

    try:
        cantidad = int(m.Attachments.Count)
    except Exception:
        return "", "", 0

    if cantidad == 0:
        return "", "", 0

    for i in range(1, cantidad + 1):
        try:
            adj = m.Attachments.Item(i)
            nombre = str(adj.FileName)
        except Exception:
            continue

        nombres.append(nombre)

        if not leer_contenido:
            continue

        ext = os.path.splitext(nombre)[1].lower()
        if ext not in EXTENSIONES_LEIBLES:
            continue

        # Adjuntos muy pesados: solo nos quedamos con el nombre
        try:
            tamano_mb = int(adj.Size) / (1024 * 1024)
            if tamano_mb > config.MAX_MB_ADJUNTO:
                print(f"      * Adjunto pesado, se salta el contenido: {nombre} ({tamano_mb:.1f} MB)")
                continue
        except Exception:
            pass

        # Nombre temporal seguro (los nombres de archivo de correos vienen sucios)
        nombre_seguro = re.sub(r"[^A-Za-z0-9._-]", "_", nombre)[:80]
        ruta = os.path.join(os.path.abspath(config.CARPETA_TEMP_ADJUNTOS), f"{i}_{nombre_seguro}")

        try:
            adj.SaveAsFile(ruta)
        except Exception as e:
            print(f"      ! No se pudo guardar el adjunto {nombre}: {e}")
            continue

        try:
            if ext == ".pdf":
                textos.append(_leer_pdf(ruta, usar_ocr=usar_ocr))
            elif ext == ".docx":
                textos.append(_leer_docx(ruta))
            elif ext in (".xlsx", ".xls"):
                textos.append(_leer_excel(ruta))
            else:
                textos.append(_leer_plano(ruta))
        finally:
            try:
                if os.path.exists(ruta):
                    os.remove(ruta)
            except Exception:
                pass

        # Si ya tenemos suficiente texto, no seguimos abriendo adjuntos
        if sum(len(t) for t in textos) > config.MAX_CHARS_ADJUNTOS * 2:
            break

    return (
        " | ".join(nombres),
        limpiar_texto(" ".join(textos), config.MAX_CHARS_ADJUNTOS),
        cantidad,
    )


# ----------------------------------------------------------------------
# 5. EXTRACCION COMPLETA DE UN CORREO
# ----------------------------------------------------------------------
def extraer_correo(m, usar_ocr=False, recortar=True, leer_adjuntos=True):
    """
    Convierte un correo de Outlook en un diccionario plano.

    recortar=True  -> aplica los topes de config (para la base historica)
    recortar=False -> guarda todo completo (para la base cruda de 10 correos)
    """
    max_cuerpo = config.MAX_CHARS_CUERPO if recortar else None

    destinatario, destinatarios_todos = obtener_destinatarios(m)
    nombres_adj, texto_adj, n_adj = procesar_adjuntos(
        m, usar_ocr=usar_ocr, leer_contenido=leer_adjuntos
    )

    if not recortar:
        # En la base cruda no recortamos el texto de adjuntos
        pass

    fecha = None
    for campo in ("SentOn", "ReceivedTime", "CreationTime"):
        try:
            fecha = str(getattr(m, campo))
            if fecha:
                break
        except Exception:
            continue

    remitente = ""
    try:
        remitente = str(m.SenderEmailAddress or "").lower()
    except Exception:
        pass

    return {
        "email_id": obtener_id_correo(m),
        "entry_id": obtener_entry_id(m),
        "conversation_id": _intento(lambda: str(m.ConversationID)),
        "fecha": fecha,
        "remitente": remitente,
        "destinatario": destinatario,
        "destinatarios_todos": destinatarios_todos,
        "asunto": obtener_asunto(m),
        "cuerpo": obtener_cuerpo(m, max_cuerpo),
        "nombres_adjuntos": nombres_adj,
        "texto_adjuntos": texto_adj,
        "n_adjuntos": n_adj,
    }


def _intento(fn, por_defecto=""):
    try:
        return fn()
    except Exception:
        return por_defecto
