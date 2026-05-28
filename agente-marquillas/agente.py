#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGENTE MARQUILLAS S.A.S — Verificador automático de facturas
=============================================================
Conecta con el correo corporativo de Outlook via Microsoft Graph API,
descarga ZIPs adjuntos, extrae el PDF interno, lo verifica con Claude AI
y reenvía el PDF al destinatario configurado si la factura es aprobada.

Ejecutar con: python agente.py
"""

import os
import io
import json
import base64
import logging
import zipfile
import time
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

import fitz           # PyMuPDF — extrae texto de PDFs
import msal           # Microsoft Authentication Library — autenticación Azure
import requests       # Llamadas HTTP a Microsoft Graph API
import schedule       # Planificador de tareas periódicas
from anthropic import Anthropic   # Cliente oficial de Claude AI
from dotenv import load_dotenv    # Carga variables desde el archivo .env

# Cargar las variables de entorno desde el archivo .env antes de cualquier otra cosa
load_dotenv()

# ═══════════════════════════════════════════════════════════════
# ═══ CONFIGURACIÓN — todas las variables del .env se leen aquí ═══
# ═══════════════════════════════════════════════════════════════
# El resto del código usa estas variables directamente, nunca os.getenv() interno.

ANTHROPIC_API_KEY   = os.getenv("ANTHROPIC_API_KEY", "")
AZURE_CLIENT_ID     = os.getenv("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "")
AZURE_TENANT_ID     = os.getenv("AZURE_TENANT_ID", "")
EMAIL_MONITOREAR           = os.getenv("EMAIL_MONITOREAR", "")
EMAIL_SABANETA_PRINCIPAL   = os.getenv("EMAIL_SABANETA_PRINCIPAL", "")
EMAIL_SABANETA_COPIA       = os.getenv("EMAIL_SABANETA_COPIA", "")
EMAIL_RIONEGRO_PRINCIPAL   = os.getenv("EMAIL_RIONEGRO_PRINCIPAL", "")
EMAIL_RIONEGRO_COPIA       = os.getenv("EMAIL_RIONEGRO_COPIA", "")
EMAIL_COPIA_SIEMPRE        = [
    correo.strip()
    for correo in os.getenv("EMAIL_COPIA_SIEMPRE", "").split(",")
    if correo.strip()
]
INTERVALO_MINUTOS   = int(os.getenv("INTERVALO_MINUTOS", "5"))
ARCHIVO_PROVEEDORES = "proveedores.json"

# ═══ CONSTANTES ═══
NIT_ESPERADO          = "890900314"
RAZON_SOCIAL_ESPERADA = "MARQUILLAS S.A.S"
MODELO_CLAUDE         = "claude-opus-4-5"
MAXIMO_CARACTERES_PDF = 4000          # Límite para no exceder tokens de Claude
SCOPES_MICROSOFT      = ["https://graph.microsoft.com/.default"]
URL_GRAPH_API         = "https://graph.microsoft.com/v1.0"
ARCHIVO_INSTRUCCIONES = "agente.md"
TAMANO_MAXIMO_LOG             = 5 * 1024 * 1024   # 5 MB — cada archivo rota al llegar aquí

# ── Rutas de las 5 carpetas y archivos de log especializados ──
RUTA_LOG_ERRORES                    = os.path.join("logs", "errores",                    "errores.log")
RUTA_LOG_APROBADOS_AGENTE           = os.path.join("logs", "aprobados_agente",           "aprobados_agente.log")
RUTA_LOG_RECHAZADOS_AGENTE          = os.path.join("logs", "rechazados_agente",          "rechazados_agente.log")
RUTA_LOG_APROBADOS_HUMANOS          = os.path.join("logs", "aprobados_area_responsable", "aprobados_area_responsable.log")
RUTA_LOG_PROVEEDORES_NO_ENCONTRADOS = os.path.join("logs", "proveedores_no_encontrados", "proveedores_no_encontrados.log")
RUTA_LOG_RECHAZADOS_HUMANOS         = os.path.join("logs", "rechazados_area_responsable", "rechazados_area_responsable.log")

# ── Fase 2 / 5: detección y clasificación de respuestas humanas ──
CARPETA_FACTURAS_APROBADAS  = os.getenv("CARPETA_APROBADAS",  "APROBADAS")
CARPETA_FACTURAS_RECHAZADAS = os.getenv("CARPETA_RECHAZADAS", "RECHAZADAS")
ARCHIVO_CLASIFICADOR = "clasificador.md"


# ═══════════════════════════════════════════════════════════════
# CONFIGURACIÓN DEL SISTEMA DE LOGS
# ═══════════════════════════════════════════════════════════════

def _crear_logger_archivo(nombre: str, ruta: str) -> logging.Logger:
    """Crea un logger dedicado que escribe en un archivo rotativo de máximo 5MB."""
    logger = logging.getLogger(nombre)
    logger.setLevel(logging.INFO)
    logger.propagate = False
    manejador = RotatingFileHandler(ruta, maxBytes=TAMANO_MAXIMO_LOG, backupCount=3, encoding="utf-8")
    manejador.setFormatter(logging.Formatter(fmt="[%(asctime)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    logger.addHandler(manejador)
    return logger


def _configurar_sistema_de_logs() -> tuple:
    """Crea las 6 carpetas de log y configura un logger de errores (consola+archivo) y 5 loggers de archivo."""
    for ruta in [RUTA_LOG_ERRORES, RUTA_LOG_APROBADOS_AGENTE, RUTA_LOG_RECHAZADOS_AGENTE, RUTA_LOG_APROBADOS_HUMANOS, RUTA_LOG_PROVEEDORES_NO_ENCONTRADOS, RUTA_LOG_RECHAZADOS_HUMANOS]:
        Path(os.path.dirname(ruta)).mkdir(parents=True, exist_ok=True)
    logger_main = logging.getLogger("agente_marquillas")
    logger_main.setLevel(logging.ERROR)
    logger_main.propagate = False
    manejador_consola = logging.StreamHandler()
    manejador_consola.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S"))
    logger_main.addHandler(manejador_consola)
    manejador_errores = RotatingFileHandler(RUTA_LOG_ERRORES, maxBytes=TAMANO_MAXIMO_LOG, backupCount=3, encoding="utf-8")
    manejador_errores.setFormatter(logging.Formatter("[%(asctime)s] ERROR CRÍTICO | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    logger_main.addHandler(manejador_errores)
    return (logger_main,
            _crear_logger_archivo("aprobados_agente",             RUTA_LOG_APROBADOS_AGENTE),
            _crear_logger_archivo("rechazados_agente",            RUTA_LOG_RECHAZADOS_AGENTE),
            _crear_logger_archivo("aprobados_humanos",            RUTA_LOG_APROBADOS_HUMANOS),
            _crear_logger_archivo("proveedores_no_encontrados",   RUTA_LOG_PROVEEDORES_NO_ENCONTRADOS),
            _crear_logger_archivo("rechazados_humanos",           RUTA_LOG_RECHAZADOS_HUMANOS))


# Inicializar los loggers al momento de importar el módulo
log, log_aprobados_agente, log_rechazados_agente, log_aprobados_humanos, log_proveedores_no_encontrados, log_rechazados_humanos = _configurar_sistema_de_logs()


def _registrar_aprobado_agente(correo_id: str, asunto: str, resultado: dict, nit_emisor_limpio: str = "", lista_almacen: str = "") -> None:
    """Registra en el log especializado cuando el agente aprueba automáticamente una factura."""
    _NOMBRE_ALMACEN = {
        "almacenSabaneta":         "Sabaneta",
        "almacenRionegro":         "Rionegro",
        "almacenRionegroSabaneta": "Sabaneta y Rionegro",
    }
    emisor          = resultado.get("razon_social_emisor", "N/A")
    numero_factura  = resultado.get("numero_factura", "N/A")
    nombre_almacen  = _NOMBRE_ALMACEN.get(lista_almacen, lista_almacen)
    log_aprobados_agente.info(
        f"APROBADO | Correo: {asunto} | Proveedor: {emisor} | "
        f"NIT Proveedor: {nit_emisor_limpio} | Factura: {numero_factura} | "
        f"Almacén destino: {nombre_almacen}"
    )


def _registrar_rechazado_agente(correo_id: str, asunto: str, resultado: dict) -> None:
    """Registra en el log especializado cuando el agente rechaza una factura."""
    nit_encontrado        = resultado.get("nit_encontrado", "N/A")
    razon_social_encontrada = resultado.get("razon_social_encontrada", "N/A")
    motivo                = resultado.get("motivo", "Sin motivo especificado")
    log_rechazados_agente.info(
        f"RECHAZADO | Correo: {asunto} | NIT Marquillas encontrado: {nit_encontrado} | "
        f"Razón social encontrada: {razon_social_encontrada} | Motivo: {motivo}"
    )


def _registrar_aprobado_humano(correo_id: str, nombre_pdf: str) -> None:
    """Registra en el log especializado cuando un humano aprueba y el original se mueve al archivo."""
    log_aprobados_humanos.info(
        f"APROBADO POR EL ÁREA ENCARGADA | Proveedor: {nombre_pdf}"
    )


def _registrar_rechazado_humano(correo_id: str, nombre_proveedor: str) -> None:
    """Registra en el log especializado cuando un humano rechaza una factura."""
    log_rechazados_humanos.info(
        f"RECHAZADO POR EL ÁREA ENCARGADA | Proveedor: {nombre_proveedor}"
    )


def _registrar_proveedor_no_encontrado(nit_limpio: str, razon_social: str) -> None:
    """Registra en el log especializado cuando el NIT del emisor no existe en proveedores.json."""
    log_proveedores_no_encontrados.info(
        f"NO ENCONTRADO | NIT Proveedor: {nit_limpio} | Nombre Proveedor: {razon_social}"
    )


# ═══════════════════════════════════════════════════════════════
# FUNCIONES PRINCIPALES
# ═══════════════════════════════════════════════════════════════

def cargar_instrucciones_agente() -> str:
    """
    Lee el archivo agente.md que contiene las instrucciones para Claude AI.
    Este archivo define el rol del agente, qué buscar en el PDF y el formato de respuesta.
    Debe existir en la misma carpeta desde donde se ejecuta el programa.
    No recibe parámetros — busca el archivo según la constante ARCHIVO_INSTRUCCIONES.
    Retorna: texto completo del archivo agente.md como string.
    Lanza: FileNotFoundError si el archivo no existe en la carpeta del proyecto.
    """
    try:
        ruta_archivo = Path(ARCHIVO_INSTRUCCIONES)

        if not ruta_archivo.exists():
            raise FileNotFoundError(
                f"No se encontró '{ARCHIVO_INSTRUCCIONES}'. "
                "Asegúrate de ejecutar el agente desde la carpeta del proyecto."
            )

        contenido = ruta_archivo.read_text(encoding="utf-8")
        log.info(f"📋 Instrucciones del agente cargadas ({len(contenido)} caracteres)")
        return contenido

    except FileNotFoundError:
        raise
    except Exception as error:
        log.error(f"💥 Error inesperado al leer {ARCHIVO_INSTRUCCIONES}: {error}")
        raise


def cargar_instrucciones_clasificador() -> str:
    """
    Lee el archivo clasificador.md con las instrucciones para el clasificador de respuestas humanas.
    Usado en Fase 5 para determinar si una respuesta es APROBADO, RECHAZADO o NINGUNO.
    No recibe parámetros — usa la constante ARCHIVO_CLASIFICADOR.
    Retorna: texto completo del archivo clasificador.md como string.
    Lanza: FileNotFoundError si el archivo no existe en la carpeta del proyecto.
    """
    try:
        ruta_archivo = Path(ARCHIVO_CLASIFICADOR)
        if not ruta_archivo.exists():
            raise FileNotFoundError(
                f"No se encontró '{ARCHIVO_CLASIFICADOR}'. "
                "Asegúrate de ejecutar el agente desde la carpeta del proyecto."
            )
        return ruta_archivo.read_text(encoding="utf-8")
    except FileNotFoundError:
        raise
    except Exception as error:
        log.error(f"💥 Error inesperado al leer {ARCHIVO_CLASIFICADOR}: {error}")
        raise


def obtener_token_microsoft() -> str:
    """
    Obtiene un token de acceso temporal de Microsoft Azure usando credenciales de cliente.
    Este token funciona como una llave de seguridad para acceder al correo corporativo.
    Usa el flujo 'client credentials' (máquina a máquina, sin intervención de usuario).
    El token tiene validez de 1 hora; se obtiene uno nuevo en cada ciclo de revisión.
    No recibe parámetros — usa las constantes AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID.
    Retorna: string con el token de acceso de Microsoft Graph API.
    Lanza: Exception si las credenciales son incorrectas o Azure rechaza la solicitud.
    """
    try:
        # Crear la aplicación cliente confidencial de MSAL
        aplicacion_azure = msal.ConfidentialClientApplication(
            client_id=AZURE_CLIENT_ID,
            client_credential=AZURE_CLIENT_SECRET,
            authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
        )

        # Solicitar el token de acceso al servicio de Azure
        resultado_token = aplicacion_azure.acquire_token_for_client(scopes=SCOPES_MICROSOFT)

        if "access_token" not in resultado_token:
            descripcion = resultado_token.get("error_description", "Sin descripción del error")
            raise Exception(f"Azure rechazó las credenciales: {descripcion}")

        log.info("🔑 Token de Microsoft obtenido exitosamente")
        return resultado_token["access_token"]

    except Exception as error:
        log.error(f"💥 Error al obtener token de Microsoft: {error}")
        raise


def obtener_correos_nuevos(token: str) -> list:
    """
    Consulta la bandeja de entrada via Microsoft Graph API buscando correos no leídos con adjuntos.
    Solo recupera correos que tienen hasAttachments=true para no procesar correos vacíos.
    Incluye metadatos de los adjuntos (nombre, tipo, tamaño) pero NO su contenido binario.
    Recibe: token (str) — token de acceso obtenido con obtener_token_microsoft().
    Retorna: lista de diccionarios, cada uno con los datos de un correo y sus adjuntos.
    Retorna lista vacía [] si no hay correos nuevos con adjuntos.
    Lanza: Exception si la llamada a Graph API falla (token inválido, sin conexión, etc.).
    """
    try:
        url_correos  = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/mailFolders/inbox/messages"
        encabezados  = {"Authorization": f"Bearer {token}"}
        parametros   = {
            "$filter":  "isRead eq false and hasAttachments eq true",
            "$select":  "id,subject,from,receivedDateTime,hasAttachments",
            "$expand":  "attachments($select=id,name,contentType,size)",
            "$top":     "10",
        }

        respuesta = requests.get(
            url_correos, headers=encabezados, params=parametros, timeout=30
        )
        respuesta.raise_for_status()
        print(f"[DIAGNÓSTICO] URL consultada: {respuesta.url}")
        print(f"[DIAGNÓSTICO] Total correos retornados: {len(respuesta.json().get('value', []))}")

        lista_correos = respuesta.json().get("value", [])
        return lista_correos

    except requests.exceptions.RequestException as error:
        log.error(f"💥 Error al consultar correos en Microsoft Graph: {error}")
        raise


def encontrar_adjunto_zip(correo: dict) -> dict | None:
    """
    Busca entre los adjuntos de un correo alguno que sea un archivo ZIP.
    Verifica tanto la extensión del nombre del archivo como el tipo de contenido MIME.
    Recibe: correo (dict) — objeto correo tal como lo devuelve Microsoft Graph API,
            que incluye la lista 'attachments' con los metadatos de cada adjunto.
    Retorna: diccionario con los datos del primer adjunto ZIP encontrado.
    Retorna None si el correo no tiene ningún adjunto ZIP.
    No lanza excepciones — simplemente retorna None ante cualquier problema.
    """
    try:
        lista_adjuntos = correo.get("attachments", [])

        for adjunto in lista_adjuntos:
            nombre_archivo = adjunto.get("name", "").lower()
            tipo_mime      = adjunto.get("contentType", "").lower()

            # Identificar el ZIP por extensión de nombre o por tipo MIME
            es_zip_por_nombre = nombre_archivo.endswith(".zip")
            es_zip_por_tipo   = "zip" in tipo_mime or "compressed" in tipo_mime

            if es_zip_por_nombre or es_zip_por_tipo:
                log.info(f"📦 ZIP encontrado: {adjunto.get('name', 'sin nombre')}")
                return adjunto

        return None  # No se encontró ningún adjunto ZIP

    except Exception as error:
        log.warning(f"⚠️  Error al buscar adjunto ZIP en el correo: {error}")
        return None


def descargar_adjunto(token: str, correo_id: str, adjunto_id: str) -> bytes:
    """
    Descarga el contenido binario de un adjunto específico usando Microsoft Graph API.
    El adjunto se mantiene completamente en memoria, sin guardar nada en disco.
    Graph API devuelve el contenido codificado en base64, que esta función decodifica.
    Recibe:
      - token (str): token de acceso de Microsoft para autorizar la descarga.
      - correo_id (str): identificador único del correo en Microsoft Graph.
      - adjunto_id (str): identificador único del adjunto dentro del correo.
    Retorna: bytes con el contenido binario crudo del archivo adjunto.
    Lanza: Exception si el adjunto no existe o hay error de red.
    """
    try:
        url_adjunto = (
            f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}"
            f"/messages/{correo_id}/attachments/{adjunto_id}"
        )
        encabezados = {"Authorization": f"Bearer {token}"}

        respuesta = requests.get(url_adjunto, headers=encabezados, timeout=60)
        respuesta.raise_for_status()

        datos_adjunto   = respuesta.json()
        contenido_b64   = datos_adjunto.get("contentBytes", "")

        if not contenido_b64:
            raise ValueError("El adjunto no tiene contenido (contentBytes vacío)")

        bytes_adjunto = base64.b64decode(contenido_b64)
        log.info(f"📥 Adjunto descargado: {len(bytes_adjunto):,} bytes en memoria")
        return bytes_adjunto

    except Exception as error:
        log.error(f"💥 Error al descargar el adjunto: {error}")
        raise


def extraer_pdf_del_zip(bytes_zip: bytes) -> tuple:
    """
    Abre un archivo ZIP desde memoria (sin guardarlo en disco) y extrae el primer PDF que encuentre.
    Usa io.BytesIO para crear un archivo virtual en RAM que la librería zipfile puede leer.
    Si el ZIP contiene varios PDFs, siempre toma el primero en orden alfabético.
    Recibe: bytes_zip (bytes) — contenido binario del archivo ZIP descargado del correo.
    Retorna: tupla (nombre_pdf: str, bytes_pdf: bytes) con el nombre y contenido del PDF.
    Retorna (None, None) si el ZIP no contiene ningún archivo con extensión .pdf.
    Lanza: zipfile.BadZipFile si el contenido no es un ZIP válido o está corrupto.
    """
    try:
        # Abrir el ZIP directamente desde los bytes en memoria usando un buffer virtual
        with zipfile.ZipFile(io.BytesIO(bytes_zip), "r") as archivo_zip:
            nombres_internos = archivo_zip.namelist()

            for nombre_interno in nombres_internos:
                if nombre_interno.lower().endswith(".pdf"):
                    bytes_pdf = archivo_zip.read(nombre_interno)
                    log.info(f"📄 PDF extraído: {nombre_interno} ({len(bytes_pdf):,} bytes)")
                    return nombre_interno, bytes_pdf

        log.warning("⚠️  El ZIP no contiene ningún archivo PDF")
        return None, None

    except zipfile.BadZipFile:
        log.error("💥 El archivo descargado no es un ZIP válido o está corrompido")
        raise
    except Exception as error:
        log.error(f"💥 Error al extraer el PDF del ZIP: {error}")
        raise


def leer_texto_del_pdf(bytes_pdf: bytes) -> str:
    """
    Extrae el texto visible de un PDF usando PyMuPDF con estrategia de doble extracción.
    Primera pasada: get_text("text") — extracción estándar por página.
    Si la primera página devuelve menos de 200 caracteres, segunda pasada con get_text("blocks")
    que captura áreas que el método simple omite; ambos resultados se combinan sin duplicados.
    Limita el resultado final a MAXIMO_CARACTERES_PDF para no exceder los tokens de Claude.
    Recibe: bytes_pdf (bytes) — contenido binario del PDF extraído del ZIP.
    Retorna: string con el texto más completo posible, truncado a MAXIMO_CARACTERES_PDF.
    Retorna string vacío si el PDF es una imagen escaneada sin texto seleccionable.
    Lanza: Exception si los bytes no corresponden a un PDF válido.
    """
    try:
        documento = fitz.open(stream=bytes_pdf, filetype="pdf")

        texto_total = ""
        for numero_pagina in range(len(documento)):
            pagina = documento[numero_pagina]

            # Primera pasada: extracción estándar
            texto_simple = pagina.get_text("text")

            # Segunda pasada solo en primera página si el texto simple es escaso
            if numero_pagina == 0 and len(texto_simple) < 200:
                bloques = pagina.get_text("blocks")
                # Cada bloque es una tupla; el índice 4 contiene el texto
                texto_bloques = "\n".join(
                    b[4] for b in bloques if isinstance(b[4], str)
                )
                # Combinar: agregar líneas de bloques que no estén ya en el texto simple
                lineas_simples  = set(texto_simple.splitlines())
                lineas_extra    = [
                    linea for linea in texto_bloques.splitlines()
                    if linea.strip() and linea not in lineas_simples
                ]
                texto_pagina = texto_simple + "\n".join(lineas_extra)
            else:
                texto_pagina = texto_simple

            texto_total += texto_pagina

        documento.close()

        # Avisar si la primera página tiene muy poco texto — posible PDF con imágenes
        primera_pagina_chars = len(texto_total.split("\f")[0]) if "\f" in texto_total else len(texto_total)
        if primera_pagina_chars < 300:
            log.error(
                f"Módulo: leer_texto_del_pdf | "
                f"Descripción: primera página con solo {primera_pagina_chars} caracteres — "
                f"el PDF puede tener encabezado en imagen o formato especial | "
                f"Detalle técnico: considerar extracción OCR"
            )

        texto_truncado = texto_total[:MAXIMO_CARACTERES_PDF]
        log.info(
            f"📝 Texto extraído: {len(texto_total)} caracteres "
            f"(enviando {len(texto_truncado)} a Claude)"
        )

        return texto_truncado

    except Exception as error:
        log.error(f"💥 Error al leer el texto del PDF: {error}")
        raise


def verificar_documento_con_claude(texto_pdf: str, instrucciones: str) -> dict:
    """
    Envía el texto del PDF a Claude AI para verificar si corresponde a Marquillas S.A.S.
    Claude analiza el texto buscando el NIT y la razón social según las reglas del agente.md.
    Espera una respuesta en formato JSON puro con los campos: nit_encontrado,
    razon_social_encontrada, aprobado y motivo.
    Recibe:
      - texto_pdf (str): texto extraído del PDF con leer_texto_del_pdf().
      - instrucciones (str): contenido completo del archivo agente.md.
    Retorna: diccionario Python con el resultado de la verificación.
    Lanza: json.JSONDecodeError si Claude no responde con JSON válido.
    Lanza: Exception si hay error de conexión con la API de Anthropic.
    """
    try:
        log.info("🤖 Consultando a Claude...")
        cliente_anthropic = Anthropic(api_key=ANTHROPIC_API_KEY)

        contenido_usuario = (
            "Analiza el siguiente texto extraído de un PDF de factura "
            "y responde siguiendo exactamente las instrucciones del sistema.\n\n"
            f"TEXTO DEL PDF:\n{texto_pdf}"
        )

        respuesta_claude = cliente_anthropic.messages.create(
            model=MODELO_CLAUDE,
            max_tokens=500,
            system=instrucciones,
            messages=[{"role": "user", "content": contenido_usuario}]
        )

        texto_respuesta = respuesta_claude.content[0].text.strip()

        # Limpiar bloques de código Markdown que Claude podría agregar por error
        if "```" in texto_respuesta:
            partes = texto_respuesta.split("```")
            texto_respuesta = partes[1] if len(partes) > 1 else texto_respuesta
            if texto_respuesta.startswith("json"):
                texto_respuesta = texto_respuesta[4:].strip()

        resultado_json = json.loads(texto_respuesta)
        return resultado_json

    except json.JSONDecodeError:
        log.error("⚠️  Claude no respondió con JSON válido — se omitirá este correo")
        raise
    except Exception as error:
        log.error(f"💥 Error al consultar a Claude AI: {error}")
        raise


def convertir_pdf_a_imagenes(bytes_pdf: bytes) -> list:
    """
    Convierte las primeras 2 páginas de un PDF a imágenes en base64.

    Usa PyMuPDF que ya está instalado. Se usa como respaldo cuando
    Claude no encuentra el NIT o razón social del proveedor en el
    texto del PDF, lo que indica que están en una imagen.

    Recibe:
    - bytes_pdf: contenido del PDF en memoria como bytes

    Retorna: lista de strings en base64, máximo 2 páginas.
    Resolución 150 DPI — suficiente para leer texto sin gastar créditos.
    """
    try:
        documento  = fitz.open(stream=bytes_pdf, filetype="pdf")
        paginas    = min(len(documento), 2)
        imagenes   = []
        matriz_dpi = fitz.Matrix(150 / 72, 150 / 72)  # 150 DPI

        for numero_pagina in range(paginas):
            pagina       = documento[numero_pagina]
            pixmap       = pagina.get_pixmap(matrix=matriz_dpi)
            bytes_imagen = pixmap.tobytes("png")
            imagenes.append(base64.b64encode(bytes_imagen).decode("utf-8"))

        documento.close()
        log.info(f"🖼️  PDF convertido a {len(imagenes)} imagen(es) para visión")
        return imagenes

    except Exception as error:
        log.error(f"💥 Error al convertir PDF a imágenes: {error}")
        raise


def verificar_pdf_con_imagenes(bytes_pdf: bytes, instrucciones: str) -> dict:
    """
    Verifica un PDF enviando imágenes de sus páginas a Claude
    en vez del texto extraído.

    Solo se llama cuando verificar_documento_con_claude() no encontró
    el NIT o razón social del proveedor — señal de PDF con imágenes.

    Recibe:
    - bytes_pdf: contenido del PDF en memoria como bytes
    - instrucciones: contenido del archivo agente.md

    Retorna: diccionario con el mismo formato JSON que
    verificar_documento_con_claude()
    """
    try:
        log.info("🤖 Consultando a Claude con imágenes del PDF...")
        imagenes_b64      = convertir_pdf_a_imagenes(bytes_pdf)
        cliente_anthropic = Anthropic(api_key=ANTHROPIC_API_KEY)

        contenido_usuario = [
            {
                "type": "image",
                "source": {
                    "type":       "base64",
                    "media_type": "image/png",
                    "data":       imagen_b64,
                },
            }
            for imagen_b64 in imagenes_b64
        ]
        contenido_usuario.append({
            "type": "text",
            "text": "Analiza estas imágenes del PDF de factura y extrae la información solicitada.",
        })

        respuesta_claude = cliente_anthropic.messages.create(
            model=MODELO_CLAUDE,
            max_tokens=500,
            system=instrucciones,
            messages=[{"role": "user", "content": contenido_usuario}]
        )

        texto_respuesta = respuesta_claude.content[0].text.strip()

        # Limpiar bloques de código Markdown que Claude podría agregar por error
        if "```" in texto_respuesta:
            partes = texto_respuesta.split("```")
            texto_respuesta = partes[1] if len(partes) > 1 else texto_respuesta
            if texto_respuesta.startswith("json"):
                texto_respuesta = texto_respuesta[4:].strip()

        return json.loads(texto_respuesta)

    except json.JSONDecodeError:
        log.error("⚠️  Claude (visión) no respondió con JSON válido")
        raise
    except Exception as error:
        log.error(f"💥 Error al verificar PDF con imágenes: {error}")
        raise


def enviar_correo_aprobado(
    token: str, bytes_pdf: bytes, nombre_pdf: str, resultado: dict,
    destinatarios: dict
) -> None:
    """
    Envía un correo electrónico con el PDF adjunto usando destinatarios principales y en copia.
    Solo se llama cuando Claude verifica que el NIT y la razón social son correctos (aprobado=True).
    El asunto es '{razon_social_emisor} - {numero_factura}' y el cuerpo incluye REF-AGENTE para detección de respuestas.
    Se envía UN SOLO correo con todos los destinatarios — principales en toRecipients, copia en ccRecipients.
    Recibe:
      - token (str): token de acceso de Microsoft.
      - bytes_pdf (bytes): contenido binario del PDF a adjuntar en el correo.
      - nombre_pdf (str): nombre del archivo PDF para mostrarlo como adjunto.
      - resultado (dict): diccionario de Claude con razon_social_emisor y numero_factura.
      - destinatarios (dict): diccionario con claves "principales" y "copia", cada una lista de correos.
    No retorna nada. Lanza: Exception si hay error al enviar via Microsoft Graph.
    """
    try:
        principales = destinatarios.get("principales", [])
        copia       = destinatarios.get("copia", [])

        # Codificar el PDF en base64 — formato requerido por Graph API para adjuntos
        contenido_pdf_b64 = base64.b64encode(bytes_pdf).decode("utf-8")
        emisor            = resultado.get("razon_social_emisor", "N/A")
        numero_factura    = resultado.get("numero_factura", "")

        asunto_correo = f"{emisor} - {numero_factura}"
        cuerpo_html   = (
            "<p>Buen día,</p>"
            "<p>Comparto la siguiente factura para su revisión y gestión/aceptación.</p>"
            "<br>"
            f"<p style=\"color: white; font-size: 1px;\">REF-AGENTE: {emisor} - {numero_factura}</p>"
            "<br>"
            "<p><i>Correo enviado por agente de automatización Marquillas.</i></p>"
        )

        estructura_correo = {
            "message": {
                "subject": asunto_correo,
                "body": {"contentType": "HTML", "content": cuerpo_html},
                "toRecipients": [{"emailAddress": {"address": e}} for e in principales],
                "ccRecipients": [{"emailAddress": {"address": e}} for e in copia if e],
                "attachments": [{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": nombre_pdf,
                    "contentType": "application/pdf",
                    "contentBytes": contenido_pdf_b64,
                }],
            }
        }

        url_envio   = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/sendMail"
        encabezados = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

        respuesta = requests.post(url_envio, headers=encabezados, json=estructura_correo, timeout=30)
        respuesta.raise_for_status()
        log.info(f"📤 Correo enviado a {principales} con copia a {copia}")

    except Exception as error:
        log.error(f"💥 Error al enviar correo aprobado: {error}")
        raise


def marcar_correo_como_leido(token: str, correo_id: str) -> None:
    """
    Marca un correo específico como 'leído' en Outlook para evitar reprocesarlo.
    Esta es siempre la última operación de cada ciclo de procesamiento de correo.
    Se ejecuta tanto si el correo fue aprobado, rechazado o si hubo un error.
    Recibe:
      - token (str): token de acceso de Microsoft.
      - correo_id (str): identificador único del correo en Microsoft Graph API.
    No retorna nada.
    Lanza: Exception si hay error al actualizar el estado del correo en Outlook.
    """
    try:
        url_correo  = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/messages/{correo_id}"
        encabezados = {
            "Authorization":  f"Bearer {token}",
            "Content-Type":   "application/json",
        }
        datos_actualizacion = {"isRead": True}

        respuesta = requests.patch(
            url_correo, headers=encabezados, json=datos_actualizacion, timeout=15
        )
        respuesta.raise_for_status()
        log.info("👁️  Correo marcado como leído")

    except Exception as error:
        log.error(f"⚠️  Error al marcar correo como leído: {error}")
        raise


def _registrar_y_actuar(
    token: str, resultado: dict, bytes_pdf: bytes, nombre_pdf: str,
    correo_id: str = "", asunto: str = ""
) -> None:
    """
    Registra el resultado de la verificación de Claude y ejecuta la acción correspondiente.
    Si aprobado=True: registra en log de aprobados y envía el correo con el PDF adjunto.
    Si aprobado=False: registra en log de rechazados sin enviar ningún correo.
    Recibe:
      - token (str): token de acceso de Microsoft para enviar el correo si aplica.
      - resultado (dict): respuesta de Claude con nit_encontrado, razon_social, aprobado, motivo.
      - bytes_pdf (bytes): contenido binario del PDF para adjuntar si es aprobado.
      - nombre_pdf (str): nombre del archivo PDF.
      - correo_id (str): ID del correo para trazabilidad en el log.
      - asunto (str): asunto del correo para trazabilidad en el log.
    No retorna nada.
    """
    if resultado.get("aprobado"):
        _registrar_aprobado_agente(correo_id, asunto, resultado)
        enviar_correo_aprobado(token, bytes_pdf, nombre_pdf, resultado)
    else:
        _registrar_rechazado_agente(correo_id, asunto, resultado)


# ═══════════════════════════════════════════════════════════════
# FASE 2 — DETECCIÓN Y PROCESAMIENTO DE RESPUESTAS DE APROBACIÓN
# ═══════════════════════════════════════════════════════════════

def es_respuesta_humana(correo: dict) -> bool:
    """
    Determina si un correo es una respuesta humana a una factura enviada por el agente.
    Para ser considerado respuesta humana debe cumplir DOS condiciones al mismo tiempo:
      1. El asunto contiene "RE:" — confirma que es una respuesta, no un correo nuevo.
      2. El cuerpo contiene "REF-AGENTE:" — confirma que es respuesta a un correo del agente.
    La clasificación de aprobación/rechazo/ninguno la realiza clasificar_respuesta_humana().
    Recibe: correo (dict) — objeto correo de Graph API con los campos 'subject' y 'body'.
    Retorna: True si el correo cumple las DOS condiciones al mismo tiempo.
    Retorna: False en cualquier otro caso o si ocurre algún error.
    No lanza excepciones.
    """
    try:
        asunto    = correo.get("subject", "").upper()
        contenido = correo.get("body", {}).get("content", "")

        # Condición 1: el asunto es una respuesta (RE:)
        if "RE:" not in asunto:
            return False

        # Condición 2: el cuerpo contiene la marca del agente
        if "REF-AGENTE:" not in contenido:
            return False

        return True

    except Exception:
        return False


def _encontrar_adjunto_pdf(correo: dict) -> dict | None:
    """
    Busca el primer adjunto PDF en la lista de adjuntos de un correo.
    Identifica el PDF por extensión del nombre del archivo o por tipo MIME.
    Recibe: correo (dict) — objeto correo con la lista 'attachments' expandida de Graph API.
    Retorna: diccionario con los metadatos del primer adjunto PDF encontrado.
    Retorna None si el correo no contiene ningún adjunto de tipo PDF.
    No lanza excepciones.
    """
    try:
        for adjunto in correo.get("attachments", []):
            nombre = adjunto.get("name", "").lower()
            tipo   = adjunto.get("contentType", "").lower()
            if nombre.endswith(".pdf") or "pdf" in tipo:
                return adjunto
        return None
    except Exception:
        return None


def obtener_correo_original_del_hilo(token: str, conversation_id: str) -> dict | None:
    """
    Busca todos los mensajes del mismo hilo de conversación y retorna el más antiguo.
    El correo más antiguo es el original — el que contiene el PDF que fue enviado al aprobador.
    Incluye los metadatos de los adjuntos para poder identificar el PDF del correo original.
    Recibe:
      - token (str): token de acceso de Microsoft Graph API.
      - conversation_id (str): ID de la conversación obtenido del correo de respuesta.
    Retorna: diccionario con los datos del correo más antiguo del hilo, incluyendo adjuntos.
    Retorna None si no se encontró ningún correo con ese conversationId en el buzón.
    Lanza: Exception si hay error de red o la llamada a Graph API falla.
    """
    try:
        url_correos = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/messages"
        encabezados = {"Authorization": f"Bearer {token}"}
        parametros  = {
            "$filter": f"conversationId eq '{conversation_id}'",
            "$select": "id,subject,receivedDateTime,hasAttachments",
            "$expand": "attachments($select=id,name,contentType,size)",
            "$top":    "50",
        }

        respuesta = requests.get(url_correos, headers=encabezados, params=parametros, timeout=30)
        respuesta.raise_for_status()

        correos_hilo = respuesta.json().get("value", [])
        if not correos_hilo:
            return None

        # Ordenar por fecha de recepción ascendente — el primero es el correo original
        correo_original = sorted(
            correos_hilo, key=lambda c: c.get("receivedDateTime", "")
        )[0]
        return correo_original

    except Exception as error:
        log.error(f"💥 Error al buscar el correo original del hilo: {error}")
        raise


def obtener_id_carpeta_outlook(token: str, nombre_carpeta: str) -> str | None:
    """
    Busca una carpeta de Outlook por su nombre y retorna su ID interno de Graph API.
    La comparación del nombre es insensible a mayúsculas y minúsculas.
    Recibe:
      - token (str): token de acceso de Microsoft Graph API.
      - nombre_carpeta (str): nombre exacto de la carpeta a buscar (ej: "FACTURAS APROBADAS").
    Retorna: string con el ID de la carpeta si existe en el buzón del usuario.
    Retorna None si no existe ninguna carpeta con ese nombre — el agente registra el error.
    Lanza: Exception si hay error de red o la llamada a Graph API falla.
    """
    try:
        url_carpetas = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/mailFolders"
        encabezados  = {"Authorization": f"Bearer {token}"}
        parametros   = {"$top": "50", "$select": "id,displayName"}

        respuesta = requests.get(url_carpetas, headers=encabezados, params=parametros, timeout=15)
        respuesta.raise_for_status()

        for carpeta in respuesta.json().get("value", []):
            if carpeta.get("displayName", "").upper() == nombre_carpeta.upper():
                return carpeta.get("id")

        return None

    except Exception as error:
        log.error(f"💥 Error al buscar la carpeta '{nombre_carpeta}' en Outlook: {error}")
        raise


def mover_correo_a_carpeta(token: str, correo_id: str, carpeta_id: str) -> None:
    """
    Mueve un correo de su ubicación actual a una carpeta específica de Outlook.
    Usa el endpoint POST /messages/{id}/move de Microsoft Graph API.
    El correo desaparece de su carpeta de origen y aparece en la carpeta destino.
    Recibe:
      - token (str): token de acceso de Microsoft Graph API.
      - correo_id (str): ID único del correo a mover.
      - carpeta_id (str): ID de la carpeta destino obtenido con obtener_id_carpeta_outlook().
    No retorna nada.
    Lanza: Exception si hay error de red o si alguno de los IDs no es válido.
    """
    try:
        url_mover   = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/messages/{correo_id}/move"
        encabezados = {
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        }
        respuesta = requests.post(
            url_mover, headers=encabezados, json={"destinationId": carpeta_id}, timeout=15
        )
        respuesta.raise_for_status()

    except Exception as error:
        log.error(f"💥 Error al mover el correo a la carpeta de destino: {error}")
        raise


def procesar_aprobacion(token: str, correo_respuesta: dict) -> None:
    """
    Orquesta todo el proceso cuando se detecta una respuesta de aprobación humana.
    Pasos: busca el correo original del hilo → identifica el PDF → mueve el correo a 'FACTURAS APROBADAS'.
    Si la carpeta 'FACTURAS APROBADAS' no existe, registra el error claramente y no falla el programa.
    Recibe:
      - token (str): token de acceso de Microsoft.
      - correo_respuesta (dict): el correo de respuesta detectado como aprobación.
    No retorna nada — todos los eventos quedan registrados en el log.
    No propaga excepciones — los errores son capturados y registrados internamente.
    """
    asunto = correo_respuesta.get("subject", "Sin asunto")
    log.info(f"📨 Respuesta de aprobación detectada: '{asunto}'")
    try:
        log.info("🔍 Buscando correo original en el hilo...")
        correo_original = obtener_correo_original_del_hilo(
            token, correo_respuesta.get("conversationId", "")
        )
        if not correo_original:
            log.error("⚠️  No se encontró el correo original del hilo — se omite la aprobación")
            return

        # Identificar el PDF adjunto en el correo original y descargarlo
        adjunto_pdf = _encontrar_adjunto_pdf(correo_original)
        nombre_pdf  = adjunto_pdf.get("name", "factura.pdf") if adjunto_pdf else "factura.pdf"
        if adjunto_pdf:
            log.info(f"📄 PDF encontrado en correo original: {nombre_pdf}")
            descargar_adjunto(token, correo_original["id"], adjunto_pdf["id"])

        # Obtener la carpeta destino y mover el correo original
        log.info(f"📁 Moviendo a carpeta {CARPETA_FACTURAS_APROBADAS}...")
        carpeta_id = obtener_id_carpeta_outlook(token, CARPETA_FACTURAS_APROBADAS)
        if not carpeta_id:
            log.error(f"❌ La carpeta '{CARPETA_FACTURAS_APROBADAS}' no existe en Outlook. Créala manualmente.")
            return

        mover_correo_a_carpeta(token, correo_original["id"], carpeta_id)
        _registrar_aprobado_humano(correo_original["id"], nombre_pdf)
        log.info(f"✅ Factura {nombre_pdf} movida a {CARPETA_FACTURAS_APROBADAS} exitosamente")

    except Exception as error:
        log.error(f"💥 Error al procesar la aprobación del correo '{asunto}': {error}")


def obtener_correos_aprobacion(token: str) -> list:
    """
    Obtiene correos no leídos cuyo asunto comienza con 'RE:' para candidatos a aprobación.
    Usar startswith(subject,'RE:') en el filtro de Graph API pre-filtra en el servidor:
    solo trae respuestas, nunca correos originales enviados por el agente.
    La validación definitiva (REF-AGENTE: + palabra de aprobación) la hace es_respuesta_de_aprobacion().
    Registra en el log todos los correos candidatos encontrados para facilitar el diagnóstico.
    Recibe: token (str) — token de acceso obtenido con obtener_token_microsoft().
    Retorna: lista de todos los correos no leídos cuyo asunto comienza con 'RE:'.
    Retorna lista vacía [] si no hay candidatos.
    Lanza: Exception si hay error de red o la llamada a Graph API falla.
    """
    try:
        url_correos = f"{URL_GRAPH_API}/users/{EMAIL_MONITOREAR}/mailFolders/inbox/messages"
        encabezados = {"Authorization": f"Bearer {token}"}
        # startswith('RE:') está soportado en Graph API — pre-filtra en el servidor
        # sin filtrar por hasAttachments para no perder respuestas con firma de imagen
        parametros  = {
            "$filter": "isRead eq false and startswith(subject,'RE:')",
            "$select": "id,subject,conversationId,body,from,receivedDateTime,hasAttachments",
            "$top":    "20",
        }

        respuesta = requests.get(url_correos, headers=encabezados, params=parametros, timeout=30)
        respuesta.raise_for_status()

        candidatos = respuesta.json().get("value", [])

        # Log de diagnóstico — mostrar qué correos con RE: se encontraron
        log.info(f"🔎 Correos no leídos con 'RE:' en asunto: {len(candidatos)}")
        for correo in candidatos:
            asunto_diag   = correo.get("subject", "Sin asunto")
            tiene_adjunto = correo.get("hasAttachments", False)
            icono_adjunto = "📎" if tiene_adjunto else "  "
            log.info(f"   {icono_adjunto} '{asunto_diag}'")

        if candidatos:
            log.info(f"📨 {len(candidatos)} correo(s) con 'RE:' encontrado(s) — se validarán con REF-AGENTE")
        else:
            log.info("📭 Ningún correo con 'RE:' en el asunto encontrado")

        return candidatos

    except requests.exceptions.RequestException as error:
        log.error(f"💥 Error al consultar correos candidatos a aprobación: {error}")
        raise


# ═══════════════════════════════════════════════════════════════
# FASE 5 — CLASIFICACIÓN DE RESPUESTAS HUMANAS
# ═══════════════════════════════════════════════════════════════

def clasificar_con_claude(texto: str, instrucciones_clasificador: str) -> str:
    """
    Clasifica el cuerpo de un correo de respuesta humana usando Claude Haiku.
    Antes de enviar a Claude, recorta el historial citado de Outlook para que el modelo
    analice únicamente el texto nuevo escrito por el humano.
    Recibe:
      - texto (str): cuerpo completo del correo (puede incluir historial de Outlook).
      - instrucciones_clasificador (str): contenido de clasificador.md.
    Retorna: "APROBADO", "RECHAZADO" o "NINGUNO" (según responda Claude).
    Retorna "NINGUNO" si Claude responde con algo inesperado o hay un error.
    """
    # Cortar el historial citado de Outlook — solo analizar lo que escribió el humano
    marcadores = ["De:", "From:", "Enviado:", "Sent:", "________________________________"]
    lineas = texto.splitlines()
    for i, linea in enumerate(lineas):
        linea_strip = linea.strip()
        if any(linea_strip.startswith(m) for m in marcadores):
            texto = "\n".join(lineas[:i]).strip()
            break

    if not texto:
        return "NINGUNO"

    try:
        cliente_anthropic = Anthropic(api_key=ANTHROPIC_API_KEY)
        respuesta_claude  = cliente_anthropic.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=10,
            system=instrucciones_clasificador,
            messages=[{"role": "user", "content": texto}]
        )
        resultado = respuesta_claude.content[0].text.strip().upper()
        if resultado in ("APROBADO", "RECHAZADO", "NINGUNO"):
            return resultado
        log.error(f"⚠️  Clasificador Claude retornó valor inesperado: '{resultado}' — se trata como NINGUNO")
        return "NINGUNO"
    except Exception as error:
        log.error(f"💥 Error al consultar clasificador Claude: {error}")
        return "NINGUNO"


def clasificar_respuesta_humana(correo: dict, instrucciones_clasificador: str) -> str:
    """
    Clasifica la intención del cuerpo de un correo de respuesta humana usando Claude Haiku.
    Recibe:
      - correo (dict): objeto correo de Graph API con el campo 'body'.
      - instrucciones_clasificador (str): contenido de clasificador.md.
    Retorna: "APROBADO", "RECHAZADO" o "NINGUNO".
    """
    texto = correo.get("body", {}).get("content", "")
    log.info("🤖 Consultando clasificador Claude...")
    clasificacion = clasificar_con_claude(texto, instrucciones_clasificador)
    log.info(f"🏷️  Clasificación: {clasificacion}")
    return clasificacion


def procesar_rechazo(token: str, correo_respuesta: dict) -> None:
    """
    Orquesta todo el proceso cuando se detecta una respuesta de rechazo humana.
    Pasos: busca el correo original del hilo → mueve el correo a CARPETA_FACTURAS_RECHAZADAS.
    Si la carpeta no existe, registra el error claramente y no falla el programa.
    Recibe:
      - token (str): token de acceso de Microsoft.
      - correo_respuesta (dict): el correo de respuesta detectado como rechazo.
    No retorna nada — todos los eventos quedan registrados en el log.
    No propaga excepciones — los errores son capturados y registrados internamente.
    """
    asunto = correo_respuesta.get("subject", "Sin asunto")
    log.info(f"📨 Respuesta de rechazo detectada: '{asunto}'")
    try:
        correo_original = obtener_correo_original_del_hilo(
            token, correo_respuesta.get("conversationId", "")
        )
        if not correo_original:
            log.error("⚠️  No se encontró el correo original del hilo — se omite el rechazo")
            return

        adjunto_pdf  = _encontrar_adjunto_pdf(correo_original)
        nombre_pdf   = adjunto_pdf.get("name", "factura.pdf") if adjunto_pdf else "factura.pdf"

        log.info(f"📁 Moviendo a carpeta {CARPETA_FACTURAS_RECHAZADAS}...")
        carpeta_id = obtener_id_carpeta_outlook(token, CARPETA_FACTURAS_RECHAZADAS)
        if not carpeta_id:
            log.error(f"❌ La carpeta '{CARPETA_FACTURAS_RECHAZADAS}' no existe en Outlook. Créala manualmente.")
            return

        mover_correo_a_carpeta(token, correo_original["id"], carpeta_id)
        _registrar_rechazado_humano(correo_original["id"], nombre_pdf)
        log.info(f"✅ Factura {nombre_pdf} movida a {CARPETA_FACTURAS_RECHAZADAS} exitosamente")

    except Exception as error:
        log.error(f"💥 Error al procesar el rechazo del correo '{asunto}': {error}")


def procesar_un_correo(token: str, correo: dict, instrucciones: str) -> None:
    """
    Orquesta el procesamiento de un correo individual identificando cuál de los dos casos aplica.
    CASO 1 — Factura nueva: el correo tiene adjunto ZIP → verifica el PDF con Claude AI.
    CASO 2 — Aprobación humana: el correo es una respuesta con palabra de aprobación → mueve el original.
    El correo se marca como leído ÚNICAMENTE cuando el flujo completo termina exitosamente:
      - Caso 1: Claude aprobó + proveedor encontrado + correo enviado.
      - Caso 2: aprobación detectada + correo original encontrado + movido a carpeta.
    En cualquier otro caso (rechazo, proveedor no encontrado, error) el correo queda sin leer.
    Recibe:
      - token (str): token de acceso de Microsoft.
      - correo (dict): datos del correo con asunto, cuerpo y adjuntos según corresponda.
      - instrucciones (str): contenido del agente.md para enviarlo a Claude (solo Caso 1).
    No retorna nada — todos los resultados quedan registrados en el log.
    """
    correo_id = correo.get("id", "")
    asunto    = correo.get("subject", "Sin asunto")
    try:
        # ── Caso 2: Respuesta humana a una factura del agente ───────────────────
        if es_respuesta_humana(correo):
            instrucciones_clasificador = cargar_instrucciones_clasificador()
            clasificacion = clasificar_respuesta_humana(correo, instrucciones_clasificador)
            if clasificacion == "APROBADO":
                procesar_aprobacion(token, correo)
                marcar_correo_como_leido(token, correo_id)
            elif clasificacion == "RECHAZADO":
                procesar_rechazo(token, correo)
                marcar_correo_como_leido(token, correo_id)
            else:
                log.info(f"⏭️  Respuesta clasificada como NINGUNO — se ignora: '{asunto}'")
            return

        # ── Caso 1: Factura nueva con adjunto ZIP ────────────────────────────────
        adjunto_zip = encontrar_adjunto_zip(correo)
        if adjunto_zip:
            log.info(f"📧 Procesando: '{asunto}'")
            bytes_zip             = descargar_adjunto(token, correo_id, adjunto_zip.get("id", ""))
            nombre_pdf, bytes_pdf = extraer_pdf_del_zip(bytes_zip)
            if not bytes_pdf:
                log.warning(f"⚠️  El ZIP del correo '{asunto}' no contiene PDF — se omite")
                return  # No marcar como leído — ZIP sin PDF no es un éxito

            texto_pdf = leer_texto_del_pdf(bytes_pdf)
            resultado = verificar_documento_con_claude(texto_pdf, instrucciones)

            # Respaldo de visión cuando el texto no contiene datos del emisor
            if (resultado.get("nit_emisor") == "NO ENCONTRADO"
                    or resultado.get("razon_social_emisor") == "NO ENCONTRADA"):
                log.warning(
                    "Módulo: procesar_un_correo | "
                    "Descripción: PDF con texto insuficiente, usando respaldo de imágenes | "
                    "Detalle técnico: nit_emisor o razon_social_emisor no encontrados en texto"
                )
                resultado = verificar_pdf_con_imagenes(bytes_pdf, instrucciones)

            if not resultado.get("aprobado"):
                _registrar_rechazado_agente(correo_id, asunto, resultado)
                return  # No marcar como leído — factura rechazada por Claude

            # ── Fase 4: renombrar el PDF con proveedor y número de factura ──
            nuevo_nombre, bytes_pdf = renombrar_pdf(
                resultado.get("razon_social_emisor", ""),
                resultado.get("numero_factura", ""),
                bytes_pdf,
            )

            # ── Fase 3: buscar el proveedor y determinar destinatarios ──
            nit_limpio                      = limpiar_nit(resultado.get("nit_emisor", ""))
            proveedores                     = cargar_proveedores()
            nombre_proveedor, lista_almacen = buscar_proveedor_en_lista(nit_limpio, proveedores)

            if lista_almacen is None:
                log.warning(f"⚠️  NIT {nit_limpio} no encontrado en proveedores.json — no se envía correo")
                _registrar_proveedor_no_encontrado(nit_limpio, resultado.get("razon_social_emisor", "N/A"))
                return  # No marcar como leído — proveedor no está en el listado

            _registrar_aprobado_agente(correo_id, asunto, resultado, nit_limpio, lista_almacen)
            destinatarios = determinar_destinatarios(lista_almacen)
            enviar_correo_aprobado(token, bytes_pdf, nuevo_nombre, resultado, destinatarios)
            # Marcar como leído solo después de enviar exitosamente
            marcar_correo_como_leido(token, correo_id)
            return

        # ── Sin coincidencia: el correo no es de ninguno de los dos casos ────────
        log.info(f"⏭️  Correo ignorado (no es factura nueva ni aprobación): '{asunto}'")
        # No marcar como leído — correo ignorado no corresponde a ningún flujo conocido

    except Exception as error:
        log.error(f"💥 Error procesando el correo '{asunto}': {error}")
        # No marcar como leído — el error puede ser transitorio y conviene reintentar


# ═══════════════════════════════════════════════════════════════
# FASE 3 — ENRUTAMIENTO POR PROVEEDOR
# ═══════════════════════════════════════════════════════════════

def cargar_proveedores() -> dict:
    """
    Lee el archivo proveedores.json y retorna el diccionario completo con las tres listas de almacén.
    Si el archivo no existe o no se puede leer, registra un error crítico y retorna un diccionario vacío.
    No recibe parámetros — usa la constante ARCHIVO_PROVEEDORES.
    Retorna: diccionario con las claves almacenSabaneta, almacenRionegro y almacenRionegroSabaneta.
    """
    try:
        ruta = Path(ARCHIVO_PROVEEDORES)
        if not ruta.exists():
            log.error(f"💥 No se encontró '{ARCHIVO_PROVEEDORES}' — no se puede determinar el destinatario")
            return {}
        return json.loads(ruta.read_text(encoding="utf-8"))
    except Exception as error:
        log.error(f"💥 Error al cargar {ARCHIVO_PROVEEDORES}: {error}")
        return {}


def limpiar_nit(nit: str) -> str:
    """
    Normaliza un NIT a solo sus dígitos base, sin prefijos de texto, puntos ni dígito de verificación.

    Entradas y salidas esperadas:
      "890.300.234-3"       →  "890300234"
      "NIT. 860.028.580-2"  →  "860028580"
      "900718257-1"         →  "900718257"
      "890900314-9"         →  "890900314"
      "860028580"           →  "860028580"
      "890.900.314"         →  "890900314"
      "NIT:890300234"       →  "890300234"

    Recibe: nit (str) — NIT tal como lo retornó Claude, con cualquier formato.
    Retorna: string con solo los dígitos numéricos limpios.
    Retorna string vacío si el valor recibido es None, vacío o no contiene dígitos.
    """
    if not nit:
        return ""
    nit = str(nit).strip()

    # 2. Eliminar texto no numérico del inicio (ej: "NIT.", "Nit:", "nit ")
    primer_digito = next((i for i, c in enumerate(nit) if c.isdigit()), None)
    if primer_digito is None:
        return ""
    nit = nit[primer_digito:]

    # 3. Eliminar todos los puntos separadores de miles
    nit = nit.replace(".", "")

    # 4. Cortar desde el primer carácter no numérico en adelante
    #    Maneja guiones ASCII (-), en dash (–), em dash (—) y cualquier otro separador
    resultado = ""
    for c in nit:
        if not c.isdigit():
            break
        resultado += c

    return resultado.strip()   # 5. Eliminar espacios residuales


def buscar_proveedor_en_lista(nit_proveedor: str, proveedores: dict) -> tuple:
    """
    Busca un NIT en las tres listas del diccionario de proveedores.
    Recorre almacenSabaneta, almacenRionegro y almacenRionegroSabaneta en ese orden.
    Recibe:
      - nit_proveedor (str): NIT limpio del emisor obtenido con limpiar_nit().
      - proveedores (dict): diccionario completo cargado con cargar_proveedores().
    Retorna: tupla (nombre_proveedor, lista_almacen) si se encontró el NIT.
    Retorna: (None, None) si el NIT no existe en ninguna de las tres listas.
    Ejemplos: ("CORRUMED S.A.S", "almacenRionegroSabaneta"), ("ASHE S.A.S.", "almacenSabaneta"), (None, None).
    """
    for lista_nombre in ["almacenSabaneta", "almacenRionegro", "almacenRionegroSabaneta"]:
        for entrada in proveedores.get(lista_nombre, []):
            if entrada.get("nit_proveedor") == nit_proveedor:
                return entrada.get("nombre_proveedor"), lista_nombre
    return None, None


def determinar_destinatarios(lista_almacen: str) -> dict:
    """
    Determina los correos destino según el almacén al que pertenece el proveedor.
    Recibe: lista_almacen (str) — nombre de la lista retornado por buscar_proveedor_en_lista().
    Retorna diccionario con claves "principales" y "copia":
      - almacenSabaneta         → principales: [EMAIL_SABANETA_PRINCIPAL], copia: [EMAIL_SABANETA_COPIA]
      - almacenRionegro         → principales: [EMAIL_RIONEGRO_PRINCIPAL], copia: [EMAIL_RIONEGRO_COPIA]
      - almacenRionegroSabaneta → principales: [EMAIL_SABANETA_PRINCIPAL, EMAIL_RIONEGRO_PRINCIPAL],
                                   copia:       [EMAIL_SABANETA_COPIA, EMAIL_RIONEGRO_COPIA]
    Retorna dict vacío si el valor recibido no coincide con ninguna de las tres listas.
    """
    if lista_almacen == "almacenSabaneta":
        return {"principales": [EMAIL_SABANETA_PRINCIPAL], "copia": [EMAIL_SABANETA_COPIA] + EMAIL_COPIA_SIEMPRE}
    if lista_almacen == "almacenRionegro":
        return {"principales": [EMAIL_RIONEGRO_PRINCIPAL], "copia": [EMAIL_RIONEGRO_COPIA] + EMAIL_COPIA_SIEMPRE}
    if lista_almacen == "almacenRionegroSabaneta":
        return {
            "principales": [EMAIL_SABANETA_PRINCIPAL, EMAIL_RIONEGRO_PRINCIPAL],
            "copia":       [EMAIL_SABANETA_COPIA, EMAIL_RIONEGRO_COPIA] + EMAIL_COPIA_SIEMPRE,
        }
    return {}


# ═══════════════════════════════════════════════════════════════
# FASE 4 — RENOMBRADO DE PDF Y PRESENTACIÓN DEL CORREO
# ═══════════════════════════════════════════════════════════════

def renombrar_pdf(nombre_proveedor: str, numero_factura: str, bytes_pdf: bytes) -> tuple:
    """
    Renombra el PDF usando el formato: NOMBRE PROVEEDOR - NUMERO FACTURA.pdf

    Recibe:
      - nombre_proveedor (str): razón social del proveedor emisor.
      - numero_factura (str): número de factura extraído por Claude.
      - bytes_pdf (bytes): contenido del PDF en memoria, se retorna sin modificar.

    Retorna: tupla (nuevo_nombre, bytes_pdf) donde nuevo_nombre es el nombre del archivo.

    Ejemplos:
      ("ASHE S.A.S", "MDVA-61995", ...)   → "ASHE S.A.S - MDVA-61995.pdf"
      ("DISPAPELES S.A.S.", "NO ENCONTRADO", ...) → "DISPAPELES S.A.S..pdf"

    Si numero_factura es "NO ENCONTRADO" usa solo el nombre del proveedor.
    Elimina caracteres inválidos para nombres de archivo en Windows: \\ / : * ? " < > |
    """
    CARACTERES_INVALIDOS = r'\/:*?"<>|'

    def _limpiar(texto: str) -> str:
        for c in CARACTERES_INVALIDOS:
            texto = texto.replace(c, "")
        return texto.strip()

    proveedor_limpio = _limpiar(nombre_proveedor or "PROVEEDOR")
    factura_limpia   = _limpiar(numero_factura or "")

    if factura_limpia and factura_limpia.upper() != "NO ENCONTRADO":
        nuevo_nombre = f"{proveedor_limpio} - {factura_limpia}.pdf"
    else:
        nuevo_nombre = f"{proveedor_limpio}.pdf"

    return nuevo_nombre, bytes_pdf


def procesar_correos() -> None:
    """
    Función principal del ciclo de revisión automática. Se ejecuta cada INTERVALO_MINUTOS.
    Obtiene un token fresco y busca DOS tipos de correos no leídos:
      - Con adjuntos (candidatos a Caso 1: facturas nuevas en ZIP).
      - Sin adjuntos (candidatos a Caso 2: respuestas de aprobación humana).
    Llama a procesar_un_correo() para cada correo encontrado — esa función decide el caso.
    No recibe parámetros — usa las constantes de configuración globales.
    No retorna nada — todos los resultados se registran en el log.
    Esta función captura todas las excepciones para que el scheduler no se detenga nunca.
    """
    try:
        print("\nSe inicia proceso de validación de correos.")
        log.info("🔍 Revisando correos nuevos...")

        instrucciones = cargar_instrucciones_agente()
        token_acceso  = obtener_token_microsoft()
        print(f"[DIAGNÓSTICO] Token obtenido correctamente")

        # Correos con adjuntos — candidatos a Caso 1 (facturas nuevas en ZIP)
        correos_con_adjunto = obtener_correos_nuevos(token_acceso)
        print(f"[DIAGNÓSTICO] Correos nuevos con ZIP encontrados: {len(correos_con_adjunto)}")
        for c in correos_con_adjunto:
            print(f"[DIAGNÓSTICO] - Asunto: {c.get('subject', 'sin asunto')} | Adjuntos: {c.get('hasAttachments', False)}")

        # Correos candidatos a Caso 2 (respuestas de aprobación, con o sin adjunto)
        correos_aprobacion  = obtener_correos_aprobacion(token_acceso)
        print(f"[DIAGNÓSTICO] Correos de aprobación encontrados: {len(correos_aprobacion)}")
        for c in correos_aprobacion:
            print(f"[DIAGNÓSTICO] - Asunto: {c.get('subject', 'sin asunto')}")

        # Deduplicar: si un correo ya está en correos_con_adjunto no se procesa dos veces
        # (ocurre cuando el aprobador responde y su cliente agrega imágenes de firma)
        ids_ya_vistos     = {c["id"] for c in correos_con_adjunto}
        aprobacion_nuevos = [c for c in correos_aprobacion if c["id"] not in ids_ya_vistos]

        todos_los_correos = correos_con_adjunto + aprobacion_nuevos
        cantidad          = len(todos_los_correos)

        if cantidad == 0:
            log.info("📭 No hay correos nuevos")
        else:
            log.info(f"📧 {cantidad} correo(s) nuevo(s) encontrado(s)")
            for correo in todos_los_correos:
                procesar_un_correo(token_acceso, correo, instrucciones)

        log.info(f"⏰ Próxima revisión en {INTERVALO_MINUTOS} minutos")
        print(f"Se finalizó la revisión de correos, se hará nuevamente en {INTERVALO_MINUTOS} minutos.")

    except Exception as error:
        log.error(f"💥 Error en el ciclo de revisión: {error}")


def _verificar_configuracion() -> None:
    """
    Verifica que todas las variables de entorno críticas estén configuradas en el archivo .env.
    Muestra un mensaje claro por cada variable faltante para facilitar la depuración.
    No recibe parámetros ni retorna nada.
    Lanza: SystemExit con código 1 si falta alguna variable crítica (el agente no puede operar).
    """
    variables_requeridas = {
        "ANTHROPIC_API_KEY":        ANTHROPIC_API_KEY,
        "AZURE_CLIENT_ID":          AZURE_CLIENT_ID,
        "AZURE_CLIENT_SECRET":      AZURE_CLIENT_SECRET,
        "AZURE_TENANT_ID":          AZURE_TENANT_ID,
        "EMAIL_MONITOREAR":         EMAIL_MONITOREAR,
        "EMAIL_SABANETA_PRINCIPAL": EMAIL_SABANETA_PRINCIPAL,
        "EMAIL_RIONEGRO_PRINCIPAL": EMAIL_RIONEGRO_PRINCIPAL,
    }

    variables_faltantes = [
        nombre for nombre, valor in variables_requeridas.items() if not valor
    ]

    if variables_faltantes:
        log.error("💥 Faltan las siguientes variables en el archivo .env:")
        for variable in variables_faltantes:
            log.error(f"   ❌ {variable}")
        log.error("📋 Copia .env.example a .env y completa todos los valores.")
        raise SystemExit(1)

    log.info("✅ Configuración verificada correctamente")


def main() -> None:
    """
    Punto de entrada del programa. Se ejecuta al correr: python agente.py
    Muestra el banner de bienvenida en pantalla, verifica la configuración del .env,
    ejecuta la primera revisión de correos de inmediato y luego programa revisiones
    automáticas cada INTERVALO_MINUTOS usando el planificador de tareas 'schedule'.
    El bucle infinito se interrumpe limpiamente con Ctrl+C.
    No recibe parámetros ni retorna nada.
    """
    # ── Banner de bienvenida ──
    print("╔══════════════════════════════════════╗")
    print("║   AGENTE MARQUILLAS S.A.S            ║")
    print("║   Verificador automático de facturas ║")
    print("╚══════════════════════════════════════╝")
    print(f"\nIniciando... revisaré el correo cada {INTERVALO_MINUTOS} minutos.")
    print("Presiona Ctrl+C para detener el agente.\n")

    # Verificar que el .env esté completo antes de empezar
    _verificar_configuracion()

    # Primera revisión inmediata al arrancar (sin esperar el intervalo)
    procesar_correos()

    # Programar las revisiones automáticas cada INTERVALO_MINUTOS minutos
    schedule.every(INTERVALO_MINUTOS).minutes.do(procesar_correos)

    # Bucle principal — se ejecuta indefinidamente hasta que el usuario presione Ctrl+C
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("🛑 Agente detenido por el usuario (Ctrl+C). ¡Hasta luego!")


# Punto de entrada estándar de Python
if __name__ == "__main__":
    main()
