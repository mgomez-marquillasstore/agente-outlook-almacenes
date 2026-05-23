#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PRUEBA LOCAL — Agente Marquillas S.A.S
=======================================
Permite probar la verificación de facturas SIN necesitar credenciales de Microsoft.
Solo requiere ANTHROPIC_API_KEY configurada en el archivo .env.

Reutiliza exactamente las mismas funciones de agente.py para garantizar
que lo que funcione aquí funcionará igual en producción.

Uso: python prueba_local.py
"""

import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv

# Importar funciones reutilizables del agente principal
# (se importan individualmente para no ejecutar el ciclo completo)
from agente import (
    cargar_instrucciones_agente,
    extraer_pdf_del_zip,
    leer_texto_del_pdf,
    verificar_documento_con_claude,
)

# Cargar variables de entorno desde .env
load_dotenv()


def mostrar_bienvenida() -> None:
    """
    Muestra en pantalla el encabezado y las instrucciones de uso del script de prueba.
    Explica qué se necesita para ejecutarlo correctamente.
    No recibe parámetros ni retorna nada.
    """
    print("=" * 62)
    print("  PRUEBA LOCAL — Verificador de facturas Marquillas S.A.S")
    print("=" * 62)
    print()
    print("Este script prueba la verificación sin conexión a Outlook.")
    print("Solo necesitas:")
    print("  1. Un archivo .zip que contenga un PDF de factura")
    print("  2. ANTHROPIC_API_KEY configurada en el archivo .env")
    print()


def verificar_api_key() -> None:
    """
    Comprueba que la clave de la API de Anthropic esté configurada en el archivo .env.
    Muestra un mensaje de error claro y termina el programa si no está presente.
    No recibe parámetros ni retorna nada.
    Lanza: SystemExit con código 1 si falta la clave de API.
    """
    clave_api = os.getenv("ANTHROPIC_API_KEY", "")

    if not clave_api:
        print("❌ ERROR: No encontré ANTHROPIC_API_KEY en el archivo .env")
        print()
        print("   Solución:")
        print("   1. Abre el archivo .env en la carpeta del proyecto")
        print("   2. Agrega tu clave de Anthropic así:")
        print("      ANTHROPIC_API_KEY=sk-ant-api03-tu-clave-aqui")
        print("   3. Obtén tu clave en: https://console.anthropic.com")
        sys.exit(1)

    print("✅ ANTHROPIC_API_KEY encontrada")


def solicitar_ruta_zip() -> Path:
    """
    Pide al usuario que ingrese la ruta del archivo ZIP a analizar.
    Limpia automáticamente las comillas que Windows agrega al arrastrar archivos.
    Verifica que el archivo exista antes de continuar.
    No recibe parámetros.
    Retorna: objeto Path con la ruta validada del archivo ZIP.
    Lanza: SystemExit con código 1 si el archivo no existe.
    """
    print()
    ruta_ingresada = input(
        "📁 Ruta del archivo ZIP (puedes arrastrarlo aquí):\n   > "
    ).strip()

    # Quitar comillas que Windows agrega al arrastrar archivos
    ruta_ingresada = ruta_ingresada.strip('"').strip("'")

    archivo_zip = Path(ruta_ingresada)

    if not archivo_zip.exists():
        print()
        print(f"❌ ERROR: No encontré el archivo en la ruta especificada:")
        print(f"   {ruta_ingresada}")
        print()
        print("   Verifica que la ruta sea correcta y que el archivo exista.")
        sys.exit(1)

    tamano_kb = archivo_zip.stat().st_size // 1024
    print(f"\n✅ Archivo encontrado: {archivo_zip.name} ({tamano_kb:,} KB)")
    return archivo_zip


def leer_bytes_zip(ruta_zip: Path) -> bytes:
    """
    Lee el contenido binario del archivo ZIP desde el disco.
    Recibe: ruta_zip (Path) — ruta validada al archivo ZIP.
    Retorna: bytes con el contenido completo del archivo ZIP.
    Lanza: SystemExit si hay error de permisos o el archivo está bloqueado.
    """
    try:
        with open(ruta_zip, "rb") as archivo:
            bytes_zip = archivo.read()
        print(f"📦 ZIP leído en memoria: {len(bytes_zip):,} bytes")
        return bytes_zip
    except PermissionError:
        print(f"\n❌ ERROR: Sin permisos para leer el archivo: {ruta_zip}")
        print("   Cierra el archivo si está abierto en otra aplicación.")
        sys.exit(1)
    except Exception as error:
        print(f"\n❌ ERROR inesperado al leer el archivo: {error}")
        sys.exit(1)


def mostrar_vista_previa_texto(texto_pdf: str) -> bool:
    """
    Muestra los primeros 500 caracteres del texto extraído del PDF para que el usuario
    confirme que la extracción fue correcta antes de enviar a Claude.
    Recibe: texto_pdf (str) — texto extraído completo del PDF.
    Retorna: True si el usuario quiere continuar, False si decide cancelar.
    """
    print()
    if not texto_pdf.strip():
        print("⚠️  ADVERTENCIA: El PDF no tiene texto seleccionable.")
        print("   Puede ser una imagen escaneada. Claude intentará analizarlo")
        print("   pero los resultados pueden no ser precisos.")
    else:
        print(f"📝 Texto extraído: {len(texto_pdf)} caracteres")
        print()
        print("─" * 62)
        print("VISTA PREVIA — primeros 500 caracteres del texto del PDF:")
        print("─" * 62)
        print(texto_pdf[:500])
        print("─" * 62)

    print()
    respuesta = input("¿Continuar con la verificación de Claude? (s/n): ").strip().lower()
    return respuesta == "s"


def mostrar_resultado_final(resultado: dict) -> None:
    """
    Muestra el resultado de la verificación de Claude de forma clara y legible en pantalla.
    Recibe: resultado (dict) — diccionario con los campos nit_encontrado,
            razon_social_encontrada, aprobado y motivo.
    No retorna nada.
    """
    aprobado = resultado.get("aprobado", False)
    estado   = "✅ APROBADA" if aprobado else "❌ RECHAZADA"

    print()
    print("=" * 62)
    print(f"  RESULTADO DE LA VERIFICACIÓN: {estado}")
    print("=" * 62)
    print()
    print(f"  NIT encontrado:          {resultado.get('nit_encontrado', 'N/A')}")
    print(f"  Razón social encontrada: {resultado.get('razon_social_encontrada', 'N/A')}")
    print(f"  Aprobado:                {'SÍ' if aprobado else 'NO'}")
    print(f"  Motivo:                  {resultado.get('motivo', 'Sin motivo')}")
    print()
    print("  Respuesta completa de Claude (JSON):")
    print(json.dumps(resultado, ensure_ascii=False, indent=4))
    print("=" * 62)


def main_prueba() -> None:
    """
    Función principal del script de prueba local. Guía al usuario paso a paso:
      1. Muestra bienvenida y verifica la API key de Anthropic.
      2. Pide la ruta del ZIP y lo lee desde disco.
      3. Extrae el PDF del ZIP usando la misma función de agente.py.
      4. Lee el texto del PDF usando la misma función de agente.py.
      5. Muestra vista previa y pide confirmación al usuario.
      6. Consulta a Claude con las instrucciones de agente.md.
      7. Muestra el resultado completo en pantalla.
    No recibe parámetros ni retorna nada.
    """
    mostrar_bienvenida()
    verificar_api_key()

    # Paso 1: Obtener el archivo ZIP del usuario
    ruta_zip  = solicitar_ruta_zip()
    bytes_zip = leer_bytes_zip(ruta_zip)

    # Paso 2: Extraer el PDF del ZIP (misma función que usa agente.py)
    print("\n📄 Extrayendo PDF del ZIP...")
    nombre_pdf, bytes_pdf = extraer_pdf_del_zip(bytes_zip)

    if not bytes_pdf:
        print("\n❌ ERROR: El ZIP no contiene ningún archivo PDF.")
        print("   Verifica que el ZIP incluya al menos un archivo con extensión .pdf")
        sys.exit(1)

    print(f"✅ PDF encontrado: {nombre_pdf} ({len(bytes_pdf):,} bytes)")

    # Paso 3: Leer el texto del PDF (misma función que usa agente.py)
    print("\n📝 Extrayendo texto del PDF...")
    texto_pdf = leer_texto_del_pdf(bytes_pdf)

    # Paso 4: Mostrar vista previa y pedir confirmación
    continuar = mostrar_vista_previa_texto(texto_pdf)
    if not continuar:
        print("\n🛑 Verificación cancelada por el usuario.")
        sys.exit(0)

    # Paso 5: Cargar instrucciones y consultar a Claude (mismas funciones que agente.py)
    print("\n📋 Cargando instrucciones del agente...")
    instrucciones = cargar_instrucciones_agente()

    print("🤖 Consultando a Claude AI...")
    resultado = verificar_documento_con_claude(texto_pdf, instrucciones)

    # Paso 6: Mostrar resultado
    mostrar_resultado_final(resultado)


# Punto de entrada estándar de Python
if __name__ == "__main__":
    main_prueba()
