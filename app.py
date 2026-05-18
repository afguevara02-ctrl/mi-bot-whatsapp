from flask import Flask, request, jsonify, redirect, render_template_string, url_for, send_file
from pathlib import Path
from threading import Timer
from datetime import datetime
from werkzeug.utils import secure_filename
from io import BytesIO
from openpyxl import Workbook
import json
import re
import requests
import time
import os
import unicodedata

app = Flask(__name__)

# ==========================================
# ⚙️ CONFIGURACION
# ==========================================
TOKEN_META_DEFAULT = "EAALkHb9ZBdFwBRBq5ZAZBZA9zYpYy5vbd4Esk7AzfbqLOOehck21nSZADaXC80aBEtM39MlXGsJnTHwJZBkdJUPGOjjm6UzqIZCKLpe63d0wYbXRlfCTy5MlazSfc2Smf9vBJw6zTEberthmSv1IiZBp4ZCMOGBchWqXtEtaKN2SzBs3IPSxwYwC51mJQT3l28QZDZD"
ID_TELEFONO_DEFAULT = "1052565307946835"
TOKEN_VERIFICACION = "Kiratsu.52310299*"
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "bot_config.json"
RUNTIME_STATE_PATH = BASE_DIR / "runtime_state.json"
VENTAS_PATH = BASE_DIR / "ventas.json"
CLIENTES_VENDIDOS_PATH = BASE_DIR / "clientes_vendidos.json"
CATALOGO_PATH = BASE_DIR / "catalogo.json"
MEDIA_DIR = BASE_DIR / "media"
MEDIA_DIR.mkdir(exist_ok=True)
COMPROBANTES_DIR = BASE_DIR / "comprobantes"
COMPROBANTES_DIR.mkdir(exist_ok=True)
ALLOWED_EXTENSIONS = {'mp4', 'mov', 'm4v', 'webm', 'pdf', 'jpg', 'jpeg', 'png', 'gif'}
DESCUENTO_RECORDATORIO_SEGUNDOS = 1800
VIDEO_SEGUIMIENTO_SEGUNDOS = 240
ULTIMA_OPORTUNIDAD_SEGUNDOS = 43200
PDF_SEGUIMIENTO_SEGUNDOS = 120

DEFAULT_FLOWS = {
    "video_followup_seconds": 240,
    "pdf_followup_seconds": 120,
    "discount_reminder_seconds": 1800,
    "last_chance_seconds": 43200,
    "send_pdf_after_video": True,
    "send_buy_after_pdf": True,
    "usar_lista_pago": True
}

DEFAULT_DATOS_BOT = {
    "token_meta": TOKEN_META_DEFAULT,
    "id_telefono": ID_TELEFONO_DEFAULT,
    "nombre_vendedor": "Tu nombre",
    "numero_admin": "573000000000",
    "nequi": "3000000000",
    "daviplata": "3000000000",
    "llave": "Llave 0000000000",
    "link_video": "",
    "link_pdf_demo": "",
    "link_recursos_final": "",
    "link_canal_whatsapp": "",
    "descripcion_producto": "Pack educativo digital con acceso inmediato.",
    "mensaje_bienvenida": "Hola, soy tu asistente. Escribe menu para ver el catalogo.",
    "mensaje_despues_demo": "Si te gusta el contenido, activa tu acceso hoy.",
    "mensaje_descuento": "Hoy tenemos una promocion especial por tiempo limitado.",
    "mensaje_descuento_ultima_oportunidad": "Ultima oportunidad para activar tu acceso con promocion.",
    "mensaje_respuesta_gracias": "Gracias por confirmar. Si necesitas algo mas, aqui estoy.",
    "mensaje_cuentas_cobro": "Medios de pago disponibles: Nequi, Daviplata y Llave.",
    "mensaje_confirmacion_pago": "Envia tu comprobante por este chat para validar el acceso.",
    "mensaje_entrega_final": "Acceso enviado. Disfruta el material.",
    "visitas_info": 0,
    "flujos_automaticos": DEFAULT_FLOWS.copy(),
    "catalogo_plantillas_meta_custom": {},
    "ocultar_tutorial_panel": False,
    "auto_iniciar_tutorial": True,
    "tutorial_habilitado": True,  # False = no arrancar tutorial al cargar la pagina
    "test_template_nombre": "hello_world",
    "test_template_idioma": "en_US",
    "test_template_var1": "",
    "test_template_var2": "",
}

# ==============================================================
# CATALOGO: UNICA FUENTE DE VERDAD
# Cada producto debe incluir todos los campos.
# Los links vacios son placeholders; reemplazarlos antes de ir
# a produccion. NO rompen el codigo si estan vacios.
# Payloads de botones Meta generados automaticamente como:
#   video_[id], pdf_[id], comprar_[id], descuento_[id],
#   pagar_nequi_[id], pagar_daviplata_[id]
# ==============================================================
DEFAULT_CATALOGO = {
    "finanzas": {
        "id": "finanzas",
        "opcion": "1",
        "titulo": "Pack Completo de Cartillas de Finanzas",
        # Nombre EXACTO de la plantilla aprobada en Meta Business Manager
        "plantilla_info_meta": "info_finanzas_v1",
        # texto_info_meta: cuerpo EXACTO que se copia a Meta Business Manager para la plantilla unica
        # del producto. Este texto aparece en el textarea 'Texto plantilla unica de informacion para Meta'.
        "texto_info_meta": """📚 *TRANSFORMA LA EDUCACIÓN FINANCIERA DE TUS ESTUDIANTES E HIJOS DESDE HOY* ✨

Imagina *enseñarles a manejar el dinero* de forma fácil y práctica, y además *tener el sistema ideal organizado por edades y con acceso de por vida!* 💰👧👦

👉 *¿Qué hace especial a este pack?*
Con nuestro *Pack Completo de Cartillas de Finanzas desde primero de primaria hasta undécimo,* podrás guiarlos paso a paso y según su edad, para que desarrollen hábitos que les servirán para siempre. ✨ 100% *descargable e imprimible.*""",
        "precio_normal": "15000",
        "precio_descuento": "10000",
        # link_video: URL publica del video demo (YouTube, Drive, CDN)
        "link_video": "",  # TODO: reemplazar con URL real
        # link_pdf: URL publica del PDF demo
        "link_pdf": "",    # TODO: reemplazar con URL real
        # link_drive_final: carpeta con todo el material; se entrega tras aprobar pago
        "link_drive_final": "https://drive.google.com/drive/folders/1hX5B6Q0Z2NM7D1MadDVui7LMvNis0Vd5",
        # link_canal_whatsapp: canal privado de actualizaciones
        "link_canal_whatsapp": "https://whatsapp.com/channel/0029VbBQlcRLSmbiqKysJp3q",
        # palabras_clave / aliases: palabras que el cliente puede escribir para llegar a este producto
        "palabras_clave": ["finanzas", "dinero", "ahorro", "cartillas", "economia"],
        "mensajes": {}
    },
    "ingles": {
        "id": "ingles",
        "opcion": "2",
        "titulo": "Planeaciones Unificadas de Ingles",
        "plantilla_info_meta": "info_ingles_v1",
        "texto_info_meta": """🇬🇧 *¡AHORRA CIENTOS DE HORAS EN TUS PLANEACIONES DE INGLÉS!* ✨

Sabemos lo difícil que es encontrar material organizado, actualizado y que realmente capte la atención de los estudiantes. ¡Por eso creamos esta solución para ti! 👩‍🏫👨‍🏫

👉 *¿Qué incluye este súper material?*
Un *Pack Completo de Inglés* diseñado para facilitar tu labor docente. Actividades, guías y planeaciones listas para aplicar, adaptadas a los estándares educativos. ✨ Todo el material es *100% descargable, imprimible y con acceso de por vida.*""",
        "precio_normal": "20000",
        "precio_descuento": "15000",
        "link_video": "",  # TODO: reemplazar con URL real
        "link_pdf": "",    # TODO: reemplazar con URL real
        "link_drive_final": "",  # TODO: reemplazar con URL real del Drive de ingles
        "link_canal_whatsapp": "",  # TODO: reemplazar con link del canal de ingles
        "palabras_clave": ["ingles", "planeaciones", "dba", "ebc", "english"],
        "mensajes": {}
    },
    "biologia": {
        "id": "biologia",
        "opcion": "3",
        "titulo": "Guias de Biologia",
        "plantilla_info_meta": "info_biologia_v1",
        "texto_info_meta": """🧬 *¡HAZ QUE TUS CLASES DE CIENCIAS Y BIOLOGÍA SEAN INOLVIDABLES!* 🌿

Olvídate de la teoría aburrida y del tiempo perdido buscando recursos en internet. Despierta la curiosidad de tus estudiantes con material visual y práctico. 🔬🌎

👉 *¿Qué te vas a llevar hoy?*
El *Pack Completo de Biología* con guías didácticas, esquemas y actividades que facilitan el aprendizaje de forma interactiva. Todo organizado por temas para que solo tengas que elegir y enseñar. ✨ 100% descargable, imprimible y tuyo para siempre.""",
        "precio_normal": "12000",
        "precio_descuento": "10000",
        "link_video": "",  # TODO: reemplazar con URL real
        "link_pdf": "",    # TODO: reemplazar con URL real
        "link_drive_final": "",  # TODO: reemplazar con URL real del Drive de biologia
        "link_canal_whatsapp": "",  # TODO: reemplazar con link del canal de biologia
        "palabras_clave": ["biologia", "laboratorio", "microscopio", "celula", "ciencias"],
        "mensajes": {}
    }
}

# Texto base para nuevos productos (plantilla genérica editable)
TEXTO_INFO_META_BASE_NUEVO_PRODUCTO = """🚀 *¡DESCUBRE EL MEJOR MATERIAL PARA TUS CLASES DE [NOMBRE DEL PRODUCTO]!* ✨

Diseñado específicamente para ahorrarte tiempo y potenciar el aprendizaje de tus estudiantes de forma práctica y divertida. 💡📚

👉 *¿Qué hace especial a este pack?*
Con nuestro *Pack Completo de [NOMBRE DEL PRODUCTO],* tendrás acceso inmediato a material listo para usar. ✨ *100% descargable, imprimible y con acceso de por vida.*"""

DEFAULT_MENSAJES_PRODUCTO = {
    "bienvenida": "Hola, te comparto informacion de {titulo}.",
    "bienvenida_a": "Hola, te comparto informacion de {titulo}.",
    "bienvenida_b": "Bienvenido, aqui tienes la informacion clave de {titulo}.",
    "info": "Producto: {titulo}. Precio: ${precio}.",
    "despues_demo": "Si te gusto la demo, activa tu acceso con un clic.",
    "descuento": "Hoy puedes activar {titulo} con precio promocional ${precio_descuento}.",
    "confirmacion_pago": "Perfecto, envia comprobante para validar tu acceso.",
    "entrega_final": "Acceso enviado para {titulo}. Disfrutalo."
}

ETAPAS_CLIENTE = [
    "nuevo",
    "viendo_info",
    "viendo_video",
    "viendo_pdf",
    "interesado",
    "medio_pago",
    "esperando_comprobante",
    "en_revision",
    "vendido"
]

BOTONES_RESPUESTA = {
    "ver_video": "Ver video",
    "ver_pdf": "Ver PDF",
    "comprar": "Comprar",
    "nequi": "Nequi",
    "daviplata": "Daviplata",
    "llave": "Llave",
    "descuento": "Descuento",
    "ya_pague": "Ya pague",
    "menu_principal": "Menu"
}

CATALOGO_PLANTILLAS_META = {
    # ---------------------------------------------------------------
    # plantilla_menu_general
    # {{1}} = lista dinamica de productos
    # Sin botones quick-reply obligatorios.
    # Disparo: textos hola / menu / menu_principal o numero de opcion
    # ---------------------------------------------------------------
    "plantilla_menu_general": {
        "nombre_meta": "plantilla_menu_general",
        "idioma": "es",
        "texto_referencia": "Hola 👋\n\nGracias por escribir. Estos son los packs digitales disponibles:\n\n{{1}}\n\nResponde con el número o nombre del pack que quieres conocer.",
        "variables_ordenadas": ["lista_productos"]
    },
    # ---------------------------------------------------------------
    # plantilla_compra_universal
    # Se envia cuando el cliente pulsa 'Comprar' desde cualquier producto.
    # {{1}} = titulo del producto  /  {{2}} = precio_normal
    # Botones DINÁMICOS (fuera de plantilla): pagar_nequi_[id], pagar_daviplata_[id], menu_principal
    # IMPORTANTE: La plantilla de compra NO ofrece descuento.
    # El descuento solo se activa por abandono o solicitud explícita (payload descuento_[id]).
    # ---------------------------------------------------------------
    "plantilla_compra_universal": {
        "nombre_meta": "plantilla_compra_universal",
        "idioma": "es",
        "texto_referencia": "Sin mensualidades.\n\n📲 *¿CÓMO FUNCIONA?*\n\n1️⃣ Realizas el pago por el método que prefieras\n2️⃣ Envías el comprobante por aquí\n3️⃣ Recibes TODO el material al instante por WhatsApp\n\n💳 *MEDIOS DE PAGO DISPONIBLES:*\n\n☑ Nequi\n☑ Daviplata\n☑ Llave\n\n🔥 Este material te ahorra horas de trabajo y te da recursos listos para usar.\n\nProducto: {{1}}\nPrecio normal: ${{2}}\n\nElige tu medio de pago.",
        "variables_ordenadas": ["titulo", "precio_normal"]
    },
    # ---------------------------------------------------------------
    # plantilla_seguimiento_pdf
    # Se envia INMEDIATAMENTE despues de enviar el PDF demo.
    # {{1}} = titulo del producto
    # Botones dinámicos: comprar_[id], video_[id], menu_principal. Sin descuento.
    # ---------------------------------------------------------------
    "plantilla_seguimiento_pdf": {
        "nombre_meta": "plantilla_seguimiento_pdf",
        "idioma": "es",
        "texto_referencia": "Ya tienes el PDF demo de {{1}} ✅\n\nRevísalo con calma. Si quieres acceder al material completo, puedes continuar con la compra.",
        "variables_ordenadas": ["titulo"]
    },
    # ---------------------------------------------------------------
    # plantilla_descuento_universal
    # SOLO se usa por abandono o solicitud explícita. NO en comprar/pdf/video.
    # {{1}} = titulo  /  {{2}} = precio_normal  /  {{3}} = precio_descuento
    # Botones dinámicos: pagar_nequi_[id], pagar_daviplata_[id], menu_principal
    # ---------------------------------------------------------------
    "plantilla_descuento_universal": {
        "nombre_meta": "plantilla_descuento_universal",
        "idioma": "es",
        "texto_referencia": "Te entiendo completamente 🙌, y es genial que estés interesado en {{1}}.\n\n✨ *¡Hoy tenemos una promoción especial!* Si ya viste que el pack completo está a ${{2}}, puedes aprovechar el *descuento a solo ${{3}}* por tiempo limitado.\n\nEs una gran oportunidad para acceder al material y empezar a usarlo de inmediato.\n\nSi decides llevarlo por ${{3}}, solo necesitas elegir tu método de pago y te compartiré los datos para activar tu acceso.",
        "variables_ordenadas": ["titulo", "precio_normal", "precio_descuento"]
    },
    # ---------------------------------------------------------------
    # plantilla_cuentas_cobro
    # Se envia al elegir metodo de pago (pagar_nequi_[id] etc.)
    # {{1}} = titulo  /  {{2}} = metodo  /  {{3}} = valor  /  {{4}} = datos_pago
    # Solicita comprobante al final.
    # ---------------------------------------------------------------
    "plantilla_cuentas_cobro": {
        "nombre_meta": "plantilla_cuentas_cobro",
        "idioma": "es",
        "texto_referencia": "Perfecto ✅\n\nPara adquirir {{1}}, realiza el pago por {{2}}.\n\nValor a pagar: ${{3}}\n\nDatos de pago:\n{{4}}\n\nCuando termines, envía aquí la imagen o PDF del comprobante para validar tu compra.",
        "variables_ordenadas": ["titulo", "metodo", "valor", "datos_pago"]
    },
    # ---------------------------------------------------------------
    # plantilla_entrega_final
    # Se envia cuando el admin aprueba el comprobante.
    # {{1}} = link_drive_final  /  {{2}} = link_canal_whatsapp
    # ---------------------------------------------------------------
    "plantilla_entrega_final": {
        "nombre_meta": "plantilla_entrega_final",
        "idioma": "es",
        "texto_referencia": "🔑 *¡Excelente!* 😃 Tu pago fue aprobado.\n\nAquí está el enlace con acceso a tu material:\n{{1}}\n\nSi tu compra incluye canal privado de WhatsApp, entra aquí:\n{{2}}\n\n*Gracias por tu compra. ¡Disfrútalo al máximo!* 🚀",
        "variables_ordenadas": ["link_drive_final", "link_canal_whatsapp"]
    }
}

REFERENCIA_META_DESDE_MENSAJE_BOT = {
    "mensaje_bienvenida": "plantilla_menu_general",
    "mensaje_descuento": "plantilla_descuento_universal",
    "mensaje_cuentas_cobro": "plantilla_cuentas_cobro",
    "mensaje_entrega_final": "plantilla_entrega_final"
}


def obtener_catalogo_plantillas_meta():
    base = json.loads(json.dumps(CATALOGO_PLANTILLAS_META))
    custom = datos_bot.get("catalogo_plantillas_meta_custom", {}) if isinstance(datos_bot, dict) else {}
    if isinstance(custom, dict):
        for clave, plantilla in custom.items():
            if isinstance(plantilla, dict):
                base[clave] = {
                    "nombre_meta": str(plantilla.get("nombre_meta", "") or clave).strip(),
                    "idioma": str(plantilla.get("idioma", "es") or "es").strip(),
                    "texto_referencia": str(plantilla.get("texto_referencia", "") or "").strip(),
                    "variables_ordenadas": list(plantilla.get("variables_ordenadas", []) or [])
                }

    fuente_catalogo = catalogo_productos if isinstance(globals().get("catalogo_productos", {}), dict) and catalogo_productos else DEFAULT_CATALOGO
    for clave, producto in (fuente_catalogo or {}).items():
        producto_id = str(producto.get("id", clave) or clave).strip()
        clave_info = f"info_{clave}"
        # Usar texto_info_meta real del producto si existe; si no, usar plantilla generica
        texto_info = str(producto.get("texto_info_meta", "") or "").strip()
        if not texto_info:
            titulo_p = producto.get("titulo", clave)
            precio_p = producto.get("precio_normal", "")
            precio_desc_p = producto.get("precio_descuento", "")
            texto_info = f"Hola, te comparto informacion de {titulo_p}."
            if precio_p:
                texto_info += f"\n\nPrecio: ${precio_p}"
            if precio_desc_p and precio_desc_p != precio_p:
                texto_info += f"\nDescuento especial: ${precio_desc_p}"
        base[clave_info] = {
            "nombre_meta": str(producto.get("plantilla_info_meta", "") or f"info_{producto_id}_v1").strip(),
            "idioma": "es",
            "texto_referencia": texto_info,
            "variables_ordenadas": []
        }
    return base


def construir_vista_previa_plantillas_meta(catalogo):
    vista = {}
    for clave, plantilla in (catalogo or {}).items():
        variables = plantilla.get("variables_ordenadas", [])
        texto_referencia = str(plantilla.get("texto_referencia", "") or "").strip()
        texto_meta = texto_referencia
        ejemplo = texto_referencia
        for idx, variable in enumerate(variables, start=1):
            texto_meta = f"{texto_meta}\n{{{{{idx}}}}}" if texto_meta else f"{{{{{idx}}}}}"
            ejemplo = f"{ejemplo}\n[{variable}]" if ejemplo else f"[{variable}]"
        vista[clave] = {
            "texto_meta": texto_meta,
            "ejemplo": ejemplo,
            "variables": variables,
        }
    return vista


def construir_manual_meta_rows():
    filas = [
        {
            "fase": "Menu inicial",
            "plantilla": "plantilla_menu_general",
            "texto_referencia": "Menu principal de productos",
            "variables": "{{1}} = lista_productos",
            "botones": "Sin botones obligatorios",
            "payloads": "Entrada por texto: hola, menu, menu_principal"
        },
        {
            "fase": "Descuento universal",
            "plantilla": "plantilla_descuento_universal",
            "texto_referencia": "Oferta de descuento por producto",
            "variables": "{{1}} = precio_normal | {{2}} = precio_descuento",
            "botones": "Pagar Nequi | Pagar Daviplata",
            "payloads": "pagar_nequi_[idproducto] | pagar_daviplata_[idproducto]"
        },
        {
            "fase": "Cuentas de cobro",
            "plantilla": "plantilla_cuentas_cobro",
            "texto_referencia": "Datos bancarios y solicitud de comprobante",
            "variables": "{{1}} = nequi | {{2}} = daviplata | {{3}} = llave | {{4}} = instruccion_pago",
            "botones": "Sin botones obligatorios",
            "payloads": "Flujo posterior: ya_pague"
        },
        {
            "fase": "Entrega final",
            "plantilla": "plantilla_entrega_final",
            "texto_referencia": "Entrega de acceso final",
            "variables": "{{1}} = link_drive_final | {{2}} = link_canal_whatsapp",
            "botones": "Sin botones obligatorios",
            "payloads": "Se dispara tras aprobacion admin"
        },
    ]

    for clave, producto in (catalogo_productos or {}).items():
        producto_id = producto.get("id", clave)
        plantilla_real = producto.get("plantilla_info_meta") or f"info_{producto_id}_v1"
        # Usar texto_info_meta real como texto_referencia
        texto_real = str(producto.get("texto_info_meta", "") or "").strip()
        if not texto_real:
            texto_real = f"Informacion del producto {producto.get('titulo', clave)}. (Edita el campo texto_info_meta en el catálogo para usar el texto persuasivo real.)"
        filas.append(
            {
                "fase": f"Info producto {producto.get('titulo', clave)}",
                "plantilla": plantilla_real,
                "texto_referencia": texto_real,
                "variables": "Sin variables {{N}} en el body (texto fijo por producto). Idioma: es. Categoría: MARKETING.",
                "botones": f"Ver video demo | Ver PDF demo | Comprar",
                "payloads": f"video_{producto_id} | pdf_{producto_id} | comprar_{producto_id} | (descuento_{producto_id}: solo solicitud explícita)"
            }
        )
    return filas


def exportar_catalogo_plantillas_meta_json():
    payload = {
        "templates": obtener_catalogo_plantillas_meta(),
        "manual_meta_rows": construir_manual_meta_rows(),
        "catalogo_productos": catalogo_productos,
    }
    buffer = BytesIO()
    buffer.write(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"))
    buffer.seek(0)
    return buffer


PUBLIC_TUNNEL_URL = None


# ==========================================
# 🧰 ESTABILIDAD / INSTALACION
# ==========================================
REQUIRED_PACKAGES = {
    "flask": "flask",
    "requests": "requests",
    "openpyxl": "openpyxl",
    "werkzeug": "werkzeug",
}
OPTIONAL_PACKAGES = {
    "reportlab": "reportlab"
}

def asegurar_json_archivo(path_obj, contenido_default):
    try:
        if not path_obj.exists():
            path_obj.write_text(json.dumps(contenido_default, ensure_ascii=False, indent=2), encoding="utf-8")
            return
        contenido = path_obj.read_text(encoding="utf-8").strip()
        if not contenido:
            path_obj.write_text(json.dumps(contenido_default, ensure_ascii=False, indent=2), encoding="utf-8")
            return
        cargado = json.loads(contenido)
        if isinstance(contenido_default, dict) and not isinstance(cargado, dict):
            raise ValueError("JSON invalido para dict")
        if isinstance(contenido_default, list) and not isinstance(cargado, list):
            raise ValueError("JSON invalido para list")
    except Exception:
        respaldo = path_obj.with_suffix(path_obj.suffix + ".bak")
        try:
            if path_obj.exists():
                path_obj.replace(respaldo)
        except Exception:
            pass
        path_obj.write_text(json.dumps(contenido_default, ensure_ascii=False, indent=2), encoding="utf-8")

def bootstrap_archivos_base():
    MEDIA_DIR.mkdir(exist_ok=True)
    COMPROBANTES_DIR.mkdir(exist_ok=True)
    asegurar_json_archivo(CONFIG_PATH, DEFAULT_DATOS_BOT)
    asegurar_json_archivo(RUNTIME_STATE_PATH, {"estados_clientes": {}, "solicitudes_comprobante": {}})
    asegurar_json_archivo(VENTAS_PATH, [])
    asegurar_json_archivo(CLIENTES_VENDIDOS_PATH, {})
    asegurar_json_archivo(CATALOGO_PATH, DEFAULT_CATALOGO)

def verificar_dependencias():
    faltantes = []
    opcionales = []
    import importlib
    for modulo, paquete in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(modulo)
        except Exception:
            faltantes.append(paquete)
    for modulo, paquete in OPTIONAL_PACKAGES.items():
        try:
            importlib.import_module(modulo)
        except Exception:
            opcionales.append(paquete)
    return faltantes, opcionales

def resumen_arranque():
    faltantes, opcionales = verificar_dependencias()
    estado = {
        "base_dir": str(BASE_DIR),
        "faltantes": faltantes,
        "opcionales": opcionales,
        "archivos": {
            "bot_config.json": CONFIG_PATH.exists(),
            "runtime_state.json": RUNTIME_STATE_PATH.exists(),
            "ventas.json": VENTAS_PATH.exists(),
            "clientes_vendidos.json": CLIENTES_VENDIDOS_PATH.exists(),
            "catalogo.json": CATALOGO_PATH.exists(),
            "media": MEDIA_DIR.exists(),
            "comprobantes": COMPROBANTES_DIR.exists(),
        }
    }
    return estado

def cargar_estado_runtime():
    estado = {"estados_clientes": {}, "solicitudes_comprobante": {}}
    if RUNTIME_STATE_PATH.exists():
        try:
            cargado = json.loads(RUNTIME_STATE_PATH.read_text(encoding="utf-8"))
            if isinstance(cargado, dict):
                estado["estados_clientes"] = cargado.get("estados_clientes", {})
                estado["solicitudes_comprobante"] = cargado.get("solicitudes_comprobante", {})
        except (json.JSONDecodeError, OSError):
            pass
    return estado


def guardar_estado_runtime():
    estado = {
        "estados_clientes": estados_clientes,
        "solicitudes_comprobante": solicitudes_comprobante
    }
    RUNTIME_STATE_PATH.write_text(
        json.dumps(estado, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def cargar_datos_bot():
    datos = DEFAULT_DATOS_BOT.copy()
    if CONFIG_PATH.exists():
        try:
            datos_guardados = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(datos_guardados, dict):
                datos.update(datos_guardados)
        except (json.JSONDecodeError, OSError):
            pass
    # Corrección de seguridad: si una configuración vieja conserva el
    # identificador anterior del número de WhatsApp, usar el ID correcto por
    # defecto. Esto evita que la prueba Meta falle aunque bot_config.json haya
    # quedado con el valor antiguo tras un despliegue o copia local.
    if str(datos.get("id_telefono", "")).strip() in {"", "1067511759782672"}:
        datos["id_telefono"] = ID_TELEFONO_DEFAULT
    for campo_obsoleto in (
        "auto_ngrok",
        "auto_pinggy",
        "precio_normal",
        "precio_descuento",
        "bancolombia",
        "sim_cliente_custom_apertura",
        "sim_cliente_custom_interes",
        "sim_cliente_custom_objecion",
        "sim_cliente_custom_confirmacion",
        "sim_cliente_custom_comprobante",
        "sim_cliente_custom_cierre",
    ):
        datos.pop(campo_obsoleto, None)
    return datos


def guardar_datos_bot():
    CONFIG_PATH.write_text(
        json.dumps(datos_bot, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def cargar_catalogo():
    catalogo = DEFAULT_CATALOGO.copy()
    if CATALOGO_PATH.exists():
        try:
            data = json.loads(CATALOGO_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                catalogo = data
        except (json.JSONDecodeError, OSError):
            pass
    return normalizar_catalogo(catalogo)


def guardar_catalogo():
    global catalogo_productos
    catalogo_productos = normalizar_catalogo(catalogo_productos)
    CATALOGO_PATH.write_text(
        json.dumps(catalogo_productos, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def construir_contexto_producto(clave, producto):
    titulo = producto.get("titulo", clave)
    precio = str(producto.get("precio_normal", "0"))
    precio_descuento = str(producto.get("precio_descuento", "") or precio)
    link_video = str(producto.get("link_video", "") or datos_bot.get("link_video", ""))
    link_pdf = str(producto.get("link_pdf", "") or datos_bot.get("link_pdf_demo", ""))
    contexto = dict(datos_bot)
    contexto.update(
        {
            "clave": clave,
            "titulo": titulo,
            "precio": precio,
            "precio_descuento": precio_descuento,
            "link_video": link_video,
            "link_pdf": link_pdf,
        }
    )
    return contexto


def obtener_mensaje_producto(clave, producto, tipo):
    plantilla_default = DEFAULT_MENSAJES_PRODUCTO.get(tipo, "")
    mensajes_producto = producto.get("mensajes", {}) if isinstance(producto, dict) else {}
    plantilla = (mensajes_producto.get(tipo, "") or "").strip() or plantilla_default
    contexto = construir_contexto_producto(clave, producto)
    try:
        return (plantilla or "").format(**contexto)
    except Exception:
        return plantilla


def obtener_mensaje_bienvenida_ab(numero_cliente, clave, producto):
    mensajes_producto = producto.get("mensajes", {}) if isinstance(producto, dict) else {}

    # Compatibilidad con configuraciones antiguas donde solo existe "bienvenida".
    bienvenida_base = str(mensajes_producto.get("bienvenida", "") or "").strip() or DEFAULT_MENSAJES_PRODUCTO.get("bienvenida", "")
    bienvenida_a = str(mensajes_producto.get("bienvenida_a", "") or "").strip() or bienvenida_base
    bienvenida_b = str(mensajes_producto.get("bienvenida_b", "") or "").strip() or bienvenida_base

    variante = "a"
    if numero_cliente:
        valor = sum(ord(ch) for ch in str(numero_cliente))
        variante = "a" if (valor % 2 == 0) else "b"

    tipo = "bienvenida_a" if variante == "a" else "bienvenida_b"
    plantilla = bienvenida_a if variante == "a" else bienvenida_b
    contexto = construir_contexto_producto(clave, producto)

    try:
        mensaje = (plantilla or DEFAULT_MENSAJES_PRODUCTO.get(tipo, "")).format(**contexto)
        return mensaje, variante
    except Exception:
        return plantilla or bienvenida_base, variante


def normalizar_producto_catalogo(clave, producto):
    p = producto if isinstance(producto, dict) else {}
    producto_id = normalizar_clave_catalogo(p.get("id", clave)) or clave
    opcion = str(p.get("opcion", "")).strip()
    titulo = str(p.get("titulo", clave)).strip() or clave
    plantilla_info_meta = str(p.get("plantilla_info_meta", "") or p.get("plantilla_meta", "") or f"info_{producto_id}_v1").strip()
    precio_normal = str(p.get("precio_normal", "") or p.get("precio", "0")).strip() or "0"
    precio_descuento = str(p.get("precio_descuento", "")).strip() or precio_normal
    link_video = str(p.get("link_video", "")).strip()
    link_pdf = str(p.get("link_pdf", "")).strip()
    link_drive_final = str(p.get("link_drive_final", "") or p.get("link_entrega", "")).strip()
    link_canal_whatsapp = str(p.get("link_canal_whatsapp", "") or datos_bot.get("link_canal_whatsapp", "")).strip() if "datos_bot" in globals() else str(p.get("link_canal_whatsapp", "")).strip()
    palabras = p.get("palabras_clave", [])
    if not isinstance(palabras, list):
        palabras = [str(palabras)]
    palabras_limpias = [str(x).strip() for x in palabras if str(x).strip()]
    if not palabras_limpias:
        palabras_limpias = [clave]

    mensajes_raw = p.get("mensajes", {})
    if not isinstance(mensajes_raw, dict):
        mensajes_raw = {}
    mensajes = {}
    for tipo, plantilla_default in DEFAULT_MENSAJES_PRODUCTO.items():
        mensajes[tipo] = str(mensajes_raw.get(tipo, "")).strip() or plantilla_default

    legacy_bienvenida = str(mensajes_raw.get("bienvenida", "")).strip()
    if legacy_bienvenida:
        if not str(mensajes_raw.get("bienvenida_a", "")).strip():
            mensajes["bienvenida_a"] = legacy_bienvenida
        if not str(mensajes_raw.get("bienvenida_b", "")).strip():
            mensajes["bienvenida_b"] = legacy_bienvenida
        mensajes["bienvenida"] = legacy_bienvenida

    # texto_info_meta: cuerpo de la plantilla unica del producto para Meta.
    # Prioridad: campo texto_info_meta > mensajes.info (si no es generico) > plantilla base genérica
    texto_info_meta_raw = str(p.get("texto_info_meta", "") or "").strip()
    if not texto_info_meta_raw:
        msg_info = str(mensajes_raw.get("info", "") or "").strip()
        default_info = DEFAULT_MENSAJES_PRODUCTO.get("info", "")
        if msg_info and msg_info != default_info and "{titulo}" not in msg_info:
            texto_info_meta_raw = msg_info
        else:
            texto_info_meta_raw = ""

    return {
        "id": producto_id,
        "opcion": opcion,
        "titulo": titulo,
        "plantilla_info_meta": plantilla_info_meta,
        "texto_info_meta": texto_info_meta_raw,
        "precio_normal": precio_normal,
        "precio_descuento": precio_descuento,
        "link_video": link_video,
        "link_pdf": link_pdf,
        "link_drive_final": link_drive_final,
        "link_canal_whatsapp": link_canal_whatsapp,
        "palabras_clave": palabras_limpias,
        "mensajes": mensajes,
    }


def normalizar_catalogo(catalogo):
    entrada = catalogo if isinstance(catalogo, dict) else {}
    if not entrada:
        entrada = DEFAULT_CATALOGO
    normalizado = {}
    for clave_raw, producto in entrada.items():
        clave = normalizar_clave_catalogo(clave_raw) or "producto"
        base = clave
        suf = 2
        while clave in normalizado:
            clave = f"{base}_{suf}"
            suf += 1
        normalizado[clave] = normalizar_producto_catalogo(clave, producto)

    # Asegurar opciones consecutivas y unicas para el formato de carrete.
    opciones_usadas = set()
    siguiente = 1
    for clave in normalizado:
        opcion = str(normalizado[clave].get("opcion", "")).strip()
        if not opcion or opcion in opciones_usadas:
            while str(siguiente) in opciones_usadas:
                siguiente += 1
            opcion = str(siguiente)
        opciones_usadas.add(opcion)
        normalizado[clave]["opcion"] = opcion

    return normalizado


def normalizar_clave_catalogo(texto):
    base = str(texto or "").strip().lower().replace(" ", "_")
    return re.sub(r"[^a-z0-9_]", "", base)


def buscar_producto_catalogo_por_texto(texto):
    def _limpiar(valor):
        s = str(valor or "").lower().strip()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = re.sub(r"[^a-z0-9\s]", " ", s)
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    t = _limpiar(texto)
    if not t:
        return None, None

    tokens_texto = set(t.split())
    palabras_ruido = {
        "quiero", "info", "informacion", "sobre", "del", "de", "el", "la", "los", "las",
        "un", "una", "pack", "curso", "producto", "por", "para", "me", "das", "ver"
    }
    mejor_clave = None
    mejor_producto = None
    mejor_puntaje = 0

    for clave, producto in (catalogo_productos or {}).items():
        opcion = _limpiar(producto.get("opcion", ""))
        if opcion and t == opcion:
            return clave, producto

        producto_id = _limpiar(producto.get("id", clave))
        clave_norm = _limpiar(clave.replace("_", " "))
        titulo_norm = _limpiar(producto.get("titulo", clave))
        palabras_clave = [_limpiar(p) for p in (producto.get("palabras_clave", []) or []) if _limpiar(p)]

        # Coincidencias fuertes por frase completa (funciona para textos largos de anuncios).
        for frase in [producto_id, clave_norm, titulo_norm, *palabras_clave]:
            if not frase or frase in palabras_ruido:
                continue
            if t == frase or frase in t:
                return clave, producto

        # Coincidencia por tokens relevantes.
        tokens_producto = set()
        for valor in [clave_norm, producto_id, *palabras_clave]:
            for tk in valor.split():
                if len(tk) >= 4 and tk not in palabras_ruido:
                    tokens_producto.add(tk)

        if not tokens_producto:
            continue
        puntaje = len(tokens_texto & tokens_producto)
        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_clave = clave
            mejor_producto = producto

    if mejor_puntaje >= 1:
        return mejor_clave, mejor_producto
    return None, None


def parsear_accion_catalogo(texto):
    t = normalizar_texto(texto)
    for prefijo in ("video_", "pdf_", "comprar_"):
        if t.startswith(prefijo):
            return prefijo[:-1], normalizar_clave_catalogo(t[len(prefijo):])
    for prefijo in ("video ", "pdf ", "comprar "):
        if t.startswith(prefijo):
            return prefijo.strip(), normalizar_clave_catalogo(t[len(prefijo):])
    return None, None


def catalogo_json_pretty():
    return json.dumps(catalogo_productos, ensure_ascii=False, indent=2)


def cargar_ventas():
    if not VENTAS_PATH.exists():
        return []
    try:
        data = json.loads(VENTAS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def guardar_ventas():
    VENTAS_PATH.write_text(
        json.dumps(ventas_registradas, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def cargar_clientes_vendidos():
    if not CLIENTES_VENDIDOS_PATH.exists():
        return {}
    try:
        data = json.loads(CLIENTES_VENDIDOS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def guardar_clientes_vendidos():
    CLIENTES_VENDIDOS_PATH.write_text(
        json.dumps(clientes_vendidos, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


def extraer_numeros_texto(texto):
    candidatos = re.findall(r"\d{10,15}", texto or "")
    numeros = []
    for c in candidatos:
        n = normalizar_numero_whatsapp(c)
        if len(n) >= 10:
            numeros.append(n)
    return numeros


def extraer_numeros_archivo(file_obj):
    if not file_obj or file_obj.filename == "":
        return []
    try:
        contenido = file_obj.read().decode("utf-8", errors="ignore")
    except Exception:
        return []
    return extraer_numeros_texto(contenido)


def convertir_monto(texto_monto):
    limpio = "".join(ch for ch in str(texto_monto or "") if ch.isdigit())
    return int(limpio) if limpio else 0


def es_archivo_permitido(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def guardar_archivo_media(file_obj, file_prefix):
    """Guarda un archivo en la carpeta media y retorna la ruta relativa"""
    if not file_obj or file_obj.filename == '':
        return None
    
    if not es_archivo_permitido(file_obj.filename):
        return None
    
    ext = file_obj.filename.rsplit('.', 1)[1].lower()
    filename = secure_filename(f"{file_prefix}.{ext}")
    filepath = MEDIA_DIR / filename
    
    file_obj.save(str(filepath))
    return f"/media/{filename}"


def normalizar_numero_whatsapp(numero):
    return "".join(ch for ch in str(numero or "") if ch.isdigit())


def construir_url_media(link):
    if not link:
        return ""
    if link.startswith("/media/") and PUBLIC_TUNNEL_URL:
        return f"{PUBLIC_TUNNEL_URL}{link}"
    return link


def obtener_estado_recurso(link):
    estado = {
        "link": link or "",
        "es_local": False,
        "existe": False,
        "archivo": "",
        "tamano_kb": 0,
        "url_visual": construir_url_media(link or ""),
        "url_local": ""
    }
    if not link:
        return estado

    if link.startswith("/media/"):
        estado["es_local"] = True
        estado["url_local"] = link
        ruta_local = BASE_DIR / link.lstrip("/")
        estado["archivo"] = ruta_local.name
        estado["existe"] = ruta_local.exists()
        if estado["existe"]:
            estado["tamano_kb"] = round(ruta_local.stat().st_size / 1024, 1)
    else:
        estado["url_local"] = link
    return estado


def programar_recordatorio_descuento(numero_cliente):
    def _recordatorio():
        estado = estados_clientes.get(numero_cliente, {})
        if not estado:
            return
        if estado.get("recordatorio_descuento_enviado"):
            return
        if estado.get("hizo_click_comprar") or estado.get("esperando_comprobante") or estado.get("en_revision"):
            return
        ultima_actividad = estado.get("ultima_actividad", 0)
        if time.time() - ultima_actividad < DESCUENTO_RECORDATORIO_SEGUNDOS:
            return

        clave = estado.get("catalogo_activo", "")
        producto = catalogo_productos.get(clave, {}) if clave else {}
        kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave, producto)
        enviar_mensaje_plantilla(numero_cliente, "descuento_producto", **kwargs_tpl)
        estado["oferta_descuento_activa"] = True
        enviar_mensaje_botones(
            numero_cliente,
            "Si quieres, te ayudo de una vez con la compra.",
            ["comprar", "ver_pdf", "menu_principal"],
            pie="Promocion activa"
        )
        estado["recordatorio_descuento_enviado"] = True
        guardar_estado_runtime()

    Timer(DESCUENTO_RECORDATORIO_SEGUNDOS, _recordatorio).start()


def cliente_no_ha_comprado(numero_cliente):
    estado = estados_clientes.get(numero_cliente, {})
    return not (
        estado.get("hizo_click_comprar")
        or estado.get("esperando_comprobante")
        or estado.get("en_revision")
    )


def programar_ultima_oportunidad_descuento(numero_cliente):
    def _recordatorio_ultima_oportunidad():
        estado = estados_clientes.get(numero_cliente, {})
        if not estado:
            return
        if estado.get("ultima_oportunidad_enviada"):
            return
        if not cliente_no_ha_comprado(numero_cliente):
            return
        ultima_actividad = estado.get("ultima_actividad", 0)
        if time.time() - ultima_actividad < ULTIMA_OPORTUNIDAD_SEGUNDOS:
            return

        clave = estado.get("catalogo_activo", "")
        producto = catalogo_productos.get(clave, {}) if clave else {}
        kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave, producto)
        enviar_mensaje_plantilla(numero_cliente, "descuento_ultima_oportunidad", **kwargs_tpl)
        estado["oferta_descuento_activa"] = True
        enviar_mensaje_botones(
            numero_cliente,
            "Si quieres cerrar tu compra ahora, te acompano en el proceso.",
            ["comprar", "ver_pdf", "menu_principal"],
            pie="Ultima oportunidad"
        )
        estado["ultima_oportunidad_enviada"] = True
        guardar_estado_runtime()

    Timer(ULTIMA_OPORTUNIDAD_SEGUNDOS, _recordatorio_ultima_oportunidad).start()


def programar_seguimiento_video(numero_cliente):
    def _seguimiento():
        estado = estados_clientes.get(numero_cliente, {})
        if not estado:
            return
        if estado.get("mensaje_pdf_sugerido_enviado"):
            return
        if estado.get("hizo_click_comprar") or estado.get("esperando_comprobante") or estado.get("en_revision"):
            return
        ultima_actividad = estado.get("ultima_actividad", 0)
        if time.time() - ultima_actividad < VIDEO_SEGUIMIENTO_SEGUNDOS:
            return

        enviar_mensaje_botones(
            numero_cliente,
            "Cuando termines de ver el video, tambien puedo enviarte el demo PDF para que compares el contenido.",
            ["ver_pdf", "comprar", "menu_principal"],
            pie="Siguiente paso"
        )
        estado["mensaje_pdf_sugerido_enviado"] = True
        guardar_estado_runtime()

    Timer(VIDEO_SEGUIMIENTO_SEGUNDOS, _seguimiento).start()


def renderizar_plantilla(texto):
    contexto = dict(datos_bot)
    contexto.update(obtener_config_flujos())
    return (texto or "").format(**contexto)

def obtener_config_flujos():
    flujos = datos_bot.get("flujos_automaticos", {})
    if not isinstance(flujos, dict):
        flujos = {}
    combinado = DEFAULT_FLOWS.copy()
    combinado.update(flujos)
    return combinado

def actualizar_etapa_cliente(numero_cliente, etapa, **extras):
    estado = estados_clientes.get(numero_cliente, {})
    estado["etapa"] = etapa
    estado["ultima_actividad"] = time.time()
    for k, v in extras.items():
        estado[k] = v
    estados_clientes[numero_cliente] = estado
    guardar_estado_runtime()
    return estado

def obtener_mes_venta(fecha_texto):
    try:
        return datetime.strptime(fecha_texto, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m")
    except Exception:
        return ""

def obtener_resumen_mensual(ventas):
    resumen = {}
    for venta in ventas:
        mes = obtener_mes_venta(venta.get("fecha_aprobacion", "")) or "sin-fecha"
        item = resumen.setdefault(mes, {"mes": mes, "ventas": 0, "ingresos": 0})
        item["ventas"] += 1
        item["ingresos"] += int(venta.get("valor_venta", 0) or 0)
    return sorted(resumen.values(), key=lambda x: x["mes"], reverse=True)

def obtener_embudo_clientes():
    conteo = {etapa: 0 for etapa in ETAPAS_CLIENTE}
    for estado in estados_clientes.values():
        etapa = estado.get("etapa", "nuevo")
        if etapa not in conteo:
            conteo[etapa] = 0
        conteo[etapa] += 1
    return conteo


def obtener_clientes_por_etapa():
    filas = []
    for numero, estado in estados_clientes.items():
        filas.append({
            "numero": numero,
            "etapa": estado.get("etapa", "nuevo"),
            "catalogo_activo": estado.get("catalogo_activo", ""),
            "producto_titulo": estado.get("producto_titulo", ""),
            "metodo_pago": estado.get("metodo_pago", ""),
            "ultima_actividad": datetime.fromtimestamp(estado.get("ultima_actividad", 0)).strftime("%Y-%m-%d %H:%M:%S") if estado.get("ultima_actividad") else "",
            "origen_lead": estado.get("origen_lead", "organico"),
            "campaign_name": estado.get("campaign_name", "")
        })
    return sorted(filas, key=lambda x: x.get("ultima_actividad", ""), reverse=True)

def obtener_biblioteca_demo():
    items = []
    for clave, producto in catalogo_productos.items():
        items.append({
            "clave": clave,
            "titulo": producto.get("titulo", clave),
            "video": obtener_estado_recurso(producto.get("link_video", "") or datos_bot.get("link_video", "")),
            "pdf": obtener_estado_recurso(producto.get("link_pdf", "") or datos_bot.get("link_pdf_demo", ""))
        })
    return items

def guardar_flujos_desde_form(form):
    datos_bot["flujos_automaticos"] = {
        "video_followup_seconds": int(form.get("video_followup_seconds", 240) or 240),
        "pdf_followup_seconds": int(form.get("pdf_followup_seconds", 120) or 120),
        "discount_reminder_seconds": int(form.get("discount_reminder_seconds", 1800) or 1800),
        "last_chance_seconds": int(form.get("last_chance_seconds", 43200) or 43200),
        "send_pdf_after_video": form.get("send_pdf_after_video") == "on",
        "send_buy_after_pdf": form.get("send_buy_after_pdf") == "on",
        "usar_lista_pago": form.get("usar_lista_pago") == "on"
    }
    guardar_datos_bot()

def resumen_ads(ventas):
    filas = {}
    for venta in ventas:
        clave = venta.get("campaign_name") or venta.get("campaign_id") or venta.get("origen_lead") or "organico"
        item = filas.setdefault(clave, {"campana": clave, "ventas": 0, "ingresos": 0})
        item["ventas"] += 1
        item["ingresos"] += int(venta.get("valor_venta", 0) or 0)
    return sorted(filas.values(), key=lambda x: x["ingresos"], reverse=True)

def obtener_origen_cliente(numero_cliente):
    estado = estados_clientes.get(numero_cliente, {})
    return {
        "origen_lead": estado.get("origen_lead", "organico"),
        "campaign_id": estado.get("campaign_id", ""),
        "campaign_name": estado.get("campaign_name", ""),
        "adset_id": estado.get("adset_id", ""),
        "ad_id": estado.get("ad_id", ""),
        "producto_interes_inicial": estado.get("producto_interes_inicial", estado.get("catalogo_activo", ""))
    }

def registrar_origen_desde_texto(numero_cliente, texto):
    texto = texto or ""
    pares = re.findall(r"(src|source|campaign|campaign_id|campaign_name|adset_id|ad_id|producto)=([^\s,;]+)", texto, flags=re.I)
    if not pares:
        return
    estado = estados_clientes.get(numero_cliente, {})
    mapa = {
        "src": "origen_lead",
        "source": "origen_lead",
        "campaign": "campaign_name",
        "campaign_id": "campaign_id",
        "campaign_name": "campaign_name",
        "adset_id": "adset_id",
        "ad_id": "ad_id",
        "producto": "producto_interes_inicial"
    }
    for k, v in pares:
        estado[mapa.get(k.lower(), k.lower())] = v
    estados_clientes[numero_cliente] = estado
    guardar_estado_runtime()

def programar_seguimiento_pdf(numero_cliente):
    def _seguimiento_pdf():
        estado = estados_clientes.get(numero_cliente, {})
        if not estado:
            return
        if estado.get("hizo_click_comprar") or estado.get("esperando_comprobante") or estado.get("en_revision"):
            return
        if estado.get("mensaje_compra_sugerida_enviado"):
            return
        cfg = obtener_config_flujos()
        espera = int(cfg.get("pdf_followup_seconds", PDF_SEGUIMIENTO_SEGUNDOS))
        ultima_actividad = estado.get("ultima_actividad", 0)
        if time.time() - ultima_actividad < espera:
            return
        clave = estado.get("catalogo_activo", "")
        producto = catalogo_productos.get(clave, {}) if clave else {}
        texto = "¿Ya revisaste el PDF? Si quieres, te ayudo a activar tu acceso ahora mismo."
        botones = [(f"comprar_{clave}", "Comprar"), (f"video_{clave}", "Ver video"), ("menu_principal", "Menu")] if producto else ["comprar", "ver_video", "menu_principal"]
        enviar_mensaje_botones(numero_cliente, texto, botones, pie="Siguiente paso")
        estado["mensaje_compra_sugerida_enviado"] = True
        guardar_estado_runtime()
    espera = int(obtener_config_flujos().get("pdf_followup_seconds", PDF_SEGUIMIENTO_SEGUNDOS))
    Timer(espera, _seguimiento_pdf).start()



bootstrap_archivos_base()
datos_bot = cargar_datos_bot()
if not CONFIG_PATH.exists():
    guardar_datos_bot()

estado_runtime = cargar_estado_runtime()
estados_clientes = estado_runtime["estados_clientes"]
solicitudes_comprobante = estado_runtime["solicitudes_comprobante"]
ventas_registradas = cargar_ventas()
clientes_vendidos = cargar_clientes_vendidos()
catalogo_productos = cargar_catalogo()
if not RUNTIME_STATE_PATH.exists():
    guardar_estado_runtime()
if not CATALOGO_PATH.exists():
    guardar_catalogo()

if not clientes_vendidos and ventas_registradas:
    for venta in ventas_registradas:
        numero = normalizar_numero_whatsapp(venta.get("numero_cliente", ""))
        if not numero:
            continue
        actual = clientes_vendidos.get(numero, {"numero": numero, "ventas": 0})
        actual["ventas"] = int(actual.get("ventas", 0)) + 1
        actual["ultima_venta"] = venta.get("fecha_aprobacion", "")
        actual["ultimo_metodo_pago"] = venta.get("metodo_pago", "")
        actual["ultimo_tipo_precio"] = venta.get("tipo_precio", "")
        actual["ultimo_producto"] = venta.get("producto_titulo", venta.get("catalogo_clave", ""))
        clientes_vendidos[numero] = actual
    guardar_clientes_vendidos()


# ==========================================
# 🛠️ FUNCIONES DE WHATSAPP
# ==========================================
def enviar_peticion_whatsapp(data):
    id_telefono = datos_bot.get("id_telefono", ID_TELEFONO_DEFAULT).strip()
    token_meta = datos_bot.get("token_meta", TOKEN_META_DEFAULT).strip()
    # API v19.0: version estable y reciente de Meta Graph API
    url = f"https://graph.facebook.com/v19.0/{id_telefono}/messages"
    headers = {
        "Authorization": f"Bearer {token_meta}",
        "Content-Type": "application/json"
    }
    return requests.post(url, headers=headers, json=data, timeout=30)


def enviar_mensaje(numero_destino, texto):
    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "text",
        "text": {"body": texto}
    }
    return enviar_peticion_whatsapp(data)


def enviar_mensaje_plantilla(numero_destino, clave_plantilla, **kwargs):
    catalogo_plantillas = obtener_catalogo_plantillas_meta()
    plantilla = catalogo_plantillas.get(clave_plantilla)
    if not plantilla:
        return enviar_mensaje(numero_destino, f"Plantilla no configurada: {clave_plantilla}")

    variables_ordenadas = plantilla.get("variables_ordenadas", [])
    parameters = []
    for variable in variables_ordenadas:
        valor = limpiar_parametro_meta(kwargs.get(variable, ""))
        parameters.append({
            "type": "text",
            "text": valor
        })

    template_data = {
        "name": plantilla.get("nombre_meta", ""),
        "language": {
            "code": plantilla.get("idioma", "es")
        }
    }
    if parameters:
        template_data["components"] = [
            {
                "type": "body",
                "parameters": parameters
            }
        ]

    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "template",
        "template": template_data
    }
    return enviar_peticion_whatsapp(data)


def enviar_video(numero_destino, url_video, texto_pie):
    """Envia un video a WhatsApp.
    - URLs http/https (GitHub Raw, CDN, etc.): se usan directamente.
    - URLs locales /media/...: se convierten con PUBLIC_TUNNEL_URL si está disponible.
    - No descarga el archivo localmente.
    """
    if not url_video:
        return None
    # Solo convertir si es ruta local /media/
    if url_video.startswith('/media/'):
        if PUBLIC_TUNNEL_URL:
            url_video = f"{PUBLIC_TUNNEL_URL}{url_video}"
        else:
            # Sin tunnel, no se puede enviar URL local a Meta
            return None
    # URLs http/https: usar directamente (GitHub Raw, GDrive direct, CDN, etc.)
    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "video",
        "video": {
            "link": url_video,
            "caption": texto_pie or ""
        }
    }
    return enviar_peticion_whatsapp(data)


def enviar_documento(numero_destino, url_documento, nombre_archivo):
    """Envia un documento/PDF a WhatsApp.
    - URLs http/https (GitHub Raw, CDN, etc.): se usan directamente.
    - URLs locales /media/...: se convierten con PUBLIC_TUNNEL_URL si está disponible.
    - No descarga el archivo localmente.
    """
    if not url_documento:
        return None
    # Solo convertir si es ruta local /media/
    if url_documento.startswith('/media/'):
        if PUBLIC_TUNNEL_URL:
            url_documento = f"{PUBLIC_TUNNEL_URL}{url_documento}"
        else:
            # Sin tunnel, no se puede enviar URL local a Meta
            return None
    # URLs http/https: usar directamente (GitHub Raw, GDrive direct, CDN, etc.)
    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "document",
        "document": {
            "link": url_documento,
            "filename": nombre_archivo or "documento.pdf"
        }
    }
    return enviar_peticion_whatsapp(data)


def enviar_media_por_id(numero_destino, tipo_media, media_id, nombre_archivo="Comprobante"):
    if not media_id:
        return None

    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": tipo_media
    }

    if tipo_media == "image":
        data["image"] = {"id": media_id}
    elif tipo_media == "document":
        data["document"] = {
            "id": media_id,
            "filename": nombre_archivo or "Comprobante"
        }
    else:
        return None

    return enviar_peticion_whatsapp(data)


def enviar_material_difusion(numero_destino, link_material):
    if not link_material:
        return None
    link_norm = normalizar_texto(link_material)
    if any(link_norm.endswith(ext) for ext in (".mp4", ".mov", ".m4v", ".webm")):
        return enviar_video(numero_destino, link_material, "Material adicional")
    nombre_archivo = Path(link_material).name if link_material.startswith("/media/") else "Material.pdf"
    return enviar_documento(numero_destino, link_material, nombre_archivo)


def descargar_media_comprobante(media_id, nombre_archivo=""):
    if not media_id:
        return ""

    token_meta = datos_bot.get("token_meta", TOKEN_META_DEFAULT).strip()
    headers = {"Authorization": f"Bearer {token_meta}"}
    try:
        meta = requests.get(f"https://graph.facebook.com/v19.0/{media_id}", headers=headers, timeout=30)
        if not meta.ok:
            return ""
        media_url = meta.json().get("url", "")
        if not media_url:
            return ""
        contenido = requests.get(media_url, headers=headers, timeout=60)
        if not contenido.ok:
            return ""

        base_name = secure_filename(nombre_archivo or f"comprobante_{media_id}")
        marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
        archivo_local = COMPROBANTES_DIR / f"{marca_tiempo}_{base_name}"
        archivo_local.write_bytes(contenido.content)
        return str(archivo_local)
    except requests.RequestException:
        return ""


def registrar_venta_aprobada(numero_cliente, solicitud, aprobado_desde="panel"):
    catalogo_clave = solicitud.get("catalogo_clave", "")
    producto = catalogo_productos.get(catalogo_clave, {}) if catalogo_clave else {}
    producto_titulo = solicitud.get("producto_titulo") or producto.get("titulo", catalogo_clave or "General")
    tipo_precio = solicitud.get("tipo_precio", "normal")
    valor_venta = solicitud.get("valor_venta", 0)
    if not valor_venta:
        if producto:
            precio_base = producto.get("precio_descuento" if tipo_precio == "descuento" else "precio_normal", "0")
            valor_venta = convertir_monto(precio_base)
        else:
            estado_cliente = estados_clientes.get(numero_cliente, {})
            clave_estado = estado_cliente.get("catalogo_activo", "")
            producto_estado = catalogo_productos.get(clave_estado, {}) if clave_estado else {}
            if producto_estado:
                precio_base = producto_estado.get("precio_descuento" if tipo_precio == "descuento" else "precio_normal", "0")
                valor_venta = convertir_monto(precio_base)
            else:
                valor_venta = int(estado_cliente.get("valor_venta", 0) or 0)

    origen = obtener_origen_cliente(numero_cliente)
    estado_actual = estados_clientes.get(numero_cliente, {})
    venta = {
        "fecha_aprobacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "numero_cliente": numero_cliente,
        "metodo_pago": solicitud.get("metodo_pago", "No definido"),
        "catalogo_clave": catalogo_clave,
        "producto_titulo": producto_titulo,
        "tipo_precio": tipo_precio,
        "bienvenida_variante": solicitud.get("bienvenida_variante") or estado_actual.get("bienvenida_variante", ""),
        "valor_venta": int(valor_venta),
        "tipo_archivo": solicitud.get("tipo_archivo", ""),
        "nombre_archivo": solicitud.get("nombre_archivo", ""),
        "comprobante_local": solicitud.get("comprobante_local", ""),
        "aprobado_desde": aprobado_desde,
        "mes": datetime.now().strftime("%Y-%m"),
        **origen
    }
    ventas_registradas.append(venta)
    guardar_ventas()

    numero_norm = normalizar_numero_whatsapp(numero_cliente)
    if numero_norm:
        cliente = clientes_vendidos.get(numero_norm, {"numero": numero_norm, "ventas": 0})
        cliente["ventas"] = int(cliente.get("ventas", 0)) + 1
        cliente["ultima_venta"] = venta.get("fecha_aprobacion", "")
        cliente["ultimo_metodo_pago"] = venta.get("metodo_pago", "")
        cliente["ultimo_tipo_precio"] = venta.get("tipo_precio", "")
        cliente["ultimo_producto"] = venta.get("producto_titulo", venta.get("catalogo_clave", ""))
        clientes_vendidos[numero_norm] = cliente
        guardar_clientes_vendidos()


def enviar_mensaje_botones(numero_destino, texto, botones, pie="Selecciona una opcion"):
    botones_json = []
    for boton in botones[:3]:
        # Soportar botones dinámicos con formato (id, titulo) o solo id
        if isinstance(boton, tuple):
            boton_id, titulo = boton
        else:
            boton_id = boton
            # Manejar botones dinamicos de aprobacion/rechazo
            if boton_id.startswith("aprobar_"):
                titulo = "Aprobar"
            elif boton_id.startswith("rechazar_"):
                titulo = "Rechazar"
            else:
                titulo = BOTONES_RESPUESTA.get(boton_id, boton_id)[:20]

        botones_json.append(
            {
                "type": "reply",
                "reply": {
                    "id": boton_id,
                    "title": titulo[:20]
                }
            }
        )

    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "interactive",
        "interactive": {
            "type": "button",
            "body": {"text": texto},
            "footer": {"text": pie},
            "action": {"buttons": botones_json}
        }
    }
    respuesta = enviar_peticion_whatsapp(data)
    if not respuesta.ok:
        opciones_texto_list = []
        for boton in botones:
            if isinstance(boton, tuple):
                _, titulo = boton
                opciones_texto_list.append(titulo)
            else:
                if boton.startswith("aprobar_"):
                    opciones_texto_list.append("Aprobar")
                elif boton.startswith("rechazar_"):
                    opciones_texto_list.append("Rechazar")
                else:
                    opciones_texto_list.append(BOTONES_RESPUESTA.get(boton, boton))
        opciones_texto = ", ".join(opciones_texto_list)
        enviar_mensaje(numero_destino, f"{texto}\n\nOpciones: {opciones_texto}")
    return respuesta


def enviar_mensaje_lista(numero_destino, texto, opciones, titulo_boton="Ver opciones", pie="Selecciona una opcion"):
    filas = []
    for opcion_id in opciones:
        filas.append(
            {
                "id": opcion_id,
                "title": BOTONES_RESPUESTA[opcion_id][:24],
                "description": f"Pagar por {BOTONES_RESPUESTA[opcion_id]}"
            }
        )

    data = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "interactive",
        "interactive": {
            "type": "list",
            "body": {"text": texto},
            "footer": {"text": pie},
            "action": {
                "button": titulo_boton,
                "sections": [
                    {
                        "title": "Medios de pago disponibles",
                        "rows": filas
                    }
                ]
            }
        }
    }
    respuesta = enviar_peticion_whatsapp(data)
    if not respuesta.ok:
        opciones_texto = "\n".join(f"- {BOTONES_RESPUESTA[o]}" for o in opciones)
        enviar_mensaje(
            numero_destino,
            f"{texto}\n\nEscribe una de estas opciones para continuar:\n{opciones_texto}"
        )
    return respuesta


FLUJO_SIGUIENTE_PASO = {
    "categoria_no_encontrada": {
        "texto": "Te ayudo a retomar el flujo. Elige una opcion:",
        "botones": ["ver_video", "ver_pdf", "menu_principal"],
        "pie": "Menu rapido"
    },
    "agradecimiento": {
        "texto": "Si quieres, te acompano con el siguiente paso:",
        "botones": ["menu_principal", "ver_pdf", "comprar"],
        "pie": "Siguiente paso"
    },
    "confirmacion_pago": {
        "texto": "Cuando lo tengas listo, envia el comprobante por este chat.",
        "botones": ["menu_principal", "ver_pdf", "ya_pague"],
        "pie": "Validacion de pago"
    },
    "comprobante_recibido": {
        "texto": "Estamos validando tu comprobante. Te avisamos por aqui.",
        "botones": ["menu_principal", "ver_pdf", "comprar"],
        "pie": "En revision"
    },
    "comprobante_en_revision": {
        "texto": "Tu validacion sigue en proceso. Gracias por esperar.",
        "botones": ["menu_principal", "ver_pdf", "comprar"],
        "pie": "Revision"
    },
    "comprobante_fuera_flujo": {
        "texto": "Primero selecciona compra para asociar correctamente el pago.",
        "botones": ["comprar", "ver_pdf", "menu_principal"],
        "pie": "Orden del proceso"
    },
    "entrega_final": {
        "texto": "Listo. Si todo salio bien, confirma recibido o vuelve al menu.",
        "botones": ["menu_principal", "ver_pdf", "comprar"],
        "pie": "Soporte y seguimiento"
    },
}


def enviar_template_con_siguiente_paso(numero_destino, clave_plantilla, clave_flujo="", **kwargs):
    """Envia una plantilla y opcionalmente un mensaje de botones de seguimiento (sin sleep)."""
    respuesta = enviar_mensaje_plantilla(numero_destino, clave_plantilla, **kwargs)
    flujo = FLUJO_SIGUIENTE_PASO.get(clave_flujo or clave_plantilla)
    if not flujo:
        return respuesta
    # Sin sleep: envio secuencial inmediato
    enviar_mensaje_botones(
        numero_destino,
        flujo.get("texto", "Selecciona una opcion:"),
        flujo.get("botones", ["menu_principal"]),
        pie=flujo.get("pie", "Siguiente paso")
    )
    return respuesta


def normalizar_texto(texto):
    limpio = str(texto or "")
    limpio = limpio.replace("\u00a0", " ")
    limpio = re.sub(r"\s+", " ", limpio)
    return limpio.strip().lower()


def limpiar_parametro_meta(valor):
    texto = str(valor or "")
    texto = texto.replace("\n", ", ")
    texto = texto.replace("\t", "")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def construir_kwargs_plantilla(numero_cliente="", clave_catalogo="", producto=None, **extras):
    estado = estados_clientes.get(numero_cliente, {}) if numero_cliente else {}
    clave = clave_catalogo or estado.get("catalogo_activo", "")
    prod = producto or (catalogo_productos.get(clave, {}) if clave else {})

    tipo_precio_estado = "descuento" if estado.get("oferta_descuento_activa") else estado.get("tipo_precio", "normal")
    precio_base = prod.get("precio_descuento" if tipo_precio_estado == "descuento" else "precio_normal", "") if prod else ""
    precio_desc = prod.get("precio_descuento", precio_base) if prod else ""

    payload = {
        "nombre_cliente": estado.get("nombre_cliente", "Cliente"),
        "numero_cliente": numero_cliente,
        "catalogo_clave": clave,
        "nombre_producto": prod.get("titulo", estado.get("producto_titulo", clave or "material")),
        "precio": str(precio_base or ""),
        "precio_normal": str(prod.get("precio_normal", "") or ""),
        "precio_descuento": str(precio_desc or ""),
        "nequi": datos_bot.get("nequi", ""),
        "daviplata": datos_bot.get("daviplata", ""),
        "llave": datos_bot.get("llave", ""),
        "link_video": prod.get("link_video", "") or datos_bot.get("link_video", ""),
        "link_pdf": prod.get("link_pdf", "") or datos_bot.get("link_pdf_demo", ""),
        "link_drive_final": prod.get("link_drive_final", "") or datos_bot.get("link_recursos_final", ""),
        "link_canal_whatsapp": prod.get("link_canal_whatsapp", "") or datos_bot.get("link_canal_whatsapp", ""),
        "instruccion_pago": "Despues del pago, envia el comprobante por este chat para validar tu acceso.",
        "metodo_pago": estado.get("metodo_pago", ""),
    }
    payload.update(extras)
    return payload


def construir_lista_menu_productos():
    lineas = []
    for clave, producto in (catalogo_productos or {}).items():
        opcion = str(producto.get("opcion", "")).strip() or "-"
        titulo = str(producto.get("titulo", clave)).strip() or clave
        precio = str(producto.get("precio_normal", "")).strip()
        if precio:
            lineas.append(f"{opcion}. {titulo} (${precio})")
        else:
            lineas.append(f"{opcion}. {titulo}")
    return "\n".join(lineas) if lineas else "Sin productos configurados"


def resolver_producto_por_id_o_estado(numero_cliente, id_producto=""):
    id_normalizado = normalizar_clave_catalogo(id_producto)
    if id_normalizado and id_normalizado in catalogo_productos:
        return id_normalizado, catalogo_productos[id_normalizado]

    for clave, producto in (catalogo_productos or {}).items():
        if normalizar_clave_catalogo(producto.get("id", clave)) == id_normalizado:
            return clave, producto

    estado = estados_clientes.get(numero_cliente, {}) if numero_cliente else {}
    clave_estado = estado.get("catalogo_activo", "")
    if clave_estado and clave_estado in catalogo_productos:
        return clave_estado, catalogo_productos[clave_estado]
    return "", {}


def contiene_intencion(texto, opciones):
    texto_normalizado = normalizar_texto(texto)
    return any(opcion in texto_normalizado for opcion in opciones)


def es_link_video_directo(url):
    url_normalizada = normalizar_texto(url)
    return any(url_normalizada.endswith(ext) for ext in (".mp4", ".mov", ".m4v", ".webm"))


def cliente_puede_pedir_descuento(numero_cliente):
    """Compatibilidad: esta validacion ya no se usa para bloquear descuento manual."""
    estado = estados_clientes.get(numero_cliente, {})
    ultima_actividad = estado.get("ultima_actividad")
    
    if not ultima_actividad:
        return False
    
    tiempo_transcurrido = time.time() - ultima_actividad
    una_hora_en_segundos = 3600
    
    return tiempo_transcurrido >= una_hora_en_segundos


def extraer_texto_mensaje(mensaje_info):
    tipo = mensaje_info.get("type", "")
    if tipo == "text":
        return normalizar_texto(mensaje_info.get("text", {}).get("body")), tipo
    if tipo == "interactive":
        interactive = mensaje_info.get("interactive", {})
        button_reply = interactive.get("button_reply", {})
        list_reply = interactive.get("list_reply", {})
        return normalizar_texto(button_reply.get("id") or list_reply.get("id")), tipo
    return "", tipo


def extraer_media_mensaje(mensaje_info):
    tipo = mensaje_info.get("type", "")
    if tipo == "image":
        imagen = mensaje_info.get("image", {})
        return tipo, imagen.get("id"), "Comprobante.jpg"
    if tipo == "document":
        documento = mensaje_info.get("document", {})
        return tipo, documento.get("id"), documento.get("filename") or "Comprobante.pdf"
    return "", None, ""


def mostrar_menu_principal(numero_cliente):
    """Envia el menu general usando plantilla_menu_general.
    Resetea flags de embudo para que el cliente pueda reiniciar el flujo.
    """
    datos_bot["visitas_info"] = int(datos_bot.get("visitas_info", 0)) + 1
    guardar_datos_bot()
    estado = estados_clientes.get(numero_cliente, {})
    estado["hizo_click_comprar"] = False
    estado["recordatorio_descuento_enviado"] = False
    estado["ultima_oportunidad_enviada"] = False
    estado["oferta_descuento_activa"] = False
    estados_clientes[numero_cliente] = estado
    guardar_estado_runtime()
    lista_productos = construir_lista_menu_productos()
    enviar_mensaje_plantilla(numero_cliente, "plantilla_menu_general", lista_productos=lista_productos)


def enviar_opciones_compra(numero_cliente):
    estado = estados_clientes.get(numero_cliente, {})
    clave = estado.get("catalogo_activo", "")
    producto = catalogo_productos.get(clave, {}) if clave else {}
    producto_id = producto.get("id", clave)
    enviar_mensaje_botones(
        numero_cliente,
        "Para continuar con la compra, elige tu medio de pago.",
        [
            (f"pagar_nequi_{producto_id}", "Pagar Nequi"),
            (f"pagar_daviplata_{producto_id}", "Pagar Daviplata"),
            ("menu_principal", "Menú")
        ],
        pie="Compra"
    )


def enviar_cuentas_cobro(numero_cliente, clave_catalogo, metodo_pago):
    """Envia plantilla_cuentas_cobro e inmediatamente pide el comprobante."""
    producto = catalogo_productos.get(clave_catalogo, {}) if clave_catalogo else {}
    kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave_catalogo, producto, metodo_pago=metodo_pago)
    enviar_mensaje_plantilla(numero_cliente, "plantilla_cuentas_cobro", **kwargs_tpl)
    # Sin sleep: el envio secuencial es inmediato
    enviar_mensaje_botones(
        numero_cliente,
        "Cuando realices el pago, envia el comprobante (foto o PDF) por este chat.",
        ["ya_pague", "menu_principal"],
        pie="Confirmacion de pago"
    )


def enviar_datos_pago(numero_cliente, metodo_pago):
    estado_actual = estados_clientes.get(numero_cliente, {})
    clave = estado_actual.get("catalogo_activo", "")
    producto = catalogo_productos.get(clave, {}) if clave else {}
    if not producto:
        enviar_mensaje(
            numero_cliente,
            "Primero elige un producto del catálogo. Escribe menu para continuar."
        )
        return
    tipo_precio = "descuento" if estado_actual.get("oferta_descuento_activa") else "normal"
    precio_producto = producto.get("precio_descuento" if tipo_precio == "descuento" else "precio_normal", "0")
    valor_venta = convertir_monto(precio_producto)
    estado_nuevo = dict(estado_actual)
    estado_nuevo.update({
        "etapa": "medio_pago",
        "esperando_comprobante": True,
        "metodo_pago": metodo_pago,
        "ultima_actividad": estado_actual.get("ultima_actividad", time.time()),
        "hizo_click_comprar": True,
        "recordatorio_descuento_enviado": True,
        "tipo_precio": tipo_precio,
        "valor_venta": valor_venta,
        "oferta_descuento_activa": estado_actual.get("oferta_descuento_activa", False),
        "catalogo_activo": clave,
        "producto_titulo": producto.get("titulo", estado_actual.get("producto_titulo", "General"))
    })
    estados_clientes[numero_cliente] = estado_nuevo
    guardar_estado_runtime()
    enviar_cuentas_cobro(numero_cliente, clave, metodo_pago)


def enviar_recursos_finales(numero_cliente):
    estado = estados_clientes.get(numero_cliente, {})
    clave = estado.get("catalogo_activo", "")
    producto = catalogo_productos.get(clave, {}) if clave else {}
    estado_final = dict(estado)
    estado_final.update({"esperando_comprobante": False, "catalogo_activo": clave, "en_revision": False})
    estados_clientes[numero_cliente] = estado_final
    solicitudes_comprobante.pop(numero_cliente, None)
    guardar_estado_runtime()
    kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave, producto)
    # Envia plantilla_entrega_final con {{1}}=link_drive_final, {{2}}=link_canal_whatsapp
    enviar_mensaje_plantilla(numero_cliente, "plantilla_entrega_final", **kwargs_tpl)
    # Sin sleep: envio secuencial inmediato
    producto_id = producto.get("id", clave)
    enviar_mensaje_botones(
        numero_cliente,
        "Si necesitas soporte adicional, responde por este chat.",
        [
            ("menu_principal", "Menu"),
            (f"pdf_{producto_id}", "Ver PDF"),
            (f"video_{producto_id}", "Ver video")
        ],
        pie="Soporte"
    )


def registrar_comprobante_pendiente(numero_cliente, tipo_mensaje, media_id=None, nombre_archivo=""):
    estado_cliente = estados_clientes.get(numero_cliente, {})
    clave = estado_cliente.get("catalogo_activo", "")
    producto = catalogo_productos.get(clave, {}) if clave else {}
    comprobante_local = descargar_media_comprobante(media_id, nombre_archivo)
    solicitudes_comprobante[numero_cliente] = {
        "numero": numero_cliente,
        "metodo_pago": estado_cliente.get("metodo_pago", "No definido"),
        "catalogo_clave": clave,
        "producto_titulo": producto.get("titulo", estado_cliente.get("producto_titulo", "General")),
        "bienvenida_variante": estado_cliente.get("bienvenida_variante", ""),
        "tipo_archivo": tipo_mensaje,
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "media_id": media_id or "",
        "nombre_archivo": nombre_archivo or "",
        "comprobante_local": comprobante_local,
        "tipo_precio": estado_cliente.get("tipo_precio", "normal"),
        "valor_venta": estado_cliente.get("valor_venta", 0)
    }
    estado_revision = dict(estado_cliente)
    estado_revision.update({
        "etapa": "vendido",
        "esperando_comprobante": False,
        "en_revision": True,
        "metodo_pago": estado_cliente.get("metodo_pago", "No definido"),
        "ultima_actividad": time.time(),
        "hizo_click_comprar": True,
        "recordatorio_descuento_enviado": True
    })
    estados_clientes[numero_cliente] = estado_revision
    guardar_estado_runtime()
    
    # Notificar al admin
    numero_admin = normalizar_numero_whatsapp(datos_bot.get("numero_admin", ""))
    if numero_admin:
        metodo_pago = estado_cliente.get("metodo_pago", "No definido")
        mensaje_admin = (
            f"Nuevo comprobante pendiente de aprobacion.\n"
            f"Cliente: {numero_cliente}\n"
            f"Producto: {producto.get('titulo', estado_cliente.get('producto_titulo', 'General'))}\n"
            f"Metodo: {metodo_pago}\n"
            f"Tipo: {tipo_mensaje}\n"
            f"Hora: {datetime.now().strftime('%H:%M:%S')}"
        )
        
        enviar_mensaje_botones(
            numero_admin,
            mensaje_admin,
            [f"aprobar_{numero_cliente}", f"rechazar_{numero_cliente}"]
        )

        # Tambien enviar el comprobante para validacion visual del admin
        enviar_media_por_id(numero_admin, tipo_mensaje, media_id, nombre_archivo)


def obtener_solicitudes_ordenadas():
    return sorted(
        solicitudes_comprobante.values(),
        key=lambda solicitud: solicitud.get("fecha", ""),
        reverse=True
    )


def obtener_solicitudes_agrupadas():
    grupos = {}
    for solicitud in obtener_solicitudes_ordenadas():
        clave = solicitud.get("catalogo_clave") or "general"
        titulo = solicitud.get("producto_titulo") or clave.title()
        if clave not in grupos:
            grupos[clave] = {"clave": clave, "titulo": titulo, "items": []}
        grupos[clave]["items"].append(solicitud)
    return list(grupos.values())


def obtener_resumen_ventas():
    total = len(ventas_registradas)
    ventas_normal = [v for v in ventas_registradas if v.get("tipo_precio") == "normal"]
    ventas_descuento = [v for v in ventas_registradas if v.get("tipo_precio") == "descuento"]
    ingresos_total = sum(int(v.get("valor_venta", 0) or 0) for v in ventas_registradas)
    ingresos_normal = sum(int(v.get("valor_venta", 0) or 0) for v in ventas_normal)
    ingresos_descuento = sum(int(v.get("valor_venta", 0) or 0) for v in ventas_descuento)
    return {
        "total": total,
        "ventas_normal": len(ventas_normal),
        "ventas_descuento": len(ventas_descuento),
        "ingresos_total": ingresos_total,
        "ingresos_normal": ingresos_normal,
        "ingresos_descuento": ingresos_descuento,
    }


def parsear_fecha_filtro(texto):
    if not texto:
        return None
    try:
        return datetime.strptime(texto, "%Y-%m-%d").date()
    except ValueError:
        return None


def filtrar_ventas_por_fecha(desde_texto, hasta_texto, catalogo_clave=""):
    desde = parsear_fecha_filtro(desde_texto)
    hasta = parsear_fecha_filtro(hasta_texto)
    clave_filtro = normalizar_clave_catalogo(catalogo_clave)
    if not desde and not hasta and not clave_filtro:
        return list(ventas_registradas)

    filtradas = []
    for venta in ventas_registradas:
        fecha_aprobacion = venta.get("fecha_aprobacion", "")
        try:
            fecha_venta = datetime.strptime(fecha_aprobacion, "%Y-%m-%d %H:%M:%S").date()
        except ValueError:
            continue

        if desde and fecha_venta < desde:
            continue
        if hasta and fecha_venta > hasta:
            continue
        if clave_filtro and normalizar_clave_catalogo(venta.get("catalogo_clave", "")) != clave_filtro:
            continue
        filtradas.append(venta)
    return filtradas


def obtener_resumen_ventas_de(ventas):
    total = len(ventas)
    ventas_normal = [v for v in ventas if v.get("tipo_precio") == "normal"]
    ventas_descuento = [v for v in ventas if v.get("tipo_precio") == "descuento"]
    ingresos_total = sum(int(v.get("valor_venta", 0) or 0) for v in ventas)
    ingresos_normal = sum(int(v.get("valor_venta", 0) or 0) for v in ventas_normal)
    ingresos_descuento = sum(int(v.get("valor_venta", 0) or 0) for v in ventas_descuento)
    return {
        "total": total,
        "ventas_normal": len(ventas_normal),
        "ventas_descuento": len(ventas_descuento),
        "ingresos_total": ingresos_total,
        "ingresos_normal": ingresos_normal,
        "ingresos_descuento": ingresos_descuento,
    }


def obtener_resumen_por_producto(ventas):
    resumen = {}
    for venta in ventas:
        clave = venta.get("catalogo_clave", "") or "general"
        titulo = venta.get("producto_titulo", "") or clave.title()
        if clave not in resumen:
            resumen[clave] = {
                "clave": clave,
                "titulo": titulo,
                "ventas": 0,
                "ingresos": 0,
            }
        resumen[clave]["ventas"] += 1
        resumen[clave]["ingresos"] += int(venta.get("valor_venta", 0) or 0)
    return sorted(resumen.values(), key=lambda r: r["ingresos"], reverse=True)


def obtener_analitica_clientes_por_producto(ventas):
    analitica = {}
    for venta in ventas:
        clave = venta.get("catalogo_clave", "") or "general"
        titulo = venta.get("producto_titulo", "") or clave.title()
        numero = normalizar_numero_whatsapp(venta.get("numero_cliente", ""))
        if clave not in analitica:
            analitica[clave] = {
                "clave": clave,
                "titulo": titulo,
                "ventas": 0,
                "ingresos": 0,
                "clientes": set(),
            }
        analitica[clave]["ventas"] += 1
        analitica[clave]["ingresos"] += int(venta.get("valor_venta", 0) or 0)
        if numero:
            analitica[clave]["clientes"].add(numero)

    filas = []
    for _, item in analitica.items():
        filas.append(
            {
                "clave": item["clave"],
                "titulo": item["titulo"],
                "ventas": item["ventas"],
                "ingresos": item["ingresos"],
                "clientes_unicos": len(item["clientes"]),
            }
        )
    return sorted(filas, key=lambda x: x["ingresos"], reverse=True)


def obtener_numeros_segmentados(tipo_precio="", metodo_pago="", catalogo_clave=""):
    numeros = set()
    for venta in ventas_registradas:
        if tipo_precio and normalizar_texto(venta.get("tipo_precio", "")) != normalizar_texto(tipo_precio):
            continue
        if metodo_pago and normalizar_texto(venta.get("metodo_pago", "")) != normalizar_texto(metodo_pago):
            continue
        if catalogo_clave and normalizar_texto(venta.get("catalogo_clave", "")) != normalizar_texto(catalogo_clave):
            continue
        n = normalizar_numero_whatsapp(venta.get("numero_cliente", ""))
        if n:
            numeros.add(n)
    return numeros


def generar_excel_ventas(ventas=None):
    ventas = ventas if ventas is not None else ventas_registradas
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"
    ws.append([
        "Fecha aprobacion",
        "Cliente",
        "Producto",
        "Clave producto",
        "Metodo pago",
        "Tipo precio",
        "Variante bienvenida",
        "Valor venta",
        "Tipo archivo",
        "Nombre archivo",
        "Comprobante local",
        "Aprobado desde",
    ])

    for venta in ventas:
        ws.append([
            venta.get("fecha_aprobacion", ""),
            venta.get("numero_cliente", ""),
            venta.get("producto_titulo", ""),
            venta.get("catalogo_clave", ""),
            venta.get("metodo_pago", ""),
            venta.get("tipo_precio", ""),
            venta.get("bienvenida_variante", ""),
            int(venta.get("valor_venta", 0) or 0),
            venta.get("tipo_archivo", ""),
            venta.get("nombre_archivo", ""),
            venta.get("comprobante_local", ""),
            venta.get("aprobado_desde", ""),
        ])

    resumen = obtener_resumen_ventas_de(ventas)
    ws2 = wb.create_sheet("Resumen")
    ws2.append(["Metricas", "Valor"])
    ws2.append(["Total ventas", resumen["total"]])
    ws2.append(["Ventas precio normal", resumen["ventas_normal"]])
    ws2.append(["Ventas con descuento", resumen["ventas_descuento"]])
    ws2.append(["Ingresos total", resumen["ingresos_total"]])
    ws2.append(["Ingresos precio normal", resumen["ingresos_normal"]])
    ws2.append(["Ingresos con descuento", resumen["ingresos_descuento"]])

    ws3 = wb.create_sheet("Por producto")
    ws3.append(["Producto", "Clave", "Ventas", "Ingresos"])
    for r in obtener_resumen_por_producto(ventas):
        ws3.append([r["titulo"], r["clave"], r["ventas"], r["ingresos"]])

    ws4 = wb.create_sheet("AB Bienvenida")
    ws4.append(["Variante", "Ventas", "Porcentaje"])
    total_ventas = len(ventas)
    conteo_ab = {"a": 0, "b": 0, "sin_dato": 0}
    for venta in ventas:
        variante = str(venta.get("bienvenida_variante", "") or "").strip().lower()
        if variante not in {"a", "b"}:
            variante = "sin_dato"
        conteo_ab[variante] += 1
    for variante in ("a", "b", "sin_dato"):
        cantidad = conteo_ab[variante]
        porcentaje = f"{(cantidad / total_ventas * 100):.2f}%" if total_ventas else "0.00%"
        ws4.append([variante, cantidad, porcentaje])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def generar_pdf_consolidado_comprobantes(ventas=None):
    ventas = ventas if ventas is not None else ventas_registradas
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas
    except Exception as exc:
        raise RuntimeError("ReportLab no esta instalado. Ejecuta: pip install reportlab") from exc

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    resumen = obtener_resumen_ventas_de(ventas)
    y = height - 40
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(40, y, "Consolidado de ventas y comprobantes")
    y -= 22
    pdf.setFont("Helvetica", 10)
    pdf.drawString(40, y, f"Fecha de generacion: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 18
    pdf.drawString(40, y, f"Total ventas: {resumen['total']} | Normal: {resumen['ventas_normal']} | Descuento: {resumen['ventas_descuento']}")
    y -= 16
    pdf.drawString(40, y, f"Ingresos total: ${resumen['ingresos_total']}")
    y -= 26

    for idx, venta in enumerate(ventas, start=1):
        if y < 130:
            pdf.showPage()
            y = height - 40
            pdf.setFont("Helvetica", 10)

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(40, y, f"Venta {idx} - Cliente {venta.get('numero_cliente', '')}")
        y -= 14
        pdf.setFont("Helvetica", 9)
        pdf.drawString(40, y, f"Producto: {venta.get('producto_titulo', '')} ({venta.get('catalogo_clave', '')})")
        y -= 12
        pdf.drawString(40, y, f"Fecha: {venta.get('fecha_aprobacion', '')} | Metodo: {venta.get('metodo_pago', '')}")
        y -= 12
        pdf.drawString(40, y, f"Tipo precio: {venta.get('tipo_precio', '')} | Valor: ${int(venta.get('valor_venta', 0) or 0)}")
        y -= 12
        comprobante_local = venta.get("comprobante_local", "")
        pdf.drawString(40, y, f"Comprobante: {comprobante_local or 'No disponible'}")
        y -= 10

        if comprobante_local and Path(comprobante_local).exists():
            ext = Path(comprobante_local).suffix.lower()
            if ext in {".jpg", ".jpeg", ".png", ".webp"}:
                try:
                    img = ImageReader(comprobante_local)
                    pdf.drawImage(img, 40, max(40, y - 200), width=220, height=140, preserveAspectRatio=True, anchor='sw')
                    y -= 150
                except Exception:
                    y -= 8
            elif ext == ".pdf":
                pdf.drawString(40, y, "Comprobante PDF adjunto en archivo local (referencia incluida).")
                y -= 12
        y -= 10

    pdf.save()
    buffer.seek(0)
    return buffer


# ==========================================
# 🤖 WEBHOOK
# ==========================================
@app.route('/', methods=['GET'])
def inicio():
    return redirect(url_for('admin'))


# ==============================================================
# WEBHOOK REFACTORIZADO
# Prefijos validos: menu_, video_, pdf_, comprar_, descuento_,
#                   pagar_, aprobar_, rechazar_
# Textos naturales: hola, menu, info, info [producto], [numero]
# ==============================================================
@app.route('/webhook', methods=['GET', 'POST'])
def webhook():
    # ----- verificacion GET de Meta -----
    if request.method == 'GET':
        if request.args.get("hub.verify_token") == TOKEN_VERIFICACION:
            return request.args.get("hub.challenge", "")
        return "Error de autenticacion", 403

    # ----- procesamiento POST -----
    data = request.get_json(silent=True) or {}
    try:
        valor = data["entry"][0]["changes"][0]["value"]
        mensajes = valor.get("messages", [])
        if not mensajes:
            return jsonify({"status": "ignored"}), 200

        mensaje_info = mensajes[0]
        numero_cliente = mensaje_info.get("from", "")
        if not numero_cliente:
            return jsonify({"status": "ignored"}), 200

        tipo_mensaje = str(mensaje_info.get("type", "") or "").strip().lower()

        # ------------------------------------------------------------------
        # EXTRACCION ROBUSTA DEL PAYLOAD / TEXTO
        # Prioridad: button payload > interactive button_reply/list_reply > text
        # ------------------------------------------------------------------
        PREFIJOS_PERMITIDOS = (
            "menu_", "video_", "pdf_", "comprar_",
            "descuento_", "pagar_", "aprobar_", "rechazar_"
        )
        PAYLOADS_ESPECIALES = {
            "hola", "menu", "menu", "menu_principal", "ya_pague", "info"
        }

        texto_recibido = ""

        if tipo_mensaje == "text":
            # Texto libre del cliente
            texto_recibido = normalizar_texto(
                mensaje_info.get("text", {}).get("body", "")
            )

        elif tipo_mensaje in ["interactive", "button"]:
            # 1. CAPTURAR EL TEXTO DEL BOTÓN (Sea plantilla o interactivo)
            accion = ""
            if tipo_mensaje == "interactive":
                interactivo = mensaje_info.get("interactive", {}) or {}
                if interactivo.get("type") == "button_reply":
                    accion = (interactivo.get("button_reply", {}) or {}).get("title", "")
                elif interactivo.get("type") == "list_reply":
                    accion = (interactivo.get("list_reply", {}) or {}).get("title", "")
            elif tipo_mensaje == "button":
                accion = (mensaje_info.get("button", {}) or {}).get("text", "")

            accion = normalizar_texto(accion)
            tokens_accion = set(accion.replace("-", " ").replace("/", " ").split())

            if not accion:
                app.logger.info(
                    "[WEBHOOK] No se pudo extraer accion de boton/interactivo (numero=%s)",
                    numero_cliente
                )
                return "EVENT_RECEIVED", 200

            # 2. RECUPERAR EL PRODUCTO DE LA MEMORIA DEL BOT
            estado_actual = estados_clientes.setdefault(numero_cliente, {})
            id_prod = estado_actual.get("catalogo_activo")

            # Si presionan pagar, no necesitamos catalogo activo, solo procesar el pago
            if "nequi" in tokens_accion or "daviplata" in tokens_accion:
                metodo = "nequi" if "nequi" in accion else "daviplata"
                estados_clientes[numero_cliente]["metodo_pago"] = metodo
                estados_clientes[numero_cliente]["esperando_comprobante"] = True
                guardar_estado_runtime()
                enviar_datos_pago(numero_cliente, metodo)
                return "EVENT_RECEIVED", 200

            # Si no hay producto activo y no están pagando, enviamos al menú
            if not id_prod:
                enviar_mensaje_plantilla(numero_cliente, "plantilla_menu_general")
                return "EVENT_RECEIVED", 200

            producto = catalogo_productos.get(id_prod, {})

            # ==========================================
            # RAMAS DE BOTONES (Buscan por palabra clave)
            # ==========================================

            # RAMA: VIDEO
            if "video" in tokens_accion:
                link_video = producto.get("link_video", "")
                if link_video:
                    enviar_video(numero_cliente, link_video, "Video demo")

                enviar_mensaje_botones(numero_cliente, "¿Qué deseas hacer a continuación?", [
                    (f"pdf_{id_prod}", "Ver PDF demo"),
                    (f"comprar_{id_prod}", "Comprar"),
                    ("menu_principal", "Menú principal")
                ])
                return "EVENT_RECEIVED", 200

            # RAMA: PDF
            elif "pdf" in tokens_accion:
                link_pdf = producto.get("link_pdf", "")
                if link_pdf:
                    enviar_documento(numero_cliente, link_pdf, f"Demo_{id_prod}.pdf")

                enviar_mensaje_botones(numero_cliente, "¿Qué deseas hacer a continuación?", [
                    (f"video_{id_prod}", "Ver video demo"),
                    (f"comprar_{id_prod}", "Comprar"),
                    ("menu_principal", "Menú principal")
                ])
                return "EVENT_RECEIVED", 200

            # RAMA: COMPRAR
            elif "comprar" in tokens_accion:
                enviar_mensaje_plantilla(
                    numero_cliente,
                    "plantilla_compra_universal",
                    titulo=producto.get("titulo", "el producto"),
                    precio_normal=str(producto.get("precio_normal", ""))
                )
                return "EVENT_RECEIVED", 200

            # RAMA: MENÚ
            elif "menu" in tokens_accion:
                enviar_mensaje_plantilla(numero_cliente, "plantilla_menu_general")
                return "EVENT_RECEIVED", 200

        # Extraer adjunto (comprobante de pago)
        tipo_media, media_id, nombre_archivo = extraer_media_mensaje(mensaje_info)

        # Log de entrada
        contenido_log = (texto_recibido or tipo_media or "")[:200]
        print(f"[WEBHOOK] numero={numero_cliente} tipo={tipo_mensaje} contenido={contenido_log}")
        app.logger.info(
            "[WEBHOOK] numero=%s tipo_mensaje=%s contenido_recibido=%s",
            numero_cliente, tipo_mensaje, contenido_log
        )

        # Registrar actividad
        if numero_cliente not in estados_clientes:
            estados_clientes[numero_cliente] = {}
        estados_clientes[numero_cliente]["ultima_actividad"] = time.time()
        guardar_estado_runtime()

        # ------------------------------------------------------------------
        # FLUJO: COMPROBANTE (imagen o documento)
        # ------------------------------------------------------------------
        if tipo_mensaje in {"image", "document"}:
            estado_cliente = estados_clientes.get(numero_cliente, {})
            if estado_cliente.get("esperando_comprobante"):
                registrar_comprobante_pendiente(
                    numero_cliente, tipo_media or tipo_mensaje, media_id, nombre_archivo
                )
                enviar_mensaje(
                    numero_cliente,
                    "Comprobante recibido. Lo estamos validando y te confirmamos en breve."
                )
                actualizar_etapa_cliente(numero_cliente, "en_revision")
            elif estado_cliente.get("en_revision"):
                enviar_mensaje(
                    numero_cliente,
                    "Tu comprobante sigue en revision. Te avisamos cuando finalice."
                )
            else:
                enviar_mensaje(
                    numero_cliente,
                    "Primero selecciona un producto y pulsa Comprar para asociar correctamente el comprobante."
                )
            return jsonify({"status": "success"}), 200

        # ------------------------------------------------------------------
        # PARSEO DEL PAYLOAD EN accion + metodo_pago + objetivo
        # Formato: accion_metodo_idproducto  o  accion_idproducto
        # Ejemplos:
        #   video_finanzas      -> accion=video,  objetivo=finanzas
        #   pagar_nequi_ingles  -> accion=pagar,  metodo=nequi, objetivo=ingles
        #   aprobar_573001234567 -> accion=aprobar, objetivo=573001234567
        # ------------------------------------------------------------------
        accion = ""
        metodo_pago = ""
        objetivo_payload = ""

        if texto_recibido:
            partes = texto_recibido.split("_")
            if len(partes) >= 2:
                accion = partes[0]
                if accion == "pagar" and len(partes) >= 3:
                    metodo_pago = partes[1]
                    objetivo_payload = "_".join(partes[2:])
                else:
                    objetivo_payload = "_".join(partes[1:])

        # ------------------------------------------------------------------
        # HELPERS LOCALES
        # ------------------------------------------------------------------
        def _enviar_menu_general():
            """Envia plantilla_menu_general con la lista dinamica de productos."""
            lista_productos = construir_lista_menu_productos()
            enviar_mensaje_plantilla(
                numero_cliente, "plantilla_menu_general",
                lista_productos=lista_productos
            )

        def _enviar_info_producto(clave_producto, producto):
            """Actualiza estado y envia plantilla de info del producto.
            Usa el nombre exacto del campo plantilla_info_meta como nombre Meta.
            Si el campo existe, hace una llamada de plantilla directa con ese nombre.
            Si no existe, cae en texto plano.
            """
            actualizar_etapa_cliente(
                numero_cliente, "viendo_info",
                catalogo_activo=clave_producto,
                producto_titulo=producto.get("titulo", clave_producto)
            )
            # plantilla_info_meta contiene el NOMBRE APROBADO en Meta (ej: "info_finanzas_v1")
            plantilla_info_nombre = str(producto.get("plantilla_info_meta", "")).strip()
            if plantilla_info_nombre:
                # Enviar directamente como template con el nombre aprobado
                id_telefono = datos_bot.get("id_telefono", ID_TELEFONO_DEFAULT).strip()
                token_meta = datos_bot.get("token_meta", TOKEN_META_DEFAULT).strip()
                url = f"https://graph.facebook.com/v19.0/{id_telefono}/messages"
                headers = {
                    "Authorization": f"Bearer {token_meta}",
                    "Content-Type": "application/json"
                }
                payload_meta = {
                    "messaging_product": "whatsapp",
                    "to": numero_cliente,
                    "type": "template",
                    "template": {
                        "name": plantilla_info_nombre,
                        "language": {"code": "es"}
                    }
                }
                try:
                    resp = requests.post(url, headers=headers, json=payload_meta, timeout=30)
                    if not resp.ok:
                        # Fallback: enviar botones interactivos con info del producto
                        producto_id = producto.get("id", clave_producto)
                        titulo = producto.get("titulo", clave_producto)
                        precio = producto.get("precio_normal", "")
                        texto_fallback = f"{titulo}"
                        if precio:
                            texto_fallback += f"\nPrecio: ${precio}"
                        enviar_mensaje(numero_cliente, texto_fallback)
                        enviar_mensaje_botones(
                            numero_cliente,
                            "¿Qué deseas hacer?",
                            [
                                (f"video_{producto_id}", "Ver video"),
                                (f"pdf_{producto_id}", "Ver PDF"),
                                (f"comprar_{producto_id}", "Comprar")
                            ],
                            pie="Selecciona una opción"
                        )
                except Exception:
                    producto_id = producto.get("id", clave_producto)
                    enviar_mensaje_botones(
                        numero_cliente,
                        f"Información sobre {producto.get('titulo', clave_producto)}",
                        [
                            (f"video_{producto_id}", "Ver video"),
                            (f"pdf_{producto_id}", "Ver PDF"),
                            (f"comprar_{producto_id}", "Comprar")
                        ],
                        pie="Selecciona una opción"
                    )
            else:
                enviar_mensaje(
                    numero_cliente,
                    "Este producto no tiene plantilla de informacion configurada. Escribe menu para ver otras opciones."
                )

        def _limpiar_intencion_texto(valor):
            texto = str(valor or "").lower().strip()
            texto = unicodedata.normalize("NFKD", texto)
            texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
            texto = re.sub(r"[^a-z0-9\s]", " ", texto)
            texto = re.sub(r"\s+", " ", texto)
            return texto.strip()

        # ------------------------------------------------------------------
        # FLUJO PRINCIPAL
        # ------------------------------------------------------------------
        texto_limpio_intencion = _limpiar_intencion_texto(texto_recibido)
        clave_detectada_texto, producto_detectado_texto = (None, None)
        if tipo_mensaje == "text":
            clave_detectada_texto, producto_detectado_texto = buscar_producto_catalogo_por_texto(texto_recibido)

        # 0. DESCUENTO OCULTO: si hay catalogo_activo con oferta activa y el cliente responde
        # cualquier texto que no sea un payload estructurado, activar la plantilla de descuento.
        # Prioridad: ANTES de menú y de detección de producto por texto.
        estado_cliente_actual = estados_clientes.get(numero_cliente, {})
        if (
            tipo_mensaje == "text"
            and estado_cliente_actual.get("oferta_descuento_activa")
            and not accion  # No es un payload estructurado
            and not texto_limpio_intencion in {"hola", "menu_principal", "menu", "info", "inicio"}
        ):
            clave_descuento = estado_cliente_actual.get("catalogo_activo", "")
            producto_descuento = catalogo_productos.get(clave_descuento, {}) if clave_descuento else {}
            if producto_descuento:
                kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave_descuento, producto_descuento)
                enviar_mensaje_plantilla(numero_cliente, "plantilla_descuento_universal", **kwargs_tpl)
                producto_id_desc = producto_descuento.get("id", clave_descuento)
                enviar_mensaje_botones(
                    numero_cliente,
                    "Elige tu metodo de pago para activar el precio promocional.",
                    [
                        (f"pagar_nequi_{producto_id_desc}", "Pagar Nequi"),
                        (f"pagar_daviplata_{producto_id_desc}", "Pagar Daviplata"),
                        ("menu_principal", "Ver menu")
                    ],
                    pie="Pago con descuento"
                )
                return jsonify({"status": "success"}), 200

        # 1. MENU / HOLA / Saludos generales
        PALABRAS_SALUDO_MENU = {
            "hola", "holas", "buenas", "buenos dias", "buenas tardes", "buenas noches",
            "menu_principal", "menu", "info", "inicio", "empezar", "start", "hi", "hello"
        }
        if (
            accion == "menu"
            or texto_recibido in {"hola", "menu_principal"}
            or (
                tipo_mensaje == "text"
                and texto_limpio_intencion in PALABRAS_SALUDO_MENU
                and not producto_detectado_texto
            )
        ):
            datos_bot["visitas_info"] = int(datos_bot.get("visitas_info", 0)) + 1
            guardar_datos_bot()
            _enviar_menu_general()

        # 2. INFO PRODUCTO por texto (deteccion inteligente)
        elif tipo_mensaje == "text" and producto_detectado_texto:
            _enviar_info_producto(clave_detectada_texto, producto_detectado_texto)

        # 3. VER VIDEO
        elif accion == "video":
            clave, producto = resolver_producto_por_id_o_estado(numero_cliente, objetivo_payload)
            if not producto:
                enviar_mensaje(
                    numero_cliente,
                    "No encontre el producto del video. Escribe menu para continuar."
                )
            else:
                link_video = (
                    producto.get("link_video", "")
                    or datos_bot.get("link_video", "")
                )
                actualizar_etapa_cliente(
                    numero_cliente, "viendo_video",
                    catalogo_activo=clave,
                    producto_titulo=producto.get("titulo", clave)
                )
                if link_video:
                    # Soportar URLs externas (GitHub Raw, CDN, etc.) directamente
                    enviar_video(
                        numero_cliente, link_video,
                        f"Video demo: {producto.get('titulo', clave)}"
                    )
                    producto_id = producto.get("id", clave)
                    # Solo pdf/comprar/menu. SIN descuento en esta etapa.
                    enviar_mensaje_botones(
                        numero_cliente,
                        "¿Qué deseas hacer ahora?",
                        [
                            (f"pdf_{producto_id}", "Ver PDF"),
                            (f"comprar_{producto_id}", "Comprar"),
                            ("menu_principal", "Menú")
                        ],
                        pie="Siguiente paso"
                    )
                else:
                    producto_id = producto.get("id", clave)
                    enviar_mensaje(
                        numero_cliente,
                        "Este producto no tiene video configurado todavia."
                    )
                    enviar_mensaje_botones(
                        numero_cliente,
                        "¿Qué deseas hacer?",
                        [
                            (f"comprar_{producto_id}", "Comprar"),
                            ("menu_principal", "Menú")
                        ],
                        pie="Opciones"
                    )

        # 4. VER PDF  -> envia PDF y botones permitidos
        elif accion == "pdf":
            clave, producto = resolver_producto_por_id_o_estado(numero_cliente, objetivo_payload)
            if not producto:
                enviar_mensaje(
                    numero_cliente,
                    "No encontre el producto del PDF. Escribe menu para continuar."
                )
            else:
                link_pdf = (
                    producto.get("link_pdf", "")
                    or datos_bot.get("link_pdf_demo", "")
                )
                actualizar_etapa_cliente(
                    numero_cliente, "viendo_pdf",
                    catalogo_activo=clave,
                    producto_titulo=producto.get("titulo", clave)
                )
                producto_id = producto.get("id", clave)
                if link_pdf:
                    # Soportar URLs externas (GitHub Raw, CDN, etc.) directamente
                    enviar_documento(
                        numero_cliente, link_pdf,
                        f"Demo_{producto_id}.pdf"
                    )
                    # Solo video/comprar/menu. SIN descuento en esta etapa.
                    enviar_mensaje_botones(
                        numero_cliente,
                        "¿Qué deseas hacer ahora?",
                        [
                            (f"video_{producto_id}", "Ver Video"),
                            (f"comprar_{producto_id}", "Comprar"),
                            ("menu_principal", "Menú")
                        ],
                        pie="Siguiente paso"
                    )
                else:
                    enviar_mensaje(
                        numero_cliente,
                        "Este producto aun no tiene PDF demo configurado."
                    )
                    enviar_mensaje_botones(
                        numero_cliente,
                        "¿Qué deseas hacer?",
                        [
                            (f"comprar_{producto_id}", "Comprar"),
                            ("menu_principal", "Menú")
                        ],
                        pie="Opciones"
                    )

        # 5. COMPRAR -> plantilla_compra_universal + solo pagar_nequi/pagar_daviplata/menu
        elif accion == "comprar":
            clave, producto = resolver_producto_por_id_o_estado(numero_cliente, objetivo_payload)
            if not producto:
                enviar_mensaje(
                    numero_cliente,
                    "Primero elige un producto. Escribe menu para ver el catalogo."
                )
            else:
                actualizar_etapa_cliente(
                    numero_cliente, "interesado",
                    catalogo_activo=clave,
                    producto_titulo=producto.get("titulo", clave),
                    hizo_click_comprar=True,
                    oferta_descuento_activa=False,
                    tipo_precio="normal",
                    valor_venta=convertir_monto(producto.get("precio_normal", "0"))
                )
                producto_id = producto.get("id", clave)
                titulo_prod = producto.get("titulo", clave)
                precio_prod = producto.get("precio_normal", "")
                # Enviar plantilla_compra_universal si existe, o texto plano con precio
                kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave, producto)
                resp_tpl = enviar_mensaje_plantilla(
                    numero_cliente, "plantilla_compra_universal", **kwargs_tpl
                )
                # Si la plantilla no existe en Meta, enviar texto de respaldo
                if not resp_tpl or not resp_tpl.ok:
                    texto_compra = f"{titulo_prod}"
                    if precio_prod:
                        texto_compra += f"\nPrecio: ${precio_prod}"
                    enviar_mensaje(numero_cliente, texto_compra)
                # SOLO pagar_nequi/pagar_daviplata/menu. SIN descuento.
                enviar_mensaje_botones(
                    numero_cliente,
                    "Selecciona tu metodo de pago para activar tu compra.",
                    [
                        (f"pagar_nequi_{producto_id}", "Pagar Nequi"),
                        (f"pagar_daviplata_{producto_id}", "Pagar Daviplata"),
                        ("menu_principal", "Menú")
                    ],
                    pie="Activa tu compra"
                )

        # 6. DESCUENTO -> inyecta precios del catalogo en plantilla_descuento_universal
        # Solo se activa por payload explícito 'descuento_[id]' o por lógica de abandono.
        elif accion == "descuento":
            clave, producto = resolver_producto_por_id_o_estado(numero_cliente, objetivo_payload)
            if not producto:
                enviar_mensaje(
                    numero_cliente,
                    "No encontre el producto del descuento. Escribe menu para continuar."
                )
            else:
                actualizar_etapa_cliente(
                    numero_cliente, "interesado",
                    catalogo_activo=clave,
                    producto_titulo=producto.get("titulo", clave),
                    hizo_click_comprar=True,
                    oferta_descuento_activa=True,
                    tipo_precio="descuento",
                    valor_venta=convertir_monto(producto.get("precio_descuento", "0"))
                )
                # Inyectar titulo, precio_normal y precio_descuento desde el catalogo
                kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave, producto)
                enviar_mensaje_plantilla(
                    numero_cliente, "plantilla_descuento_universal", **kwargs_tpl
                )
                producto_id = producto.get("id", clave)
                enviar_mensaje_botones(
                    numero_cliente,
                    "Elige tu metodo de pago para activar el precio promocional.",
                    [
                        (f"pagar_nequi_{producto_id}", "Pagar Nequi"),
                        (f"pagar_daviplata_{producto_id}", "Pagar Daviplata"),
                        ("menu_principal", "Ver menu")
                    ],
                    pie="Pago con descuento"
                )

        # 7. PAGAR_[METODO]_[ID] -> datos bancarios + solicitud comprobante
        elif accion == "pagar":
            clave, producto = resolver_producto_por_id_o_estado(numero_cliente, objetivo_payload)
            if not producto or metodo_pago not in {"nequi", "daviplata", "llave"}:
                enviar_mensaje(
                    numero_cliente,
                    "No pude identificar el metodo de pago o el producto. Escribe menu para retomar."
                )
            else:
                # Actualizar catalogo activo antes de enviar datos
                estados_clientes.setdefault(numero_cliente, {}).update({
                    "catalogo_activo": clave,
                    "producto_titulo": producto.get("titulo", clave)
                })
                guardar_estado_runtime()
                enviar_datos_pago(numero_cliente, metodo_pago)

        # 8. METODO DE PAGO ESCRITO DIRECTAMENTE (texto libre compatible)
        elif texto_recibido in {"nequi", "daviplata", "llave"}:
            enviar_datos_pago(numero_cliente, texto_recibido)

        # 9. YA_PAGUE -> marcar esperando comprobante
        elif texto_recibido == "ya_pague":
            estado_actual = estados_clientes.get(numero_cliente, {})
            estados_clientes[numero_cliente] = dict(estado_actual)
            estados_clientes[numero_cliente].update({
                "esperando_comprobante": True,
                "metodo_pago": estado_actual.get("metodo_pago", "No definido"),
                "ultima_actividad": time.time(),
                "hizo_click_comprar": True,
                "recordatorio_descuento_enviado": True,
                "tipo_precio": estado_actual.get("tipo_precio", "normal"),
                "valor_venta": estado_actual.get("valor_venta", 0),
                "oferta_descuento_activa": estado_actual.get("oferta_descuento_activa", False)
            })
            guardar_estado_runtime()
            enviar_mensaje(
                numero_cliente,
                "Perfecto. Envia foto o PDF del comprobante por este chat para validar y liberar tu acceso."
            )
            clave = estado_actual.get("catalogo_activo", "")
            producto_ya_pague = catalogo_productos.get(clave, {}) if clave else {}
            actualizar_etapa_cliente(
                numero_cliente, "esperando_comprobante",
                catalogo_activo=clave,
                producto_titulo=producto_ya_pague.get(
                    "titulo", estado_actual.get("producto_titulo", "General")
                )
            )

        # 10. APROBAR_[NUMERO] -> solo admin; envia recursos finales al cliente
        elif accion == "aprobar":
            numero_admin = normalizar_numero_whatsapp(datos_bot.get("numero_admin", ""))
            if normalizar_numero_whatsapp(numero_cliente) != numero_admin:
                app.logger.warning(
                    "[WEBHOOK] aprobar rechazado: no es admin. numero=%s", numero_cliente
                )
                return jsonify({"status": "ignored"}), 200
            numero_cliente_objetivo = normalizar_numero_whatsapp(objetivo_payload)
            if numero_cliente_objetivo in solicitudes_comprobante:
                solicitud = solicitudes_comprobante.pop(numero_cliente_objetivo, {}).copy()
                guardar_estado_runtime()
                registrar_venta_aprobada(
                    numero_cliente_objetivo, solicitud, aprobado_desde="whatsapp"
                )
                if solicitud.get("catalogo_clave"):
                    estados_clientes.setdefault(numero_cliente_objetivo, {})[
                        "catalogo_activo"
                    ] = solicitud["catalogo_clave"]
                # Envia plantilla_entrega_final con link_drive_final y link_canal_whatsapp
                enviar_recursos_finales(numero_cliente_objetivo)
                enviar_mensaje(
                    numero_cliente,
                    f"Comprobante aprobado para {numero_cliente_objetivo}. Acceso enviado correctamente."
                )
            else:
                enviar_mensaje(
                    numero_cliente,
                    f"No hay solicitud pendiente para {numero_cliente_objetivo}."
                )

        # 11. RECHAZAR_[NUMERO] -> solo admin; avisa al cliente
        elif accion == "rechazar":
            numero_admin = normalizar_numero_whatsapp(datos_bot.get("numero_admin", ""))
            if normalizar_numero_whatsapp(numero_cliente) != numero_admin:
                app.logger.warning(
                    "[WEBHOOK] rechazar rechazado: no es admin. numero=%s", numero_cliente
                )
                return jsonify({"status": "ignored"}), 200
            numero_cliente_objetivo = normalizar_numero_whatsapp(objetivo_payload)
            if numero_cliente_objetivo in solicitudes_comprobante:
                solicitudes_comprobante.pop(numero_cliente_objetivo, None)
                estado_rechazo = dict(estados_clientes.get(numero_cliente_objetivo, {}))
                estado_rechazo.update({
                    "esperando_comprobante": True,
                    "en_revision": False,
                    "metodo_pago": estados_clientes.get(
                        numero_cliente_objetivo, {}
                    ).get("metodo_pago", "No definido")
                })
                estados_clientes[numero_cliente_objetivo] = estado_rechazo
                guardar_estado_runtime()
                enviar_mensaje(
                    numero_cliente_objetivo,
                    "Tu comprobante no pudo ser validado. Por favor reenvialo con mejor calidad."
                )
                enviar_mensaje(
                    numero_cliente,
                    f"Comprobante rechazado para {numero_cliente_objetivo}. El cliente fue notificado."
                )
            else:
                enviar_mensaje(
                    numero_cliente,
                    f"No hay solicitud pendiente para {numero_cliente_objetivo}."
                )

        # 12. DESCUENTO por texto explícito del cliente
        # Se activa cuando el cliente escribe exactamente palabras relacionadas con descuento.
        elif (
            tipo_mensaje == "text"
            and texto_limpio_intencion in {
                "descuento", "promo", "promocion", "oferta", "rebaja",
                "precio especial", "caro", "me haces descuento"
            }
        ):
            estado_actual = estados_clientes.get(numero_cliente, {})
            clave_estado = estado_actual.get("catalogo_activo", "")
            producto_estado = catalogo_productos.get(clave_estado, {}) if clave_estado else {}
            if producto_estado:
                actualizar_etapa_cliente(
                    numero_cliente, "interesado",
                    catalogo_activo=clave_estado,
                    producto_titulo=producto_estado.get("titulo", clave_estado),
                    hizo_click_comprar=True,
                    oferta_descuento_activa=True,
                    tipo_precio="descuento",
                    valor_venta=convertir_monto(producto_estado.get("precio_descuento", "0"))
                )
                kwargs_tpl = construir_kwargs_plantilla(numero_cliente, clave_estado, producto_estado)
                enviar_mensaje_plantilla(
                    numero_cliente, "plantilla_descuento_universal", **kwargs_tpl
                )
                producto_id = producto_estado.get("id", clave_estado)
                # Incluir menu_principal como tercera opción
                enviar_mensaje_botones(
                    numero_cliente,
                    "Elige tu metodo de pago para activar el precio promocional.",
                    [
                        (f"pagar_nequi_{producto_id}", "Pagar Nequi"),
                        (f"pagar_daviplata_{producto_id}", "Pagar Daviplata"),
                        ("menu_principal", "Ver menu")
                    ],
                    pie="Pago con descuento"
                )
            else:
                _enviar_menu_general()

        # 13. AGRADECIMIENTOS
        elif contiene_intencion(
            texto_recibido,
            {"gracias", "recibido", "recibida", "confirmo recibido", "confirmado"}
        ):
            enviar_mensaje(
                numero_cliente,
                datos_bot.get(
                    "mensaje_respuesta_gracias",
                    "Gracias por confirmar. Si necesitas algo mas, aqui estoy."
                )
            )

        # 14. FALLBACK -> buscar por texto libre en catalogo o devolver menu
        else:
            clave, producto = buscar_producto_catalogo_por_texto(texto_recibido)
            if producto:
                _enviar_info_producto(clave, producto)
            else:
                app.logger.info(
                    "[WEBHOOK] mensaje sin ruta reconocida numero=%s tipo=%s texto=%s -> enviando menu",
                    numero_cliente, tipo_mensaje, texto_recibido
                )
                _enviar_menu_general()

    except Exception as exc:  # noqa: BLE001
        # Capturar cualquier error sin romper el webhook;
        # Meta requiere 200 para no reintentar indefinidamente.
        app.logger.exception(
            "[WEBHOOK] Error no controlado procesando mensaje: %s", exc
        )

    return jsonify({"status": "success"}), 200


# ==========================================
# 🎛️ PANEL DE ADMINISTRACION
# ==========================================
PANEL_HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel de Control - Bot de Ventas</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <style>
        :root {
            --bg: #eef4f8;
            --card: #ffffff;
            --line: #d6e0e8;
            --text: #18313f;
            --muted: #537080;
            --brand: #0f9d7a;
            --brand-dark: #0a765b;
            --accent: #dff5ec;
        }

        * { box-sizing: border-box; }
        body {
            margin: 0;
            font-family: "Segoe UI", sans-serif;
            color: var(--text);
            background:
                radial-gradient(circle at top right, #dff5ec 0, transparent 25%),
                linear-gradient(135deg, #f8fbfd 0%, #eef4f8 100%);
            padding: 24px;
        }

        .wrapper {
            max-width: 1100px;
            margin: 0 auto;
            display: grid;
            gap: 20px;
        }

        .hero,
        .card {
            background: var(--card);
            border: 1px solid rgba(15, 157, 122, 0.12);
            border-radius: 20px;
            box-shadow: 0 18px 50px rgba(24, 49, 63, 0.08);
            position: relative;
            overflow: hidden;
        }

        .hero::before,
        .card::before {
            content: "";
            position: absolute;
            inset: -1px;
            border-radius: 20px;
            padding: 1px;
            background: conic-gradient(from 0deg, rgba(0, 255, 231, 0.0), rgba(0, 255, 231, 0.85), rgba(255, 170, 51, 0.85), rgba(0, 255, 231, 0.0));
            -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
            -webkit-mask-composite: xor;
            mask-composite: exclude;
            animation: electricRotate 8s linear infinite;
            opacity: 0.36;
            pointer-events: none;
        }

        @keyframes electricRotate {
            to { transform: rotate(360deg); }
        }

        .hero {
            padding: 28px;
            display: grid;
            gap: 14px;
        }

        .hero h1 {
            margin: 0;
            font-size: 2rem;
        }

        .hero p {
            margin: 0;
            color: var(--muted);
            line-height: 1.5;
        }

        .stats {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
        }

        .stat {
            background: var(--accent);
            border-radius: 14px;
            padding: 14px 18px;
            min-width: 220px;
        }

        .stat strong {
            display: block;
            font-size: 1.5rem;
            margin-top: 4px;
        }

        .card {
            padding: 24px;
        }

        .grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 18px;
        }

        .full {
            grid-column: 1 / -1;
        }

        label {
            display: block;
            font-weight: 700;
            margin-bottom: 6px;
        }

        input,
        textarea,
        select {
            width: 100%;
            border: 1px solid var(--line);
            border-radius: 12px;
            padding: 12px 14px;
            font: inherit;
            color: var(--text);
            background: #fbfdff;
        }

        textarea {
            min-height: 110px;
            resize: vertical;
        }

        .helper {
            display: block;
            margin-top: 6px;
            color: var(--muted);
            font-size: 0.92rem;
        }

        .actions {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-top: 10px;
        }

        .button {
            border: none;
            border-radius: 999px;
            background: var(--brand);
            color: white;
            padding: 14px 22px;
            font-size: 1rem;
            font-weight: 700;
            cursor: pointer;
        }

        .button:hover {
            background: var(--brand-dark);
        }

        .button-danger {
            background: #d54949;
        }

        .button-danger:hover {
            background: #b63636;
        }

        .alerta {
            padding: 14px 16px;
            border-radius: 14px;
            background: #e8fff6;
            color: #0d6b4f;
            border: 1px solid #bcebd8;
            margin-bottom: 18px;
        }

        .preview {
            border: 1px dashed var(--line);
            border-radius: 16px;
            padding: 18px;
            background: #f8fbfd;
            line-height: 1.5;
            white-space: pre-wrap;
        }

        .tabla {
            width: 100%;
            border-collapse: collapse;
        }

        .tabla th,
        .tabla td {
            border-bottom: 1px solid var(--line);
            text-align: left;
            padding: 10px 8px;
            vertical-align: top;
        }

        .tabla th {
            color: var(--muted);
            font-size: 0.92rem;
        }

        .acciones-pendiente {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }

        .acciones-pendiente form {
            margin: 0;
        }

        .btn-small {
            border: none;
            border-radius: 999px;
            padding: 8px 12px;
            font-size: 0.85rem;
            font-weight: 700;
            cursor: pointer;
            color: white;
            background: var(--brand);
        }

        .btn-small.button-danger {
            background: #d54949;
        }

        .btn-link {
            display: inline-block;
            text-decoration: none;
            background: #f0f6fb;
            color: var(--text);
            border: 1px solid var(--line);
            border-radius: 999px;
            padding: 8px 14px;
            font-weight: 700;
            font-size: 0.9rem;
            cursor: pointer;
        }

        .btn-link.active {
            background: var(--brand);
            color: #fff;
            border-color: var(--brand);
        }

        .tabs-nav {
            display: flex;
            flex-wrap: nowrap;
            overflow-x: auto;
            border: 1px solid #bdc9d3;
            background: #eef3f7;
            border-radius: 8px;
        }

        .tab-link {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            text-decoration: none;
            color: #11222f;
            border-right: 1px solid #bdc9d3;
            padding: 10px 16px;
            font-weight: 700;
            white-space: nowrap;
            background: #edf2f6;
            transition: 0.2s ease;
        }

        .tab-link:last-child {
            border-right: none;
        }

        .tab-link:hover {
            background: #e4edf4;
            color: #0b1822;
        }

        .tab-link.active {
            background: #ffffff;
            color: #0b1822;
            box-shadow: inset 0 -3px 0 #0f9d7a;
        }

        .tab-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 20px;
            height: 20px;
            padding: 0 6px;
            border-radius: 999px;
            background: #d54949;
            color: #fff;
            font-size: 0.75rem;
            font-weight: 700;
        }

        .panel-section.is-hidden {
            display: none;
        }

        .token-row {
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 8px;
            align-items: center;
        }

        .estado-media {
            border: 1px solid var(--line);
            border-radius: 12px;
            background: #f8fbfd;
            padding: 12px 14px;
            margin-top: 8px;
        }

        .estado-media p {
            margin: 4px 0;
        }

        .catalogo-item {
            cursor: grab;
        }

        .catalogo-item.dragging {
            opacity: 0.55;
            border-style: dashed;
        }

        .drag-handle {
            user-select: none;
        }

        .wizard {
            border: 1px solid var(--line);
            border-radius: 14px;
            background: #f6fbff;
            padding: 14px;
            margin-top: 10px;
        }

        .wizard-step.is-hidden {
            display: none;
        }

        .is-hidden {
            display: none !important;
        }

        .wizard-nav {
            display: flex;
            gap: 10px;
            margin-top: 10px;
            justify-content: space-between;
        }

        .catalog-preview {
            font-family: "Consolas", "Courier New", monospace;
            font-size: 0.85rem;
            background: #f3f8fc;
            border-color: #c5d6e3;
            max-height: 320px;
            overflow: auto;
        }

        .chat-simulacion {
            border: 1px solid var(--line);
            border-radius: 14px;
            background: linear-gradient(180deg, #f4fbff 0%, #eef6fc 100%);
            padding: 12px;
            max-height: 380px;
            overflow: auto;
            display: grid;
            gap: 8px;
        }

        .chat-burbuja {
            max-width: 88%;
            border-radius: 12px;
            padding: 8px 10px;
            font-size: 0.92rem;
            line-height: 1.4;
            white-space: pre-wrap;
            word-break: break-word;
        }

        .chat-burbuja.bot {
            justify-self: start;
            background: #ffffff;
            border: 1px solid #cddce8;
            color: #143248;
        }

        .chat-burbuja.cliente {
            justify-self: end;
            background: #dcf8c6;
            border: 1px solid #bfe7a1;
            color: #17321f;
        }

        .chat-etiqueta {
            display: block;
            font-size: 0.75rem;
            font-weight: 700;
            margin-bottom: 3px;
            opacity: 0.8;
            text-transform: uppercase;
        }

        .tutorial-fab {
            position: fixed;
            right: 20px;
            bottom: 20px;
            z-index: 1400;
            border: none;
            border-radius: 999px;
            background: #18313f;
            color: #fff;
            font-weight: 700;
            padding: 12px 18px;
            box-shadow: 0 10px 24px rgba(24, 49, 63, 0.25);
            cursor: pointer;
        }

        .tutorial-fab:hover {
            background: #0f9d7a;
        }

        .tutorial-overlay {
            position: fixed;
            inset: 0;
            z-index: 1500;
            background: rgba(7, 15, 22, 0.58);
        }

        .tutorial-card {
            position: fixed;
            z-index: 1600;
            width: min(420px, calc(100vw - 24px));
            background: #ffffff;
            border: 1px solid #c8d6df;
            border-radius: 16px;
            box-shadow: 0 24px 60px rgba(11, 24, 34, 0.28);
            padding: 16px;
        }

        .tutorial-card h3 {
            margin: 0 0 8px 0;
            font-size: 1.06rem;
        }

        .tutorial-card p {
            margin: 0;
            color: #35505f;
            line-height: 1.45;
        }

        .tutorial-meta {
            margin-top: 8px;
            color: #6a8594;
            font-size: 0.85rem;
        }

        .tutorial-actions {
            display: flex;
            justify-content: space-between;
            gap: 8px;
            margin-top: 14px;
        }

        .tutorial-focus {
            position: relative;
            z-index: 1550;
            outline: 3px solid #0f9d7a;
            outline-offset: 3px;
            border-radius: 10px;
            scroll-margin-top: 16px;
            background: #ffffff;
        }

        .meta-advanced summary {
            cursor: pointer;
            font-weight: 700;
            color: #173646;
        }

        .meta-advanced[open] {
            border: 1px solid var(--line);
            border-radius: 14px;
            padding: 12px;
            background: #f8fbfd;
        }

        .catalogo-carrete {
            display: grid;
            grid-auto-flow: column;
            grid-auto-columns: minmax(320px, 360px);
            gap: 14px;
            overflow-x: auto;
            padding-bottom: 6px;
            scroll-snap-type: x proximity;
        }

        .catalogo-slide {
            scroll-snap-align: start;
            border: 1px solid #d5e2e9;
            border-radius: 18px;
            background: linear-gradient(180deg, #ffffff 0%, #f5fafc 100%);
            padding: 16px;
            box-shadow: 0 12px 28px rgba(18, 43, 59, 0.08);
        }

        .catalogo-slide h3 {
            margin: 6px 0 8px 0;
            font-size: 1.06rem;
        }

        .catalogo-kicker {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 4px 10px;
            border-radius: 999px;
            background: #e7f7f1;
            color: #0d7b5f;
            font-size: 0.82rem;
            font-weight: 700;
        }

        .catalogo-precio {
            margin: 10px 0;
            font-size: 0.95rem;
            color: #2b4a5a;
        }

        .payload-grid {
            display: grid;
            gap: 8px;
            margin-top: 10px;
        }

        .payload-chip {
            border: 1px solid #d6e2ea;
            border-radius: 12px;
            padding: 10px 12px;
            background: #fff;
        }

        .payload-chip strong {
            display: block;
            margin-bottom: 3px;
        }

        .payload-chip code,
        .payload-inline code {
            display: inline-block;
            background: #edf4f7;
            color: #153849;
            border-radius: 6px;
            padding: 2px 6px;
            font-size: 0.84rem;
        }

        @media (max-width: 780px) {
            body { padding: 16px; }
            .grid { grid-template-columns: 1fr; }
        }

        /* =============================================
           MEJORAS VISUALES FASE 3 UNIFICADO
           Carrete WhatsApp-card + Guía Meta en Catálogos
           ============================================= */

        .catalogo-slide.whatsapp-card {
            padding: 0;
            overflow: hidden;
            background: linear-gradient(180deg, #dff5ea 0%, #eff8f4 100%);
        }

        .whatsapp-topbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 14px;
            background: #0f6a57;
            color: #fff;
            font-weight: 700;
            font-size: 0.92rem;
        }

        .whatsapp-body {
            padding: 14px;
            background:
                radial-gradient(circle at top left, rgba(255,255,255,0.55), transparent 28%),
                linear-gradient(180deg, #e8f5ee 0%, #f5fbf8 100%);
        }

        .wa-bubble {
            background: #ffffff;
            border-radius: 16px 16px 16px 6px;
            padding: 14px;
            box-shadow: 0 6px 16px rgba(15, 71, 54, 0.08);
            border: 1px solid rgba(15, 106, 87, 0.12);
        }

        .wa-bubble p {
            margin: 0 0 8px 0;
            color: #173646;
            line-height: 1.45;
        }

        .wa-buttons {
            display: grid;
            gap: 8px;
            margin-top: 12px;
        }

        .wa-button {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            padding: 10px 12px;
            border-radius: 12px;
            border: 1px solid #cde4da;
            background: #f8fffb;
            color: #0d7b5f;
            font-weight: 700;
        }

        .wa-button small {
            color: #45636f;
            font-weight: 600;
        }

        .catalogo-meta-footer {
            border-top: 1px solid #d8e7df;
            margin-top: 12px;
            padding-top: 12px;
        }

        .payload-copy-row {
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 8px;
            align-items: center;
            margin-top: 6px;
        }

        .payload-copy-row input {
            margin: 0;
            font-family: Consolas, monospace;
            background: #f3f8fb;
        }

        .meta-guia-catalogo {
            margin-top: 14px;
            border: 1px solid #d6e5ea;
            border-radius: 16px;
            background: #fbfdff;
            padding: 14px;
        }

        .meta-guia-grid {
            display: grid;
            gap: 12px;
            margin-top: 10px;
        }

        .meta-guia-card {
            border: 1px solid #dbe8ee;
            border-radius: 14px;
            padding: 12px;
            background: #fff;
        }

        .meta-guia-card h4 {
            margin: 0 0 6px 0;
            color: #173646;
        }

        .meta-guia-card p {
            margin: 0 0 8px 0;
            color: #47616d;
        }

    </style>
</head>
<body>
    <div class="wrapper">
        <section class="hero">
            <h1>Panel de control del bot</h1>
            <p>Servidor profesional organizado por pestañas. Configura mensajes generales, productos, comprobantes, reportes y difusión sin tocar código.</p>
            <nav class="tabs-nav" aria-label="Pestanas principales">
                <a href="#configuracion" class="tab-link active" data-tab="configuracion"><i class="bi bi-sliders"></i>1. Configuración</a>
                <a href="#pendientes" class="tab-link" data-tab="pendientes"><i class="bi bi-hourglass-split"></i>2. Pendientes <span class="tab-badge" id="tab-pendientes-count">{{ pendientes_total }}</span></a>
                <a href="#catalogos" class="tab-link" data-tab="catalogos"><i class="bi bi-journal-richtext"></i>3. Catálogos</a>
                <a href="#reportes" class="tab-link" data-tab="reportes"><i class="bi bi-bar-chart-line"></i>4. Reportes</a>
                <a href="#clientes" class="tab-link" data-tab="clientes"><i class="bi bi-people"></i>5. Clientes</a>
                <a href="#difusion" class="tab-link" data-tab="difusion"><i class="bi bi-broadcast"></i>6. Difusión</a>
                <a href="#prueba_meta" class="tab-link" data-tab="prueba_meta"><i class="bi bi-send-check"></i>7. Prueba Meta</a>
            </nav>
            <div class="actions" style="margin-top: 10px;">
                <button class="btn-small" type="button" onclick="iniciarTutorialPanel()">Reactivar guía</button>
            </div>
            <div class="stats">
                <div class="stat">
                    Solicitudes de info
                    <strong>{{ visitas_info }}</strong>
                </div>
            </div>
            <div class="estado-media">
                <p><strong>Estado rapido de archivos</strong></p>
                <p>Video demo: {% if estado_video.link %}Guardado{% else %}No configurado{% endif %}</p>
                <p>PDF demo: {% if estado_pdf.link %}Guardado{% else %}No configurado{% endif %}</p>
                <p>Recursos finales: {% if link_recursos_final %}Guardado{% else %}No configurado{% endif %}</p>
                <div class="actions" style="margin-top:8px;">
                    {% if estado_video.url_local %}
                    <a class="btn-link" href="{{ estado_video.url_local }}" target="_blank">Abrir video demo</a>
                    {% endif %}
                    {% if estado_pdf.url_local %}
                    <a class="btn-link" href="{{ estado_pdf.url_local }}" target="_blank">Abrir PDF demo</a>
                    {% endif %}
                </div>
            </div>
        </section>

        <section class="card panel-section" id="configuracion">
            {% if mensaje %}<div class="alerta">{{ mensaje }}</div>{% endif %}

            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="accion_form" value="configuracion">
                <input type="hidden" name="tab_destino" value="configuracion">
                <div class="grid">
                    <div class="full" id="tutorial-config-base">
                        <label for="id_telefono">ID de telefono de Meta</label>
                        <input id="id_telefono" type="text" name="id_telefono" value="{{ id_telefono }}" placeholder="1052565307946835">
                        <span class="helper">Puedes cambiar este valor cuando pases del entorno de pruebas al real.</span>
                    </div>
                    <div class="full">
                        <label for="token_meta">Token de acceso Meta</label>
                        <div class="token-row">
                            <input id="token_meta" type="password" name="token_meta" value="{{ token_meta }}" placeholder="EAA...">
                            <button class="btn-small" type="button" onclick="alternarToken()">Mostrar/Ocultar</button>
                        </div>
                        <span class="helper">Si cambias de app, numero o expira el token, actualizalo aqui y guarda.</span>
                    </div>
                    <div>
                        <label for="nequi">Nequi</label>
                        <input id="nequi" type="text" name="nequi" value="{{ nequi }}">
                    </div>
                    <div>
                        <label for="daviplata">Daviplata</label>
                        <input id="daviplata" type="text" name="daviplata" value="{{ daviplata }}">
                    </div>
                    <div>
                        <label for="llave">Llave</label>
                        <input id="llave" type="text" name="llave" value="{{ llave }}">
                    </div>
                    <div class="full" id="tutorial-plantillas-meta">
                        <details class="meta-advanced">
                            <summary>Editor avanzado de plantillas Meta</summary>
                            <span class="helper">Aquí puedes editar nombre, idioma, texto y variables persistentes de tus plantillas de Meta.</span>
                            <div class="actions" style="margin: 8px 0 12px 0;">
                                <a class="btn-link" href="{{ url_for('admin_exportar_plantillas_meta') }}" target="_blank">Exportar catálogo JSON</a>
                            </div>
                            <div class="estado-media" style="margin-top: 10px;">
                                {% for clave, plantilla in catalogo_plantillas_meta.items() %}
                                <div style="padding: 10px 0; border-bottom: 1px solid var(--line);">
                                    <p style="margin: 0 0 8px 0;"><strong>{{ clave }}</strong></p>
                                    <div class="grid">
                                        <div>
                                            <label style="margin-bottom:4px;">Nombre en Meta</label>
                                            <input type="text" name="tpl_nombre_{{ clave }}" value="{{ plantilla.nombre_meta }}" placeholder="nombre_plantilla_aprobada">
                                        </div>
                                        <div>
                                            <label style="margin-bottom:4px;">Idioma</label>
                                            <input type="text" name="tpl_idioma_{{ clave }}" value="{{ plantilla.idioma }}" placeholder="es">
                                        </div>
                                        <div class="full">
                                            <label style="margin-bottom:4px;">Texto de referencia para replicar en Meta</label>
                                            <textarea name="tpl_texto_{{ clave }}" style="min-height:92px;">{{ plantilla.texto_referencia }}</textarea>
                                            <small class="helper">Usa placeholders de Meta como {{ '{{1}}' }}, {{ '{{2}}' }}, etc., segun el orden de variables.</small>
                                        </div>
                                        <div class="full">
                                            <label style="margin-bottom:4px;">Variables ordenadas (separadas por coma)</label>
                                            <input type="text" name="tpl_variables_{{ clave }}" value="{{ plantilla.variables_ordenadas|join(', ') }}" placeholder="nombre_cliente, nombre_producto, precio">
                                            <small class="helper"><strong>Orden actual:</strong> {{ plantilla.variables_ordenadas|join(', ') }}</small>
                                        </div>
                                        <div class="full estado-media" style="margin-top: 6px;">
                                            <label style="margin-bottom:4px;">Vista previa para crear en Meta (copiar y pegar)</label>
                                            <textarea id="tpl_preview_meta_{{ clave }}" readonly style="min-height:86px;">{{ vista_previa_plantillas_meta[clave].texto_meta }}</textarea>
                                            <div class="actions" style="margin-top:8px;">
                                                <button class="btn-small" type="button" onclick="copiarTextoPorId('tpl_preview_meta_{{ clave }}')">Copiar vista Meta</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                {% endfor %}
                            </div>
                        </details>
                    </div>
                </div>

                <div class="full" style="margin-top:14px; padding: 14px; border: 1px solid var(--line); border-radius: 12px; background: #f8fbfd;">
                    <label style="font-weight:700; margin-bottom: 8px; display:block;">Control del tutorial de inicio</label>
                    <div style="display:flex; flex-wrap:wrap; gap:14px; align-items:center;">
                        <label style="font-weight:400; display:flex; align-items:center; gap:8px; cursor:pointer;">
                            <input type="checkbox" name="tutorial_habilitado" id="chk_tutorial_habilitado" {% if tutorial_habilitado %}checked{% endif %}>
                            Mostrar tutorial automáticamente al cargar el panel
                        </label>
                        <label style="font-weight:400; display:flex; align-items:center; gap:8px; cursor:pointer;">
                            <input type="checkbox" name="auto_iniciar_tutorial" id="chk_auto_iniciar" {% if auto_iniciar_tutorial %}checked{% endif %}>
                            Auto-iniciar tutorial (arrancar en el paso 1 al cargar)
                        </label>
                        <label style="font-weight:400; display:flex; align-items:center; gap:8px; cursor:pointer;">
                            <input type="checkbox" name="ocultar_tutorial_panel" id="chk_ocultar_tutorial" {% if ocultar_tutorial_panel %}checked{% endif %}>
                            Ocultar botón flotante de tutorial
                        </label>
                    </div>
                    <span class="helper" style="margin-top:6px; display:block;">Si desmarcar "Mostrar tutorial automáticamente" deshabilita el arranque al recargar. Puedes volver a abrirlo manualmente con el botón "Reactivar guía" de arriba o con el botón flotante.</span>
                </div>

                <div class="actions">
                    <button class="button" type="submit">Guardar configuración general</button>
                    <button class="btn-small button-danger" type="submit" name="reset_plantillas_meta" value="1" onclick="return confirm('Esto restaura la configuracion de plantillas Meta (nombre, idioma, texto y variables) a valores por defecto. ¿Continuar?');">Restaurar configuración de plantillas por defecto</button>
                </div>
            </form>
        </section>

        <section class="card panel-section is-hidden" id="catalogos">
            <h2>Catálogos y plantillas Meta por producto</h2>
            <p class="helper">Aquí configuras productos, precios, URLs de demo y las plantillas Meta. <strong>Precio normal</strong> es la variable {{2}} en plantilla_compra_universal; <strong>precio descuento</strong> es {{3}} en plantilla_descuento_universal. Si cambias estos valores aquí, el bot los usa automáticamente sin necesidad de editar Meta. El campo verde <em>"Texto plantilla única de información"</em> es el cuerpo que debes copiar a Meta para la plantilla exclusiva del producto.</p>
            <div class="aviso-render" style="background:#fff8e1; border:1px solid #f0c040; border-radius:10px; padding:12px 16px; margin-bottom:16px; font-size:0.93rem;">
                <strong>⚠️ Estás en Render (entorno cloud):</strong> No subas archivos locales aquí. Para video y PDF demo, pega una <strong>URL pública HTTPS</strong> (GitHub Raw, Google Drive directo, CDN, etc.) para que Meta pueda descargar el archivo correctamente.
            </div>
            <div class="estado-media" style="margin-bottom: 12px;">
                <label>Carrete visual del catálogo y botones Meta</label>
                <small class="helper">Este carrete simula la tarjeta que verías en WhatsApp. Cada producto muestra el mensaje, los botones visibles y el payload exacto que debe existir en Meta. Usa los botones "Copiar" para llevarlo directo a la consola de Meta.</small>
                <div class="catalogo-carrete" style="margin-top:10px;">
                    {% for item in catalogo_items %}
                    <article class="catalogo-slide whatsapp-card">
                        <div class="whatsapp-topbar">
                            <span>{{ item.titulo }}</span>
                            <span>Opción {{ item.opcion }}</span>
                        </div>
                        <div class="whatsapp-body">
                            <span class="catalogo-kicker">{{ item.id }} · {{ item.plantilla_info_meta }}</span>
                            <div class="wa-bubble" style="margin-top:10px;">
                                <p><strong>{{ item.titulo }}</strong></p>
                                <p>Precio: <strong>${{ item.precio_normal }}</strong></p>
                                <p style="color:#537080; font-size:0.88rem;">Plantilla Meta: <code>{{ item.plantilla_info_meta }}</code></p>
                                <div class="wa-buttons">
                                    <div class="wa-button"><span>Ver video</span><small>video_{{ item.id }}</small></div>
                                    <div class="wa-button"><span>Ver PDF</span><small>pdf_{{ item.id }}</small></div>
                                    <div class="wa-button"><span>Comprar</span><small>comprar_{{ item.id }}</small></div>
                                </div>
                                <p style="margin-top:10px; color:#537080; font-size:0.82rem;">Después de Comprar &rarr; pagar_nequi / pagar_daviplata</p>
                            </div>
                            <div class="catalogo-meta-footer">
                                <div class="payload-inline"><strong>Plantilla info Meta:</strong> <code>{{ item.plantilla_info_meta }}</code></div>
                                <div class="payload-grid">
                                    <div class="payload-chip">
                                        <strong>Video</strong>
                                        <div class="payload-copy-row">
                                            <input id="payload_video_{{ loop.index0 }}" type="text" readonly value="video_{{ item.id }}">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('payload_video_{{ loop.index0 }}')">Copiar</button>
                                        </div>
                                    </div>
                                    <div class="payload-chip">
                                        <strong>PDF</strong>
                                        <div class="payload-copy-row">
                                            <input id="payload_pdf_{{ loop.index0 }}" type="text" readonly value="pdf_{{ item.id }}">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('payload_pdf_{{ loop.index0 }}')">Copiar</button>
                                        </div>
                                    </div>
                                    <div class="payload-chip">
                                        <strong>Comprar</strong>
                                        <div class="payload-copy-row">
                                            <input id="payload_comprar_{{ loop.index0 }}" type="text" readonly value="comprar_{{ item.id }}">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('payload_comprar_{{ loop.index0 }}')">Copiar</button>
                                        </div>
                                    </div>
                                    <div class="payload-chip">
                                        <strong>Descuento</strong>
                                        <div class="payload-copy-row">
                                            <input id="payload_descuento_{{ loop.index0 }}" type="text" readonly value="descuento_{{ item.id }}">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('payload_descuento_{{ loop.index0 }}')">Copiar</button>
                                        </div>
                                    </div>
                                    <div class="payload-chip">
                                        <strong>Pagar Nequi</strong>
                                        <div class="payload-copy-row">
                                            <input id="payload_nequi_{{ loop.index0 }}" type="text" readonly value="pagar_nequi_{{ item.id }}">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('payload_nequi_{{ loop.index0 }}')">Copiar</button>
                                        </div>
                                    </div>
                                    <div class="payload-chip">
                                        <strong>Pagar Daviplata</strong>
                                        <div class="payload-copy-row">
                                            <input id="payload_daviplata_{{ loop.index0 }}" type="text" readonly value="pagar_daviplata_{{ item.id }}">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('payload_daviplata_{{ loop.index0 }}')">Copiar</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </article>
                    {% endfor %}
                </div>
            </div>
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="accion_form" value="catalogo">
                <input type="hidden" name="tab_destino" value="catalogos">
                <div class="grid">
                    <div class="full">
                        <label>Catálogo de productos (editor visual)</label>
                        <div class="actions" style="margin-top: 6px; margin-bottom: 8px;">
                            <select id="selector_catalogo" onchange="enfocarCatalogo(this.value)">
                                <option value="">Ir a producto...</option>
                                {% for item in catalogo_items %}
                                <option value="catalogo_item_{{ loop.index0 }}">{{ item.titulo }} ({{ item.key }})</option>
                                {% endfor %}
                            </select>
                            <button class="btn-small" type="button" onclick="agregarProductoCatalogo()">+ Agregar producto</button>
                            <button class="btn-small" type="button" onclick="mostrarWizardProducto()">Asistente 3 pasos</button>
                            <button class="button" type="submit">Guardar cambios de catálogos</button>
                        </div>
                        <div id="wizard-producto" class="wizard is-hidden">
                            <h3 style="margin-top:0;">Asistente de producto (3 pasos)</h3>
                            <div id="wizard-step-1" class="wizard-step">
                                <div class="grid">
                                    <div>
                                        <label for="wizard_key">Clave interna</label>
                                        <input id="wizard_key" type="text" placeholder="ej: matematica">
                                    </div>
                                    <div>
                                        <label for="wizard_titulo">Título</label>
                                        <input id="wizard_titulo" type="text" placeholder="Pack de Matemática">
                                    </div>
                                    <div>
                                        <label for="wizard_precio">Precio</label>
                                        <input id="wizard_precio" type="text" placeholder="20.000">
                                    </div>
                                    <div>
                                        <label for="wizard_precio_desc">Precio descuento</label>
                                        <input id="wizard_precio_desc" type="text" placeholder="15.000">
                                    </div>
                                    <div class="full">
                                        <label for="wizard_palabras">Palabras clave</label>
                                        <input id="wizard_palabras" type="text" placeholder="matematica, algebra, fracciones">
                                    </div>
                                </div>
                            </div>
                            <div id="wizard-step-2" class="wizard-step is-hidden">
                                <div class="grid">
                                    <div class="full">
                                        <label for="wizard_video">Link video demo</label>
                                        <input id="wizard_video" type="text" placeholder="https://... o /media/...">
                                    </div>
                                    <div class="full">
                                        <label for="wizard_pdf">Link PDF demo</label>
                                        <input id="wizard_pdf" type="text" placeholder="https://... o /media/...">
                                    </div>
                                </div>
                                <span class="helper" style="color:#8a5e2a;">En Render, usa siempre URLs públicas HTTPS (GitHub Raw, CDN, Google Drive directo). No subas archivos locales.</span>
                            </div>
                            <div id="wizard-step-3" class="wizard-step is-hidden">
                                <div class="grid">
                                    <div class="full">
                                        <label for="wizard_texto_info_meta">Texto plantilla única de información para Meta <span style="color:#2e7d32; font-size:0.82rem;">(cuerpo de info_[id]_v1)</span></label>
                                        <small class="helper" style="color:#2e7d32;">Este texto se copia en Meta Business Manager como cuerpo de la plantilla única del producto. Sin variables {{N}} en el body.</small>
                                        <textarea id="wizard_texto_info_meta" style="min-height:120px;" placeholder="Escribe el texto persuasivo del producto..."></textarea>
                                    </div>
                                    <div class="full">
                                        <label for="wizard_msg_info">Mensaje info (respaldo texto plano)</label>
                                        <textarea id="wizard_msg_info" placeholder="Mensaje inicial del producto"></textarea>
                                    </div>
                                    <div class="full">
                                        <label for="wizard_msg_demo">Mensaje después demo</label>
                                        <textarea id="wizard_msg_demo" placeholder="Mensaje de proceso y pago"></textarea>
                                    </div>
                                    <div class="full">
                                        <label for="wizard_msg_desc">Mensaje descuento</label>
                                        <textarea id="wizard_msg_desc" placeholder="Mensaje de promoción"></textarea>
                                    </div>
                                </div>
                            </div>
                            <div class="wizard-nav">
                                <button class="btn-small" type="button" id="wizard_prev" onclick="wizardPrev()">Anterior</button>
                                <div style="display:flex; gap:8px;">
                                    <button class="btn-small" type="button" onclick="ocultarWizardProducto()">Cancelar</button>
                                    <button class="btn-small" type="button" id="wizard_next" onclick="wizardNext()">Siguiente</button>
                                    <button class="button button-danger is-hidden" type="button" id="wizard_finish" onclick="crearProductoDesdeWizard()">Crear producto</button>
                                </div>
                            </div>
                        </div>
                        <div id="catalogo-editor" class="full">
                            {% for item in catalogo_items %}
                            <div class="estado-media catalogo-item" id="catalogo_item_{{ loop.index0 }}" data-index="{{ loop.index0 }}" draggable="true">
                                <div class="actions" style="margin-top:0; margin-bottom:8px;">
                                    <span class="btn-link drag-handle" title="Arrastra para reordenar">Mover</span>
                                    <strong>Producto {{ loop.index }}</strong>
                                    <button class="btn-small" type="button" onclick="duplicarProductoCatalogo(this)">Duplicar</button>
                                    <button class="btn-small button-danger" type="button" onclick="eliminarProductoCatalogo(this)">Eliminar</button>
                                </div>
                                <input type="hidden" name="catalog_index" value="{{ loop.index0 }}">
                                <input type="hidden" name="catalog_original_key" value="{{ item.key }}">
                                <div class="grid">
                                    <div>
                                        <label>Clave interna</label>
                                        <input type="text" name="catalog_key" value="{{ item.key }}" placeholder="ej: finanzas">
                                    </div>
                                    <div>
                                        <label>ID / alias Meta</label>
                                        <input type="text" name="catalog_id" value="{{ item.id or item.key }}" placeholder="ej: finanzas">
                                    </div>
                                    <div>
                                        <label>Opción del carrete</label>
                                        <input type="text" name="catalog_opcion" value="{{ item.opcion or loop.index }}" placeholder="1">
                                    </div>
                                    <div>
                                        <label>Título del producto</label>
                                        <input type="text" name="catalog_titulo" value="{{ item.titulo }}" placeholder="Titulo visible al cliente">
                                    </div>
                                    <div>
                                        <label>Plantilla Meta info</label>
                                        <input type="text" name="catalog_plantilla_meta" value="{{ item.plantilla_info_meta or ('info_' ~ (item.id or item.key) ~ '_v1') }}" placeholder="info_finanzas_v1">
                                    </div>
                                    <div>
                                        <label>Precio normal <span style="font-weight:400; color:#0d7b5f; font-size:0.82rem;">→ variable {{2}} en plantilla_compra_universal</span></label>
                                        <input type="text" name="catalog_precio" value="{{ item.precio_normal }}" placeholder="15000">
                                        <small class="helper" style="color:#0d7b5f;">Si cambias este valor, las plantillas universales lo usan automáticamente. No necesitas editar Meta.</small>
                                    </div>
                                    <div>
                                        <label>Precio descuento <span style="font-weight:400; color:#c0392b; font-size:0.82rem;">→ variable {{3}} en plantilla_descuento_universal</span></label>
                                        <input type="text" name="catalog_precio_descuento" value="{{ item.precio_descuento }}" placeholder="9000">
                                        <small class="helper" style="color:#c0392b;">Oculto al cliente hasta abandono o solicitud explícita. No aparece en la plantilla de compra normal.</small>
                                    </div>
                                    <div>
                                        <label>Palabras clave</label>
                                        <input type="text" name="catalog_palabras" value="{{ item.palabras_clave|join(', ') }}" placeholder="finanzas, ahorro, dinero">
                                    </div>
                                    <div>
                                        <label>URL pública del video demo <span style="color:#c0392b; font-size:0.82rem;">(GitHub Raw, CDN o HTTPS directo)</span></label>
                                        <input type="text" name="catalog_link_video" value="{{ item.link_video }}" placeholder="https://raw.githubusercontent.com/tu_usuario/repo/main/video_demo.mp4">
                                        <small class="helper" style="color:#8a5e2a;">En Render no subas archivos locales aquí. Pega una URL pública HTTPS para que Meta pueda descargar el archivo.</small>
                                    </div>
                                    <div>
                                        <label>URL pública del PDF demo <span style="color:#c0392b; font-size:0.82rem;">(GitHub Raw, CDN o HTTPS directo)</span></label>
                                        <input type="text" name="catalog_link_pdf" value="{{ item.link_pdf }}" placeholder="https://raw.githubusercontent.com/tu_usuario/repo/main/demo.pdf">
                                        <small class="helper" style="color:#8a5e2a;">En Render no subas archivos locales aquí. Pega una URL pública HTTPS para que Meta pueda descargar el archivo.</small>
                                    </div>
                                    <div>
                                        <label>Link de Drive final</label>
                                        <input type="text" name="catalog_link_entrega" value="{{ item.link_drive_final or '' }}" placeholder="https://drive.google.com/...">
                                    </div>
                                    {# Inputs de subida local ocultados: no usar en Render. El backend los soporta pero no se muestran en UI. #}
                                    <input type="hidden" name="catalog_archivo_video_{{ loop.index0 }}" value="">
                                    <input type="hidden" name="catalog_archivo_pdf_{{ loop.index0 }}" value="">

                                    {# Sección: texto_info_meta editable - CAMPO PRINCIPAL V3 #}
                                    <div class="full" style="margin-top:16px; padding:14px; background:#e8f5e9; border:2px solid #2e7d32; border-radius:12px;">
                                        <strong style="font-size:1rem; color:#1b5e20;">📝 Texto plantilla única de información para Meta</strong>
                                        <p style="margin:6px 0 6px 0; font-size:0.86rem; color:#1b5e20;">Este es el cuerpo EXACTO de la plantilla <code>{{ item.plantilla_info_meta or ('info_' ~ (item.id or item.key) ~ '_v1') }}</code> que debes crear en Meta Business Manager. Redáctalo aquí, guarda, y luego cópialo directo a Meta.</p>
                                        <p style="margin:0 0 8px 0; font-size:0.83rem; color:#388e3c;">⚠️ Esta plantilla NO usa variables <code>&#123;&#123;1&#125;&#125;</code> en el cuerpo: el texto es fijo por producto. Los botones sí llevan payloads dinámicos.</p>
                                        <textarea name="catalog_texto_info_meta" id="tpl_texto_info_meta_{{ loop.index0 }}" style="min-height:160px; width:100%; font-size:0.88rem; background:#f1f8f2; border:1px solid #81c784; border-radius:8px; padding:10px;">{{ item.texto_info_meta or '' }}</textarea>
                                        <div class="actions" style="margin-top:6px;">
                                            <button class="btn-small" type="button" onclick="copiarTextoPorId('tpl_texto_info_meta_{{ loop.index0 }}')">Copiar texto para Meta</button>
                                        </div>
                                    </div>

                                    {# Sección: Plantillas Meta para este producto #}
                                    <div class="full" style="margin-top:12px; padding:14px; background:#eaf5ff; border:1.5px solid #2196f3; border-radius:10px;">
                                        <strong style="font-size:1rem; color:#1565c0;">&#x1F4CB; Plantillas Meta que debes crear para: {{ item.titulo }}</strong>
                                        <p style="margin:6px 0 6px 0; font-size:0.88rem; color:#1a5276;">Estas plantillas deben crearse en <strong>Meta Business Manager &gt; WhatsApp &gt; Plantillas</strong>.</p>
                                        <div style="background:#fff8e1; border:1px solid #ffc107; border-radius:8px; padding:10px; margin-bottom:12px; font-size:0.84rem; color:#6d4c00;">
                                            <strong>💡 ¿Por qué las plantillas universales usan variables {{1}}, {{2}}?</strong><br>
                                            Las plantillas universales (compra, descuento, cuentas, entrega) usan variables porque aplican a <em>todos los productos</em>. El bot reemplaza esas variables con los datos del catálogo al enviarlas. <strong>Si cambias el precio o el nombre en el catálogo, NO necesitas editar la plantilla en Meta.</strong>
                                        </div>

                                        <div style="margin-bottom:12px; padding:10px; background:#fff; border-radius:8px; border:1px solid #90caf9;">
                                            <p style="margin:0 0 4px 0; font-weight:700; color:#1565c0;">A) Plantilla única del producto: <code>{{ item.plantilla_info_meta or ('info_' ~ (item.id or item.key) ~ '_v1') }}</code></p>
                                            <p style="margin:0 0 4px 0; font-size:0.84rem; color:#537080;">Categoría sugerida: <strong>MARKETING</strong> &bull; Idioma: <strong>es</strong></p>
                                            <p style="margin:0 0 4px 0; font-size:0.84rem; color:#537080;">Precio normal actual: <strong>${{ item.precio_normal }}</strong> &bull; Precio descuento: <strong>${{ item.precio_descuento }}</strong></p>
                                            <p style="margin:0 0 4px 0; font-size:0.83rem; color:#1b5e20;"><em>El texto de arriba (campo verde) es el cuerpo de esta plantilla. Si cambias precio o título, edita el campo verde y guarda.</em></p>
                                            <textarea id="tpl_producto_texto_{{ loop.index0 }}" readonly style="min-height:110px; width:100%; font-size:0.85rem; background:#f8fbff;">{{ item.texto_info_meta or ('Hola, te comparto informacion de ' ~ item.titulo ~ '.

Precio: $' ~ item.precio_normal ~ '
Descuento especial: $' ~ item.precio_descuento) }}</textarea>
                                            <div class="actions" style="margin-top:6px;">
                                                <button class="btn-small" type="button" onclick="sincronizarTextoInfoMeta({{ loop.index0 }})">Sincronizar desde campo editable</button>
                                                <button class="btn-small" type="button" onclick="copiarTextoPorId('tpl_producto_texto_{{ loop.index0 }}')">Copiar texto plantilla</button>
                                            </div>
                                            <p style="margin:8px 0 4px 0; font-size:0.84rem; color:#537080;"><strong>Botones obligatorios en Meta (Quick Reply):</strong></p>
                                            <div class="payload-grid">
                                                <div class="payload-chip" style="background:#e8f5e9; border-color:#66bb6a;"><strong>Botón 1 – Ver video demo</strong><div class="payload-copy-row"><input id="btn1_video_{{ loop.index0 }}" type="text" readonly value="video_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('btn1_video_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip" style="background:#e8f5e9; border-color:#66bb6a;"><strong>Botón 2 – Ver PDF demo</strong><div class="payload-copy-row"><input id="btn2_pdf_{{ loop.index0 }}" type="text" readonly value="pdf_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('btn2_pdf_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip" style="background:#e8f5e9; border-color:#66bb6a;"><strong>Botón 3 – Comprar</strong><div class="payload-copy-row"><input id="btn3_comprar_{{ loop.index0 }}" type="text" readonly value="comprar_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('btn3_comprar_{{ loop.index0 }}')">Copiar</button></div></div>
                                            </div>
                                        </div>

                                        <div style="margin-bottom:12px; padding:10px; background:#fff; border-radius:8px; border:1px solid #90caf9;">
                                            <p style="margin:0 0 6px 0; font-weight:700; color:#1565c0;">B) Todos los payloads del producto <code>{{ item.id or item.key }}</code></p>
                                            <div class="payload-grid">
                                                <div class="payload-chip"><strong>video</strong><div class="payload-copy-row"><input id="all_video_{{ loop.index0 }}" type="text" readonly value="video_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('all_video_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip"><strong>pdf</strong><div class="payload-copy-row"><input id="all_pdf_{{ loop.index0 }}" type="text" readonly value="pdf_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('all_pdf_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip"><strong>comprar</strong><div class="payload-copy-row"><input id="all_comprar_{{ loop.index0 }}" type="text" readonly value="comprar_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('all_comprar_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip"><strong>pagar_nequi</strong><div class="payload-copy-row"><input id="all_nequi_{{ loop.index0 }}" type="text" readonly value="pagar_nequi_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('all_nequi_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip"><strong>pagar_daviplata</strong><div class="payload-copy-row"><input id="all_daviplata_{{ loop.index0 }}" type="text" readonly value="pagar_daviplata_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('all_daviplata_{{ loop.index0 }}')">Copiar</button></div></div>
                                                <div class="payload-chip" style="border-color:#e8a85a; background:#fffbf3;"><strong>descuento <small style="color:#8a5e2a;">(solo rescate/solicitud explícita)</small></strong><div class="payload-copy-row"><input id="all_descuento_{{ loop.index0 }}" type="text" readonly value="descuento_{{ item.id or item.key }}"><button class="btn-small" type="button" onclick="copiarTextoPorId('all_descuento_{{ loop.index0 }}')">Copiar</button></div><small style="color:#8a5e2a; font-size:0.78rem;">No va en la plantilla de info. Solo se activa por abandono o solicitud explícita. La plantilla de compra NO ofrece descuento.</small></div>
                                            </div>
                                        </div>

                                        <div style="padding:10px; background:#fff; border-radius:8px; border:1px solid #90caf9;">
                                            <p style="margin:0 0 6px 0; font-weight:700; color:#1565c0;">C) Guía completa del producto (para copiar)</p>
                                            <textarea id="guia_producto_{{ loop.index0 }}" readonly style="min-height:150px; width:100%; font-size:0.82rem; background:#f8fbff;">PLANTILLA: {{ item.plantilla_info_meta or ('info_' ~ (item.id or item.key) ~ '_v1') }}
Producto: {{ item.titulo }}
Precio normal: ${{ item.precio_normal }} | Descuento: ${{ item.precio_descuento }}
URL video: {{ item.link_video or '(pegar URL publica HTTPS)' }}
URL PDF: {{ item.link_pdf or '(pegar URL publica HTTPS)' }}

BOTONES (Quick Reply en Meta):
  Boton 1 | Texto: Ver video demo | Payload: video_{{ item.id or item.key }}
  Boton 2 | Texto: Ver PDF demo   | Payload: pdf_{{ item.id or item.key }}
  Boton 3 | Texto: Comprar        | Payload: comprar_{{ item.id or item.key }}

OTROS PAYLOADS (no van en plantilla info):
  pagar_nequi_{{ item.id or item.key }}
  pagar_daviplata_{{ item.id or item.key }}
  descuento_{{ item.id or item.key }} <- solo rescate/abandono</textarea>
                                            <div class="actions" style="margin-top:6px;">
                                                <button class="btn-small" type="button" onclick="copiarTextoPorId('guia_producto_{{ loop.index0 }}')">Copiar guía completa</button>
                                            </div>
                                        </div>
                                    </div>

                                    {# Acordeón: textos de referencia / respaldo #}
                                    <div class="full" style="margin-top:12px;">
                                        <details>
                                            <summary style="cursor:pointer; color:#537080; font-size:0.9rem;">&#x1F4AC; Textos de referencia / respaldo (bienvenida A/B y mensajes planos)</summary>
                                            <p style="font-size:0.82rem; color:#7f8c8d; margin:8px 0;">Estos textos son de respaldo para flujos de texto plano (sin plantilla Meta activa). La pantalla principal prioriza las plantillas Meta de arriba.</p>
                                            <div class="grid" style="margin-top:8px;">
                                            <div class="full">
                                                <label>Bienvenida A <span style="font-weight:400; color:#7f8c8d;">(texto de respaldo si plantilla Meta no está activa)</span></label>
                                                <small class="helper">Variante A/B: el bot alterna automáticamente por cliente. Usa {'{'}titulo{'}'} y {'{'}precio{'}'}.</small>
                                                <div class="actions" style="margin-top:4px; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'bienvenida_a')">Restaurar default A</button>
                                                </div>
                                                <textarea name="catalog_msg_bienvenida_a" data-msg-type="bienvenida_a" placeholder="Hola, te cuento sobre {titulo}...">{{ item.mensajes.bienvenida_a }}</textarea>
                                            </div>
                                            <div class="full">
                                                <label>Bienvenida B <span style="font-weight:400; color:#7f8c8d;">(texto alternativo de respaldo)</span></label>
                                                <div class="actions" style="margin-top:4px; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'bienvenida_b')">Restaurar default B</button>
                                                </div>
                                                <textarea name="catalog_msg_bienvenida_b" data-msg-type="bienvenida_b" placeholder="Version B para probar conversion con {titulo}...">{{ item.mensajes.bienvenida_b }}</textarea>
                                            </div>
                                            <div class="full">
                                                <label>Texto info producto <span style="font-weight:400; color:#7f8c8d;">(respaldo)</span></label>
                                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'info')">Restaurar default</button>
                                                </div>
                                                <textarea name="catalog_msg_info" data-msg-type="info" placeholder="Material: {titulo}...">{{ item.mensajes.info }}</textarea>
                                            </div>
                                            <div class="full">
                                                <label>Texto después de demo <span style="font-weight:400; color:#7f8c8d;">(respaldo)</span></label>
                                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'despues_demo')">Restaurar default</button>
                                                </div>
                                                <textarea name="catalog_msg_despues_demo" data-msg-type="despues_demo" placeholder="Viste la demo de {titulo}...">{{ item.mensajes.despues_demo }}</textarea>
                                            </div>
                                            <div class="full">
                                                <label>Texto de descuento <span style="font-weight:400; color:#7f8c8d;">(respaldo – solo abandono/solicitud explícita)</span></label>
                                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'descuento')">Restaurar default</button>
                                                </div>
                                                <textarea name="catalog_msg_descuento" data-msg-type="descuento" placeholder="Hoy {titulo} tiene oferta...">{{ item.mensajes.descuento }}</textarea>
                                            </div>
                                            <div class="full">
                                                <label>Texto confirmación pago <span style="font-weight:400; color:#7f8c8d;">(respaldo)</span></label>
                                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'confirmacion_pago')">Restaurar default</button>
                                                </div>
                                                <textarea name="catalog_msg_confirmacion_pago" data-msg-type="confirmacion_pago" placeholder="Envia el comprobante de {titulo}...">{{ item.mensajes.confirmacion_pago }}</textarea>
                                            </div>
                                            <div class="full">
                                                <label>Texto entrega final <span style="font-weight:400; color:#7f8c8d;">(respaldo)</span></label>
                                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'entrega_final')">Restaurar default</button>
                                                </div>
                                                <textarea name="catalog_msg_entrega_final" data-msg-type="entrega_final" placeholder="Compra aprobada de {titulo}...">{{ item.mensajes.entrega_final }}</textarea>
                                            </div>
                                            </div>
                                        </details>
                                    </div>
                                    <div class="full preview catalog-preview"></div>
                                </div>
                            </div>
                            {% endfor %}
                        </div>
                        <span class="helper">Variables disponibles: {'{'}titulo{'}'}, {'{'}precio{'}'}, {'{'}precio_descuento{'}'}, {'{'}link_video{'}'}, {'{'}link_pdf{'}'}, {'{'}link_recursos_final{'}'}, {'{'}link_canal_whatsapp{'}'}.</span>
                    </div>
                    <div class="full meta-guia-catalogo">
                        <label>Guía de plantillas Meta por producto activo</label>
                        <small class="helper">Aparece debajo del catálogo para copiar rápidamente el nombre de plantilla requerido y cada payload exacto.</small>
                        <div class="meta-guia-grid">
                            {% for item in catalogo_items %}
                            <div class="meta-guia-card">
                                <h4>{{ item.titulo }} ({{ item.id }})</h4>
                                <p><strong>Plantilla única requerida:</strong> <code>{{ item.plantilla_info_meta }}</code></p>
                                <p style="font-size:0.85rem; color:#537080;">Esta plantilla debe tener botones con payloads: <code>video_{{ item.id }}</code>, <code>pdf_{{ item.id }}</code>, <code>comprar_{{ item.id }}</code></p>
                                <div class="payload-grid">
                                    <div class="payload-chip"><strong>video (botón 1)</strong><div class="payload-copy-row"><code>video_{{ item.id }}</code><button class="btn-small" type="button" aria-label="Copiar payload video_{{ item.id }}" onclick="copiarPayload('video_{{ item.id }}', this)">Copiar</button></div></div>
                                    <div class="payload-chip"><strong>pdf (botón 2)</strong><div class="payload-copy-row"><code>pdf_{{ item.id }}</code><button class="btn-small" type="button" aria-label="Copiar payload pdf_{{ item.id }}" onclick="copiarPayload('pdf_{{ item.id }}', this)">Copiar</button></div></div>
                                    <div class="payload-chip"><strong>comprar (botón 3)</strong><div class="payload-copy-row"><code>comprar_{{ item.id }}</code><button class="btn-small" type="button" aria-label="Copiar payload comprar_{{ item.id }}" onclick="copiarPayload('comprar_{{ item.id }}', this)">Copiar</button></div></div>
                                    <div class="payload-chip"><strong>pagar_nequi (después de comprar)</strong><div class="payload-copy-row"><code>pagar_nequi_{{ item.id }}</code><button class="btn-small" type="button" aria-label="Copiar payload pagar_nequi_{{ item.id }}" onclick="copiarPayload('pagar_nequi_{{ item.id }}', this)">Copiar</button></div></div>
                                    <div class="payload-chip"><strong>pagar_daviplata (después de comprar)</strong><div class="payload-copy-row"><code>pagar_daviplata_{{ item.id }}</code><button class="btn-small" type="button" aria-label="Copiar payload pagar_daviplata_{{ item.id }}" onclick="copiarPayload('pagar_daviplata_{{ item.id }}', this)">Copiar</button></div></div>
                                    <div class="payload-chip" style="border-color:#e8a85a; background:#fffbf3;"><strong>descuento (solo rescate/solicitud explícita)</strong><div class="payload-copy-row"><code>descuento_{{ item.id }}</code><button class="btn-small" type="button" aria-label="Copiar payload descuento_{{ item.id }}" onclick="copiarPayload('descuento_{{ item.id }}', this)">Copiar</button></div><small style="color:#8a5e2a; font-size:0.8rem;">Solo se activa por payload explícito o por lógica de abandono, no aparece como botón en la plantilla de info.</small></div>
                                </div>
                            </div>
                            {% endfor %}
                        </div>
                        <div class="full estado-media" style="margin-top:12px;">
                            <a class="btn-link" href="{{ url_for('admin_exportar_plantillas_meta') }}">Descargar JSON de plantillas y payloads</a>
                        </div>
                    </div>
                </div>
                <div class="actions">
                    <button class="button" type="submit">Guardar catálogos</button>
                </div>
            </form>

            {# Sección global: Plantillas universales en Meta #}
            <div style="margin-top:24px; padding:16px; background:#f0f9ff; border:2px solid #1976d2; border-radius:12px;">
                <h3 style="margin-top:0; color:#1565c0;">&#x1F310; Plantillas universales en Meta</h3>
                <p style="font-size:0.9rem; color:#1a5276; margin-bottom:14px;">Estas plantillas se usan para todos los productos. Créalas en Meta Business Manager con los cuerpos indicados. <strong>La plantilla de compra NO ofrece descuento; el descuento solo se activa por abandono o solicitud explícita.</strong></p>

                {% set tpls_universales = [
                    {'nombre': 'plantilla_menu_general', 'categoria': 'UTILITY', 'vars': '{{1}} = lista dinámica de productos', 'botones': 'Sin botones obligatorios', 'nota': 'Disparo: hola / menú / menu_principal o número de opción'},
                    {'nombre': 'plantilla_compra_universal', 'categoria': 'MARKETING', 'vars': '{{1}} = titulo, {{2}} = precio_normal', 'botones': 'Dinámicos: pagar_nequi_[id], pagar_daviplata_[id], menu_principal', 'nota': 'NO incluir botón de descuento. Descuento solo por abandono/solicitud explícita.'},
                    {'nombre': 'plantilla_descuento_universal', 'categoria': 'MARKETING', 'vars': '{{1}} = titulo, {{2}} = precio_normal, {{3}} = precio_descuento', 'botones': 'Dinámicos: pagar_nequi_[id], pagar_daviplata_[id], menu_principal', 'nota': 'Solo para abandono o solicitud explícita. NO usar en flujo normal de compra.'},
                    {'nombre': 'plantilla_cuentas_cobro', 'categoria': 'UTILITY', 'vars': '{{1}} = titulo, {{2}} = metodo, {{3}} = valor, {{4}} = datos_pago', 'botones': 'Sin botones obligatorios. Flujo posterior: ya_pague.', 'nota': ''},
                    {'nombre': 'plantilla_entrega_final', 'categoria': 'UTILITY', 'vars': '{{1}} = link_drive_final, {{2}} = link_canal_whatsapp', 'botones': 'Sin botones obligatorios', 'nota': 'Se dispara tras aprobación admin.'},
                    {'nombre': 'plantilla_seguimiento_pdf', 'categoria': 'MARKETING', 'vars': '{{1}} = titulo', 'botones': 'Dinámicos: comprar_[id], video_[id], menu_principal', 'nota': 'Sin botón de descuento. Se envía inmediatamente tras enviar el PDF demo.'}
                ] %}

                {% for tpl in tpls_universales %}
                <div style="margin-bottom:14px; padding:12px; background:#fff; border-radius:8px; border:1px solid #90caf9;">
                    <p style="margin:0 0 6px 0; font-weight:700; color:#1565c0;">{{ tpl.nombre }}</p>
                    <p style="margin:0 0 4px 0; font-size:0.84rem;"><strong>Categoría:</strong> {{ tpl.categoria }} &bull; <strong>Idioma:</strong> es</p>
                    <p style="margin:0 0 4px 0; font-size:0.84rem;"><strong>Variables:</strong> {{ tpl.vars }}</p>
                    <p style="margin:0 0 4px 0; font-size:0.84rem;"><strong>Botones:</strong> {{ tpl.botones }}</p>
                    {% if tpl.nota %}<p style="margin:0 0 6px 0; font-size:0.82rem; color:#8a5e2a;"><strong>Nota:</strong> {{ tpl.nota }}</p>{% endif %}
                    <div style="margin-top:8px;">
                        <label style="font-size:0.83rem; color:#537080;">Texto con variables (el bot reemplaza {{1}}, {{2}}, etc. desde el catálogo — si cambias precio o nombre en catálogo, <strong>no necesitas editar Meta</strong>):</label>
                        <textarea id="tpl_universal_{{ tpl.nombre }}" readonly style="min-height:90px; width:100%; font-size:0.83rem; background:#f8fbff; margin-top:4px;">{% if tpl.nombre in catalogo_plantillas_meta %}{{ catalogo_plantillas_meta[tpl.nombre].texto_referencia }}{% else %}(No disponible){% endif %}</textarea>
                        <button class="btn-small" type="button" style="margin-top:4px;" onclick="copiarTextoPorId('tpl_universal_{{ tpl.nombre }}')">Copiar texto con variables</button>
                    </div>
                </div>
                {% endfor %}
            </div>
        </section>

        <section class="card panel-section is-hidden" id="pendientes">
            <h2>Comprobantes pendientes de aprobacion</h2>
            <p class="helper">Esta tabla se actualiza automaticamente cada 8 segundos.</p>
            <div id="pendientes-container">
                {% if pendientes_agrupados %}
                {% for grupo in pendientes_agrupados %}
                <h3 style="margin-bottom: 8px;">{{ grupo.titulo }} ({{ grupo.clave }})</h3>
                <table class="tabla">
                    <thead>
                        <tr>
                            <th>Numero</th>
                            <th>Producto</th>
                            <th>Metodo</th>
                            <th>Archivo</th>
                            <th>Fecha</th>
                            <th>Acciones</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for solicitud in grupo["items"] %}
                        <tr>
                            <td>{{ solicitud.numero }}</td>
                            <td>{{ solicitud.producto_titulo or solicitud.catalogo_clave or 'General' }}</td>
                            <td>{{ solicitud.metodo_pago }}</td>
                            <td>{{ solicitud.tipo_archivo }}</td>
                            <td>{{ solicitud.fecha }}</td>
                            <td>
                                <div class="acciones-pendiente">
                                    <form method="POST" action="{{ url_for('admin_aprobar', numero=solicitud.numero) }}">
                                        <button class="btn-small" type="submit">Aprobar y entregar</button>
                                    </form>
                                    <form method="POST" action="{{ url_for('admin_rechazar', numero=solicitud.numero) }}">
                                        <button class="btn-small button-danger" type="submit">Rechazar</button>
                                    </form>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% endfor %}
                {% else %}
                <p>No hay comprobantes pendientes en este momento.</p>
                {% endif %}
            </div>
        </section>

        <section class="card panel-section is-hidden" id="reportes">
            <h2>Reportes de ventas y contabilidad</h2>
            <form id="tutorial-reportes-filtros" method="GET" action="{{ url_for('admin') }}#reportes" class="actions" style="margin-bottom: 10px;">
                <div>
                    <label for="fecha_desde">Desde</label>
                    <input id="fecha_desde" type="date" name="fecha_desde" value="{{ fecha_desde }}">
                </div>
                <div>
                    <label for="fecha_hasta">Hasta</label>
                    <input id="fecha_hasta" type="date" name="fecha_hasta" value="{{ fecha_hasta }}">
                </div>
                <div>
                    <label for="producto_reporte">Producto</label>
                    <select id="producto_reporte" name="producto_reporte">
                        <option value="">Todos</option>
                        {% for item in catalogo_items %}
                        <option value="{{ item.key }}" {% if producto_reporte == item.key %}selected{% endif %}>{{ item.titulo }} ({{ item.key }})</option>
                        {% endfor %}
                    </select>
                </div>
                <div style="display:flex; align-items:flex-end; gap:8px;">
                    <button class="btn-small" type="submit">Filtrar</button>
                    <a class="btn-link" href="{{ url_for('admin') }}#reportes">Limpiar</a>
                </div>
            </form>
            <div class="stats">
                <div class="stat">
                    Total ventas
                    <strong>{{ resumen_ventas.total }}</strong>
                </div>
                <div class="stat">
                    Ventas normales
                    <strong>{{ resumen_ventas.ventas_normal }}</strong>
                </div>
                <div class="stat">
                    Ventas con descuento
                    <strong>{{ resumen_ventas.ventas_descuento }}</strong>
                </div>
                <div class="stat">
                    Ingreso total
                    <strong>${{ resumen_ventas.ingresos_total }}</strong>
                </div>
            </div>
            <div class="actions" style="margin-top: 14px;">
                <a class="btn-link" href="{{ url_for('admin_reporte_excel', fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, producto_reporte=producto_reporte) }}">Descargar Excel</a>
                <a class="btn-link" href="{{ url_for('admin_reporte_pdf', fecha_desde=fecha_desde, fecha_hasta=fecha_hasta, producto_reporte=producto_reporte) }}">Descargar PDF consolidado</a>
            </div>
            {% if ventas %}
            {% if resumen_producto %}
            <table class="tabla" style="margin-top: 14px;">
                <thead>
                    <tr>
                        <th>Producto</th>
                        <th>Clave</th>
                        <th>Ventas</th>
                        <th>Ingresos</th>
                    </tr>
                </thead>
                <tbody>
                    {% for rp in resumen_producto %}
                    <tr>
                        <td>{{ rp.titulo }}</td>
                        <td>{{ rp.clave }}</td>
                        <td>{{ rp.ventas }}</td>
                        <td>${{ rp.ingresos }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% endif %}
            <table class="tabla" style="margin-top: 14px;">
                <thead>
                    <tr>
                        <th>Fecha</th>
                        <th>Cliente</th>
                        <th>Producto</th>
                        <th>Metodo</th>
                        <th>Tipo precio</th>
                        <th>Valor</th>
                        <th>Comprobante</th>
                    </tr>
                </thead>
                <tbody>
                    {% for venta in ventas %}
                    <tr>
                        <td>{{ venta.fecha_aprobacion }}</td>
                        <td>{{ venta.numero_cliente }}</td>
                        <td>{{ venta.producto_titulo or venta.catalogo_clave or 'General' }}</td>
                        <td>{{ venta.metodo_pago }}</td>
                        <td>{{ venta.tipo_precio }}</td>
                        <td>${{ venta.valor_venta }}</td>
                        <td>{{ venta.nombre_archivo or venta.tipo_archivo }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p class="helper">Aun no hay ventas aprobadas para mostrar.</p>
            {% endif %}
        </section>

        <section class="card panel-section is-hidden" id="clientes">
            <h2>Clientes vendidos</h2>
            <p class="helper">Aqui se guardan automaticamente los numeros de clientes con compra aprobada.</p>
            <div class="stats" id="tutorial-clientes-resumen">
                <div class="stat">
                    Total clientes unicos
                    <strong>{{ clientes_total }}</strong>
                </div>
            </div>
            {% if analitica_clientes_producto %}
            <table class="tabla" style="margin-top: 14px;">
                <thead>
                    <tr>
                        <th>Producto</th>
                        <th>Clave</th>
                        <th>Clientes unicos</th>
                        <th>Ventas</th>
                        <th>Ingresos</th>
                    </tr>
                </thead>
                <tbody>
                    {% for fila in analitica_clientes_producto %}
                    <tr>
                        <td>{{ fila.titulo }}</td>
                        <td>{{ fila.clave }}</td>
                        <td>{{ fila.clientes_unicos }}</td>
                        <td>{{ fila.ventas }}</td>
                        <td>${{ fila.ingresos }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% endif %}
            {% if clientes_lista %}
            <table class="tabla" style="margin-top: 14px;">
                <thead>
                    <tr>
                        <th>Numero</th>
                        <th>Ventas</th>
                        <th>Ultima venta</th>
                        <th>Ultimo producto</th>
                        <th>Ultimo metodo</th>
                        <th>Ultimo tipo precio</th>
                    </tr>
                </thead>
                <tbody>
                    {% for c in clientes_lista %}
                    <tr>
                        <td>{{ c.numero }}</td>
                        <td>{{ c.ventas }}</td>
                        <td>{{ c.ultima_venta }}</td>
                        <td>{{ c.ultimo_producto }}</td>
                        <td>{{ c.ultimo_metodo_pago }}</td>
                        <td>{{ c.ultimo_tipo_precio }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p class="helper">Aun no hay clientes vendidos registrados.</p>
            {% endif %}
        </section>

        <section class="card panel-section is-hidden" id="difusion">
            <h2>Difusion de materiales y promociones</h2>
            <p class="helper">Envia mensajes masivos a clientes vendidos y/o numeros cargados manualmente.</p>
            <form id="tutorial-difusion-form" method="POST" action="{{ url_for('admin_difusion') }}" enctype="multipart/form-data">
                <div class="grid">
                    <div class="full">
                        <label for="mensaje_difusion">Mensaje de difusion</label>
                        <textarea id="mensaje_difusion" name="mensaje_difusion" placeholder="Escribe la promocion o informacion del nuevo material..."></textarea>
                    </div>
                    <div class="full">
                        <label for="numeros_manual">Numeros manuales (opcional)</label>
                        <textarea id="numeros_manual" name="numeros_manual" placeholder="Pega numeros separados por coma, espacio o salto de linea"></textarea>
                    </div>
                    <div class="full">
                        <label for="archivo_numeros">Cargar archivo de numeros (TXT/CSV)</label>
                        <input id="archivo_numeros" type="file" name="archivo_numeros" accept=".txt,.csv">
                    </div>
                    <div class="full">
                        <label for="link_material_difusion">Link de material (opcional)</label>
                        <input id="link_material_difusion" type="text" name="link_material_difusion" placeholder="https://... o /media/...">
                        <span class="helper">Si agregas un link de video o PDF, se enviara despues del mensaje.</span>
                    </div>
                    <div class="full">
                        <label for="archivo_material_difusion">Cargar archivo de material (opcional)</label>
                        <input id="archivo_material_difusion" type="file" name="archivo_material_difusion" accept="video/*,.pdf,.doc,.docx,.ppt,.pptx,.zip,.jpg,.jpeg,.png">
                        <span class="helper">El archivo se guardara y se enviara como material adicional.</span>
                    </div>
                    <div class="full">
                        <label><input type="checkbox" name="usar_clientes_vendidos" checked> Incluir automaticamente clientes vendidos</label>
                    </div>
                    <div>
                        <label for="segmento_tipo_precio">Segmento por tipo de precio</label>
                        <select id="segmento_tipo_precio" name="segmento_tipo_precio">
                            <option value="">Todos</option>
                            <option value="normal">Normal</option>
                            <option value="descuento">Descuento</option>
                        </select>
                    </div>
                    <div>
                        <label for="segmento_metodo_pago">Segmento por metodo de pago</label>
                        <select id="segmento_metodo_pago" name="segmento_metodo_pago">
                            <option value="">Todos</option>
                            <option value="nequi">Nequi</option>
                            <option value="daviplata">Daviplata</option>
                            <option value="llave">Llave</option>
                        </select>
                    </div>
                    <div>
                        <label for="segmento_producto">Segmento por producto</label>
                        <select id="segmento_producto" name="segmento_producto">
                            <option value="">Todos</option>
                            {% for item in catalogo_items %}
                            <option value="{{ item.key }}">{{ item.titulo }} ({{ item.key }})</option>
                            {% endfor %}
                        </select>
                    </div>
                </div>
                <div class="actions">
                    <button class="button" type="submit">Enviar difusion</button>
                </div>
            </form>
        </section>
        <section class="card panel-section is-hidden" id="prueba_meta">
            <h2><i class="bi bi-send-check"></i> Prueba de Conexión de Plantillas Meta</h2>
            <p class="helper">Envía una plantilla con variables dinámicas directamente a la API de Meta para verificar la conexión. La plantilla de referencia es: <em>"Hi &#123;&#123;1&#125;&#125;, we need to reschedule your &#123;&#123;2&#125;&#125;. Reply Reschedule to pick a new time."</em></p>
            {% if mensaje %}<div class="alerta">{{ mensaje }}</div>{% endif %}
            <form method="POST" action="{{ url_for('test_meta') }}">
                <div class="grid">
                    <div>
                        <label for="test_numero_destino">Número de destino</label>
                        <input id="test_numero_destino" type="text" name="numero_destino" placeholder="ej. 573001234567" required>
                        <span class="helper">Número WhatsApp sin + ni espacios (código de país incluido).</span>
                    </div>
                    <div>
                        <label for="test_template_nombre">Nombre de la plantilla</label>
                        <input id="test_template_nombre" type="text" name="test_template_nombre" value="{{ test_template_nombre or 'hello_world' }}" required>
                    </div>
                    <div>
                        <label for="test_template_idioma">Idioma</label>
                        <input id="test_template_idioma" type="text" name="test_template_idioma" value="{{ test_template_idioma or 'en_US' }}" required>
                    </div>
                    <div>
                        <label for="test_template_var1">Variable &#123;&#123;1&#125;&#125; (ej. nombre del cliente)</label>
                        <input id="test_template_var1" type="text" name="test_template_var1" value="{{ test_template_var1 or '' }}" placeholder="ej. Juan">
                    </div>
                    <div>
                        <label for="test_template_var2">Variable &#123;&#123;2&#125;&#125; (ej. motivo de la cita)</label>
                        <input id="test_template_var2" type="text" name="test_template_var2" value="{{ test_template_var2 or '' }}" placeholder="ej. cita médica">
                    </div>
                </div>
                <div class="actions" style="margin-top: 16px;">
                    <button class="button" type="submit">Enviar Prueba a Meta</button>
                </div>
            </form>
        </section>
    </div>
    {% if not ocultar_tutorial_panel %}
    <button class="tutorial-fab" id="btn-tutorial-flotante" type="button" onclick="iniciarTutorialPanel()">Guía rápida</button>
    {% endif %}
    <div class="tutorial-overlay is-hidden" id="tutorial-overlay" onclick="cerrarTutorialPanel()"></div>
    <div class="tutorial-card is-hidden" id="tutorial-card" role="dialog" aria-live="polite" aria-label="Tutorial del panel">
        <h3 id="tutorial-titulo">Tutorial</h3>
        <p id="tutorial-texto"></p>
        <div class="tutorial-meta" id="tutorial-meta"></div>
        <div class="tutorial-actions">
            <button class="btn-small" type="button" onclick="tutorialAnterior()">Anterior</button>
            <div style="display:flex; gap:8px;">
                <button class="btn-small button-danger" type="button" onclick="cerrarTutorialPanel()">Cerrar</button>
                <button class="btn-small" type="button" onclick="tutorialSiguiente()">Siguiente</button>
            </div>
        </div>
    </div>
    <script>
        const MENSAJES_DEFAULT = {{ mensajes_default_json|safe }};
        const PENDIENTES_INICIAL = {{ pendientes_total }};
        // AUTO_INICIAR_TUTORIAL: true solo si tutorial_habilitado Y auto_iniciar_tutorial están activos
        const AUTO_INICIAR_TUTORIAL = {{ 'true' if (tutorial_habilitado and auto_iniciar_tutorial) else 'false' }};
        const TUTORIAL_HABILITADO = {{ 'true' if tutorial_habilitado else 'false' }};
        const TEXTO_INFO_META_BASE = {{ texto_info_meta_base_json|safe }};
    </script>
    <script>
        let wizardStepActual = 1;
        let tutorialPasoActual = 0;
        const PASOS_TUTORIAL = [
            {
                tab: "configuracion",
                selector: "#tutorial-config-base",
                titulo: "Paso 1: Configuración base",
                texto: "Empieza por este bloque: ID de teléfono, token y datos principales de cobro. Guarda siempre después de cualquier cambio importante."
            },
            {
                tab: "configuracion",
                selector: "#tutorial-plantillas-meta",
                titulo: "Paso 2: Plantillas Meta",
                texto: "Este bloque es solo para plantillas Meta: nombre, idioma, texto de referencia y variables. Si te equivocas, usa restaurar por defecto y vuelve a guardar."
            },
            {
                tab: "catalogos",
                selector: "#selector_catalogo",
                titulo: "Paso 3: Crear y editar catálogos",
                texto: "Desde este selector y editor gestionas productos, precio, links demo y mensajes por producto. Es el módulo principal para escalar el bot."
            },
            {
                tab: "pendientes",
                selector: "#pendientes-container",
                titulo: "Paso 4: Aprobar comprobantes",
                texto: "Aquí validas pagos pendientes en tiempo real. Al aprobar, el bot entrega recursos automáticamente al cliente."
            },
            {
                tab: "clientes",
                selector: "#tutorial-clientes-resumen",
                titulo: "Paso 5: Contactos y clientes",
                texto: "Este bloque resume clientes únicos y su historial por producto para seguimiento, retención y nuevas ofertas."
            },
            {
                tab: "difusion",
                selector: "#tutorial-difusion-form",
                titulo: "Paso 6: Difusiones",
                texto: "Desde este formulario envías campañas por segmento, números manuales o archivo. Ideal para reactivar ventas."
            },
            {
                tab: "reportes",
                selector: "#tutorial-reportes-filtros",
                titulo: "Paso 7: Reportes",
                texto: "Aquí filtras ventas por fecha y producto, y exportas Excel o PDF para control comercial."
            }
        ];

        function actualizarBadgePendientes(cantidad) {
            const badge = document.getElementById("tab-pendientes-count");
            if (!badge) return;
            const total = Number.isFinite(cantidad) ? cantidad : 0;
            badge.textContent = String(total);
            badge.style.background = total > 0 ? "#d54949" : "#6b7f8c";
        }

        function contarPendientesEnDOM() {
            const contenedor = document.getElementById("pendientes-container");
            if (!contenedor) return 0;
            return contenedor.querySelectorAll("tbody tr").length;
        }

        (function () {
            const contenedor = document.getElementById("pendientes-container");
            if (!contenedor) return;
            function refrescarPendientes() {
                fetch("/admin/pendientes", { cache: "no-store" })
                    .then((res) => res.text())
                    .then((html) => {
                        contenedor.innerHTML = html;
                        actualizarBadgePendientes(contarPendientesEnDOM());
                    })
                    .catch(() => {});
            }
            setInterval(function () {
                refrescarPendientes();
            }, 8000);
            window.addEventListener("focus", refrescarPendientes);
            document.addEventListener("visibilitychange", function () {
                if (!document.hidden) refrescarPendientes();
            });
            actualizarBadgePendientes(PENDIENTES_INICIAL);
        })();

        function alternarToken() {
            const input = document.getElementById("token_meta");
            if (!input) return;
            input.type = input.type === "password" ? "text" : "password";
        }

        function copiarPayload(texto, boton) {
            if (!texto) return;
            const marcarBoton = (nuevoTexto) => {
                if (!boton) return;
                const original = boton.dataset.originalText || boton.textContent;
                boton.dataset.originalText = original;
                boton.textContent = nuevoTexto;
                setTimeout(() => {
                    boton.textContent = original;
                }, 1200);
            };
            if (!navigator.clipboard || !navigator.clipboard.writeText) {
                marcarBoton("Error");
                return;
            }
            navigator.clipboard.writeText(texto)
                .then(() => marcarBoton("Copiado"))
                .catch(() => marcarBoton("Error"));
        }

        function copiarTextoPorId(id) {
            const input = document.getElementById(id);
            if (!input) return;
            const texto = input.value || "";
            if (!texto) return;
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(texto).catch(() => {});
                return;
            }
            input.select();
            document.execCommand("copy");
        }

        // Sincroniza el textarea de previsualización con el campo texto_info_meta editable
        function sincronizarTextoInfoMeta(idx) {
            const src = document.getElementById("tpl_texto_info_meta_" + idx);
            const dest = document.getElementById("tpl_producto_texto_" + idx);
            if (src && dest) {
                dest.value = src.value;
            }
        }

        function enfocarCatalogo(id) {
            if (!id) return;
            const el = document.getElementById(id);
            if (!el) return;
            el.scrollIntoView({ behavior: "smooth", block: "center" });
        }

        function eliminarProductoCatalogo(btn) {
            const card = btn.closest(".catalogo-item");
            if (card) {
                card.remove();
                refrescarSelectorCatalogo();
                refrescarPreviewCatalogo();
            }
        }

        function duplicarProductoCatalogo(btn) {
            const card = btn.closest(".catalogo-item");
            const contenedor = document.getElementById("catalogo-editor");
            if (!card || !contenedor) return;
            const idx = Date.now();
            const clon = card.cloneNode(true);
            clon.id = `catalogo_item_${idx}`;
            clon.setAttribute("data-index", String(idx));
            clon.setAttribute("draggable", "true");

            const idxInput = clon.querySelector('input[name="catalog_index"]');
            if (idxInput) idxInput.value = String(idx);
            const originalKeyInput = clon.querySelector('input[name="catalog_original_key"]');
            if (originalKeyInput) originalKeyInput.value = "";
            const keyInput = clon.querySelector('input[name="catalog_key"]');
            if (keyInput && keyInput.value) keyInput.value = `${keyInput.value}_copia`;

            // Renombrar inputs ocultos de archivo (no se muestran en UI; el backend los soporta por compatibilidad)
            clon.querySelectorAll('input[type="hidden"]').forEach((hiddenInput) => {
                if (hiddenInput.name.startsWith("catalog_archivo_video_")) {
                    hiddenInput.name = `catalog_archivo_video_${idx}`;
                } else if (hiddenInput.name.startsWith("catalog_archivo_pdf_")) {
                    hiddenInput.name = `catalog_archivo_pdf_${idx}`;
                }
                hiddenInput.value = "";
            });

            contenedor.insertBefore(clon, card.nextSibling);
            activarDnDCatalogo();
            refrescarSelectorCatalogo();
            refrescarPreviewCatalogo();
            clon.scrollIntoView({ behavior: "smooth", block: "center" });
        }

        function restaurarMensajeProducto(btn, tipo) {
            const card = btn.closest(".catalogo-item");
            if (!card) return;
            const textarea = card.querySelector(`textarea[data-msg-type="${tipo}"]`);
            if (!textarea) return;
            textarea.value = MENSAJES_DEFAULT[tipo] || "";
            refrescarPreviewCatalogo();
        }

        function agregarProductoCatalogo() {
            const contenedor = document.getElementById("catalogo-editor");
            if (!contenedor) return;
            const idx = Date.now();
            const card = document.createElement("div");
            card.className = "estado-media catalogo-item";
            card.id = `catalogo_item_${idx}`;
            card.setAttribute("data-index", String(idx));
            card.innerHTML = `
                <div class="actions" style="margin-top:0; margin-bottom:8px;">
                    <span class="btn-link drag-handle" title="Arrastra para reordenar">Mover</span>
                    <strong>Nuevo producto</strong>
                    <button class="btn-small" type="button" onclick="duplicarProductoCatalogo(this)">Duplicar</button>
                    <button class="btn-small button-danger" type="button" onclick="eliminarProductoCatalogo(this)">Eliminar</button>
                </div>
                <input type="hidden" name="catalog_index" value="${idx}">
                <input type="hidden" name="catalog_original_key" value="">
                <div class="grid">
                    <div>
                        <label>Clave interna</label>
                        <input type="text" name="catalog_key" value="" placeholder="ej: matematicas">
                    </div>
                    <div>
                        <label>Titulo del producto</label>
                        <input type="text" name="catalog_titulo" value="" placeholder="Titulo visible al cliente">
                    </div>
                    <div>
                        <label>Precio</label>
                        <input type="text" name="catalog_precio" value="" placeholder="20.000">
                    </div>
                    <div>
                        <label>Precio descuento</label>
                        <input type="text" name="catalog_precio_descuento" value="" placeholder="18.000">
                    </div>
                    <div>
                        <label>Palabras clave</label>
                        <input type="text" name="catalog_palabras" value="" placeholder="matematicas, algebra, ecuaciones">
                    </div>
                    <div>
                        <label>URL pública del video demo <span style="color:#c0392b; font-size:0.82rem;">(GitHub Raw, CDN o HTTPS directo)</span></label>
                        <input type="text" name="catalog_link_video" value="" placeholder="https://raw.githubusercontent.com/tu_usuario/repo/main/video_demo.mp4">
                        <small style="color:#8a5e2a; font-size:0.8rem;">En Render no subas archivos locales aquí. Pega URL pública HTTPS.</small>
                    </div>
                    <div>
                        <label>URL pública del PDF demo <span style="color:#c0392b; font-size:0.82rem;">(GitHub Raw, CDN o HTTPS directo)</span></label>
                        <input type="text" name="catalog_link_pdf" value="" placeholder="https://raw.githubusercontent.com/tu_usuario/repo/main/demo.pdf">
                        <small style="color:#8a5e2a; font-size:0.8rem;">En Render no subas archivos locales aquí. Pega URL pública HTTPS.</small>
                    </div>
                    <input type="hidden" name="catalog_archivo_video_${idx}" value="">
                    <input type="hidden" name="catalog_archivo_pdf_${idx}" value="">
                    <div class="full" style="margin-top:10px; padding:14px; background:#e8f5e9; border:2px solid #2e7d32; border-radius:12px;">
                        <strong style="font-size:0.9rem; color:#1b5e20;">📝 Texto plantilla única de información para Meta</strong>
                        <p style="font-size:0.82rem; color:#388e3c; margin:4px 0 6px 0;">Escribe el cuerpo de la plantilla única de este producto para Meta. Texto persuasivo, sin variables {{1}} en el body.</p>
                        <textarea name="catalog_texto_info_meta" style="min-height:120px; width:100%; font-size:0.84rem; background:#f1f8f2; border:1px solid #81c784; border-radius:8px; padding:8px;">${TEXTO_BASE_NUEVO || ""}</textarea>
                    </div>
                    <div class="full" style="margin-top:8px; padding:10px; background:#eaf5ff; border:1px solid #90caf9; border-radius:8px;">
                        <strong style="font-size:0.9rem; color:#1565c0;">Plantillas Meta (configura en Meta Business Manager)</strong>
                        <p style="font-size:0.82rem; color:#537080; margin:4px 0 4px 0;">Botones obligatorios: video_[id], pdf_[id], comprar_[id]. Descuento solo por abandono.</p>
                        <p style="font-size:0.82rem; color:#0d7b5f; margin:0;">Precio normal → {{2}} en plantilla_compra_universal. Si cambias precio aquí, no debes editar Meta.</p>
                    </div>
                    <div class="full" style="margin-top:10px;">
                        <details>
                            <summary style="cursor:pointer; color:#537080; font-size:0.9rem;">Textos de referencia / respaldo (bienvenida A/B y mensajes planos)</summary>
                            <div class="full">
                                <label>Bienvenida A (respaldo)</label>
                                <div class="actions" style="margin-top:4px; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'bienvenida_a')">Restaurar default A</button>
                                </div>
                                <textarea name="catalog_msg_bienvenida_a" data-msg-type="bienvenida_a" placeholder="Hola, te cuento sobre {titulo}...">${MENSAJES_DEFAULT.bienvenida_a || MENSAJES_DEFAULT.bienvenida || ""}</textarea>
                            </div>
                            <div class="full">
                                <label>Bienvenida B (respaldo)</label>
                                <div class="actions" style="margin-top:4px; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'bienvenida_b')">Restaurar default B</button>
                                </div>
                                <textarea name="catalog_msg_bienvenida_b" data-msg-type="bienvenida_b" placeholder="Version B para probar conversion con {titulo}...">${MENSAJES_DEFAULT.bienvenida_b || MENSAJES_DEFAULT.bienvenida || ""}</textarea>
                            </div>
                            <div class="full">
                                <label>Texto info (respaldo)</label>
                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'info')">Restaurar default</button>
                                </div>
                                <textarea name="catalog_msg_info" data-msg-type="info" placeholder="Material: {titulo}...">${MENSAJES_DEFAULT.info || ""}</textarea>
                            </div>
                            <div class="full">
                                <label>Texto después de demo (respaldo)</label>
                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'despues_demo')">Restaurar default</button>
                                </div>
                                <textarea name="catalog_msg_despues_demo" data-msg-type="despues_demo" placeholder="Viste la demo de {titulo}...">${MENSAJES_DEFAULT.despues_demo || ""}</textarea>
                            </div>
                            <div class="full">
                                <label>Texto descuento (respaldo – solo abandono)</label>
                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'descuento')">Restaurar default</button>
                                </div>
                                <textarea name="catalog_msg_descuento" data-msg-type="descuento" placeholder="Hoy {titulo} tiene oferta...">${MENSAJES_DEFAULT.descuento || ""}</textarea>
                            </div>
                            <div class="full">
                                <label>Texto confirmación pago (respaldo)</label>
                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'confirmacion_pago')">Restaurar default</button>
                                </div>
                                <textarea name="catalog_msg_confirmacion_pago" data-msg-type="confirmacion_pago" placeholder="Envia el comprobante de {titulo}...">${MENSAJES_DEFAULT.confirmacion_pago || ""}</textarea>
                            </div>
                            <div class="full">
                                <label>Texto entrega final (respaldo)</label>
                                <div class="actions" style="margin-top:0; margin-bottom:6px;">
                                    <button class="btn-small" type="button" onclick="restaurarMensajeProducto(this, 'entrega_final')">Restaurar default</button>
                                </div>
                                <textarea name="catalog_msg_entrega_final" data-msg-type="entrega_final" placeholder="Compra aprobada de {titulo}...">${MENSAJES_DEFAULT.entrega_final || ""}</textarea>
                            </div>
                        </details>
                    </div>
                    <div class="full preview catalog-preview"></div>
                </div>
            `;
            contenedor.appendChild(card);
            activarDnDCatalogo();
            refrescarSelectorCatalogo();
            refrescarPreviewCatalogo();
            card.scrollIntoView({ behavior: "smooth", block: "center" });
            return card;
        }

        function mostrarWizardProducto() {
            const wizard = document.getElementById("wizard-producto");
            if (!wizard) return;
            wizard.classList.remove("is-hidden");
            wizardStepActual = 1;
            actualizarWizardStep();

            document.getElementById("wizard_msg_info").value = MENSAJES_DEFAULT.info || "";
            document.getElementById("wizard_msg_demo").value = MENSAJES_DEFAULT.despues_demo || "";
            document.getElementById("wizard_msg_desc").value = MENSAJES_DEFAULT.descuento || "";
        }

        function ocultarWizardProducto() {
            const wizard = document.getElementById("wizard-producto");
            if (!wizard) return;
            wizard.classList.add("is-hidden");
        }

        function actualizarWizardStep() {
            [1, 2, 3].forEach((n) => {
                const step = document.getElementById(`wizard-step-${n}`);
                if (!step) return;
                step.classList.toggle("is-hidden", n !== wizardStepActual);
            });

            const prev = document.getElementById("wizard_prev");
            const next = document.getElementById("wizard_next");
            const finish = document.getElementById("wizard_finish");
            if (prev) prev.disabled = wizardStepActual === 1;
            if (next) next.classList.toggle("is-hidden", wizardStepActual === 3);
            if (finish) finish.classList.toggle("is-hidden", wizardStepActual !== 3);
        }

        function wizardNext() {
            wizardStepActual = Math.min(3, wizardStepActual + 1);
            actualizarWizardStep();
        }

        function wizardPrev() {
            wizardStepActual = Math.max(1, wizardStepActual - 1);
            actualizarWizardStep();
        }

        function crearProductoDesdeWizard() {
            const card = agregarProductoCatalogo();
            if (!card) return;

            const key = document.getElementById("wizard_key")?.value || "";
            const titulo = document.getElementById("wizard_titulo")?.value || "";
            const precio = document.getElementById("wizard_precio")?.value || "";
            const precioDesc = document.getElementById("wizard_precio_desc")?.value || "";
            const palabras = document.getElementById("wizard_palabras")?.value || "";
            const video = document.getElementById("wizard_video")?.value || "";
            const pdf = document.getElementById("wizard_pdf")?.value || "";
            const textoInfoMeta = document.getElementById("wizard_texto_info_meta")?.value || TEXTO_INFO_META_BASE || "";
            const msgInfo = document.getElementById("wizard_msg_info")?.value || "";
            const msgDemo = document.getElementById("wizard_msg_demo")?.value || "";
            const msgDesc = document.getElementById("wizard_msg_desc")?.value || "";

            const setVal = (selector, value) => {
                const el = card.querySelector(selector);
                if (el) el.value = value;
            };
            setVal('input[name="catalog_key"]', key);
            setVal('input[name="catalog_titulo"]', titulo);
            setVal('input[name="catalog_precio"]', precio);
            setVal('input[name="catalog_precio_descuento"]', precioDesc || precio);
            setVal('input[name="catalog_palabras"]', palabras);
            setVal('input[name="catalog_link_video"]', video);
            setVal('input[name="catalog_link_pdf"]', pdf);
            setVal('textarea[name="catalog_texto_info_meta"]', textoInfoMeta);
            setVal('textarea[name="catalog_msg_info"]', msgInfo);
            setVal('textarea[name="catalog_msg_despues_demo"]', msgDemo);
            setVal('textarea[name="catalog_msg_descuento"]', msgDesc);

            ocultarWizardProducto();
            refrescarSelectorCatalogo();
            refrescarPreviewCatalogo();
            card.scrollIntoView({ behavior: "smooth", block: "center" });
        }

        function refrescarSelectorCatalogo() {
            const select = document.getElementById("selector_catalogo");
            if (!select) return;
            select.innerHTML = '<option value="">Ir a producto...</option>';
            document.querySelectorAll(".catalogo-item").forEach((item, i) => {
                const titulo = item.querySelector('input[name="catalog_titulo"]')?.value || `Producto ${i + 1}`;
                const clave = item.querySelector('input[name="catalog_key"]')?.value || "sin-clave";
                const option = document.createElement("option");
                option.value = item.id;
                option.textContent = `${titulo} (${clave})`;
                select.appendChild(option);
            });
        }

        function refrescarPreviewCatalogo() {
            document.querySelectorAll(".catalogo-item").forEach((item) => {
                const titulo = item.querySelector('input[name="catalog_titulo"]')?.value || "Producto";
                const precio = item.querySelector('input[name="catalog_precio"]')?.value || "0";
                const precioDesc = item.querySelector('input[name="catalog_precio_descuento"]')?.value || precio;
                const linkVideo = item.querySelector('input[name="catalog_link_video"]')?.value || "";
                const linkPdf = item.querySelector('input[name="catalog_link_pdf"]')?.value || "";

                const msgInfo = item.querySelector('textarea[name="catalog_msg_info"]')?.value || "";
                const msgDespues = item.querySelector('textarea[name="catalog_msg_despues_demo"]')?.value || "";
                const msgDesc = item.querySelector('textarea[name="catalog_msg_descuento"]')?.value || "";
                const msgComprobante = item.querySelector('textarea[name="catalog_msg_confirmacion_pago"]')?.value || "";
                const msgEntrega = item.querySelector('textarea[name="catalog_msg_entrega_final"]')?.value || "";

                const nequi = document.getElementById("nequi")?.value || "No configurado";
                const daviplata = document.getElementById("daviplata")?.value || "No configurado";
                const llave = document.getElementById("llave")?.value || "No configurado";
                const linkRecursos = document.getElementById("link_recursos_final")?.value || "";
                const linkCanal = document.getElementById("link_canal_whatsapp")?.value || "";

                const render = (texto) => (texto || "")
                    .replaceAll("{titulo}", titulo)
                    .replaceAll("{precio}", precio)
                    .replaceAll("{precio_descuento}", precioDesc)
                    .replaceAll("{link_video}", linkVideo)
                    .replaceAll("{link_pdf}", linkPdf)
                    .replaceAll("{link_recursos_final}", linkRecursos)
                    .replaceAll("{link_canal_whatsapp}", linkCanal);

                const preview = item.querySelector(".catalog-preview");
                if (!preview) return;
                preview.textContent =
                    `Conversación de ejemplo completa:\n\n` +
                    `BOT (info):\n${render(msgInfo)}\n\n` +
                    `BOT (envía video demo):\n${linkVideo || "[Sin link de video]"}\n\n` +
                    `BOT (proceso):\n${render(msgDespues)}\n\n` +
                    `CLIENTE: Quiero descuento\n` +
                    `BOT (descuento):\n${render(msgDesc)}\n\n` +
                    `CLIENTE: Pago con Nequi\n` +
                    `BOT (datos de pago):\nNequi: ${nequi}\nDaviplata: ${daviplata}\nLlave: ${llave}\n\n` +
                    `BOT (pedir comprobante):\n${render(msgComprobante)}\n\n` +
                    `BOT (entrega final):\n${render(msgEntrega)}`;
            });
        }

        function activarDnDCatalogo() {
            const contenedor = document.getElementById("catalogo-editor");
            if (!contenedor) return;
            let dragging = null;

            contenedor.querySelectorAll(".catalogo-item").forEach((item) => {
                item.addEventListener("dragstart", (e) => {
                    dragging = item;
                    item.classList.add("dragging");
                    e.dataTransfer.effectAllowed = "move";
                });

                item.addEventListener("dragend", () => {
                    item.classList.remove("dragging");
                    dragging = null;
                    refrescarSelectorCatalogo();
                });

                item.addEventListener("dragover", (e) => {
                    e.preventDefault();
                });

                item.addEventListener("drop", (e) => {
                    e.preventDefault();
                    if (!dragging || dragging === item) return;
                    const rect = item.getBoundingClientRect();
                    const before = (e.clientY - rect.top) < rect.height / 2;
                    if (before) {
                        contenedor.insertBefore(dragging, item);
                    } else {
                        contenedor.insertBefore(dragging, item.nextSibling);
                    }
                    refrescarSelectorCatalogo();
                });
            });
        }

        function limpiarResaltadoTutorial() {
            document.querySelectorAll(".tutorial-focus").forEach((el) => el.classList.remove("tutorial-focus"));
        }

        function posicionarTarjetaTutorial(target) {
            const card = document.getElementById("tutorial-card");
            if (!card || !target) return;
            const rect = target.getBoundingClientRect();
            const margen = 12;
            let top = rect.bottom + margen;
            let left = Math.min(Math.max(margen, rect.left), window.innerWidth - card.offsetWidth - margen);
            if (top + card.offsetHeight > window.innerHeight - margen) {
                top = Math.max(margen, rect.top - card.offsetHeight - margen);
            }
            card.style.top = `${top}px`;
            card.style.left = `${left}px`;
        }

        function mostrarPasoTutorial(indice) {
            if (indice < 0 || indice >= PASOS_TUTORIAL.length) return;
            tutorialPasoActual = indice;
            const paso = PASOS_TUTORIAL[indice];
            if (window.activarTabPanel && paso.tab) {
                window.activarTabPanel(paso.tab);
            }

            const overlay = document.getElementById("tutorial-overlay");
            const card = document.getElementById("tutorial-card");
            const titulo = document.getElementById("tutorial-titulo");
            const texto = document.getElementById("tutorial-texto");
            const meta = document.getElementById("tutorial-meta");
            if (!overlay || !card || !titulo || !texto || !meta) return;

            overlay.classList.remove("is-hidden");
            card.classList.remove("is-hidden");
            titulo.textContent = paso.titulo;
            texto.textContent = paso.texto;
            meta.textContent = `Paso ${indice + 1} de ${PASOS_TUTORIAL.length}`;

            setTimeout(() => {
                limpiarResaltadoTutorial();
                const target = document.querySelector(paso.selector) || document.getElementById(paso.tab);
                if (!target) {
                    posicionarTarjetaTutorial(document.body);
                    return;
                }
                target.classList.add("tutorial-focus");
                target.scrollIntoView({ behavior: "smooth", block: "center" });
                setTimeout(() => posicionarTarjetaTutorial(target), 150);
            }, 120);
        }

        function iniciarTutorialPanel() {
            mostrarPasoTutorial(0);
        }

        function cerrarTutorialPanel() {
            const overlay = document.getElementById("tutorial-overlay");
            const card = document.getElementById("tutorial-card");
            if (overlay) overlay.classList.add("is-hidden");
            if (card) card.classList.add("is-hidden");
            limpiarResaltadoTutorial();
        }

        function tutorialSiguiente() {
            if (tutorialPasoActual >= PASOS_TUTORIAL.length - 1) {
                cerrarTutorialPanel();
                return;
            }
            mostrarPasoTutorial(tutorialPasoActual + 1);
        }

        function tutorialAnterior() {
            if (tutorialPasoActual <= 0) return;
            mostrarPasoTutorial(tutorialPasoActual - 1);
        }

        (function () {
            const tabs = Array.from(document.querySelectorAll(".tab-link"));
            const sections = {
                configuracion: document.getElementById("configuracion"),
                pendientes: document.getElementById("pendientes"),
                catalogos: document.getElementById("catalogos"),
                reportes: document.getElementById("reportes"),
                clientes: document.getElementById("clientes"),
                difusion: document.getElementById("difusion"),
                prueba_meta: document.getElementById("prueba_meta")
            };

            function activarTab(tabId) {
                Object.entries(sections).forEach(([key, section]) => {
                    if (!section) return;
                    if (key === tabId) {
                        section.classList.remove("is-hidden");
                    } else {
                        section.classList.add("is-hidden");
                    }
                });

                tabs.forEach((btn) => {
                    const activo = btn.getAttribute("data-tab") === tabId;
                    btn.classList.toggle("active", activo);
                });
            }

            window.activarTabPanel = activarTab;

            tabs.forEach((btn) => {
                btn.addEventListener("click", function (event) {
                    event.preventDefault();
                    const tabId = btn.getAttribute("data-tab");
                    activarTab(tabId);
                    if (history && history.replaceState) {
                        history.replaceState(null, "", "#" + tabId);
                    }
                });
            });

            const hash = (window.location.hash || "").replace("#", "");
            const tabInicial = "{{ tab_activa }}";
            if (sections[tabInicial]) {
                activarTab(tabInicial);
            } else if (sections[hash]) {
                activarTab(hash);
            } else {
                activarTab("configuracion");
            }

            document.addEventListener("input", function (e) {
                if (e.target.closest(".catalogo-item")) {
                    refrescarSelectorCatalogo();
                    refrescarPreviewCatalogo();
                }
            });

            activarDnDCatalogo();
            refrescarSelectorCatalogo();
            refrescarPreviewCatalogo();

            if (AUTO_INICIAR_TUTORIAL) {
                setTimeout(() => iniciarTutorialPanel(), 350);
            }
        })();
    </script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""

PENDIENTES_HTML = """
{% if pendientes_agrupados %}
{% for grupo in pendientes_agrupados %}
<h3 style="margin-bottom: 8px;">{{ grupo.titulo }} ({{ grupo.clave }})</h3>
<table class="tabla">
    <thead>
        <tr>
            <th>Numero</th>
            <th>Producto</th>
            <th>Metodo</th>
            <th>Archivo</th>
            <th>Fecha</th>
            <th>Acciones</th>
        </tr>
    </thead>
    <tbody>
        {% for solicitud in grupo["items"] %}
        <tr>
            <td>{{ solicitud.numero }}</td>
            <td>{{ solicitud.producto_titulo or solicitud.catalogo_clave or 'General' }}</td>
            <td>{{ solicitud.metodo_pago }}</td>
            <td>{{ solicitud.tipo_archivo }}</td>
            <td>{{ solicitud.fecha }}</td>
            <td>
                <div class="acciones-pendiente">
                    <form method="POST" action="{{ url_for('admin_aprobar', numero=solicitud.numero) }}">
                        <button class="btn-small" type="submit">Aprobar y entregar</button>
                    </form>
                    <form method="POST" action="{{ url_for('admin_rechazar', numero=solicitud.numero) }}">
                        <button class="btn-small button-danger" type="submit">Rechazar</button>
                    </form>
                </div>
            </td>
        </tr>
        {% endfor %}
    </tbody>
</table>
{% endfor %}
{% else %}
<p>No hay comprobantes pendientes en este momento.</p>
{% endif %}
"""


@app.route('/admin', methods=['GET', 'POST'])
def admin():
    global catalogo_productos
    mensaje_exito = request.args.get("msg")
    accion_form = request.form.get("accion_form", "") if request.method == 'POST' else ""
    fecha_desde = request.args.get("fecha_desde", "")
    fecha_hasta = request.args.get("fecha_hasta", "")
    producto_reporte = request.args.get("producto_reporte", "")
    tab_activa = request.form.get("tab_destino") if request.method == 'POST' else request.args.get("tab", "configuracion")
    if not tab_activa:
        tab_activa = request.args.get("tab", "configuracion")
    if request.method == 'POST':
        if accion_form in {"", "configuracion"}:
            for campo in DEFAULT_DATOS_BOT:
                if campo in {"visitas_info", "catalogo_plantillas_meta_custom"}:
                    continue
                if campo in {"ocultar_tutorial_panel", "auto_iniciar_tutorial"}:
                    continue
                valor = request.form.get(campo)
                if valor is not None:
                    datos_bot[campo] = valor.strip()
            datos_bot["id_telefono"] = normalizar_numero_whatsapp(datos_bot.get("id_telefono", ""))
            datos_bot["numero_admin"] = normalizar_numero_whatsapp(datos_bot.get("numero_admin", ""))
            datos_bot["ocultar_tutorial_panel"] = request.form.get("ocultar_tutorial_panel") == "on"
            datos_bot["auto_iniciar_tutorial"] = request.form.get("auto_iniciar_tutorial") == "on"
            # Nuevo campo: tutorial_habilitado. Si el checkbox está desmarcado, el tutorial no arranca al cargar.
            datos_bot["tutorial_habilitado"] = request.form.get("tutorial_habilitado") == "on"

            if request.form.get("reset_plantillas_meta") == "1":
                datos_bot["catalogo_plantillas_meta_custom"] = {}
                mensaje_exito = "Configuracion de plantillas Meta restaurada a valores por defecto."
            else:
                catalogo_actual = obtener_catalogo_plantillas_meta()
                custom_catalogo = {}
                for clave, plantilla in catalogo_actual.items():
                    nombre_form = request.form.get(f"tpl_nombre_{clave}")
                    idioma_form = request.form.get(f"tpl_idioma_{clave}")
                    texto_form = request.form.get(f"tpl_texto_{clave}")
                    variables_form = request.form.get(f"tpl_variables_{clave}")
                    if nombre_form is None and idioma_form is None and texto_form is None and variables_form is None:
                        continue
                    nombre_meta = (nombre_form or "").strip() or plantilla.get("nombre_meta", "")
                    idioma_meta = (idioma_form or "").strip() or plantilla.get("idioma", "es")
                    texto_referencia = (texto_form or "").strip() or plantilla.get("texto_referencia", "")
                    variables_ordenadas = [v.strip() for v in (variables_form or "").split(",") if v.strip()]
                    if not variables_ordenadas:
                        variables_ordenadas = plantilla.get("variables_ordenadas", [])
                    custom_catalogo[clave] = {
                        "nombre_meta": nombre_meta,
                        "idioma": idioma_meta,
                        "texto_referencia": texto_referencia,
                        "variables_ordenadas": variables_ordenadas,
                    }
                if custom_catalogo:
                    datos_bot["catalogo_plantillas_meta_custom"] = custom_catalogo

            # Procesar carga de video
            if 'archivo_video' in request.files:
                archivo = request.files['archivo_video']
                if archivo and archivo.filename != '':
                    ruta_video = guardar_archivo_media(archivo, "video_demo")
                    if ruta_video:
                        datos_bot['link_video'] = ruta_video
                        mensaje_exito = "Video demo actualizado correctamente."

            # Procesar carga de PDF
            if 'archivo_pdf_demo' in request.files:
                archivo = request.files['archivo_pdf_demo']
                if archivo and archivo.filename != '':
                    ruta_pdf = guardar_archivo_media(archivo, "pdf_demo")
                    if ruta_pdf:
                        datos_bot['link_pdf_demo'] = ruta_pdf
                        if not mensaje_exito or "Video" not in mensaje_exito:
                            mensaje_exito = "PDF demo actualizado correctamente."
                        else:
                            mensaje_exito = "Video y PDF demo actualizados correctamente."

        if accion_form in {"", "catalogo"}:
            # Procesar editor visual de catalogo
            indices_catalogo = request.form.getlist("catalog_index")
            original_keys = request.form.getlist("catalog_original_key")
            keys = request.form.getlist("catalog_key")
            ids_producto = request.form.getlist("catalog_id")
            opciones_producto = request.form.getlist("catalog_opcion")
            titulos = request.form.getlist("catalog_titulo")
            plantillas_meta = request.form.getlist("catalog_plantilla_meta")
            precios = request.form.getlist("catalog_precio")
            precios_descuento = request.form.getlist("catalog_precio_descuento")
            links_video = request.form.getlist("catalog_link_video")
            links_pdf = request.form.getlist("catalog_link_pdf")
            links_entrega = request.form.getlist("catalog_link_entrega")
            palabras = request.form.getlist("catalog_palabras")
            msg_info = request.form.getlist("catalog_msg_info")
            msg_bienvenida_a = request.form.getlist("catalog_msg_bienvenida_a")
            msg_bienvenida_b = request.form.getlist("catalog_msg_bienvenida_b")
            msg_despues_demo = request.form.getlist("catalog_msg_despues_demo")
            msg_descuento = request.form.getlist("catalog_msg_descuento")
            msg_confirmacion_pago = request.form.getlist("catalog_msg_confirmacion_pago")
            msg_entrega_final = request.form.getlist("catalog_msg_entrega_final")

            cantidad = max(
                len(keys), len(ids_producto), len(opciones_producto), len(titulos), len(plantillas_meta), len(precios), len(precios_descuento), len(links_video), len(links_pdf), len(links_entrega),
                len(palabras), len(original_keys), len(indices_catalogo), len(msg_info), len(msg_bienvenida_a), len(msg_bienvenida_b), len(msg_despues_demo),
                len(msg_descuento), len(msg_confirmacion_pago), len(msg_entrega_final)
            )
            nuevo_catalogo = {}
            for idx in range(cantidad):
                idx_archivo = indices_catalogo[idx] if idx < len(indices_catalogo) else str(idx)
                original_key = original_keys[idx] if idx < len(original_keys) else ""
                key_raw = keys[idx] if idx < len(keys) else ""
                clave = normalizar_clave_catalogo(key_raw) or normalizar_clave_catalogo(original_key) or f"producto_{idx+1}"
                producto_id = normalizar_clave_catalogo(ids_producto[idx] if idx < len(ids_producto) else "") or clave
                opcion = (opciones_producto[idx] if idx < len(opciones_producto) else "").strip() or str(idx + 1)
                titulo = (titulos[idx] if idx < len(titulos) else "").strip() or clave.replace("_", " ").title()
                plantilla_meta = (plantillas_meta[idx] if idx < len(plantillas_meta) else "").strip() or f"info_{producto_id}_v1"
                precio = (precios[idx] if idx < len(precios) else "").strip() or "0"
                precio_desc = (precios_descuento[idx] if idx < len(precios_descuento) else "").strip() or precio
                link_video = (links_video[idx] if idx < len(links_video) else "").strip()
                link_pdf = (links_pdf[idx] if idx < len(links_pdf) else "").strip()
                link_entrega = (links_entrega[idx] if idx < len(links_entrega) else "").strip()
                palabras_raw = (palabras[idx] if idx < len(palabras) else "")
                palabras_clave = [p.strip() for p in re.split(r"[,;\n]", palabras_raw) if p.strip()]

                bienvenida_a = (msg_bienvenida_a[idx] if idx < len(msg_bienvenida_a) else "").strip()
                bienvenida_b = (msg_bienvenida_b[idx] if idx < len(msg_bienvenida_b) else "").strip()
                mensajes = {
                    "bienvenida": bienvenida_a or bienvenida_b,
                    "bienvenida_a": bienvenida_a,
                    "bienvenida_b": bienvenida_b,
                    "info": (msg_info[idx] if idx < len(msg_info) else "").strip(),
                    "despues_demo": (msg_despues_demo[idx] if idx < len(msg_despues_demo) else "").strip(),
                    "descuento": (msg_descuento[idx] if idx < len(msg_descuento) else "").strip(),
                    "confirmacion_pago": (msg_confirmacion_pago[idx] if idx < len(msg_confirmacion_pago) else "").strip(),
                    "entrega_final": (msg_entrega_final[idx] if idx < len(msg_entrega_final) else "").strip(),
                }

                archivo_video = request.files.get(f"catalog_archivo_video_{idx_archivo}") or request.files.get(f"catalog_archivo_video_{original_key}")
                if archivo_video and archivo_video.filename:
                    ruta_video = guardar_archivo_media(archivo_video, f"catalogo_{clave}_video")
                    if ruta_video:
                        link_video = ruta_video

                archivo_pdf = request.files.get(f"catalog_archivo_pdf_{idx_archivo}") or request.files.get(f"catalog_archivo_pdf_{original_key}")
                if archivo_pdf and archivo_pdf.filename:
                    ruta_pdf = guardar_archivo_media(archivo_pdf, f"catalogo_{clave}_pdf")
                    if ruta_pdf:
                        link_pdf = ruta_pdf

                texto_info_meta_form = request.form.getlist("catalog_texto_info_meta")
                texto_info_meta_val = (texto_info_meta_form[idx] if idx < len(texto_info_meta_form) else "").strip()

                nuevo_catalogo[clave] = {
                    "id": producto_id,
                    "opcion": opcion,
                    "titulo": titulo,
                    "plantilla_info_meta": plantilla_meta,
                    "texto_info_meta": texto_info_meta_val,
                    "precio_normal": precio,
                    "precio_descuento": precio_desc,
                    "link_video": link_video,
                    "link_pdf": link_pdf,
                    "link_drive_final": link_entrega,
                    "link_canal_whatsapp": datos_bot.get("link_canal_whatsapp", ""),
                    "palabras_clave": palabras_clave or [clave],
                    "mensajes": mensajes,
                }

            if nuevo_catalogo:
                catalogo_productos = nuevo_catalogo
                guardar_catalogo()
            else:
                mensaje_exito = "Debes tener al menos un producto en el catalogo."
        
        guardar_datos_bot()
        if not mensaje_exito:
            mensaje_exito = "Configuración guardada correctamente."

    ventas_filtradas = filtrar_ventas_por_fecha(fecha_desde, fecha_hasta, producto_reporte)
    resumen_filtrado = obtener_resumen_ventas_de(ventas_filtradas)
    catalogo_plantillas_meta = obtener_catalogo_plantillas_meta()
    vista_previa_plantillas_meta = construir_vista_previa_plantillas_meta(catalogo_plantillas_meta)
    manual_meta_rows = construir_manual_meta_rows()
    clientes_lista = sorted(
        clientes_vendidos.values(),
        key=lambda c: c.get("ultima_venta", ""),
        reverse=True
    )

    return render_template_string(
        PANEL_HTML,
        mensaje=mensaje_exito,
        pendientes_total=len(solicitudes_comprobante),
        pendientes_agrupados=obtener_solicitudes_agrupadas(),
        resumen_ventas=resumen_filtrado,
        resumen_producto=obtener_resumen_por_producto(ventas_filtradas),
        analitica_clientes_producto=obtener_analitica_clientes_por_producto(ventas_registradas),
        ventas=sorted(ventas_filtradas, key=lambda v: v.get("fecha_aprobacion", ""), reverse=True)[:100],
        clientes_total=len(clientes_lista),
        clientes_lista=clientes_lista[:500],
        catalogo_items=[
            {
                "key": k,
                "id": v.get("id", k),
                "opcion": v.get("opcion", ""),
                "titulo": v.get("titulo", k),
                "plantilla_info_meta": v.get("plantilla_info_meta", f"info_{v.get('id', k)}_v1"),
                "texto_info_meta": v.get("texto_info_meta", ""),
                "precio_normal": v.get("precio_normal", "0"),
                "precio_descuento": v.get("precio_descuento", v.get("precio_normal", "0")),
                "link_video": v.get("link_video", ""),
                "link_pdf": v.get("link_pdf", ""),
                "link_drive_final": v.get("link_drive_final", ""),
                "palabras_clave": v.get("palabras_clave", []),
                "mensajes": v.get("mensajes", DEFAULT_MENSAJES_PRODUCTO),
            }
            for k, v in catalogo_productos.items()
        ],
        tab_activa=tab_activa,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta,
        producto_reporte=producto_reporte,
        mensajes_default_json=json.dumps(DEFAULT_MENSAJES_PRODUCTO, ensure_ascii=False),
        texto_info_meta_base_json=json.dumps(TEXTO_INFO_META_BASE_NUEVO_PRODUCTO, ensure_ascii=False),
        catalogo_plantillas_meta=catalogo_plantillas_meta,
        vista_previa_plantillas_meta=vista_previa_plantillas_meta,
        manual_meta_rows=manual_meta_rows,
        estado_video=obtener_estado_recurso(datos_bot.get("link_video", "")),
        estado_pdf=obtener_estado_recurso(datos_bot.get("link_pdf_demo", "")),
        **datos_bot
    )


@app.route('/admin/pendientes', methods=['GET'])
def admin_pendientes():
    return render_template_string(PENDIENTES_HTML, pendientes_agrupados=obtener_solicitudes_agrupadas())


@app.route('/admin/difusion', methods=['POST'])
def admin_difusion():
    mensaje_difusion = (request.form.get("mensaje_difusion") or "").strip()
    link_material = (request.form.get("link_material_difusion") or "").strip()
    usar_clientes_vendidos = request.form.get("usar_clientes_vendidos") == "on"
    segmento_tipo_precio = (request.form.get("segmento_tipo_precio") or "").strip()
    segmento_metodo_pago = (request.form.get("segmento_metodo_pago") or "").strip()
    segmento_producto = (request.form.get("segmento_producto") or "").strip()

    numeros = set()
    if usar_clientes_vendidos:
        numeros.update(obtener_numeros_segmentados(segmento_tipo_precio, segmento_metodo_pago, segmento_producto))

    numeros.update(extraer_numeros_texto(request.form.get("numeros_manual", "")))
    archivo_numeros = request.files.get("archivo_numeros")
    numeros.update(extraer_numeros_archivo(archivo_numeros))

    archivo_material = request.files.get("archivo_material_difusion")
    if archivo_material and archivo_material.filename:
        ruta_material = guardar_archivo_media(archivo_material, f"difusion_{int(time.time())}")
        if ruta_material:
            link_material = ruta_material

    if not numeros:
        return redirect(url_for('admin', tab='difusion', msg="No se encontraron numeros para difusion."))

    enviados_ok = 0
    errores = 0
    for numero in sorted(numeros):
        try:
            ok_texto = True
            if mensaje_difusion:
                resp_texto = enviar_mensaje(numero, mensaje_difusion)
                ok_texto = bool(resp_texto and resp_texto.ok)

            ok_material = True
            if link_material:
                resp_material = enviar_material_difusion(numero, link_material)
                ok_material = bool(resp_material and resp_material.ok)

            if ok_texto and ok_material:
                enviados_ok += 1
            else:
                errores += 1
        except Exception:
            errores += 1

    return redirect(
        url_for(
            'admin',
            tab='difusion',
            msg=f"Difusion finalizada. Exitosos: {enviados_ok}. Errores: {errores}. Total: {len(numeros)}. Segmento precio={segmento_tipo_precio or 'todos'}, metodo={segmento_metodo_pago or 'todos'}, producto={segmento_producto or 'todos'}."
        )
    )


@app.route('/test_meta', methods=['GET', 'POST'])
def test_meta():
    """Ruta para probar plantillas Meta.
    GET: redirige al panel en pestaña prueba_meta.
    POST: envía la plantilla y redirige con el resultado.
    Correcciones aplicadas:
    - Acepta GET para evitar error 405.
    - Variables {{1}} y {{2}} opcionales (plantillas sin variables funcionan).
    - Usa API v19.0 (más reciente y estable).
    - Muestra respuesta completa de Meta en mensaje de resultado.
    - No requiere CSRF ni dependencias externas para funcionar.
    """
    global datos_bot

    if request.method == 'GET':
        return redirect(url_for('admin', tab='prueba_meta'))

    numero_destino = normalizar_numero_whatsapp(request.form.get("numero_destino") or "")
    nombre_plantilla = (request.form.get("test_template_nombre") or "hello_world").strip()
    idioma = (request.form.get("test_template_idioma") or "en_US").strip()
    var1 = (request.form.get("test_template_var1") or "").strip()
    var2 = (request.form.get("test_template_var2") or "").strip()

    # Persistir los campos de configuración de la prueba
    datos_bot["test_template_nombre"] = nombre_plantilla
    datos_bot["test_template_idioma"] = idioma
    datos_bot["test_template_var1"] = var1
    datos_bot["test_template_var2"] = var2
    try:
        guardar_datos_bot()
    except Exception:
        pass

    if not numero_destino:
        return redirect(url_for('admin', tab='prueba_meta', msg="Error: Debes ingresar un número de destino válido (solo dígitos, código de país incluido)."))
    if not nombre_plantilla:
        return redirect(url_for('admin', tab='prueba_meta', msg="Error: Debes ingresar el nombre de la plantilla."))

    # Usar credenciales activas del bot
    id_telefono = (datos_bot.get("id_telefono") or ID_TELEFONO_DEFAULT).strip() or ID_TELEFONO_DEFAULT
    token_meta = (datos_bot.get("token_meta") or TOKEN_META_DEFAULT).strip() or TOKEN_META_DEFAULT

    # API v19.0: más reciente y compatible
    url_api = f"https://graph.facebook.com/v19.0/{id_telefono}/messages"
    headers = {
        "Authorization": f"Bearer {token_meta}",
        "Content-Type": "application/json"
    }

    # Construir componentes de variables solo si se proporcionaron
    parameters = []
    if var1:
        parameters.append({"type": "text", "text": var1})
    if var2:
        parameters.append({"type": "text", "text": var2})

    template_payload = {
        "name": nombre_plantilla,
        "language": {"code": idioma}
    }
    if parameters:
        template_payload["components"] = [
            {
                "type": "body",
                "parameters": parameters
            }
        ]

    payload = {
        "messaging_product": "whatsapp",
        "to": numero_destino,
        "type": "template",
        "template": template_payload
    }

    try:
        resp = requests.post(url_api, headers=headers, json=payload, timeout=30)
        status_code = resp.status_code
        response_text = resp.text[:500]  # Limitar para evitar URLs muy largas
        if status_code in (200, 201):
            msg = f"OK Plantilla '{nombre_plantilla}' enviada exitosamente a {numero_destino}. HTTP {status_code}."
        else:
            try:
                error_data = resp.json()
                error_info = error_data.get("error", {}) if isinstance(error_data, dict) else {}
                error_code = error_info.get("code", status_code)
                error_fbtrace = error_info.get("fbtrace_id", "")
                error_message = error_info.get("message") or response_text
                msg = f"Error Meta codigo={error_code} HTTP={status_code}: {error_message}"
                if error_fbtrace:
                    msg += f" | fbtrace_id={error_fbtrace}"
            except Exception:
                msg = f"Error Meta HTTP {status_code}: {response_text}"
    except requests.exceptions.Timeout:
        msg = "Error de conexion: timeout al contactar la API de Meta. Verifica el token y el ID de telefono."
    except Exception as exc:
        msg = f"Error de conexion al enviar la plantilla: {str(exc)[:200]}"

    return redirect(url_for('admin', tab='prueba_meta', msg=msg))


@app.route('/admin/reportes/excel', methods=['GET'])
def admin_reporte_excel():
    fecha_desde = request.args.get("fecha_desde", "")
    fecha_hasta = request.args.get("fecha_hasta", "")
    producto_reporte = request.args.get("producto_reporte", "")
    ventas_filtradas = filtrar_ventas_por_fecha(fecha_desde, fecha_hasta, producto_reporte)
    archivo = generar_excel_ventas(ventas_filtradas)
    nombre = f"reporte_ventas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        archivo,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/admin/reportes/pdf', methods=['GET'])
def admin_reporte_pdf():
    fecha_desde = request.args.get("fecha_desde", "")
    fecha_hasta = request.args.get("fecha_hasta", "")
    producto_reporte = request.args.get("producto_reporte", "")
    ventas_filtradas = filtrar_ventas_por_fecha(fecha_desde, fecha_hasta, producto_reporte)
    archivo = generar_pdf_consolidado_comprobantes(ventas_filtradas)
    nombre = f"consolidado_comprobantes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(
        archivo,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/pdf"
    )


@app.route('/admin/plantillas-meta/export', methods=['GET'])
def admin_exportar_plantillas_meta():
    archivo = exportar_catalogo_plantillas_meta_json()
    nombre = f"catalogo_plantillas_meta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    return send_file(
        archivo,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/json"
    )


@app.route('/media/<filename>', methods=['GET'])
def media(filename):
    """Sirve archivos de la carpeta media (videos, PDFs, etc)"""
    try:
        filepath = MEDIA_DIR / secure_filename(filename)
        if filepath.exists() and filepath.parent == MEDIA_DIR:
            return send_file(str(filepath), as_attachment=True)
    except:
        pass
    return "Archivo no encontrado", 404


@app.route('/admin/aprobar/<numero>', methods=['POST'])
def admin_aprobar(numero):
    solicitud = solicitudes_comprobante.pop(numero, None)
    if solicitud:
        guardar_estado_runtime()
        registrar_venta_aprobada(numero, solicitud, aprobado_desde="panel")
        if solicitud.get("catalogo_clave"):
            estados_clientes.setdefault(numero, {})["catalogo_activo"] = solicitud.get("catalogo_clave")
        enviar_recursos_finales(numero)
        return redirect(url_for('admin', msg=f"Comprobante aprobado para {numero} y acceso enviado."))
    return redirect(url_for('admin', msg=f"No habia solicitud pendiente para {numero}."))


@app.route('/admin/rechazar/<numero>', methods=['POST'])
def admin_rechazar(numero):
    solicitud = solicitudes_comprobante.pop(numero, None)
    if solicitud:
        estados_clientes[numero] = {
            "esperando_comprobante": True,
            "metodo_pago": solicitud.get("metodo_pago", "No definido")
        }
        guardar_estado_runtime()
        enviar_mensaje(numero, "Tu comprobante no se pudo validar. Por favor reenvialo con buena calidad para aprobar tu acceso.")
        return redirect(url_for('admin', msg=f"Comprobante rechazado para {numero}."))
    return redirect(url_for('admin', msg=f"No habia solicitud pendiente para {numero}."))


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.getenv("PORT", 5000)))
