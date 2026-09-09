"""Extraccion de los datos de un comprobante, sea PDF o foto.

Un solo punto de entrada, extraer(nombre, contenido), que elige el camino:

  PDF con texto  -> pdfplumber + los regex de ARCA (el camino del escritorio,
                    exacto y gratis).
  PDF escaneado  -> se renderiza la primera pagina y se trata como foto.
  Foto           -> QR de ARCA para los datos fiscales + modelo de vision para
                    el periodo, la descripcion y el domicilio.

Nada se guarda en disco: el archivo entra por HTTP, se procesa en memoria y se
descarta. El escritorio dejaba un 'decrypted.pdf' en el directorio actual.
"""

import io
import os
from datetime import date

from .centro_costo import centro_costo_de_domicilio, provincia_de_domicilio
from .modelo import Factura, Resultado
from .qr import campos_desde_qr, leer_qr
from .texto import datos_desde_texto, parece_factura_arca
from .vision import VisionNoDisponible, leer_comprobante

EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp", ".bmp", ".tif", ".tiff"}
EXTENSIONES_PDF = {".pdf"}
EXTENSIONES_ACEPTADAS = EXTENSIONES_IMAGEN | EXTENSIONES_PDF

# A que resolucion se rasteriza una pagina de PDF escaneado. 200 dpi alcanza
# para que el QR se lea y el texto quede legible sin inflar la imagen.
DPI_RASTERIZADO = 200


def _abrir_imagen(contenido):
    from PIL import Image

    try:  # los iPhone mandan HEIC; sin el plugin, Pillow no lo abre
        import pillow_heif

        pillow_heif.register_heif_opener()
    except ImportError:
        pass
    return Image.open(io.BytesIO(contenido))


def _texto_del_pdf(contenido):
    """Texto del PDF, desencriptandolo en memoria si hace falta."""
    import pdfplumber
    import pikepdf

    # Las facturas de ARCA vienen con encriptacion de solo-lectura, que pikepdf
    # saca sin contrasena. pdfplumber solo no las abre.
    buffer = io.BytesIO()
    with pikepdf.open(io.BytesIO(contenido)) as pdf:
        pdf.save(buffer)
    buffer.seek(0)
    with pdfplumber.open(buffer) as documento:
        return "".join(pagina.extract_text() or "" for pagina in documento.pages)


def _primera_pagina_como_imagen(contenido):
    """Rasteriza la primera pagina del PDF, para los que vienen escaneados."""
    import pypdfium2

    documento = pypdfium2.PdfDocument(io.BytesIO(contenido))
    try:
        return documento[0].render(scale=DPI_RASTERIZADO / 72).to_pil()
    finally:
        documento.close()


def _completar_domicilio(datos):
    """Deriva provincia y centro de costos del domicilio ya extraido."""
    domicilio = datos.get("domicilio")
    nombre, codigo = centro_costo_de_domicilio(domicilio)
    datos["provincia"] = provincia_de_domicilio(domicilio)
    datos["centro_costo"] = codigo
    datos["centro_costo_nombre"] = nombre
    return datos


def _datos_desde_imagen(imagen):
    """(datos, avisos, origen) leyendo QR + vision sobre una imagen."""
    from .modelo import TIPOS_COMPROBANTE

    avisos = []
    payload = leer_qr(imagen)
    if not payload:
        avisos.append(
            "No se pudo leer el QR de ARCA: los datos fiscales salen de la "
            "lectura de la imagen. Verificá CUIT, número, CAE e importe antes "
            "de confirmar."
        )

    leido = leer_comprobante(imagen)

    codigo = (leido.get("codigo_arca") or "").lstrip("0")
    periodo_hasta = leido.get("periodo_hasta") or None
    if not periodo_hasta:
        avisos.append(
            "No se leyó el período facturado ('Hasta:'). Quedaron sin cargar el "
            "vencimiento y el devengamiento."
        )

    datos = {
        "cuit": leido.get("cuit_emisor") or None,
        "tipo_comprobante": TIPOS_COMPROBANTE.get(codigo),
        "punto_venta": (leido.get("punto_venta") or "").lstrip("0") or None,
        "nro_factura": (leido.get("nro_comprobante") or "").lstrip("0") or None,
        "fecha_recepcion": date.today().strftime("%d/%m/%Y"),
        "fecha_emision": leido.get("fecha_emision") or None,
        "fecha_vencimiento": periodo_hasta,
        "fecha_devengamiento": periodo_hasta,
        "cae": leido.get("cae") or None,
        "descripcion": " ".join((leido.get("descripcion") or "").split()),
        "importe": leido.get("importe_total") or None,
        "domicilio": leido.get("domicilio_comercial") or None,
    }

    # El QR es dato fiscal firmado por ARCA: pisa lo que haya leido el modelo.
    datos.update(campos_desde_qr(payload))
    return _completar_domicilio(datos), avisos


def extraer(nombre, contenido):
    """Resultado de la extraccion para un archivo. Nunca lanza."""
    resultado = Resultado(archivo=nombre)
    extension = os.path.splitext(nombre)[1].lower()

    if extension not in EXTENSIONES_ACEPTADAS:
        resultado.error = f"Formato no soportado ({extension or 'sin extensión'})."
        return resultado

    try:
        if extension in EXTENSIONES_PDF:
            texto = _texto_del_pdf(contenido)
            if parece_factura_arca(texto):
                datos, avisos = datos_desde_texto(texto)
                resultado.origen = "pdf-texto"
            else:
                # PDF sin texto: es un escaneo, va por el mismo camino que la foto.
                datos, avisos = _datos_desde_imagen(_primera_pagina_como_imagen(contenido))
                resultado.origen = "pdf-imagen"
        else:
            datos, avisos = _datos_desde_imagen(_abrir_imagen(contenido))
            resultado.origen = "foto"
    except VisionNoDisponible as e:
        resultado.error = str(e)
        return resultado
    except Exception as e:
        resultado.error = f"No se pudo leer el comprobante: {e}"
        return resultado

    resultado.factura = Factura(**datos)
    resultado.avisos = avisos

    faltantes = resultado.factura.faltantes()
    if faltantes:
        etiquetas = {
            "cuit": "CUIT del emisor",
            "tipo_comprobante": "tipo de comprobante",
            "punto_venta": "punto de venta",
            "nro_factura": "número de comprobante",
        }
        resultado.error = (
            "Faltan datos que SISalud necesita sí o sí: "
            + ", ".join(etiquetas[c] for c in faltantes)
            + ". Completalos abajo o cargá el comprobante a mano."
        )
    return resultado
