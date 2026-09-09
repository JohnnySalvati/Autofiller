"""Lectura del QR que ARCA imprime en toda factura electronica.

El QR es un link 'https://www.afip.gob.ar/fe/qr/?p=<json en base64>' con los
datos fiscales exactos: CUIT del emisor, tipo, punto de venta, numero, fecha,
importe y CAE. En una foto de celular es mucho mas confiable que el OCR, asi
que cuando se lee, esos campos mandan sobre lo que devuelva el modelo de vision.

Lo que el QR NO trae y hay que sacar del texto o de la vision: el periodo
facturado, la descripcion del detalle y el domicilio comercial del emisor.
"""

import base64
import json
import re
from urllib.parse import parse_qs, urlparse

try:
    import zxingcpp
except ImportError:  # el QR es opcional: sin la libreria se sigue con vision
    zxingcpp = None

# Los nombres de campo del payload de ARCA, tal cual los documenta.
CAMPOS_QR = ("cuit", "ptoVta", "tipoCmp", "nroCmp", "fecha", "importe", "codAut")


def _payload_de_url(url):
    """El JSON del parametro ?p= de la URL del QR, o None si no es de ARCA."""
    partes = urlparse(url)
    if "afip.gob.ar" not in partes.netloc and "arca.gob.ar" not in partes.netloc:
        return None
    crudo = parse_qs(partes.query).get("p", [None])[0]
    if not crudo:
        return None
    # ARCA no siempre manda el padding del base64.
    crudo += "=" * (-len(crudo) % 4)
    try:
        return json.loads(base64.b64decode(crudo))
    except Exception:
        return None


def leer_qr(imagen):
    """Payload del QR de ARCA en una imagen PIL, o None.

    Se recorren todos los codigos de la imagen porque un comprobante puede traer
    otros QR (logos, links de pago) ademas del fiscal.
    """
    if zxingcpp is None:
        return None
    try:
        codigos = zxingcpp.read_barcodes(imagen)
    except Exception:
        return None
    for codigo in codigos:
        payload = _payload_de_url(codigo.text or "")
        if payload and "cuit" in payload:
            return payload
    return None


def formatear_importe(valor):
    """1234.5 -> '1234,50', el formato con el que ARCA lo imprime."""
    try:
        return f"{float(valor):.2f}".replace(".", ",")
    except (TypeError, ValueError):
        return None


def formatear_fecha(valor):
    """'2026-09-03' -> '03/09/2026'."""
    match = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(valor or ""))
    if not match:
        return None
    anio, mes, dia = match.groups()
    return f"{dia}/{mes}/{anio}"


def campos_desde_qr(payload):
    """Los campos de la factura que el QR da exactos, listos para la Factura."""
    from .modelo import TIPOS_COMPROBANTE

    if not payload:
        return {}
    codigo = str(payload.get("tipoCmp", "")).lstrip("0")
    campos = {
        "cuit": str(payload.get("cuit", "")).zfill(11) or None,
        "tipo_comprobante": TIPOS_COMPROBANTE.get(codigo),
        "punto_venta": str(payload.get("ptoVta", "")).lstrip("0") or None,
        "nro_factura": str(payload.get("nroCmp", "")).lstrip("0") or None,
        "fecha_emision": formatear_fecha(payload.get("fecha")),
        "cae": str(payload.get("codAut", "")) or None,
        "importe": formatear_importe(payload.get("importe")),
    }
    return {k: v for k, v in campos.items() if v}
