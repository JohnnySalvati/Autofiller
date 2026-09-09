"""Lectura de comprobantes en foto con un modelo de vision de Claude.

Existe porque la mitad de los usuarios no puede usar AutoFiller: la mayoria de
los comprobantes que reciben son fotos de celular, no PDF. Decision del cliente
(2026-09-09): los comprobantes pueden salir de la red de AOMAOSAM hacia un
proveedor externo, asi que se puede usar visión en vez de OCR local.

La API key vive solo en el servidor (variable de entorno ANTHROPIC_API_KEY); no
se distribuye en ningun ejecutable, que era uno de los motivos para partir la
app en servidor + agente.

El QR de ARCA manda sobre lo que devuelva el modelo para los datos fiscales
(ver qr.py): la vision se usa para lo que el QR no trae.
"""

import base64
import io
import json
import os

# Lado maximo recomendado por Anthropic para imagenes: mas grande solo agrega
# tokens sin mejorar la lectura.
LADO_MAXIMO = 1568

MODELO = os.environ.get("AUTOFILLER_MODELO_VISION", "claude-opus-5")
ESFUERZO = os.environ.get("AUTOFILLER_ESFUERZO_VISION", "medium")

INSTRUCCIONES = """Sos un lector de facturas electrónicas argentinas (ARCA/AFIP).
Extraés los campos de la imagen y devolvés el JSON pedido, sin comentarios.

Reglas:
- Todos los campos son texto. Si un dato no aparece en la imagen, devolvé "".
- No inventes ni completes datos: si no lo leés con certeza, devolvé "".
- Las fechas van en formato DD/MM/AAAA.
- codigo_arca es el número del recuadro del tipo de comprobante ("COD. 011" -> "11").
- cuit_emisor es el CUIT de quien EMITE la factura (arriba, junto a la razón
  social del prestador), NO el del receptor.
- punto_venta y nro_comprobante van sin ceros a la izquierda.
- periodo_hasta es la fecha rotulada "Hasta:" del período facturado.
- descripcion es el texto del detalle facturado, en una sola línea.
- domicilio_comercial es el "Domicilio Comercial:" del emisor, completo y tal
  cual está impreso (calle - localidad, provincia). No lo abrevies ni lo
  corrijas: se usa para deducir el centro de costos.
- importe_total va sin símbolo de moneda ni separador de miles, con coma
  decimal: "161023,60"."""

CAMPOS = [
    "codigo_arca",
    "cuit_emisor",
    "punto_venta",
    "nro_comprobante",
    "fecha_emision",
    "periodo_hasta",
    "cae",
    "descripcion",
    "domicilio_comercial",
    "importe_total",
]

ESQUEMA = {
    "type": "json_schema",
    "schema": {
        "type": "object",
        "properties": {c: {"type": "string"} for c in CAMPOS},
        "required": CAMPOS,
        "additionalProperties": False,
    },
}


class VisionNoDisponible(RuntimeError):
    """No hay API key o no esta instalado el SDK: se avisa, no se rompe."""


def disponible():
    try:
        import anthropic  # noqa: F401
    except ImportError:
        return False
    return bool(os.environ.get("ANTHROPIC_API_KEY"))


def a_jpeg_base64(imagen):
    """Achica la imagen y la codifica en JPEG base64 para mandarla a la API."""
    from PIL import Image

    imagen = imagen.convert("RGB")
    lado = max(imagen.size)
    if lado > LADO_MAXIMO:
        escala = LADO_MAXIMO / lado
        nuevo = (max(1, int(imagen.width * escala)), max(1, int(imagen.height * escala)))
        imagen = imagen.resize(nuevo, Image.LANCZOS)
    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=85)
    return base64.standard_b64encode(buffer.getvalue()).decode("ascii")


def leer_comprobante(imagen):
    """Campos del comprobante segun el modelo de vision. Lanza VisionNoDisponible."""
    if not disponible():
        raise VisionNoDisponible(
            "El servidor no tiene configurada la lectura de fotos "
            "(falta ANTHROPIC_API_KEY)."
        )

    import anthropic

    cliente = anthropic.Anthropic()
    respuesta = cliente.messages.create(
        model=MODELO,
        max_tokens=4000,
        system=INSTRUCCIONES,
        output_config={"format": ESQUEMA, "effort": ESFUERZO},
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": a_jpeg_base64(imagen),
                    },
                },
                {"type": "text", "text": "Extraé los campos de este comprobante."},
            ],
        }],
    )

    if respuesta.stop_reason == "refusal":
        raise RuntimeError("El modelo no pudo procesar la imagen.")

    texto = next((b.text for b in respuesta.content if b.type == "text"), "")
    if not texto:
        raise RuntimeError("El modelo no devolvió datos para la imagen.")
    return json.loads(texto)
