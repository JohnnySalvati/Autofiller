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
import logging
import os

registro = logging.getLogger(__name__)

# Lado maximo de la imagen que se manda. 1568 es el tope que recomienda
# Anthropic: mas grande solo agrega tokens. Pero es un TOPE, no un optimo: los
# tokens de imagen son ancho*alto/750, asi que bajarlo achica el costo de forma
# lineal y una factura de ARCA es texto con etiquetas, no una foto de detalle.
# Es configurable para poder medir donde empieza a fallar la lectura en fotos
# reales de celular, que es lo que decide el valor definitivo.
LADO_MAXIMO = int(os.environ.get("AUTOFILLER_LADO_MAXIMO", "1568"))

MODELO = os.environ.get("AUTOFILLER_MODELO_VISION", "claude-opus-5")
ESFUERZO = os.environ.get("AUTOFILLER_ESFUERZO_VISION", "medium")

# Precios de lista por millon de tokens (entrada, salida), en dolares.
# Referencia de sep-2026: sirve para que el log muestre el costo real medido en
# vez de una estimacion, no como fuente de verdad. Un modelo que no este aca se
# registra igual, solo sin la linea de costo.
PRECIOS = {
    "claude-opus-5": (5.0, 25.0),
    "claude-opus-4-8": (5.0, 25.0),
    "claude-sonnet-5": (2.0, 10.0),
    "claude-haiku-4-5": (1.0, 5.0),
}

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
    """(base64, ancho, alto) de la imagen achicada y codificada en JPEG."""
    from PIL import Image

    imagen = imagen.convert("RGB")
    lado = max(imagen.size)
    if lado > LADO_MAXIMO:
        escala = LADO_MAXIMO / lado
        nuevo = (max(1, int(imagen.width * escala)), max(1, int(imagen.height * escala)))
        imagen = imagen.resize(nuevo, Image.LANCZOS)
    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=85)
    datos = base64.standard_b64encode(buffer.getvalue()).decode("ascii")
    return datos, imagen.width, imagen.height


def _registrar_consumo(respuesta, ancho, alto):
    """Deja en el log los tokens y el costo REALES de la llamada.

    Existe porque el costo por comprobante se venia estimando, y la estimacion
    dependia de cuantos tokens ocupa la imagen (ancho*alto/750), que es
    justamente lo que se puede ajustar con LADO_MAXIMO. Con esto se mide en vez
    de calcular: es el numero que decide si conviene bajar la resolucion, bajar
    de modelo o irse a un OCR.
    """
    uso = getattr(respuesta, "usage", None)
    if uso is None:
        return

    entrada = getattr(uso, "input_tokens", 0) or 0
    salida = getattr(uso, "output_tokens", 0) or 0
    partes = [
        f"vision {MODELO} esfuerzo={ESFUERZO}",
        f"imagen {ancho}x{alto} (~{ancho * alto // 750} tok)",
        f"entrada {entrada}",
        f"salida {salida}",
    ]

    precio = PRECIOS.get(MODELO)
    if precio:
        costo = (entrada * precio[0] + salida * precio[1]) / 1_000_000
        partes.append(f"USD {costo:.5f} (x1000 = USD {costo * 1000:.2f})")

    registro.info(" | ".join(partes))


def leer_comprobante(imagen):
    """Campos del comprobante segun el modelo de vision. Lanza VisionNoDisponible."""
    if not disponible():
        raise VisionNoDisponible(
            "El servidor no tiene configurada la lectura de fotos "
            "(falta ANTHROPIC_API_KEY)."
        )

    import anthropic

    datos, ancho, alto = a_jpeg_base64(imagen)
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
                        "data": datos,
                    },
                },
                {"type": "text", "text": "Extraé los campos de este comprobante."},
            ],
        }],
    )
    _registrar_consumo(respuesta, ancho, alto)

    if respuesta.stop_reason == "refusal":
        raise RuntimeError("El modelo no pudo procesar la imagen.")

    texto = next((b.text for b in respuesta.content if b.type == "text"), "")
    if not texto:
        raise RuntimeError("El modelo no devolvió datos para la imagen.")
    return json.loads(texto)
