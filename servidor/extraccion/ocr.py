"""OCR de Google Cloud Vision: el camino barato para leer una imagen.

Existe porque el 2026-09-10 se midio, contra las 63 muestras, que el texto que
devuelve Cloud Vision lo parsean bien los mismos regex de ARCA que usa el camino
PDF (`texto.py`): 63/63 en periodo, centro de costos y provincia; 59/63 en
domicilio con diferencias cosmeticas; 46/63 exactas en descripcion, y de las que
difieren, once por tres caracteres o menos. Ver `pruebas/LEEME.md`.

Frente al modelo de vision tiene dos ventajas que no se compran con calidad:
Google regala 1000 paginas por mes -- al volumen de AOMAOSAM, probablemente
gratis para siempre -- y devuelve **el mismo texto siempre**, mientras que el
modelo puede variar entre corridas. Por eso este es el camino por defecto y el
modelo quedo de respaldo (ver `__init__.py`).

Dos decisiones heredadas del harness de la prueba:

- Se llama por REST con una API key, no con el SDK de Google. El SDK exige
  credenciales de service account y arrastra dependencias; con una key y el
  urllib de la biblioteca estandar alcanza, y el requirements no crece.
- No hay cache. El harness cacheaba en disco para no quemar cuota iterando
  sobre las mismas 63 muestras; aca cada comprobante se ve una sola vez y el
  servidor no guarda nada.
"""

import base64
import io
import json
import os
import urllib.error
import urllib.request

# DOCUMENT_TEXT_DETECTION es la variante para documentos (texto denso y
# estructurado). TEXT_DETECTION esta pensada para carteles y fotos de la calle.
CARACTERISTICA = "DOCUMENT_TEXT_DETECTION"

URL = "https://vision.googleapis.com/v1/images:annotate"

# Lado maximo de la imagen que se manda, en pixeles. NO es el LADO_MAXIMO de
# vision.py (1568): ahi achicar ahorra tokens, o sea plata, y aca no se paga por
# pixel sino por pagina. Achicar solo perjudica al OCR.
#
# El numero sale de no tocar lo que se midio: la prueba rasterizo las muestras a
# 200 dpi, o sea 2339 px el lado largo de una A4, y las mando enteras. 3000 deja
# eso intacto y solo recorta la foto de un celular moderno, que viene mucho mas
# grande de lo que hace falta para leer una factura.
LADO_MAXIMO = int(os.environ.get("AUTOFILLER_LADO_MAXIMO_OCR", "3000"))

CALIDAD_JPEG = 90

TIEMPO_LIMITE = int(os.environ.get("AUTOFILLER_TIMEOUT_OCR", "60"))


class OcrNoDisponible(RuntimeError):
    """No hay API key: se cae al modelo de vision, no se rompe."""


def disponible():
    return bool(os.environ.get("GOOGLE_VISION_API_KEY"))


def _a_jpeg(imagen):
    """Bytes JPEG de la imagen, achicada solo si es mas grande que LADO_MAXIMO."""
    from PIL import Image

    imagen = imagen.convert("RGB")
    lado = max(imagen.size)
    if lado > LADO_MAXIMO:
        escala = LADO_MAXIMO / lado
        nuevo = (max(1, int(imagen.width * escala)), max(1, int(imagen.height * escala)))
        imagen = imagen.resize(nuevo, Image.LANCZOS)
    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=CALIDAD_JPEG)
    return buffer.getvalue()


def texto(imagen):
    """Texto completo de la imagen segun Cloud Vision.

    Lanza OcrNoDisponible si falta la key, y RuntimeError si Google contesta un
    error. El que llama decide que hacer: en `__init__.py` cualquiera de las dos
    cosas hace que se lea con el modelo de vision.
    """
    clave = os.environ.get("GOOGLE_VISION_API_KEY")
    if not clave:
        raise OcrNoDisponible(
            "El servidor no tiene configurado el OCR (falta GOOGLE_VISION_API_KEY)."
        )

    cuerpo = json.dumps({
        "requests": [{
            "image": {"content": base64.b64encode(_a_jpeg(imagen)).decode("ascii")},
            "features": [{"type": CARACTERISTICA}],
            # Sin esto el OCR a veces 'corrige' palabras al ingles.
            "imageContext": {"languageHints": ["es"]},
        }]
    }).encode("utf-8")

    pedido = urllib.request.Request(
        f"{URL}?key={clave}",
        data=cuerpo,
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(pedido, timeout=TIEMPO_LIMITE) as respuesta:
            datos = json.loads(respuesta.read())
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "replace")[:400]
        raise RuntimeError(f"Cloud Vision devolvió {e.code}: {detalle}") from None

    primera = (datos.get("responses") or [{}])[0]
    if "error" in primera:
        raise RuntimeError(f"Cloud Vision: {primera['error'].get('message')}")

    return (primera.get("fullTextAnnotation") or {}).get("text", "")
