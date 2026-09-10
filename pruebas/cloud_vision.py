"""OCR de Google Cloud Vision, con cache en disco.

Existe para responder una sola pregunta antes de gastar un peso: el texto que
devuelve Cloud Vision, ¿lo pueden parsear las regex de ARCA que ya tenemos en
servidor/extraccion/texto.py? Si la respuesta es si, el camino de las fotos
puede ser OCR + regex (gratis hasta 1000 paginas por mes) en vez de un modelo
de vision.

Dos decisiones:

- Se llama por REST con una API key, no con el SDK de Google. El SDK exige
  credenciales de service account (GOOGLE_APPLICATION_CREDENTIALS) y arrastra
  dependencias; para una prueba, una key y urllib de la biblioteca estandar
  alcanzan y no agregan nada al requirements.
- El texto se cachea en disco. Cloud Vision regala 1000 paginas por mes y las
  muestras son 63: sin cache, iterar diez veces sobre las regex se come media
  cuota. Con cache, se paga una vez y las corridas siguientes son gratis e
  instantaneas.
"""

import base64
import hashlib
import json
import os
import urllib.error
import urllib.request
from pathlib import Path

CACHE = Path(__file__).parent / "cache_ocr"

# DOCUMENT_TEXT_DETECTION es la variante para documentos (denso, estructurado);
# TEXT_DETECTION esta pensada para carteles y fotos de la calle.
CARACTERISTICA = "DOCUMENT_TEXT_DETECTION"

URL = "https://vision.googleapis.com/v1/images:annotate"


class SinCredenciales(RuntimeError):
    pass


def disponible():
    return bool(os.environ.get("GOOGLE_VISION_API_KEY"))


def _ruta_cache(clave):
    """Ruta del cache para una clave estable provista por el que llama.

    La clave NO puede salir de los bytes del JPEG. Se probo asi y fallaba: dos
    entornos con distinta version de Pillow (12.0.0 y 12.3.0) comprimen la misma
    imagen a bytes distintos, con lo cual el cache no acertaba nunca y cada
    corrida volvia a pagarle a Google. La clave tiene que describir el ORIGEN
    (que PDF, a que dpi), no el resultado de comprimirlo.
    """
    digest = hashlib.sha256(clave.encode("utf-8")).hexdigest()[:32]
    return CACHE / f"{digest}.txt"


def en_cache(clave):
    return _ruta_cache(clave).exists()


def texto(contenido_jpeg, clave):
    """Texto completo de la imagen. Usa el cache si ya se pidio antes."""
    ruta = _ruta_cache(clave)
    if ruta.exists():
        return ruta.read_text(encoding="utf-8")

    clave = os.environ.get("GOOGLE_VISION_API_KEY")
    if not clave:
        raise SinCredenciales(
            "Falta GOOGLE_VISION_API_KEY. Ver pruebas/LEEME.md para sacarla."
        )

    cuerpo = json.dumps({
        "requests": [{
            "image": {"content": base64.b64encode(contenido_jpeg).decode("ascii")},
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
        with urllib.request.urlopen(pedido, timeout=60) as respuesta:
            datos = json.loads(respuesta.read())
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "replace")[:400]
        raise RuntimeError(f"Cloud Vision devolvio {e.code}: {detalle}") from None

    primera = (datos.get("responses") or [{}])[0]
    if "error" in primera:
        raise RuntimeError(f"Cloud Vision: {primera['error'].get('message')}")

    leido = (primera.get("fullTextAnnotation") or {}).get("text", "")

    CACHE.mkdir(exist_ok=True)
    ruta.write_text(leido, encoding="utf-8")
    return leido
