"""Servidor de extraccion de AutoFiller.

Recibe comprobantes (PDF o foto), devuelve los campos listos para la pantalla
de Carga Rapida de SISalud y sirve la interfaz web. No toca SISalud: eso lo hace
el agente que corre en la PC del operador (ver ../agente).

Por que esta partido asi: SISalud no tiene API, la carga tiene que seguir siendo
automatizacion del navegador, y el operador tiene que ver la pantalla real antes
de confirmar. Lo unico que se puede centralizar es la lectura del comprobante, y
ahi esta el valor: la clave de la API de vision queda en el servidor y no viaja
en ningun ejecutable.

No se guarda nada: cada archivo se procesa en memoria y se descarta.
"""

import asyncio
import json
import logging
import os
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from extraccion import EXTENSIONES_ACEPTADAS, MOTOR, extraer
from extraccion.modelo import CENTROS_COSTO, NOMBRES_TIPO_COMPROBANTE, Resultado
from extraccion.ocr import disponible as ocr_disponible
from extraccion.vision import MODELO, disponible as vision_disponible

# uvicorn solo configura sus propios loggers: sin esto, el logging de la
# extraccion (tokens y costo de cada llamada de vision) se descarta en silencio.
logging.basicConfig(
    level=os.environ.get("AUTOFILLER_LOG", "INFO").upper(),
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)

WEB = Path(__file__).parent / "web"

# Donde se deja el agente compilado para que cada PC se actualice sola. En la VM
# es un volumen montado (ver docker-compose.prod.yml); en desarrollo, la carpeta
# publicacion/ del repo. La misma expresion sirve para los dos: adentro de la
# imagen el codigo vive en /app, asi que /app/.. es / y esto da /publicacion.
PUBLICACION = Path(os.environ.get(
    "AUTOFILLER_PUBLICACION",
    Path(__file__).resolve().parent.parent / "publicacion"))

ZIP_AGENTE = "AutoFillerAgente.zip"
FICHA_AGENTE = "AutoFillerAgente.json"

# Tope por archivo. Una foto de celular ronda los 3-5 MB; 25 deja margen sin
# permitir que una subida gigante bloquee el servidor.
TAMANIO_MAXIMO = int(os.environ.get("AUTOFILLER_TAMANIO_MAXIMO_MB", "25")) * 1024 * 1024

app = FastAPI(title="AutoFiller", docs_url=None, redoc_url=None)


class Opciones(BaseModel):
    tipos_comprobante: dict
    centros_costo: dict
    lectura_de_fotos: bool
    motor_lectura: str
    modelo_vision: str
    extensiones: list


def _motor_lectura():
    """Con que se van a leer las fotos, en castellano, para el chip de la web.

    Hay dos motores y cualquiera de los dos alcanza, asi que el chip tiene que
    decir cual esta activo: 'Fotos: sí' con OCR y con el modelo significan la
    misma funcionalidad pero distinto costo y distinta repetibilidad.
    """
    ocr = ocr_disponible() and MOTOR in ("auto", "ocr")
    modelo = vision_disponible() and MOTOR in ("auto", "vision")

    if ocr and modelo:
        return f"OCR de Cloud Vision, con {MODELO} de respaldo"
    if ocr:
        return "OCR de Cloud Vision"
    if modelo:
        return MODELO
    return ""


@app.get("/api/opciones", response_model=Opciones)
def opciones():
    """Lo que la interfaz necesita para armar los desplegables y avisar al operador."""
    motor = _motor_lectura()
    return Opciones(
        tipos_comprobante=NOMBRES_TIPO_COMPROBANTE,
        centros_costo=CENTROS_COSTO,
        lectura_de_fotos=bool(motor),
        motor_lectura=motor,
        modelo_vision=MODELO if vision_disponible() else "",
        extensiones=sorted(EXTENSIONES_ACEPTADAS),
    )


@app.post("/api/extraer", response_model=list[Resultado])
async def api_extraer(archivos: list[UploadFile]):
    """Extrae los datos de cada comprobante subido, en el orden en que llegan."""
    resultados = []
    for archivo in archivos:
        contenido = await archivo.read()
        nombre = archivo.filename or "sin nombre"
        if len(contenido) > TAMANIO_MAXIMO:
            resultados.append(Resultado(
                archivo=nombre,
                error=f"El archivo pesa más de {TAMANIO_MAXIMO // (1024 * 1024)} MB.",
            ))
            continue
        # `extraer` es sincrónico y tarda: un PDF ocupa CPU y una foto se queda
        # esperando al modelo de visión. Llamado derecho desde acá bloquearía el
        # event loop y el servidor entero quedaría mudo mientras dura —dos operadores
        # subiendo a la vez se serializan del todo, y hasta el healthcheck de Docker
        # puede vencer y hacer que el contenedor se reinicie solo en medio de una
        # tanda. En un hilo aparte, cada pedido avanza por su cuenta.
        resultados.append(await asyncio.to_thread(extraer, nombre, contenido))
    return resultados


class Agente(BaseModel):
    version: str = ""
    url: str = ""
    tamano: int = 0
    sha256: str = ""


@app.get("/api/agente", response_model=Agente)
def agente():
    """Que version del agente publica este servidor, para que se actualice solo.

    Devolver la version vacia es una respuesta valida y quiere decir "no hay nada
    publicado": el agente sigue andando con la version que tenga. Que el
    servidor todavia no tenga el zip no puede ser un error del lado del operador.

    Esto y /descargas/ son los dos unicos endpoints que en produccion quedan
    afuera del auth_basic de nginx, porque el agente no tiene esas credenciales
    (ver docs/DEPLOYMENT.md § 3). Lo que se expone es el instalador, que no lleva
    ningun secreto adentro: las credenciales de SISalud las tipea el operador y
    la clave de lectura vive solo en este servidor.
    """
    zip_ = PUBLICACION / ZIP_AGENTE
    ficha = PUBLICACION / FICHA_AGENTE
    if not zip_.exists() or not ficha.exists():
        return Agente()
    try:
        datos = json.loads(ficha.read_text(encoding="utf-8"))
    except Exception:
        logging.warning("%s no se puede leer: no se publica ningún agente", ficha)
        return Agente()
    return Agente(
        version=str(datos.get("version", "")),
        url="/descargas/" + ZIP_AGENTE,
        tamano=zip_.stat().st_size,
        sha256=str(datos.get("sha256", "")),
    )


@app.get("/descargas/" + ZIP_AGENTE)
def descargar_agente():
    """El zip del agente. Lo baja el agente mismo al actualizarse, y sirve
    tambien para instalarlo por primera vez en una PC nueva."""
    zip_ = PUBLICACION / ZIP_AGENTE
    if not zip_.exists():
        raise HTTPException(404, "Este servidor no publica ningún agente.")
    return FileResponse(zip_, media_type="application/zip", filename=ZIP_AGENTE)


@app.get("/")
def inicio():
    return FileResponse(WEB / "index.html")


app.mount("/", StaticFiles(directory=WEB), name="web")
