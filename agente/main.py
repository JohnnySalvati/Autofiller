"""Agente local de AutoFiller.

Corre en la PC del operador y es lo unico que toca SISalud. La web le pasa un
comprobante ya leido y el agente lo carga en la pantalla de Carga Rapida del
Chrome del operador, avisa lo que quedo dudoso y espera a que el operador
Confirme o Cancele. Nunca confirma por su cuenta.

Las credenciales de SISalud entran por aca y no salen: la web las manda al
agente (127.0.0.1) y nunca al servidor de extraccion. Viven en memoria mientras
el agente esta abierto.

Arranque:  python -m uvicorn main:app --host 127.0.0.1 --port 8765
"""

import asyncio
import base64
import os
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from navegador import Navegador
from operador import (ESTADOS, Avisos, Control, avisar, esperar_resolucion,
                      mostrar_avisos_en_pantalla)
from padron import centro_costo_del_afiliado
from sisalud import (abrir_pantalla, adjuntar_comprobante, cargar_factura,
                     esperar_genexus)

# Lo que compara la actualizacion automatica (ver actualizacion.py): subirla es
# lo que hace que las PCs se actualicen. Un zip nuevo con la misma version no
# actualiza a nadie.
VERSION = "2.2"

# De donde se sirve la web. El agente escucha en 127.0.0.1, asi que cualquier
# pagina que el operador tenga abierta podria hablarle: por eso la lista de
# origenes es explicita y no un comodin.
ORIGENES = [o.strip() for o in os.environ.get(
    "AUTOFILLER_ORIGENES",
    "http://localhost:8000,http://127.0.0.1:8000",
).split(",") if o.strip()]

app = FastAPI(title="Agente AutoFiller", docs_url=None, redoc_url=None)
app.add_middleware(
    CORSMiddleware,
    allow_origins=ORIGENES,
    allow_methods=["*"],
    allow_headers=["*"],
    # Chrome pide permiso extra para que una pagina servida desde otro lado
    # llame a 127.0.0.1 (Private Network Access). Sin esto el navegador bloquea
    # la llamada al agente en el preflight y no dice por que.
    allow_private_network=True,
)


# --------------------------------------------------------------------------
# Modelos
# --------------------------------------------------------------------------

class Factura(BaseModel):
    """Espejo de servidor/extraccion/modelo.py:Factura.

    Esta duplicado a proposito: el agente se instala aparte del servidor y no
    comparte codigo con el. Si se agrega un campo alla, hay que agregarlo aca.
    """

    cuit: Optional[str] = None
    tipo_comprobante: Optional[str] = None
    punto_venta: Optional[str] = None
    nro_factura: Optional[str] = None
    fecha_recepcion: Optional[str] = None
    fecha_emision: Optional[str] = None
    fecha_vencimiento: Optional[str] = None
    fecha_devengamiento: Optional[str] = None
    cae: Optional[str] = None
    descripcion: str = ""
    importe: Optional[str] = None

    # Identificacion del afiliado, sacada del detalle facturado. No se carga en
    # la pantalla: el agente la usa para preguntarle al padron de SISalud cual es
    # la seccional del afiliado, que es el Centro de Costos de verdad.
    dni: Optional[str] = None
    nro_afiliado: Optional[str] = None
    domicilio: Optional[str] = None
    provincia: Optional[str] = None
    centro_costo: Optional[str] = None
    centro_costo_nombre: Optional[str] = None


class Credenciales(BaseModel):
    usuario: str
    clave: str


class Trabajo(BaseModel):
    factura: Factura
    archivo: str = ""

    # El comprobante mismo, en base64, para adjuntarlo en la pantalla. Lo manda
    # la web desde el archivo que ya tiene en memoria: no vuelve a pedirselo al
    # servidor y nunca sale de esta PC. Va adentro del JSON y no como multipart
    # para no sumarle al agente una dependencia (python-multipart) ni un segundo
    # endpoint; un comprobante pesa de unos KB a unos pocos MB.
    contenido: str = ""

    indice: int = 1
    total: int = 1


class Accion(BaseModel):
    accion: str  # 'saltar' | 'detener'


class Estado(BaseModel):
    fase: str  # 'libre' | 'cargando' | 'esperando' | 'terminado'
    archivo: str = ""
    indice: int = 0
    total: int = 0
    avisos: List[str] = []
    # Cuando fase == 'terminado': 'confirmar' | 'cancelar' | 'saltar' | ...
    resolucion: str = ""
    etiqueta: str = ""
    error: str = ""


class Salud(BaseModel):
    agente: str = "autofiller"
    version: str = VERSION
    con_sesion: bool
    fase: str


# --------------------------------------------------------------------------
# Estado del agente (uno solo: un operador, un Chrome, una pestaña)
# --------------------------------------------------------------------------

control = Control()
navegador = Navegador(control)
credenciales: Optional[Credenciales] = None
estado = Estado(fase="libre")
tarea: Optional[asyncio.Task] = None


def _contenido(trabajo: Trabajo) -> bytes:
    """El archivo del comprobante, o vacio si no vino o vino roto."""
    if not trabajo.contenido:
        return b""
    try:
        return base64.b64decode(trabajo.contenido)
    except Exception:
        return b""


async def _correr(trabajo: Trabajo):
    """Carga el comprobante y espera al operador. Es el ciclo completo de uno."""
    global estado

    control.nuevo_comprobante()
    # Avisos y no una lista comun: recuerda cuales avisos son leves, que es lo
    # que decide de que color se pinta el cartel (ver operador.Avisos).
    avisos = Avisos()
    cargado = False

    # Sin pestaña no hay nada que esperar: es el unico fallo que termina el
    # comprobante sin pasar por el operador (Chrome no abre, CDP no responde).
    try:
        page = await navegador.pagina()
    except Exception as e:
        estado = estado.model_copy(update={
            "fase": "terminado", "resolucion": "sin_confirmar",
            "etiqueta": ESTADOS["sin_confirmar"],
            "error": f"No se pudo abrir la pantalla de SISalud: {e}",
        })
        return

    try:
        await abrir_pantalla(page, credenciales.usuario, credenciales.clave)

        # El Centro de Costos que trae el comprobante esta deducido del domicilio
        # del prestador, que es una aproximacion: falla cuando el prestador esta
        # lejos del afiliado. Si el detalle facturado trae el DNI, el padron de
        # SISalud da la seccional exacta. Si el padron no contesta, se sigue con
        # la aproximacion, que es lo que habia antes: nunca se empeora.
        centro_de_padron = False
        if trabajo.factura.dni or trabajo.factura.nro_afiliado:
            centro = await centro_costo_del_afiliado(page, trabajo.factura, avisos)
            if centro:
                trabajo.factura.centro_costo, trabajo.factura.centro_costo_nombre = centro
                centro_de_padron = True
            # La consulta dejo la pestaña en el padron: hay que volver.
            await abrir_pantalla(page, credenciales.usuario, credenciales.clave)

        # El adjunto va ANTES de la cabecera: sobrevive entero a los postbacks
        # de la carga (verificado el 2026-09-11) y asi, si algo sale mal con los
        # popups, todavia no hay nada cargado que se pueda perder.
        await adjuntar_comprobante(
            page, trabajo.archivo, _contenido(trabajo), avisos)

        cargado = await cargar_factura(
            page, trabajo.factura, avisos, centro_de_padron=centro_de_padron)
    except Exception as e:
        avisar(
            avisos,
            f"Falló la carga automática ({e}).\n"
            "Cargalo a mano y confirmá, o tocá 'Saltar este' para saltarlo.",
            accion="No se pudo cargar: cargalo a mano, o saltalo.",
        )
    if not cargado and not avisos:
        avisar(avisos, "No se pudo cargar el comprobante: cargalo a mano o saltalo.",
               accion="No se pudo cargar: cargalo a mano, o saltalo.")

    lote = {"indice": trabajo.indice, "total": trabajo.total, "archivo": trabajo.archivo}
    await mostrar_avisos_en_pantalla(page, avisos, lote)

    control.automatico = False
    estado = estado.model_copy(update={"fase": "esperando", "avisos": avisos})

    try:
        resolucion = await esperar_resolucion(page, control)
    except Exception as e:
        control.automatico = True
        estado = estado.model_copy(update={
            "fase": "terminado", "avisos": avisos,
            "resolucion": "sin_confirmar", "etiqueta": ESTADOS["sin_confirmar"],
            "error": f"Se perdió la pantalla mientras se esperaba: {e}",
        })
        return

    control.automatico = True
    estado = estado.model_copy(update={
        "fase": "terminado", "avisos": avisos,
        "resolucion": resolucion, "etiqueta": ESTADOS[resolucion],
    })

    # Dejar que termine el postback/navegacion antes del siguiente comprobante.
    try:
        await esperar_genexus(page, timeout=10000)
    except Exception:
        pass


# --------------------------------------------------------------------------
# Endpoints
# --------------------------------------------------------------------------

@app.get("/api/salud", response_model=Salud)
def salud():
    return Salud(con_sesion=credenciales is not None, fase=estado.fase)


@app.post("/api/sesion")
def abrir_sesion(datos: Credenciales):
    """Guarda las credenciales de SISalud en memoria, solo en esta PC."""
    global credenciales
    credenciales = datos
    return {"ok": True}


@app.delete("/api/sesion")
async def cerrar_sesion():
    global credenciales
    credenciales = None
    await navegador.cerrar()
    return {"ok": True}


@app.post("/api/trabajo", status_code=202)
async def iniciar(trabajo: Trabajo):
    """Arranca la carga de un comprobante. Vuelve enseguida: se sigue por /api/trabajo."""
    global estado, tarea

    if credenciales is None:
        raise HTTPException(409, "Falta iniciar sesión en SISalud desde AutoFiller.")
    if estado.fase in ("cargando", "esperando"):
        raise HTTPException(409, "Ya hay un comprobante en pantalla esperando resolución.")
    if not trabajo.factura.cuit:
        raise HTTPException(422, "El comprobante no tiene CUIT del emisor: no se puede cargar.")

    estado = Estado(
        fase="cargando",
        archivo=trabajo.archivo,
        indice=trabajo.indice,
        total=trabajo.total,
    )
    tarea = asyncio.create_task(_correr(trabajo))
    return {"ok": True}


@app.get("/api/trabajo", response_model=Estado)
def consultar():
    return estado


@app.post("/api/accion")
def accion(datos: Accion):
    """'saltar' pasa al siguiente; 'detener' corta la cola. Los mismos botones
    que el cartel muestra dentro de SISalud."""
    if datos.accion not in ("saltar", "detener"):
        raise HTTPException(422, f"Acción desconocida: {datos.accion}")
    control.registrar(datos.accion)
    return {"ok": True}


@app.post("/api/listo")
def listo():
    """La web avisa que ya tomó el resultado y el agente queda libre."""
    global estado
    if estado.fase == "terminado":
        estado = Estado(fase="libre")
    return {"ok": True}
