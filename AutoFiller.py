# pyinstaller --add-data "C:\Users\Johnny\AppData\Local\ms-playwright;.local\ms-playwright" --onefile AutoFiller.py
# chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfile" --no-first-run --no-default-browser-check


import asyncio
from playwright.async_api import async_playwright
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import re
import pikepdf
from datetime import date 
import subprocess
import time
import socket
import sys
import unicodedata
import os
import shutil
import threading


# Tipo de comprobante: codigo de ARCA -> value del combo #vTIPOCOMPROBANTECODIGO.
# Antes se elegia por posicion en la lista, que se rompe en silencio si SISalud
# agrega o saca un tipo; los value son estables.
TIPOS_COMPROBANTE = {
    "6": "FACCC",   # Factura B -> Factura "C" - Prestador
    "11": "FACCC",  # Factura C -> Factura "C" - Prestador
    "15": "NCCC",   # Recibo C  -> Nota de Credito "C" - Prestador
}

# Combo "Centro de Costos" de SISalud (#vEXENTOCENTROCOSTOCODIGO): sus 24 opciones
# mezclan provincias con seccionales y yacimientos. Los value son los codigos del
# maestro de centros de costo, asi que se seleccionan por value.
#
# El centro de costos es en rigor la seccional del afiliado, dato que no esta en
# la factura. Se aproxima por el domicilio del prestador, que en la mayoria de los
# casos atiende a los afiliados de su zona: primero la localidad, que es la que
# corresponde cuando existe como opcion (Olavarria, Tandil, Barker), y si no hay
# coincidencia la provincia. Si no se reconoce ninguna, el combo lo carga el
# operador en la pantalla.

# Opciones del combo que son localidades o yacimientos. Solo se buscan en el
# segmento de localidad del domicilio, nunca en la calle: "28 de Octubre 123" es
# una direccion, no un centro de costos. Los hoteles y CENTRAL quedan afuera a
# proposito, no se corresponden con el domicilio de un prestador.
CENTROS_COSTO_POR_LOCALIDAD = {
    "28 DE OCTUBRE": "3",
    "COLONIA 28 DE OCTUBRE": "77",
    "FRIAS": "4",
    "BARKER": "12",
    "VILLA CACIQUE": "12",  # pegada a Barker, mismo partido (Benito Juarez)
    "OLAVARRIA": "14",
    "FORTABAT": "14",       # Villa Fortabat, partido de Olavarria (zona Loma Negra)
    "TANDIL": "22",
    "MINA AGUILAR": "61",
}

CENTROS_COSTO_POR_PROVINCIA = {
    "BUENOS AIRES": "55",
    "CATAMARCA": "71",
    "CHUBUT": "74",
    "CORDOBA": "13",
    "ENTRE RIOS": "9",
    "JUJUY": "73",
    "MENDOZA": "7",
    "NEUQUEN": "33",
    "RIO NEGRO": "64",
    "SALTA": "5",
    "SAN JUAN": "8",
    "SAN LUIS": "54",
    "SANTA CRUZ": "62",
}

# Provincias que se reconocen pero no tienen centro de costos homonimo: el combo
# se deja como esta y lo elige el usuario antes de confirmar.
PROVINCIAS_SIN_CENTRO_COSTO = {
    "CABA",
    "CHACO",
    "CORRIENTES",
    "FORMOSA",
    "LA PAMPA",
    "LA RIOJA",
    "MISIONES",
    "SANTA FE",
    "SANTIAGO DEL ESTERO",
    "TIERRA DEL FUEGO",
    "TUCUMAN",
}

# Las variantes de CABA se listan aparte para que no terminen en "BUENOS AIRES",
# que es la provincia y tiene otro centro de costos.
ALIAS_PROVINCIA = {
    "CAPITAL FEDERAL": "CABA",
    "CIUDAD AUTONOMA DE BUENOS AIRES": "CABA",
    "CIUDAD DE BUENOS AIRES": "CABA",
    "C A B A": "CABA",
    "CABA": "CABA",
    "PROVINCIA DE BUENOS AIRES": "BUENOS AIRES",
    "BS AS": "BUENOS AIRES",
    "STGO DEL ESTERO": "SANTIAGO DEL ESTERO",
    "TIERRA DEL FUEGO ANTARTIDA E ISLAS DEL ATLANTICO SUR": "TIERRA DEL FUEGO",
}


# Campos que en el PDF vienen pegados al domicilio, en la misma linea o al final
# de la continuacion, porque ARCA arma la cabecera en dos columnas.
CAMPOS_JUNTO_AL_DOMICILIO = r"(?:CUIT|Ingresos\s+Brutos|Condici|IVA|Fecha|Per[ií]odo|Domicilio)"

PROVINCIAS_CONOCIDAS = set(CENTROS_COSTO_POR_PROVINCIA) | PROVINCIAS_SIN_CENTRO_COSTO

# ARCA a veces recorta la provincia al armar la caja del domicilio ("La Punta
# (Capital), San Lui"). Se acepta el recorte solo si es prefijo de una unica
# provincia; el largo minimo evita que un resto corto matchee por casualidad.
LARGO_MINIMO_PREFIJO = 4


def normalizar(texto):
    """Mayusculas, sin acentos ni puntuacion, para comparar nombres de lugares."""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Za-z ]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip().upper()


def extraer_domicilio_comercial(text):
    """Domicilio Comercial del emisor, sin los campos que ARCA le pega al lado."""
    lineas = text.split("\n")
    for i, linea in enumerate(lineas):
        match = re.search(r"Domicilio\s+Comercial\s*:\s*(.+)", linea)
        if not match:
            continue

        domicilio = re.split(r"\s+" + CAMPOS_JUNTO_AL_DOMICILIO, match.group(1))[0].strip()

        # Si es largo, ARCA lo corta y lo sigue al principio de la linea siguiente
        # ("... - Perico Del" / "Carmen, Jujuy Ingresos Brutos: ...").
        if i + 1 < len(lineas):
            continuacion = re.split(r"\s*" + CAMPOS_JUNTO_AL_DOMICILIO, lineas[i + 1])[0].strip()
            if continuacion:
                domicilio = f"{domicilio} {continuacion}"

        return domicilio

    return None


def partir_domicilio(domicilio):
    """'Rojas 311 - Villa Cacique, Buenos Aires' -> ('VILLA CACIQUE', 'BUENOS AIRES')."""
    provincia = ""
    if "," in domicilio:
        domicilio, provincia = domicilio.rsplit(",", 1)
    localidad = domicilio.rsplit(" - ", 1)[1] if " - " in domicilio else ""
    return normalizar(localidad), normalizar(provincia)


def resolver_provincia(texto):
    """Nombre canonico de la provincia, o None si no se reconoce."""
    provincia = ALIAS_PROVINCIA.get(texto, texto)
    if provincia in PROVINCIAS_CONOCIDAS:
        return provincia

    if len(provincia) >= LARGO_MINIMO_PREFIJO:
        candidatas = [p for p in PROVINCIAS_CONOCIDAS if p.startswith(provincia)]
        if len(candidatas) == 1:
            return candidatas[0]

    return None


def extraer_provincia(text):
    """Provincia del domicilio del emisor, o None si no se reconoce.

    Solo se usa como control cruzado contra la que SISalud tiene cargada para el
    prestador; el centro de costos sale de centro_costo_desde_domicilio().
    """
    domicilio = extraer_domicilio_comercial(text)
    if not domicilio:
        return None

    _, provincia = partir_domicilio(domicilio)
    return resolver_provincia(provincia)


def centro_costo_desde_domicilio(text):
    """(nombre, codigo) del centro de costos segun el domicilio del prestador.

    La localidad manda sobre la provincia: donde el combo tiene la seccional
    (Olavarria, Tandil, Barker) es esa la que corresponde, no BUENOS AIRES.
    Devuelve (None, None) si el domicilio no coincide con ninguna opcion, para
    que el operador lo cargue a mano.
    """
    domicilio = extraer_domicilio_comercial(text)
    if not domicilio:
        return None, None

    localidad, provincia = partir_domicilio(domicilio)

    if localidad:
        codigo = CENTROS_COSTO_POR_LOCALIDAD.get(localidad)
        if codigo:
            return localidad, codigo
        # La localidad puede venir con agregados ("Villa Cacique (Barker)").
        for nombre, codigo in CENTROS_COSTO_POR_LOCALIDAD.items():
            if nombre in localidad:
                return nombre, codigo

    provincia = resolver_provincia(provincia) if provincia else None
    if provincia:
        codigo = CENTROS_COSTO_POR_PROVINCIA.get(provincia)
        if codigo:
            return provincia, codigo

    return None, None


SIN_MASCARA = "() => !document.querySelector('div.gx-mask')"


async def esperar_genexus(page, timeout=20000, reposo=400):
    """Espera a que GeneXus termine de procesar y la pantalla quede quieta.

    Mientras procesa tapa la pantalla entera con un div.gx-mask. Los clicks
    rebotan contra esa mascara y los fill() escriben igual, pero cuando vuelve la
    respuesta el form se redibuja y pisa lo escrito: el campo queda vacio.

    No alcanza con ver que la mascara no este: los postbacks vienen encadenados y
    entre uno y el siguiente hay un hueco sin mascara. Si se aprovecha ese hueco,
    el postback que arranca despues se lleva puesto lo que se acaba de cargar.
    Por eso se exige que siga sin mascara despues de un rato de reposo.

    Si la mascara no se libera, la pantalla quedo trabada: recargarla la limpia.
    """
    limite = time.monotonic() + timeout / 1000
    while time.monotonic() < limite:
        restante = max(1000, int((limite - time.monotonic()) * 1000))
        try:
            await page.wait_for_function(SIN_MASCARA, timeout=restante)
        except Exception:
            return False
        await page.wait_for_timeout(reposo)
        if await page.evaluate(SIN_MASCARA):
            return True
    return False


async def elegir_tipo_comprobante(page, valor, avisos):
    """Elige el tipo de comprobante con el mecanismo que GeneXus acepta.

    Es el del codigo original (click en el combo, select por value, change, click
    y Enter para comprometerlo), que en la pantalla real carga bien. Dos cambios
    seguros sobre el original:
      - elige por VALUE (FACCC) y no por posicion: el original tomaba el indice 2,
        que en algunos prestadores es "Factura B" y cargaba el tipo equivocado en
        silencio.
      - si el prestador no ofrece esa opcion (ej. factura C pero responsable
        inscripto que solo tiene B), no adivina: deja el combo sin tocar y avisa,
        en vez de cargar B o colgarse esperando una opcion que no existe.
    """
    async def avisar_no_disponible():
        disponibles = await page.evaluate(
            "() => { const s = document.getElementById('vTIPOCOMPROBANTECODIGO');"
            " return s ? Array.from(s.options).map(o => o.text).filter(t => t) : []; }")
        avisos.append(
            "El tipo de comprobante de la factura no está disponible para este "
            "prestador.\n"
            f"SISalud ofrece: {', '.join(disponibles) or '(ninguno)'}\n"
            "Elegí el tipo a mano antes de confirmar."
        )

    # El combo de algunos prestadores parpadea: GeneXus lo repuebla y la opcion
    # aparece y desaparece por un instante, asi que un intento aislado puede fallar
    # aunque el tipo si corresponda. Se reintenta unas veces dejando que la pantalla
    # se asiente entre medio. El timeout corto evita colgarse cuando el prestador
    # realmente no ofrece esa opcion (ej. responsable inscripto que solo tiene B):
    # en ese caso se avisa y se deja el combo sin tocar, sin cargar un tipo errado.
    combo = page.locator("#vTIPOCOMPROBANTECODIGO")
    for intento in range(3):
        await esperar_genexus(page)
        try:
            await combo.click()
            await combo.select_option(value=valor, timeout=4000)
        except Exception:
            continue
        await page.evaluate(
            'document.querySelector("#vTIPOCOMPROBANTECODIGO").dispatchEvent(new Event("change"))')
        await combo.click()
        await page.keyboard.press("Enter")
        await esperar_genexus(page)
        if await combo.input_value() == valor:
            return True

    await avisar_no_disponible()
    return False


async def llenar(page, campo_id, dato):
    """Llena un campo con fill() y SIN blur.

    Sin blur a proposito: cada blur dispara un postback de GeneXus, y esos
    postbacks intermedios barajan los valores entre campos. Se llena todo y
    GeneXus lo recibe junto al agregar la linea (#IMAGE3).

    Timeout corto y tolerante: si el campo esta oculto (p. ej. quedan campos
    invisibles cuando no se pudo fijar el tipo de comprobante), no se cuelga
    esperandolo; revisar_cabecera() detecta despues lo que quedo sin cargar.
    """
    try:
        await page.locator(f"#{campo_id}").fill(dato, timeout=8000)
        return True
    except Exception:
        return False


SELECTOR_PROMPT = "iframe[title=\"Promptentidad\\?34\\,\\,0\\,10\\,c\\,\\,gxPopupLevel\\%3D0\\%3B\"]"


async def elegir_prestador(page, cuit, avisos):
    """Busca el prestador por CUIT en el prompt de entidades y lo selecciona.

    Hay que esperar a que la busqueda traiga la fila antes de clickearla: si se
    clickea mientras GeneXus redibuja la grilla, el click cae sobre un elemento
    que ya fue reemplazado y se pierde. El popup queda abierto y su mascara
    bloquea la pantalla entera hasta que se recarga a mano.

    Se busca la fila que tenga el CUIT para no depender del orden de la grilla:
    si la busqueda devuelve varias entidades, el primer link no tiene por que ser
    la correcta.
    """
    await esperar_genexus(page)
    await page.locator("#PROMPTIMGENTIDAD").click()

    prompt = page.locator(SELECTOR_PROMPT).content_frame
    await prompt.locator("#vENTIDADCUIT").fill(cuit)
    await prompt.get_by_role("button", name="Buscar").click()

    enlace = prompt.locator(f"tr:has-text('{cuit}') a[href*='gxReturn']").first
    try:
        await enlace.wait_for(state="visible", timeout=30000)
    except Exception:
        avisos.append(
            f"El CUIT {cuit} no figura en el padrón de entidades de SISalud.\n"
            "Hay que darlo de alta, o cargar el comprobante a mano."
        )
        # Sin esto la pantalla queda bloqueada por la mascara del popup.
        try:
            await prompt.get_by_role("button", name="Cancelar").click()
            await esperar_genexus(page)
        except Exception:
            pass
        return False

    await enlace.click()
    await esperar_genexus(page)
    return True


CAMPOS_CABECERA = {
    "vTIPOCOMPROBANTECODIGO": "Tipo de comprobante",
    "vCOMPROBANTEPREFIJO": "Punto de venta",
    "vCOMPROBANTECODIGO": "Número de comprobante",
    "vCOMPROBANTEFECHAEMISION": "Fecha de emisión",
    "vCOMPROBANTECUOTAVTO": "Fecha de vencimiento",
    "vCOMPROBANTEDEVENGAMIENTO": "Fecha de devengamiento",
    "vCOMPROBANTECAE": "CAE",
}


async def mostrar_avisos_en_pantalla(page, avisos, lote=None):
    """Muestra los avisos como un cartel dentro de la propia pantalla de SISalud.

    Un messagebox de Windows aparece detras de la ventana del navegador que el
    operador esta mirando, asi que no lo ve hasta despues de confirmar y cerrar.
    El cartel va inyectado en la pagina, fijo arriba de todo, donde si lo ve
    antes de confirmar. El titulo del comprobante va abajo (el cartel arriba), y
    el boton Confirmar esta al pie, asi que no lo tapa.

    En modo carpeta (lote = {indice, total, archivo}) el cartel se muestra
    siempre: lleva el progreso, los botones "Siguiente comprobante" y "Detener
    lote", y engancha Confirmar (click o F12) y Cancelar de SISalud para saber
    cual toco el operador. Todo avisa a Python por window.autofillerAccion.
    """
    if not avisos and not lote:
        return
    try:
        await page.evaluate(
            """([avisos, lote]) => {
                const previo = document.getElementById('autofiller-avisos');
                if (previo) previo.remove();
                const hayAvisos = avisos.length > 0;
                const caja = document.createElement('div');
                caja.id = 'autofiller-avisos';
                caja.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;'
                    + 'background:' + (hayAvisos ? '#b91c1c' : '#1d4ed8') + ';color:#fff;'
                    + 'font:14px/1.45 Segoe UI,sans-serif;'
                    + 'padding:12px 46px 14px 18px;box-shadow:0 2px 10px rgba(0,0,0,.45)';
                const titulo = document.createElement('div');
                titulo.style.cssText = 'font-weight:700;font-size:15px;margin-bottom:6px';
                if (lote) {
                    titulo.textContent = 'AutoFiller \u2014 comprobante ' + lote.indice + ' de '
                        + lote.total + ': ' + lote.archivo
                        + (hayAvisos ? ' \u2014 revisar antes de Confirmar:' : '');
                } else {
                    titulo.textContent = 'AutoFiller \u2014 revisar antes de Confirmar:';
                }
                caja.appendChild(titulo);
                const ul = document.createElement('ul');
                ul.style.cssText = 'margin:0;padding-left:22px';
                avisos.forEach(a => {
                    const li = document.createElement('li');
                    li.style.cssText = 'margin-bottom:5px;white-space:pre-line';
                    li.textContent = a;
                    ul.appendChild(li);
                });
                caja.appendChild(ul);
                if (lote) {
                    const pie = document.createElement('div');
                    pie.style.cssText = 'margin-top:8px;display:flex;gap:10px;align-items:center;flex-wrap:wrap';
                    const nota = document.createElement('span');
                    nota.textContent = 'Revis\u00e1 y toc\u00e1 Confirmar o Cancelar en SISalud: '
                        + 'al terminar se carga el siguiente.';
                    pie.appendChild(nota);
                    const boton = (texto, accion) => {
                        const b = document.createElement('button');
                        b.type = 'button';
                        b.textContent = texto;
                        b.style.cssText = 'background:#fff;color:#111;border:0;border-radius:4px;'
                            + 'padding:5px 12px;font:600 13px Segoe UI,sans-serif;cursor:pointer';
                        b.onclick = () => { b.disabled = true; window.autofillerAccion(accion); };
                        return b;
                    };
                    pie.appendChild(boton('Siguiente comprobante (saltar este)', 'saltar'));
                    pie.appendChild(boton('Detener lote', 'detener'));
                    caja.appendChild(pie);
                    // Enganchar los botones de SISalud para saber cual toco el operador.
                    const confirmar = document.querySelector('input[name="CONFIRMAR"]');
                    if (confirmar) confirmar.addEventListener('click',
                        () => window.autofillerAccion('confirmar'), true);
                    const cancelar = document.querySelector('input[name="BUTTON2"]');
                    if (cancelar) cancelar.addEventListener('click',
                        () => window.autofillerAccion('cancelar'), true);
                    document.addEventListener('keydown', e => {
                        if (e.key === 'F12') window.autofillerAccion('confirmar');
                    }, true);
                }
                const cerrar = document.createElement('button');
                cerrar.textContent = '\u00d7';
                cerrar.title = lote ? 'Ocultar avisos' : 'Cerrar aviso';
                cerrar.style.cssText = 'position:absolute;top:8px;right:14px;background:transparent;'
                    + 'border:0;color:#fff;font-size:24px;line-height:1;cursor:pointer';
                // En modo carpeta solo se ocultan los avisos: el progreso y los
                // botones del lote tienen que seguir a la vista.
                cerrar.onclick = () => lote ? ul.remove() : caja.remove();
                caja.appendChild(cerrar);
                document.body.appendChild(caja);
            }""",
            [avisos, lote],
        )
    except Exception:
        # El cartel es un extra: si la pagina no permite inyectarlo, seguimos.
        pass


async def revisar_cabecera(page, avisos):
    """Avisa que campos de la cabecera quedaron sin cargar.

    Cada postback de GeneXus puede revertir un campo cargado antes, y no siempre
    pasa: conviene mirar como quedo la pantalla en vez de dar la carga por buena.
    Un campo "vacio" puede tener la mascara sin datos ("0000", "  /  /    ").
    """
    faltantes = []
    for campo_id, etiqueta in CAMPOS_CABECERA.items():
        campo = page.locator(f"#{campo_id}")
        if await campo.count() == 0:
            continue
        valor = await campo.input_value()
        if not valor.strip() or set(valor) <= set(" /0"):
            faltantes.append(etiqueta)

    if faltantes:
        avisos.append(
            "Estos campos quedaron sin cargar, completalos antes de confirmar:\n- "
            + "\n- ".join(faltantes)
        )
    return faltantes


async def provincia_en_sisalud(page):
    """Provincia del prestador segun el maestro de entidades de SISalud.

    SISalud la muestra en la pantalla apenas se elige la entidad por CUIT, ya
    normalizada. Sirve como control cruzado de la que sale del PDF.
    """
    span = page.locator("#span_vENTIDADDOMICILIOPROVINCIA")
    if await span.count() == 0:
        return None
    provincia = normalizar(await span.inner_text())
    provincia = ALIAS_PROVINCIA.get(provincia, provincia)
    if provincia in ("", "NO IDENTIFICADA"):
        return None
    return provincia


URL_CARGA = "http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/cargarapidacomprobantescompra?10,0"
CDP_URL = "http://localhost:9222"


def lanzar_chrome():
    """Levanta Chrome con el puerto de depuracion y espera a que responda.

    Si Chrome ya esta abierto con el mismo perfil, solo reutiliza esa instancia
    (no abre otra), asi que se puede llamar en cada corrida.
    """
    chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    user_data_dir = r"C:\\ChromeProfile"
    debugging_port = 9222

    subprocess.Popen([
        chrome_path,
        f"--remote-debugging-port={debugging_port}",
        f"--user-data-dir={user_data_dir}",
        "--no-first-run",
        "--no-default-browser-check"
    ], shell=True)

    def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(timeout)
            return sock.connect_ex((host, port)) == 0

    while not is_port_open("localhost", debugging_port):
        time.sleep(0.5)


async def abrir_pantalla(page, usuario, contraseña):
    """Lleva la pestaña a la Carga Rapida, logueando solo si hace falta."""
    await page.goto(URL_CARGA)

    # Login solo si aparece la pantalla de ingreso (la sesión puede seguir
    # abierta de una corrida anterior).
    caja_usuario = page.get_by_role("textbox", name="ingrese con la forma: usuario")
    if await caja_usuario.count() > 0:
        await caja_usuario.fill(usuario)
        await page.locator("input[name=\"W0030vPASS\"]").fill(contraseña)
        await page.get_by_role("button", name="entrar").click()
        await page.wait_for_load_state("networkidle")
        # Ya autenticado, se va directo a la pantalla de carga con un goto, en
        # vez de clickear el menú Prestadores → Carga Rapida Comprobantes (esos
        # clicks a veces no avanzan y el operador tiene que darlos a mano).
        await page.goto(URL_CARGA)

    await esperar_genexus(page)


async def cargar_factura(page, factura, avisos):
    """Carga la factura en la pantalla ya abierta. No confirma: eso lo hace el
    operador. Devuelve False si no se pudo ni elegir el prestador.
    """
    # ORDEN ORIGINAL: el prestador PRIMERO, luego el tipo, luego el resto de la
    # cabecera. Es el orden que carga bien en la pantalla real. (Invertirlo
    # rompia el tipo de comprobante y pisaba el punto de venta.)
    if not await elegir_prestador(page, factura.cuit, avisos):
        return False

    tipo_ok = await elegir_tipo_comprobante(page, factura.tipo_comprobante, avisos)

    # Los campos se llenan con fill() SIN blur, como el original: cada blur
    # dispara un postback de GeneXus, y esos postbacks intermedios barajaban
    # los valores entre campos (la fecha se colaba en el punto de venta). Sin
    # blur, GeneXus recibe todo junto recien al agregar la linea (#IMAGE3).
    await llenar(page, "vCOMPROBANTEPREFIJO", factura.punto_venta)
    await llenar(page, "vCOMPROBANTECODIGO", factura.nro_factura)
    await llenar(page, "vCOMPROBANTEFECHARECEPCION", factura.fecha_recepcion)
    await llenar(page, "vCOMPROBANTEFECHAEMISION", factura.fecha_emision)
    await llenar(page, "vCOMPROBANTECUOTAVTO", factura.fecha_vencimiento)
    await llenar(page, "vCOMPROBANTEDEVENGAMIENTO", factura.fecha_devengamiento)
    await llenar(page, "vCOMPROBANTECAE", factura.cae)

    # La descripcion tiene maxlength=150: lo que no entra se pierde sin aviso.
    descripcion = " ".join(factura.descripcion.split())
    limite = await page.locator("#vEXENTOCOMPROBANTEDETALLEDESCRIPCION").get_attribute("maxlength")
    if limite and len(descripcion) > int(limite):
        avisos.append(
            f"La descripción tiene {len(descripcion)} caracteres y en la pantalla "
            f"entran {limite}.\nSe cargó recortada, revisala antes de confirmar:\n"
            f"...{descripcion[int(limite):]}"
        )
    await llenar(page, "vEXENTOCOMPROBANTEDETALLEDESCRIPCION", descripcion)
    await llenar(page, "vEXENTOCOMPROBANTEDETALLEPRECIOUNITARIO", factura.importe)

    # Centro de costos: si no hay homonimo, o la provincia del PDF no coincide
    # con la que SISalud tiene para el prestador, se deja sin tocar y se avisa.
    centro_costo = factura.centro_costo
    provincia_sisalud = await provincia_en_sisalud(page)
    if centro_costo and provincia_sisalud and provincia_sisalud != factura.provincia:
        avisos.append(
            f"La provincia del PDF ({factura.provincia}) no coincide con la del "
            f"prestador en SISalud ({provincia_sisalud}).\n"
            "El Centro de Costos quedó sin cargar: elegilo antes de confirmar."
        )
        centro_costo = None
    if centro_costo:
        try:
            await page.locator("#vEXENTOCENTROCOSTOCODIGO").select_option(
                value=centro_costo, timeout=8000)
        except Exception:
            avisos.append("No se pudo fijar el Centro de Costos: elegilo a mano.")

    # Solo se agrega la linea si el tipo de comprobante quedo cargado. Sin
    # tipo, "Agregar linea" (#IMAGE3) dispara un dialogo de validacion de
    # GeneXus que bloquea la pantalla. Se deja todo cargado y el operador
    # elige el tipo y agrega la linea a mano (ya avisado en el cartel).
    if tipo_ok:
        await revisar_cabecera(page, avisos)
        # await page.locator("#IMAGE2").click()
        await esperar_genexus(page)
        await page.locator("#IMAGE3").click()
        await esperar_genexus(page)
    else:
        avisos.append(
            "No se agregó la línea del detalle porque falta el tipo de "
            "comprobante.\nElegí el tipo, revisá los datos y agregá la línea "
            "a mano."
        )
    return True


async def main(usuario, contraseña, factura):
    """Modo de un solo comprobante: lo carga y deja la pantalla al operador."""
    lanzar_chrome()
    avisos = []
    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = await context.new_page()
        await abrir_pantalla(page, usuario, contraseña)
        await cargar_factura(page, factura, avisos)
        await mostrar_avisos_en_pantalla(page, avisos)
        return avisos


# ---------------------------------------------------------------------------
# Modo carpeta: se cargan todos los PDF de una carpeta, uno por vez, en una sola
# pestaña. Despues de cada uno se espera a que el operador Confirme o Cancele en
# SISalud y recien ahi se carga el siguiente.
# ---------------------------------------------------------------------------

CARPETA_CARGADOS = "cargados"

# Verificado por CDP (2026-09-09): Cancelar (input name=BUTTON2, evento RETURN de
# GeneXus) navega fuera de la pantalla. Confirmar (input name=CONFIRMAR, atajo
# F12) no se pudo probar sin grabar un comprobante real; se asume que al grabar
# la pantalla navega o vuelve al formulario vacio. En ambos casos el comprobante
# "desaparece" de la pantalla: eso es lo que se detecta. Mientras el operador
# corrige un error de validacion, el comprobante sigue en pantalla y se espera.
# La botonera #TBL_BOTONES (Confirmar/Cancelar) esta oculta con el formulario
# vacio y aparece cuando hay comprobante cargado.
ESTADO_PANTALLA = """() => {
    const v = s => (document.querySelector(s) || {}).value;
    const botones = document.querySelector('#TBL_BOTONES');
    return {
        mascara: !!document.querySelector('div.gx-mask'),
        comprobante: !!botones && getComputedStyle(botones).display !== 'none'
            && v('#vENTIDADCODIGO') !== '00000000',
    };
}"""

# Cuanto tiene que sostenerse la ausencia del comprobante para darla por firme.
# Un redibujado de GeneXus puede ocultar la botonera un instante; sin esta
# espera se pasaba al siguiente comprobante antes de que el operador lo viera.
REPOSO_AUSENCIA = 1.5


class ControlLote:
    """Estado compartido entre el bucle del lote y los botones de la pantalla."""

    def __init__(self):
        self.ultimo_boton = None  # 'confirmar' | 'cancelar' (el que toco el operador)
        self.pedido = None        # 'saltar' | 'detener' (botones del cartel)
        self.automatico = True    # True mientras carga AutoFiller; False mientras espera al operador

    def registrar(self, accion):
        if accion in ("confirmar", "cancelar"):
            self.ultimo_boton = accion
        elif accion in ("saltar", "detener"):
            self.pedido = accion

    def nuevo_comprobante(self):
        self.ultimo_boton = None
        self.pedido = None
        self.automatico = True


def listar_pdfs(carpeta):
    return sorted(
        (f for f in os.listdir(carpeta) if f.lower().endswith(".pdf")
         and os.path.isfile(os.path.join(carpeta, f))),
        key=str.lower)


def leer_factura(ruta):
    """Devuelve (factura, motivo). Si no se pudo leer, factura es None."""
    if not decrypt_pdf(ruta):
        return None, "no se pudo abrir el PDF"
    try:
        factura = extract_information()
    except Exception as e:
        return None, f"error al leer el PDF ({e})"
    if not factura:
        return None, "no se reconoció como factura electrónica de ARCA"
    return factura, ""


def mover_a_cargados(carpeta, nombre):
    """Mueve el PDF confirmado a la subcarpeta 'cargados'. Devuelve un aviso o ''."""
    destino_dir = os.path.join(carpeta, CARPETA_CARGADOS)
    try:
        os.makedirs(destino_dir, exist_ok=True)
        base, ext = os.path.splitext(nombre)
        destino = os.path.join(destino_dir, nombre)
        n = 1
        while os.path.exists(destino):
            n += 1
            destino = os.path.join(destino_dir, f"{base} ({n}){ext}")
        shutil.move(os.path.join(carpeta, nombre), destino)
        return ""
    except Exception as e:
        return f"no se pudo mover a '{CARPETA_CARGADOS}': {e}"


async def esperar_resolucion(page, control):
    """Espera a que el operador termine con el comprobante en pantalla.

    Termina cuando el comprobante ya no esta en la pantalla (Confirmar o
    Cancelar), o cuando el operador toca "Siguiente" o "Detener" en el cartel.
    Devuelve 'confirmar', 'cancelar', 'sin_confirmar', 'saltar' o 'detener'.

    "Ya no esta" tiene que sostenerse REPOSO_AUSENCIA segundos sin mascara de
    GeneXus, porque un redibujado puede esconder la botonera un instante. Que la
    pestana haya navegado no alcanza por si solo (llegan navegaciones tardias
    del comprobante anterior): se mira siempre el estado real de la pantalla.
    Si evaluar falla es porque la pagina esta cambiando de documento, y eso
    cuenta como ausencia.

    Si la carga automatica fallo, la pantalla arranca vacia: primero se espera a
    ver el comprobante (el operador lo carga a mano) y despues a que se vaya.
    """
    visto = False
    ausente_desde = None
    while True:
        if control.pedido:
            return control.pedido
        try:
            estado = await page.evaluate(ESTADO_PANTALLA)
        except Exception:
            estado = None  # cambiando de documento (Cancelar navega)
        if estado and estado["comprobante"]:
            visto = True
            ausente_desde = None
        elif estado and estado["mascara"]:
            pass  # GeneXus procesando: no se decide nada
        elif visto:
            ausente_desde = ausente_desde or time.monotonic()
            if time.monotonic() - ausente_desde >= REPOSO_AUSENCIA:
                break
        await asyncio.sleep(0.3)
    return control.ultimo_boton or "sin_confirmar"


ESTADOS = {
    "confirmar": "CONFIRMADO",
    "cancelar": "CANCELADO",
    "sin_confirmar": "SIN CONFIRMAR (la pantalla se cerró o recargó)",
    "saltar": "SALTADO",
    "detener": "DETENIDO (quedó en pantalla, sin confirmar)",
}


async def procesar_lote(usuario, contraseña, carpeta, informar=lambda texto: None):
    """Carga todos los PDF de la carpeta. Devuelve [(archivo, estado, detalle)]."""
    pdfs = listar_pdfs(carpeta)
    resultados = []
    if not pdfs:
        return resultados

    lanzar_chrome()
    control = ControlLote()

    async with async_playwright() as p:
        browser = await p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = await context.new_page()

        # Sin un listener, Playwright cierra solo los alert() de GeneXus. Durante
        # la carga automatica se mantiene ese comportamiento (un dialogo abierto
        # bloquea la pagina y colgaria la carga). Mientras espera al operador, el
        # listener no hace nada: el dialogo queda en el navegador y lo cierra el.
        def manejar_dialogo(dialogo):
            if control.automatico:
                asyncio.ensure_future(dialogo.dismiss())
        page.on("dialog", manejar_dialogo)
        # Los botones del cartel y los de SISalud avisan a Python por aca.
        await page.expose_binding("autofillerAccion", lambda source, accion: control.registrar(accion))

        total = len(pdfs)
        for indice, nombre in enumerate(pdfs, start=1):
            informar(f"Cargando {indice} de {total}: {nombre}")
            factura, motivo = leer_factura(os.path.join(carpeta, nombre))
            if factura is None:
                resultados.append((nombre, "NO LEÍDO", motivo))
                continue

            control.nuevo_comprobante()
            avisos = []
            cargado = False
            try:
                await abrir_pantalla(page, usuario, contraseña)
                cargado = await cargar_factura(page, factura, avisos)
            except Exception as e:
                avisos.append(
                    f"Falló la carga automática ({e}).\n"
                    "Cargalo a mano y confirmá, o tocá 'Siguiente comprobante' para saltarlo."
                )
            if not cargado and not avisos:
                avisos.append("No se pudo cargar el comprobante: cargalo a mano o saltalo.")

            lote = {"indice": indice, "total": total, "archivo": nombre}
            await mostrar_avisos_en_pantalla(page, avisos, lote)
            control.automatico = False
            informar(f"Esperando Confirmar/Cancelar {indice} de {total}: {nombre}")

            resultado = await esperar_resolucion(page, control)
            control.automatico = True
            detalle = ""
            if resultado == "confirmar":
                detalle = mover_a_cargados(carpeta, nombre) or f"movido a '{CARPETA_CARGADOS}'"
            resultados.append((nombre, ESTADOS[resultado], detalle))

            if resultado == "detener":
                for pendiente in pdfs[indice:]:
                    resultados.append((pendiente, "PENDIENTE", "no se procesó"))
                break

            # Dejar que termine el postback/navegacion antes de ir al siguiente.
            try:
                await esperar_genexus(page, timeout=10000)
            except Exception:
                pass

        informar("Lote terminado")
    return resultados


def decrypt_pdf(input_path):
    """Decrypts a PDF and saves it to the decrypted_folder."""
    try:
        with pikepdf.open(input_path) as pdf:
            pdf.save("decrypted.pdf")  
            pdf.close()
        return True
    except pikepdf.PasswordError:
        print(f"Error: Could not open {input_path} - incorrect password or strong encryption.")
        error_message = f"Error: Could not open {input_path} - incorrect password or strong encryption."
        # error_console.insert(tk.END, error_message + "\n")
        # error_console.see(tk.END)
        return None
    except Exception as e:
        print(f"Error decrypting {input_path}: {e}")
        error_message = f"Error decrypting {input_path}: {e}"
        # error_console.insert(tk.END, error_message + "\n")
        # error_console.see(tk.END)
        return None
    
def extract_information():
    """Extracts information from a decrypted PDF and saves it with a new filename in output_folder."""
    tipo_comprobante, punto_venta, nro_factura, cuit, punto_venta, nro_factura, fecha_recepcion, fecha_emision, fecha_vencimiento, fecha_devengamiento, cae, descripcion, provincia, centro_costo, centro_costo_nombre = None, None, None, None, None, None, None, None, None, None, None, "", None, None, None
    try:
        with pdfplumber.open("decrypted.pdf") as pdf:

            text = "".join(page.extract_text() or "" for page in pdf.pages)
            
            # Regex to find Codigo
            codigo_match = re.search(r'COD\.?\s*0*(\d+)', text, re.IGNORECASE)
            if codigo_match:
                tipo_comprobante = codigo_match.group(1)
            else:
                codigo_match = re.search(r'Codigo\s*nº\s*(\d+)', text, re.IGNORECASE)
                tipo_comprobante = codigo_match.group(1) if codigo_match else None
            # traducir el codigo de ARCA al value del combo de SISalud
            tipo_comprobante = TIPOS_COMPROBANTE.get(tipo_comprobante)


            cuit_match = re.search(r'CUIT\:?\s*(\d{11})', text)
            if cuit_match:
                cuit = cuit_match.group(1)
                # cuit1 = cuit_match1.group(1) if cuit_match1 and len(cuit_match1.group(1))==11 else None
            else:
                cuit_match = re.search(r':\s*(\d{2})-(\d{8})-(\d)', text)
                cuit = cuit_match.group(1) + cuit_match.group(2) + cuit_match.group(3) if cuit_match else None

            # if len(cuit) != 11:
            #     cuit = None

            punto_venta_match = re.search(r'Punto de Venta:\s*(\d+)', text)
            if punto_venta_match:
                nro_factura_match = re.search(r'Comp\. Nro:\s*(\d+)', text)
                punto_venta = punto_venta_match.group(1).lstrip('0') if punto_venta_match else None
                nro_factura = nro_factura_match.group(1).lstrip('0') if nro_factura_match else None
            else:
                punto_venta_factura_match = re.search(r'\D*0*(\d{5})\s*-\s*0*(\d{8})', text)
                if punto_venta_factura_match:
                    punto_venta = punto_venta_factura_match.group(1).lstrip('0')
                    nro_factura = punto_venta_factura_match.group(2).lstrip('0')
                else:
                    punto_venta = None
                    nro_factura = None

            fecha_recepcion =  date.today().strftime("%d/%m/%Y")

            fecha_emision_match = re.search(r'Fecha de Emisión:\s*(\d{2}/\d{2}/\d{4})', text)
            if fecha_emision_match:
                fecha_emision = fecha_emision_match.group(1)

            fecha_hasta_match = re.search(r'Hasta:\s*(\d{2}/\d{2}/\d{4})', text)
            if fecha_hasta_match:
                fecha_hasta = fecha_hasta_match.group(1)

            fecha_vencimiento = fecha_hasta
            fecha_devengamiento = fecha_hasta

            cae_match = re.search(r'CAE N°:\s*(\d+)', text)
            cae = cae_match.group(1) if cae_match else None

            provincia = extraer_provincia(text)
            centro_costo_nombre, centro_costo = centro_costo_desde_domicilio(text)

            regex = r"(?<=Subtotal)(.*?)(?=Subtotal)"
            descripcion_match = re.search(regex, text, re.DOTALL | re.IGNORECASE)
            if descripcion_match:
                descripcion = descripcion_match.group(1).strip()

            # Expresión regular para capturar el "Importe Total"
            regex_importe = r"Importe Total:\s*\$?\s*([\d,]+(?:\.\d+)?)"
            importe_match = re.search(regex_importe, text)
            if importe_match:
                importe = importe_match.group(1)

        if cuit and tipo_comprobante and punto_venta and nro_factura:
            return Factura(cuit,tipo_comprobante, punto_venta, nro_factura, fecha_recepcion, fecha_emision, fecha_vencimiento, fecha_devengamiento, cae, descripcion, importe, provincia, centro_costo, centro_costo_nombre )
        else:
            print(f"Error: CUIT: {cuit} Cod: {tipo_comprobante} Punto de Venta: {punto_venta} Nro Factura: {nro_factura}")
            # error_message = f"Error: {decrypted_path} CUIT: {cuit} Cod: {codigo} Punto de Venta: {punto_venta} Nro Factura: {nro_factura}"
            # error_console.insert(tk.END, error_message + "\n")
            # error_console.see(tk.END)
            return None

    except Exception as e:
        print(f"Error extracting information: {e}")
        # error_message = f"Error extracting information: {e} {decrypted_path}"
        # error_console.insert(tk.END, error_message + "\n")
        # error_console.see(tk.END)
        return None

class Factura():
    def __init__(self,cuit,tipo_comprobante, punto_venta, nro_factura, fecha_recepcion, fecha_emision, fecha_vencimiento, fecha_devengamiento, cae, descripcion, importe, provincia=None, centro_costo=None, centro_costo_nombre=None):
        self.cuit = cuit
        self.tipo_comprobante = tipo_comprobante
        self.punto_venta = punto_venta
        self.nro_factura = nro_factura
        self.fecha_recepcion = fecha_recepcion
        self.fecha_emision = fecha_emision
        self.fecha_vencimiento = fecha_vencimiento
        self.fecha_devengamiento = fecha_devengamiento
        self.cae = cae
        self.descripcion = descripcion
        self.importe = importe
        self.provincia = provincia
        self.centro_costo = centro_costo
        self.centro_costo_nombre = centro_costo_nombre

def credenciales():
    """Usuario y contrasena que cargo el operador, o None si falta alguno.

    Antes venian hardcodeadas en el codigo: eran las de una operadora real y
    viajaban dentro del .exe. Ahora las carga quien usa el programa."""
    usuario = user_var.get().strip()
    contrasena = pass_var.get()
    if not usuario or not contrasena:
        messagebox.showerror(
            "Faltan las credenciales",
            "Cargue su usuario y contrasena de SISalud antes de procesar.")
        return None
    return usuario, contrasena


def start_processing(factura):
    # Los avisos se muestran como cartel dentro de la propia pantalla de SISalud
    # (mostrar_avisos_en_pantalla), que es donde el operador está mirando. No se
    # usa messagebox: aparecía detrás del navegador y era redundante.
    acceso = credenciales()
    if acceso is None:
        return
    asyncio.run(main(acceso[0], acceso[1], factura))
    sys.exit(0)

# Crear ventana principal
app = tk.Tk()
app.title("AutoFiller")
app.geometry("1100x420")
app.configure(bg="#f7f7f9")  # Fondo suave
app.iconbitmap("insoft.ico")

# Variables de entrada
user_var = tk.StringVar()
pass_var = tk.StringVar()
progreso_var = tk.StringVar()

# Las credenciales las carga el operador: son suyas y no van en el codigo.

# Estilo de colores
bg_color = "#ffffff"  # Blanco para los paneles
border_color = "#d1d5db"  # Gris claro para bordes
title_color = "#4b5563"  # Gris oscuro para el título
button_color = "#2563eb"  # Azul para botones
button_text_color = "#ffffff"  # Blanco para texto de botones

# Estilo de bordes
frame_style = {
    "bg": bg_color,
    "highlightthickness": 1,
    "highlightbackground": border_color
}

# Título
title_frame = tk.Frame(app, pady=10, **frame_style)
tk.Label(title_frame, text="AutoFiller", font=("Helvetica", 16), fg=title_color, bg=bg_color).pack()
tk.Label(title_frame, text="Desarrollado por InSoft", bg=bg_color).pack()
title_frame.pack(pady=10, fill="x", padx=10)

# Usuario y contraseña
user_frame = tk.Frame(app, pady=10, **frame_style)
tk.Label(user_frame, text="Usuario", bg=bg_color).grid(row=0, column=0, pady=5, padx=5, sticky="e")
tk.Entry(user_frame, textvariable=user_var).grid(row=0, column=1, pady=5, padx=5)
tk.Label(user_frame, text="Contraseña", bg=bg_color).grid(row=1, column=0, pady=5, padx=5, sticky="e")
tk.Entry(user_frame, textvariable=pass_var, show="*").grid(row=1, column=1, pady=5, padx=5)
user_frame.pack(pady=10, fill="x", padx=10)

# Selección de archivo o carpeta
file_frame = tk.Frame(app, pady=10, **frame_style)
process_button = None
file_label = None
error_label = None

def limpiar_seleccion():
    global process_button, file_label, error_label
    for w in (file_label, process_button, error_label):
        if w:
            w.pack_forget()
    file_label = process_button = error_label = None

def select_pdf():
    global process_button, file_label, error_label
    limpiar_seleccion()

    file_path = filedialog.askopenfilename(filetypes=(("Facturas", "*.pdf"),))
    if file_path:
        decrypt_pdf(file_path)
        factura = extract_information()
        if factura:
            file_label = tk.Label(file_frame, text=f"Archivo seleccionado: {file_path}", bg=bg_color)
            file_label.pack(pady=5, padx=5)
            process_button = tk.Button(app, text="Procesar", command=lambda: start_processing(factura), bg=button_color, fg=button_text_color)
            process_button.pack(pady=10)
        else:
            error_label = tk.Label(file_frame, text="Error: Factura inválida.", fg="red", bg=bg_color)
            error_label.pack(pady=5, padx=5)

def select_folder():
    global process_button, file_label, error_label
    limpiar_seleccion()

    carpeta = filedialog.askdirectory(title="Carpeta con los comprobantes en PDF")
    if not carpeta:
        return
    carpeta = os.path.normpath(carpeta)
    pdfs = listar_pdfs(carpeta)
    if not pdfs:
        error_label = tk.Label(file_frame, text="La carpeta no tiene archivos PDF.", fg="red", bg=bg_color)
        error_label.pack(pady=5, padx=5)
        return
    file_label = tk.Label(file_frame, text=f"Carpeta seleccionada: {carpeta}  ({len(pdfs)} PDF)", bg=bg_color)
    file_label.pack(pady=5, padx=5)
    process_button = tk.Button(app, text=f"Procesar los {len(pdfs)} comprobantes",
                               command=lambda: start_lote(carpeta), bg=button_color, fg=button_text_color)
    process_button.pack(pady=10)

def start_lote(carpeta):
    """Corre el lote en un hilo aparte para que la ventana siga respondiendo y
    pueda mostrar el progreso; al terminar muestra el resumen."""
    acceso = credenciales()
    if acceso is None:
        return
    for b in (process_button, boton_archivo, boton_carpeta):
        if b:
            b.config(state="disabled")
    progreso_var.set("Iniciando...")

    def informar(texto):
        app.after(0, lambda: progreso_var.set(texto))

    def correr():
        error = None
        resultados = []
        try:
            resultados = asyncio.run(procesar_lote(acceso[0], acceso[1], carpeta, informar))
        except Exception as e:
            error = e
        app.after(0, lambda: terminar_lote(carpeta, resultados, error))

    threading.Thread(target=correr, daemon=True).start()

def terminar_lote(carpeta, resultados, error):
    for b in (boton_archivo, boton_carpeta):
        b.config(state="normal")
    if process_button:
        process_button.pack_forget()
    progreso_var.set("Lote terminado" if error is None else f"El lote se interrumpió: {error}")
    mostrar_resumen(carpeta, resultados, error)

def mostrar_resumen(carpeta, resultados, error):
    ventana = tk.Toplevel(app)
    ventana.title("Resumen del lote")
    ventana.geometry("900x500")
    ventana.attributes("-topmost", True)
    ventana.lift()

    conteo = {}
    for _, estado, _ in resultados:
        clave = estado.split(" (")[0]
        conteo[clave] = conteo.get(clave, 0) + 1
    encabezado = f"Carpeta: {carpeta}\nTotal: {len(resultados)}   " + "   ".join(
        f"{k}: {v}" for k, v in conteo.items())
    if error is not None:
        encabezado += f"\n\nEl lote se interrumpió por un error: {error}"
    tk.Label(ventana, text=encabezado, justify="left", anchor="w", padx=10, pady=8).pack(fill="x")

    marco = tk.Frame(ventana)
    marco.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    texto = tk.Text(marco, wrap="none", font=("Consolas", 10))
    barra = tk.Scrollbar(marco, command=texto.yview)
    texto.config(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y")
    texto.pack(side="left", fill="both", expand=True)
    for nombre, estado, detalle in resultados:
        linea = f"{estado:<14} {nombre}"
        if detalle:
            linea += f"   [{detalle}]"
        texto.insert("end", linea + "\n")
    texto.config(state="disabled")

    tk.Button(ventana, text="Cerrar", command=ventana.destroy, bg=button_color, fg=button_text_color).pack(pady=(0, 10))

botones_frame = tk.Frame(file_frame, bg=bg_color)
boton_archivo = tk.Button(botones_frame, text="Seleccionar archivo", command=select_pdf, bg=button_color, fg=button_text_color)
boton_archivo.pack(side="left", padx=5)
boton_carpeta = tk.Button(botones_frame, text="Seleccionar carpeta", command=select_folder, bg=button_color, fg=button_text_color)
boton_carpeta.pack(side="left", padx=5)
botones_frame.pack(pady=5)
tk.Label(file_frame, textvariable=progreso_var, bg=bg_color, fg=title_color).pack(pady=(0, 5))
file_frame.pack(pady=10, fill="x", padx=10)

# Ejecutar la aplicación
app.mainloop()
