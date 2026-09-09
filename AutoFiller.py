# pyinstaller --add-data "C:\Users\Johnny\AppData\Local\ms-playwright;.local\ms-playwright" --onefile AutoFiller.py
# chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\ChromeProfile" --no-first-run --no-default-browser-check


import asyncio
from playwright.async_api import async_playwright
import tkinter as tk
from tkinter import filedialog
import pdfplumber
import re
import pikepdf
from datetime import date 
import subprocess
import time
import socket
import sys
import unicodedata


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


async def mostrar_avisos_en_pantalla(page, avisos):
    """Muestra los avisos como un cartel dentro de la propia pantalla de SISalud.

    Un messagebox de Windows aparece detras de la ventana del navegador que el
    operador esta mirando, asi que no lo ve hasta despues de confirmar y cerrar.
    El cartel va inyectado en la pagina, fijo arriba de todo, donde si lo ve
    antes de confirmar. El titulo del comprobante va abajo (el cartel arriba), y
    el boton Confirmar esta al pie, asi que no lo tapa.
    """
    if not avisos:
        return
    try:
        await page.evaluate(
            """(avisos) => {
                const previo = document.getElementById('autofiller-avisos');
                if (previo) previo.remove();
                const caja = document.createElement('div');
                caja.id = 'autofiller-avisos';
                caja.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;'
                    + 'background:#b91c1c;color:#fff;font:14px/1.45 Segoe UI,sans-serif;'
                    + 'padding:12px 46px 14px 18px;box-shadow:0 2px 10px rgba(0,0,0,.45)';
                const titulo = document.createElement('div');
                titulo.style.cssText = 'font-weight:700;font-size:15px;margin-bottom:6px';
                titulo.textContent = 'AutoFiller \\u2014 revisar antes de Confirmar:';
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
                const cerrar = document.createElement('button');
                cerrar.textContent = '\\u00d7';
                cerrar.title = 'Cerrar aviso';
                cerrar.style.cssText = 'position:absolute;top:8px;right:14px;background:transparent;'
                    + 'border:0;color:#fff;font-size:24px;line-height:1;cursor:pointer';
                cerrar.onclick = () => caja.remove();
                caja.appendChild(cerrar);
                document.body.appendChild(caja);
            }""",
            avisos,
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


async def main(usuario, contraseña, factura):
    chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    user_data_dir = r"C:\\ChromeProfile"
    debugging_port = 9222


    chrome_process = subprocess.Popen([
        chrome_path,
        f"--remote-debugging-port={debugging_port}",
        f"--user-data-dir={user_data_dir}",
        "--no-first-run",
        "--no-default-browser-check"
    ], shell=True)

    # Función para verificar si el puerto está abierto
    def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(timeout)
            return sock.connect_ex((host, port)) == 0

    # Esperar hasta que el puerto 9222 esté disponible
    while not is_port_open("localhost", debugging_port):
        time.sleep(0.5)

    avisos = []

    async with async_playwright() as p:

      # Conectarse al navegador en ejecución sin abrir una nueva ventana
        browser = await p.chromium.connect_over_cdp("http://localhost:9222")
        
        # Obtener el contexto de la sesión actual
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = await context.new_page()

        # Navegar a la aplicación
        url = "http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/cargarapidacomprobantescompra?10,0"
        await page.goto(url)

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
            await page.goto(url)

        await esperar_genexus(page)

        # ORDEN ORIGINAL: el prestador PRIMERO, luego el tipo, luego el resto de la
        # cabecera. Es el orden que carga bien en la pantalla real. (Invertirlo
        # rompia el tipo de comprobante y pisaba el punto de venta.)
        if not await elegir_prestador(page, factura.cuit, avisos):
            await mostrar_avisos_en_pantalla(page, avisos)
            return avisos

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

        await mostrar_avisos_en_pantalla(page, avisos)
        return avisos

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

def start_processing(factura):
    # Los avisos se muestran como cartel dentro de la propia pantalla de SISalud
    # (mostrar_avisos_en_pantalla), que es donde el operador está mirando. No se
    # usa messagebox: aparecía detrás del navegador y era redundante.
    asyncio.run(main(user_var.get(), pass_var.get() , factura))
    sys.exit(0)

# Crear ventana principal
app = tk.Tk()
app.title("AutoFiller")
app.geometry("1100x380")
app.configure(bg="#f7f7f9")  # Fondo suave
app.iconbitmap("insoft.ico")

# Variables de entrada
user_var = tk.StringVar()
pass_var = tk.StringVar()

user_var.set("PLEONETTI@OSAM")
pass_var.set("Rosario434")

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

# Selección de archivo
file_frame = tk.Frame(app, pady=10, **frame_style)
# tk.Label(file_frame, text="Seleccione el archivo PDF:", bg=bg_color).grid(row=0, column=0, pady=5, padx=5, sticky="w")
process_button = None
file_label= None

def select_pdf():
    global process_button, file_label  # Reference to the global variables

    # Remove previous labels and buttons
    if file_label:
        file_label.pack_forget()
    if process_button:
        process_button.pack_forget()

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

tk.Button(file_frame, text="Seleccionar", command=select_pdf, bg=button_color, fg=button_text_color).pack(pady=5, padx=5)
file_frame.pack(pady=10, fill="x", padx=10)

# Ejecutar la aplicación
app.mainloop()
