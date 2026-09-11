"""Carga de un comprobante en Prestadores -> Carga Rapida Comprobantes.

Es el corazon de AutoFiller y esta portado tal cual del escritorio: cada decision
de aca esta verificada contra la pantalla real por CDP, y la pantalla es GeneXus
y hostil a la automatizacion. No cambiar nada de esto sin volver a verificarlo
contra la pantalla.

Los tres hallazgos que hay que respetar si o si:

1. div.gx-mask: mientras GeneXus procesa un postback tapa la pantalla entera.
   Los click() rebotan y los fill() se pisan con el redibujado. Hay que esperar
   a que desaparezca Y que siga sin aparecer un rato (los postbacks vienen
   encadenados con huecos sin mascara entre medio).
2. El ORDEN de carga no es el visual: prestador PRIMERO (prompt por CUIT), luego
   el tipo, luego el resto de la cabecera. Invertirlo rompe el tipo y pisa el
   punto de venta.
3. Los campos se llenan con fill() SIN blur: cada blur dispara un postback y
   esos postbacks intermedios barajan los valores entre campos.
"""

import re
import time
import unicodedata

from operador import avisar, avisar_leve

URL_CARGA = "http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/cargarapidacomprobantescompra?10,0"

SIN_MASCARA = "() => !document.querySelector('div.gx-mask')"

SELECTOR_PROMPT = "iframe[title=\"Promptentidad\\?34\\,\\,0\\,10\\,c\\,\\,gxPopupLevel\\%3D0\\%3B\"]"

# Los dos popups del adjunto, encadenados: "Adjuntar" abre el alta de archivos
# (servlet temporalarchivos) y "Agregar Archivo" abre adentro el control de
# subida (servlet archivotemporalsubir). Se los busca por el servlet y no por el
# titulo del iframe, que lleva pegados los parametros de la llamada y cambia.
SELECTOR_ALTA_ARCHIVO = "iframe[src*='temporalarchivos']"
SELECTOR_SUBIR_ARCHIVO = "iframe[src*='archivotemporalsubir']"

# Unica opcion del combo "Tipo de archivo" de esa pantalla.
TIPO_ARCHIVO_COMPROBANTE = "1"

# Lo que va en la Descripcion del adjunto. Lo definio el operador.
DESCRIPCION_ADJUNTO = "FACTURA"

TIPOS_MIME = {
    "pdf": "application/pdf", "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "png": "image/png", "heic": "image/heic", "heif": "image/heif",
    "webp": "image/webp", "tif": "image/tiff", "tiff": "image/tiff",
}

# Cierra los popups que hayan quedado abiertos. Sin esto un adjunto a medio
# camino deja puesta la mascara del popup y la pantalla queda trabada: el
# operador no podria ni confirmar ni cancelar lo que ya estaba cargado.
CERRAR_POPUPS = """() => {
  let n = 0;
  while (gx.popup.currentPopup && n < 5) { gx.popup.currentPopup.close(); n++; }
  return n;
}"""


CAMPOS_CABECERA = {
    "vTIPOCOMPROBANTECODIGO": "Tipo de comprobante",
    "vCOMPROBANTEPREFIJO": "Punto de venta",
    "vCOMPROBANTECODIGO": "Número de comprobante",
    "vCOMPROBANTEFECHAEMISION": "Fecha de emisión",
    "vCOMPROBANTECUOTAVTO": "Fecha de vencimiento",
    "vCOMPROBANTEDEVENGAMIENTO": "Fecha de devengamiento",
    "vCOMPROBANTECAE": "CAE",
}

# Espejo reducido del de servidor/extraccion/centro_costo.py: aca solo hace falta
# para comparar la provincia del comprobante con la que SISalud tiene cargada
# para el prestador.
ALIAS_PROVINCIA = {
    "CAPITAL FEDERAL": "CABA",
    "CIUDAD AUTONOMA DE BUENOS AIRES": "CABA",
    "CIUDAD DE BUENOS AIRES": "CABA",
    "C A B A": "CABA",
    "PROVINCIA DE BUENOS AIRES": "BUENOS AIRES",
    "BS AS": "BUENOS AIRES",
    "STGO DEL ESTERO": "SANTIAGO DEL ESTERO",
    "TIERRA DEL FUEGO ANTARTIDA E ISLAS DEL ATLANTICO SUR": "TIERRA DEL FUEGO",
}


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"[^A-Za-z ]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip().upper()


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

    Click en el combo, select por VALUE, change, click y Enter para
    comprometerlo. Por value y no por posicion: la posicion cambia entre
    prestadores (en algunos el indice 2 es "Factura B") y cargaba el tipo
    equivocado en silencio.

    El combo de algunos prestadores parpadea: GeneXus lo repuebla y la opcion
    aparece y desaparece, asi que un intento aislado puede fallar aunque el tipo
    si corresponda. Se reintenta dejando que la pantalla se asiente entre medio.
    El timeout corto evita colgarse cuando el prestador realmente no ofrece esa
    opcion (ej. responsable inscripto que solo tiene B): en ese caso se avisa y
    se deja el combo sin tocar, sin cargar un tipo errado.
    """
    async def avisar_no_disponible():
        disponibles = await page.evaluate(
            "() => { const s = document.getElementById('vTIPOCOMPROBANTECODIGO');"
            " return s ? Array.from(s.options).map(o => o.text).filter(t => t) : []; }")
        avisar(
            avisos,
            "El tipo de comprobante de la factura no está disponible para este "
            "prestador.\n"
            f"SISalud ofrece: {', '.join(disponibles) or '(ninguno)'}\n"
            "Elegí el tipo a mano antes de confirmar.",
            accion="Elegí el tipo de comprobante.",
        )

    if not valor:
        await avisar_no_disponible()
        return False

    combo = page.locator("#vTIPOCOMPROBANTECODIGO")
    for _ in range(3):
        await esperar_genexus(page)

        # Mirar las opciones ANTES de intentar seleccionar. Antes se iba derecho
        # al select_option con timeout de 4 s, asi que un prestador que
        # realmente no ofrece este tipo -- CONTI, ARCE y varios mas -- se comia
        # 12 s de reloj para terminar avisando lo mismo. Se sigue reintentando
        # tres veces porque el combo parpadea mientras GeneXus lo repuebla, pero
        # ahora cada intento fallido cuesta milisegundos en vez de 4 s.
        valores = await page.evaluate(
            "() => { const s = document.getElementById('vTIPOCOMPROBANTECODIGO');"
            " return s ? Array.from(s.options).map(o => o.value) : []; }")
        if valor not in valores:
            await page.wait_for_timeout(300)
            continue

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
    GeneXus lo recibe junto al agregar la linea (#IMAGE3). GeneXus solo aplica
    su mascara (padea ceros: 5 -> 0005) sin necesidad de blur.

    Si el campo no esta a la vista no se intenta llenarlo. fill() espera a que el
    elemento sea visible y editable, asi que con un campo oculto se come el
    timeout entero para terminar fallando igual: eran 8 s tirados por campo. Y no
    es un caso raro -- SISalud tiene en el DOM campos de la cabecera que esta
    pantalla no muestra nunca, como la fecha de devengamiento.

    La espera corta (1 s) es por si GeneXus todavia lo esta dibujando; lo que no
    aparecio para entonces, con la pantalla ya en reposo, no va a aparecer.
    """
    if not dato:
        return False
    campo = page.locator(f"#{campo_id}")
    try:
        await campo.wait_for(state="visible", timeout=1000)
    except Exception:
        return False
    try:
        await campo.fill(dato, timeout=4000)
        return True
    except Exception:
        return False


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
        avisar(
            avisos,
            f"El CUIT {cuit} no figura en el padrón de entidades de SISalud.\n"
            "Hay que darlo de alta, o cargar el comprobante a mano.",
            accion="El prestador no está en SISalud: cargalo a mano.",
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


async def revisar_cabecera(page, avisos):
    """Avisa que campos de la cabecera quedaron sin cargar.

    Cada postback de GeneXus puede revertir un campo cargado antes, y no siempre
    pasa: conviene mirar como quedo la pantalla en vez de dar la carga por buena.
    Un campo "vacio" puede tener la mascara sin datos ("0000", "  /  /    ").
    """
    faltantes = []
    for campo_id, etiqueta in CAMPOS_CABECERA.items():
        campo = page.locator(f"#{campo_id}")
        # No alcanza con que no exista: SISalud deja en el DOM campos que la
        # pantalla no muestra (la fecha de devengamiento, por ejemplo). Estaban
        # vacios siempre, asi que se avisaban como faltantes en todos los
        # comprobantes, mandando al operador a completar un campo que no ve.
        if await campo.count() == 0 or not await campo.is_visible():
            continue
        valor = await campo.input_value()
        if not valor.strip() or set(valor) <= set(" /0"):
            faltantes.append(etiqueta)

    if faltantes:
        avisar(
            avisos,
            "Estos campos quedaron sin cargar, completalos antes de confirmar:\n- "
            + "\n- ".join(faltantes),
            accion="Completá: " + ", ".join(faltantes) + ".",
        )
    return faltantes


async def provincia_en_sisalud(page):
    """Provincia del prestador segun el maestro de entidades de SISalud.

    SISalud la muestra apenas se elige la entidad por CUIT, ya normalizada.
    Sirve como control cruzado de la que sale del comprobante.
    """
    span = page.locator("#span_vENTIDADDOMICILIOPROVINCIA")
    if await span.count() == 0:
        return None
    provincia = normalizar(await span.inner_text())
    provincia = ALIAS_PROVINCIA.get(provincia, provincia)
    if provincia in ("", "NO IDENTIFICADA"):
        return None
    return provincia


async def abrir_pantalla(page, usuario, contrasenia):
    """Lleva la pestaña a la Carga Rapida, logueando solo si hace falta.

    Tras el login se va directo con goto en vez de clickear el menu Prestadores
    -> Carga Rapida Comprobantes: esos clicks a veces no avanzaban y el operador
    tenia que darlos a mano. La sesion puede seguir abierta de una corrida
    anterior, asi que el login solo se hace si aparece la caja de usuario.
    """
    await page.goto(URL_CARGA)

    caja_usuario = page.get_by_role("textbox", name="ingrese con la forma: usuario")
    if await caja_usuario.count() > 0:
        await caja_usuario.fill(usuario)
        await page.locator("input[name=\"W0030vPASS\"]").fill(contrasenia)
        await page.get_by_role("button", name="entrar").click()
        await page.wait_for_load_state("networkidle")
        await page.goto(URL_CARGA)

    await esperar_genexus(page)


async def cargar_factura(page, factura, avisos, centro_de_padron=False):
    """Carga la factura en la pantalla ya abierta. No confirma: eso lo hace el
    operador. Devuelve False si no se pudo ni elegir el prestador.
    """
    # ORDEN ORIGINAL: el prestador PRIMERO, luego el tipo, luego el resto de la
    # cabecera. Es el orden que carga bien en la pantalla real. (Invertirlo
    # rompia el tipo de comprobante y pisaba el punto de venta.)
    if not await elegir_prestador(page, factura.cuit, avisos):
        return False

    tipo_ok = await elegir_tipo_comprobante(page, factura.tipo_comprobante, avisos)

    await llenar(page, "vCOMPROBANTEPREFIJO", factura.punto_venta)
    await llenar(page, "vCOMPROBANTECODIGO", factura.nro_factura)
    await llenar(page, "vCOMPROBANTEFECHARECEPCION", factura.fecha_recepcion)
    await llenar(page, "vCOMPROBANTEFECHAEMISION", factura.fecha_emision)
    await llenar(page, "vCOMPROBANTECUOTAVTO", factura.fecha_vencimiento)
    await llenar(page, "vCOMPROBANTEDEVENGAMIENTO", factura.fecha_devengamiento)
    await llenar(page, "vCOMPROBANTECAE", factura.cae)

    # La descripcion tiene maxlength=150: lo que no entra se pierde sin aviso.
    descripcion = " ".join((factura.descripcion or "").split())
    limite = await page.locator("#vEXENTOCOMPROBANTEDETALLEDESCRIPCION").get_attribute("maxlength")
    if limite and len(descripcion) > int(limite):
        # LEVE: la descripcion SI se cargo, solo hay que mirar que se perdio.
        # No impide confirmar, a diferencia de todos los demas avisos, y pasa
        # en casi la mitad de los comprobantes.
        avisar_leve(
            avisos,
            f"La descripción tiene {len(descripcion)} caracteres y en la pantalla "
            f"entran {limite}.\nSe cargó recortada, revisala antes de confirmar:\n"
            f"...{descripcion[int(limite):]}",
            accion="La descripción entró recortada: revisala.",
        )
    await llenar(page, "vEXENTOCOMPROBANTEDETALLEDESCRIPCION", descripcion)
    await llenar(page, "vEXENTOCOMPROBANTEDETALLEPRECIOUNITARIO", factura.importe)

    # Centro de costos: si no hay homonimo, o la provincia del comprobante no
    # coincide con la que SISalud tiene para el prestador, se deja sin tocar.
    #
    # El control cruzado por provincia existe para atajar los errores de la
    # APROXIMACION por domicilio del prestador. Cuando el centro de costos salio
    # del padron de afiliados no corresponde aplicarlo: ahi el dato es la
    # seccional real del afiliado, y justamente los casos que importan son
    # aquellos en que el prestador esta en otra provincia que su afiliado
    # (ANTOLA factura desde Jujuy y su afiliado es de Mina Aguilar).
    centro_costo = factura.centro_costo
    provincia_sisalud = None if centro_de_padron else await provincia_en_sisalud(page)
    if centro_costo and provincia_sisalud and provincia_sisalud != factura.provincia:
        avisar(
            avisos,
            f"La provincia del comprobante ({factura.provincia}) no coincide con "
            f"la del prestador en SISalud ({provincia_sisalud}).\n"
            "El Centro de Costos quedó sin cargar: elegilo antes de confirmar.",
            accion="Elegí el Centro de Costos.",
        )
        centro_costo = None
    if centro_costo:
        try:
            await page.locator("#vEXENTOCENTROCOSTOCODIGO").select_option(
                value=centro_costo, timeout=8000)
        except Exception:
            avisar(avisos, "No se pudo fijar el Centro de Costos: elegilo a mano.",
                   accion="Elegí el Centro de Costos.")

    # Solo se agrega la linea si el tipo de comprobante quedo cargado. Sin tipo,
    # "Agregar linea" (#IMAGE3) dispara un dialogo de validacion de GeneXus que
    # bloquea la pantalla.
    if tipo_ok:
        await revisar_cabecera(page, avisos)
        await esperar_genexus(page)
        await page.locator("#IMAGE3").click()
        await esperar_genexus(page)
    else:
        avisar(
            avisos,
            "No se agregó la línea del detalle porque falta el tipo de "
            "comprobante.\nElegí el tipo, revisá los datos y agregá la línea "
            "a mano.",
            accion="Agregá la línea del detalle.",
        )
    return True


async def limpiar_adjuntos(page, alta, avisos):
    """Borra los adjuntos que hayan quedado de un comprobante anterior.

    Los archivos temporales viven en la sesion y NO se limpian al recargar la
    pantalla (verificado el 2026-09-11): si el comprobante anterior se cancelo o
    se salto, sus adjuntos siguen ahi y se irian pegados al siguiente, que es
    cargar en SISalud un comprobante con el PDF de otro. El unico caso normal es
    que la lista este vacia, asi que lo que se encuentre se borra y se avisa.
    """
    # tr:has(td) y no tr a secas: la tabla del grid tiene una fila de titulos
    # que tambien cae adentro del tbody, y contarla daria una fila de mas
    # siempre -- incluso con la lista vacia.
    filas = alta.locator("#GridarchivosContainerTbl tbody tr:has(td)")
    borrados = 0
    while await filas.count() and borrados < 10:
        quedaban = await filas.count()
        await alta.locator("input[name^=vELIMINARARCHIVO]").first.click()
        # Se espera a que la fila se vaya del popup y no a esperar_genexus: el
        # postback es del iframe, y la mascara que deja en la pantalla de atras
        # tarda en irse lo que dura el timeout entero (20 s por adjunto).
        for _ in range(40):
            await page.wait_for_timeout(200)
            if await filas.count() < quedaban:
                break
        borrados += 1

    if borrados:
        avisar_leve(
            avisos,
            f"La pantalla tenía {borrados} archivo(s) adjunto(s) de antes y se "
            "quitaron, para que no se mezclen con este comprobante.",
            accion=f"Se quitaron {borrados} adjunto(s) que habían quedado.",
        )


async def adjuntar_comprobante(page, nombre, contenido, avisos):
    """Sube el archivo del comprobante al bloque Archivos de la pantalla.

    Son dos popups encadenados, verificados por CDP contra la pantalla real el
    2026-09-11:

    1. "Adjuntar" (input[name=BUTTON5], evento E'ADJUNTARARCHIVO') abre el alta
       de archivos: el combo Tipo, la Descripcion y "Agregar Archivo".
    2. "Agregar Archivo" (#BTNAGREGAR) NO abre el dialogo de Windows: valida
       primero la descripcion -- vacia contesta "Debe ingresar una descripcion" y
       no pasa de ahi -- y recien despues abre el segundo popup, que trae un
       control de subida de jQuery (input[type=file] de verdad). Por eso el
       archivo entra con set_input_files y sin tocar el disco: se le pasa el
       contenido que mando la web.
    3. La subida arranca sola al soltar el archivo y el popup de subida se
       cierra solo cuando termina. La senal de que termino es la fila nueva en
       la grilla del alta.
    4. "Salir" (input[name=BTNCANCEL]) cierra el alta y la fila aparece en la
       grilla Gridarchivocomprobante de la pantalla.

    Devuelve True si la fila quedo en la pantalla. Cualquier fallo se avisa y
    cierra los popups: el comprobante ya cargado tiene que quedar confirmable,
    aunque el adjunto haya que ponerlo a mano.
    """
    if not contenido:
        avisar(
            avisos,
            "No llegó el archivo del comprobante, así que no se adjuntó.\n"
            "Adjuntalo a mano antes de confirmar.",
            accion="Adjuntá el comprobante a mano.",
        )
        return False

    extension = nombre.rsplit(".", 1)[-1].lower() if "." in nombre else ""
    archivo = {
        "name": nombre or "comprobante",
        "mimeType": TIPOS_MIME.get(extension, "application/octet-stream"),
        "buffer": contenido,
    }

    try:
        await esperar_genexus(page)
        await page.locator("input[name=BUTTON5]").click()

        alta = page.locator(SELECTOR_ALTA_ARCHIVO).content_frame
        descripcion = alta.locator("#vARCHIVOTEMPORALDESCRIPCION")
        await descripcion.wait_for(state="visible", timeout=15000)

        # El combo trae "Comprobante" ya elegido y es su unica opcion. Se lo
        # toca solo si hiciera falta: seleccionar dispara el evento de GeneXus.
        tipo = alta.locator("#vARCHIVOTEMPORALTIPO")
        if await tipo.input_value() != TIPO_ARCHIVO_COMPROBANTE:
            await tipo.select_option(value=TIPO_ARCHIVO_COMPROBANTE, timeout=4000)

        await limpiar_adjuntos(page, alta, avisos)

        await descripcion.fill(DESCRIPCION_ADJUNTO)
        await alta.locator("#BTNAGREGAR").click()

        entrada = page.locator(SELECTOR_SUBIR_ARCHIVO).content_frame.locator(
            "#fileuploadUPLOADIFYContainer")
        await entrada.wait_for(state="attached", timeout=15000)
        await entrada.set_input_files(files=[archivo], timeout=60000)

        # Hasta 60 s: una foto de celular por una conexion lenta tarda.
        await alta.locator(
            "#GridarchivosContainerTbl tbody tr:has(td)").first.wait_for(
                state="visible", timeout=60000)

        await alta.locator("input[name=BTNCANCEL]").click()
        await esperar_genexus(page)

        filas = await page.locator(
            "#GridarchivocomprobanteContainerTbl tbody tr:has(td)").count()
        if filas:
            return True
        avisar(
            avisos,
            "El archivo se subió pero no quedó en la lista de Archivos de la "
            "pantalla.\nAdjuntalo a mano antes de confirmar.",
            accion="Adjuntá el comprobante a mano.",
        )
    except Exception as e:
        avisar(
            avisos,
            f"No se pudo adjuntar el comprobante ({e}).\n"
            "Adjuntalo a mano con el botón Adjuntar antes de confirmar.",
            accion="Adjuntá el comprobante a mano.",
        )

    try:
        await page.evaluate(CERRAR_POPUPS)
        await esperar_genexus(page)
    except Exception:
        pass
    return False
