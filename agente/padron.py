"""Centro de costos leido del padron de afiliados de SISalud.

Por que existe
--------------
El centro de costos que pide la Carga Rapida es la SECCIONAL DEL AFILIADO.
Hasta ahora se aproximaba por el domicilio del prestador, que atiende en general
a los afiliados de su zona. La aproximacion falla cuando el prestador esta
lejos: ANTOLA factura desde Perico (Jujuy) y su afiliado es de Mina Aguilar, a
250 km, asi que el domicilio daba JUJUY y lo correcto era MINA AGUILAR.

El 2026-09-10 se verifico contra el padron real que el dato exacto esta en
SISalud y se puede consultar con el DNI, que el propio comprobante trae en el
detalle facturado (63 de las 63 muestras traen DNI o numero de afiliado). Sobre
los cuatro casos de centro de costos confirmado por el usuario -- ANTOLA,
SANFELIU, ALCIBAR e HIDALGO -- la consulta acerto 4 de 4, incluido el de ANTOLA.

Vive en el agente y no en el servidor porque toca SISalud, y el unico que tiene
Chrome con la sesion del operador es el agente.

Regla de degradado (decision del usuario, 2026-09-10)
-----------------------------------------------------
Si el padron no contesta -- no hay DNI en el comprobante, el DNI no esta, la
pantalla cambio, lo que sea -- se vuelve a la aproximacion por domicilio, que ya
funciona. La consulta solo puede mejorar el resultado, nunca empeorarlo respecto
de lo que habia antes.

La trampa de la pantalla
------------------------
`wwafiliado` tiene los campos de filtro en el DOM desde que carga, pero el
servidor los IGNORA hasta que se elige `Orden Por`. Llenar el documento sin eso
devuelve cero filas siempre, sin ningun mensaje: se ve igual que "ese afiliado
no existe". Hay que setear `#vAFILIADOORDENPOR` primero y esperar el postback.
"""

import re
import unicodedata

from sisalud import esperar_genexus


def clave(nombre):
    """Nombre de delegacion comparable: mayusculas, sin acentos ni puntuacion.

    No se usa `normalizar` de sisalud.py, que esta hecho para provincias y
    descarta los digitos: con el, '28 DE OCTUBRE' queda en 'DE OCTUBRE' y dos
    delegaciones que solo se diferencien en un numero colisionarian en silencio.
    Hoy ninguna de las 49 colisiona, pero el padron crece.
    """
    sin_acentos = unicodedata.normalize("NFKD", (nombre or "").upper())
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", sin_acentos).split())


URL_PADRON = "http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/wwafiliado"

# Valores del combo "Orden Por", que es lo que habilita cada filtro.
ORDEN_POR_DOCUMENTO = "3"
ORDEN_POR_NUMERO_AFILIADO = "1"

# Delegacion del padron -> centro de costos de la Carga Rapida.
#
# Las 24 primeras son las que tienen delegacion homonima: se leyeron las 49
# delegaciones de `promptdelegacion` y los 24 centros de costo del combo, y
# coinciden por nombre exacto, 24 de 24.
CENTRO_COSTO_POR_DELEGACION = {
    "28 DE OCTUBRE": ("3", "28 DE OCTUBRE"),
    "FRIAS": ("4", "FRIAS"),
    "SALTA": ("5", "SALTA"),
    "MENDOZA": ("7", "MENDOZA"),
    "SAN JUAN": ("8", "SAN JUAN"),
    "ENTRE RIOS": ("9", "ENTRE RIOS"),
    "BARKER": ("12", "BARKER"),
    "CORDOBA": ("13", "CORDOBA"),
    "OLAVARRIA": ("14", "OLAVARRIA"),
    "TANDIL": ("22", "TANDIL"),
    "NEUQUEN": ("33", "NEUQUEN"),
    "SAN LUIS": ("54", "SAN LUIS"),
    "BUENOS AIRES": ("55", "BUENOS AIRES"),
    "MINA AGUILAR": ("61", "MINA AGUILAR"),
    "SANTA CRUZ": ("62", "SANTA CRUZ"),
    "RIO NEGRO": ("64", "RIO NEGRO"),
    "CATAMARCA": ("71", "CATAMARCA"),
    "JUJUY": ("73", "JUJUY"),
    "CHUBUT": ("74", "CHUBUT"),
    "HOTEL OSAM BS.AS. Nº4037": ("76", "HOTEL OSAM BS.AS. Nº4037"),
    "COLONIA 28 DE OCTUBRE": ("77", "COLONIA 28 DE OCTUBRE"),
    "HOTEL MAR DEL PLATA": ("78", "HOTEL MAR DEL PLATA"),
    "HOTEL BUENOS AIRES": ("79", "HOTEL BUENOS AIRES"),
    "CENTRAL": ("81", "CENTRAL"),

    # Las 21 que siguen son delegaciones SIN centro de costos propio, deducidas
    # del domicilio y la provincia que el padron tiene para cada una. Criterio
    # del usuario (2026-09-10): "todo lo que se pueda deducir es valido; si el
    # operador lo corrige, se cambia". El operador todavia no las confirmo.
    #
    # Zona Olavarria: LOMA NEGRA comparte el domicilio exacto de OLAVARRIA
    # (Lamadrid 2534) y CALERA AVELLANEDA el de CAL Y PIEDRA (9 de Julio 3272);
    # Sierras Bayas y Cerro Sotuyo son del partido de Olavarria.
    "LOMA NEGRA": ("14", "OLAVARRIA"),
    "CALERA AVELLANEDA": ("14", "OLAVARRIA"),
    "CAL Y PIEDRA OLAVARRIA": ("14", "OLAVARRIA"),
    "SIERRAS BAYAS": ("14", "OLAVARRIA"),
    "CERRO SOTUYO": ("14", "OLAVARRIA"),
    # OSAM esta en Rosario 434 y CENTRAL en Rosario 436: la puerta de al lado.
    "OSAM": ("81", "CENTRAL"),
    # Misma localidad (30163) que MINA AGUILAR.
    "CAJA FDO.SEGURO RETIRO": ("61", "MINA AGUILAR"),
    # El resto, por la provincia que el padron le asigna a la delegacion.
    "GENERAL SAN MARTIN": ("7", "MENDOZA"),
    "ASOC. MUTUAL HERCULES": ("7", "MENDOZA"),
    "QUILPO SUD": ("13", "CORDOBA"),
    "MALAGUENO": ("13", "CORDOBA"),
    "CORDOBA ART.23": ("13", "CORDOBA"),
    "JACHAL": ("8", "SAN JUAN"),
    "ALBARDON": ("8", "SAN JUAN"),
    "CARBOMETAL": ("8", "SAN JUAN"),
    "SAN JUAN (EMERGENCIA)": ("8", "SAN JUAN"),
    "SIERRA GRANDE": ("64", "RIO NEGRO"),
    # Gastre es de Chubut, pero el padron la tiene en Rio Negro: manda el padron.
    "GASTRE": ("64", "RIO NEGRO"),
    "PIEDRAS BLANCAS": ("9", "ENTRE RIOS"),
    "PIPINAS": ("55", "BUENOS AIRES"),
    "GENERACION PUEYRREDON": ("55", "BUENOS AIRES"),
}

# CORRIENTES, POSADAS, ROSARIO y LA RIOJA quedan afuera a proposito: sus
# provincias no tienen centro de costos, asi que no hay a que deducirlas.
DELEGACIONES_SIN_CENTRO_COSTO = {"CORRIENTES", "POSADAS", "ROSARIO", "LA RIOJA"}

# Lee la grilla buscando la columna por su nombre y no por posicion: la pantalla
# tiene 29 columnas y varias repiten encabezado, asi que contar seria fragil.
LEER_DELEGACION = """
() => {
  const g = document.getElementById('GridContainerTbl');
  if (!g || g.rows.length < 2) return null;
  const enc = [...g.rows[0].cells].map(c => (c.innerText || '').trim().toLowerCase());
  const i = enc.indexOf('delegación');
  if (i === -1) return null;
  return [...g.rows].slice(1).map(f => {
    const c = [...f.cells].map(x => (x.innerText || '').replace(/\\s+/g, ' ').trim());
    return c[i] || '';
  }).filter(Boolean);
}
"""


async def _buscar(page, orden_por, campo_id, valor):
    """Corre una busqueda en wwafiliado y devuelve las delegaciones que salieron."""
    # Primero el Orden Por: hasta que no se elige, el servidor ignora el filtro.
    await page.locator("#vAFILIADOORDENPOR").select_option(value=orden_por, timeout=8000)
    await esperar_genexus(page)

    await page.locator(f"#{campo_id}").fill(valor)
    await page.locator("input[name='SEARCHBUTTON']").click()
    await esperar_genexus(page)

    return await page.evaluate(LEER_DELEGACION) or []


async def centro_costo_del_afiliado(page, factura, avisos):
    """(codigo, nombre) del centro de costos segun el padron, o None.

    None significa "no se pudo": el que llama se queda con la aproximacion por
    domicilio. Nunca lanza.
    """
    if not factura.dni and not factura.nro_afiliado:
        return None

    try:
        await page.goto(URL_PADRON)
        await esperar_genexus(page)

        delegaciones = []
        if factura.dni:
            delegaciones = await _buscar(
                page, ORDEN_POR_DOCUMENTO, "vFILAFILIADOENTIDADDOCUMENTONUMERO",
                factura.dni)

        # El numero de afiliado es el segundo intento, no el primero: el DNI
        # identifica a la persona y el numero puede venir con el orden pegado.
        if not delegaciones and factura.nro_afiliado:
            numero, _, orden = factura.nro_afiliado.partition("/")
            await page.locator("#vAFILIADOORDENPOR").select_option(
                value=ORDEN_POR_NUMERO_AFILIADO, timeout=8000)
            await esperar_genexus(page)
            await page.locator("#vFILAFILIADONUMERO").fill(numero)
            if orden:
                await page.locator("#vFILAFILIADOORDEN").fill(orden.lstrip("0") or "0")
            await page.locator("input[name='SEARCHBUTTON']").click()
            await esperar_genexus(page)
            delegaciones = await page.evaluate(LEER_DELEGACION) or []
    except Exception as e:
        avisos.append(
            f"No se pudo consultar el padrón de afiliados ({e}).\n"
            "El Centro de Costos se dedujo del domicilio del prestador: revisalo."
        )
        return None

    if not delegaciones:
        avisos.append(
            "El afiliado del comprobante no apareció en el padrón. El Centro de "
            "Costos se dedujo del domicilio del prestador: revisalo."
        )
        return None

    # Varias filas con distinta delegacion serian dos afiliados distintos con el
    # mismo documento: no hay forma de elegir, y elegir mal es peor que no elegir.
    distintas = {clave(d) for d in delegaciones}
    if len(distintas) > 1:
        avisos.append(
            "El afiliado aparece en más de una delegación del padrón "
            f"({', '.join(sorted(distintas))}). El Centro de Costos quedó sin "
            "cargar: elegilo antes de confirmar."
        )
        return None

    delegacion = delegaciones[0].strip()
    buscada = clave(delegacion)
    for nombre, centro in CENTRO_COSTO_POR_DELEGACION.items():
        if clave(nombre) == buscada:
            return centro

    if buscada in {clave(d) for d in DELEGACIONES_SIN_CENTRO_COSTO}:
        avisos.append(
            f"El afiliado es de la delegación {delegacion}, que no tiene Centro "
            "de Costos propio. Elegilo antes de confirmar."
        )
    else:
        avisos.append(
            f"La delegación del afiliado ({delegacion}) no está en la tabla de "
            "Centros de Costo. Elegilo antes de confirmar."
        )
    return None
