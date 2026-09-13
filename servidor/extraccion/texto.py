"""Lectura de los campos de la factura sobre el texto plano de ARCA.

Son los regex del AutoFiller de escritorio, con una diferencia de fondo: aca
ningun campo faltante corta la extraccion. El escritorio reventaba con
UnboundLocalError si el PDF no traia 'Hasta:' o 'Importe Total:', y como el .exe
se compila --noconsole el operador solo veia "Factura invalida". Ahora cada
campo que no aparece se devuelve vacio con su aviso, y el operador decide.
"""

import re
from datetime import date

from .centro_costo import (
    centro_costo_de_domicilio,
    extraer_domicilio_comercial,
    provincia_de_domicilio,
)
from .modelo import TIPOS_COMPROBANTE

# ARCA escribe el total sin separador de miles y con coma decimal ("161023,60"),
# verificado sobre las 63 muestras. El punto opcional cubre los emisores que si
# lo ponen.
REGEX_IMPORTE = r"Importe Total:\s*\$?\s*([\d,]+(?:\.\d+)?)"

# Donde arranca el detalle. Es el ancla comun a todos los ordenes de lectura.
REGEX_ENCABEZADO_DETALLE = re.compile(r"Producto\s*/\s*Servicio", re.IGNORECASE)

# Los otros encabezados de columna. Segun quien lea la factura pueden venir
# todos en una linea, cada uno en la suya, antes de la descripcion o despues,
# asi que se sacan de donde esten. Dos condiciones los distinguen del texto de
# la descripcion: van SIN dos puntos (el detalle muchas veces dice "Cantidad: 8
# sesiones mensuales" y ahi no hay que tocar nada) y vienen de a varios
# seguidos. Un "Cantidad" suelto en medio de una frase no se toca.
REGEX_CABECERAS_DETALLE = re.compile(
    r"(?:\s*(?:C[oó]digo|Cantidad|U\.\s*Medida|Precio\s*Unit\.?|%\s*Bonif\.?"
    r"|Imp\.\s*Bonif\.?|Subtotal)(?!\s*:)){2,}",
    re.IGNORECASE,
)

# Donde termina el detalle: el primero de los bloques que ARCA imprime despues.
# Sin IGNORECASE a proposito — 'ORIGINAL' es el rotulo del comprobante y viene
# siempre en mayusculas; una descripcion podria decir 'original' en minuscula.
REGEX_FIN_DETALLE = re.compile(
    r"Subtotal:\s*\$|R[ée]gimen\s+de\s+Transparencia|Importe\s+Otros\s+Tributos"
    r"|IVA\s+Contenido|\bARCA\b|AGENCIA\s+DE\s+RECAUDACI"
    r"|P[áa]g\.\s*\d|CAE\s*N|\bORIGINAL\b|\bDUPLICADO\b|\bTRIPLICADO\b"
    r"|Ingresos\s+Brutos|Fecha\s+de\s+Inicio\s+de\s+Actividades"
)

# Las seis columnas del renglon: cantidad, unidad de medida, precio unitario,
# % bonif, imp. bonif y subtotal ("12,00 unidades 20127,95 0,00 0,00 241535,40").
# La unidad puede ser de una o dos palabras ("otras unidades"). Cinco numeros con
# dos decimales alrededor de una unidad: no hay descripcion que se parezca a eso.
REGEX_COLUMNAS_DETALLE = re.compile(
    r"\s*[\d.]+,\d{2}\s+(?:\S+\s+){1,2}[\d.]+,\d{2}\s+[\d.]+,\d{2}"
    r"\s+[\d.]+,\d{2}\s+[\d.]+,\d{2}\s*"
)


def extraer_descripcion(text):
    """Descripcion del detalle, sirviendo cualquier orden de lectura.

    La misma factura da textos distintos segun quien la lea. Verificado sobre
    las 63 muestras (pdfplumber) y sus 62 lecturas por OCR (Cloud Vision),
    aparecen tres acomodos:

      POR FILAS (pdfplumber): el encabezado entero en una linea, y las columnas
      numericas pegadas al final de la PRIMERA linea de la descripcion.

      POR BLOQUES: primero toda la descripcion junta, y recien despues, en otro
      bloque, los encabezados y los numeros.

      MIXTO: los encabezados juntos arriba, y despues la descripcion con las
      columnas numericas intercaladas en el medio.

    En vez de una regla por acomodo —que fue el primer intento y se rompio en
    los tres— se recorta una sola vez de punta a punta y se le RESTA lo que no
    es descripcion: los encabezados de columna y el renglon numerico. Lo que
    queda es la descripcion, venga en el orden que venga.
    """
    encabezado = REGEX_ENCABEZADO_DETALLE.search(text)
    if not encabezado:
        return ""

    cuerpo = text[encabezado.end():]

    fin = REGEX_FIN_DETALLE.search(cuerpo)
    if fin:
        cuerpo = cuerpo[:fin.start()]

    cuerpo = REGEX_CABECERAS_DETALLE.sub(" ", cuerpo)

    columnas = REGEX_COLUMNAS_DETALLE.search(cuerpo)
    if columnas:
        partes = columnas.group(0).split()
        cuerpo = cuerpo[:columnas.start()] + " " + cuerpo[columnas.end():]
        # "otras unidades" es la unica unidad de medida de dos palabras (21 de
        # las 63 muestras). Cuando la columna se parte en dos filas, pdfplumber
        # deja "unidades" suelta mas adelante y arriba solo se llevo "otras".
        if partes[1:-4] == ["otras"]:
            cuerpo = re.sub(r"\bunidades\b", " ", cuerpo, count=1, flags=re.IGNORECASE)

    return " ".join(cuerpo.split())


def codigo_arca(text):
    """Codigo de comprobante de ARCA ('011' -> '11').

    El punto de "COD." es opcional y admite coma: sobre una foto, el OCR lee
    "COD, 011" bastante seguido (FOTO.pdf, 2026-09-13). Sin esto el tipo de
    comprobante queda sin cargar, que es de los cuatro datos que SISalud
    necesita si o si, y el comprobante no se puede cargar.
    """
    match = re.search(r"COD[.,]?\s*0*(\d+)", text, re.IGNORECASE)
    if match:
        return match.group(1)
    match = re.search(r"Codigo\s*nº\s*(\d+)", text, re.IGNORECASE)
    return match.group(1) if match else None


def cuit_emisor(text):
    """CUIT del emisor: es el primero que aparece, arriba del todo."""
    match = re.search(r"CUIT\:?\s*(\d{11})", text)
    if match:
        return match.group(1)
    match = re.search(r":\s*(\d{2})-(\d{8})-(\d)", text)
    return "".join(match.groups()) if match else None


def punto_venta_y_numero(text):
    """(punto de venta, numero) sin ceros a la izquierda.

    Se despojan los ceros porque asi lo hacia el escritorio; la pantalla los
    vuelve a padear al ancho del maxlength cuando se carga.
    """
    pv = re.search(r"Punto de Venta:\s*(\d+)", text)
    if pv:
        nro = re.search(r"Comp\. Nro:\s*(\d+)", text)
        return pv.group(1).lstrip("0"), (nro.group(1).lstrip("0") if nro else None)

    junto = re.search(r"\D*0*(\d{5})\s*-\s*0*(\d{8})", text)
    if junto:
        return junto.group(1).lstrip("0"), junto.group(2).lstrip("0")
    return None, None


def buscar(text, patron, bandera=0):
    match = re.search(patron, text, bandera)
    return match.group(1).strip() if match else None


def datos_desde_texto(text):
    """(dict de campos, avisos) a partir del texto de la factura de ARCA."""
    avisos = []

    codigo = codigo_arca(text)
    tipo_comprobante = TIPOS_COMPROBANTE.get(codigo)
    if codigo and not tipo_comprobante:
        avisos.append(
            f"El código de comprobante de ARCA ({codigo}) no tiene equivalente "
            "conocido en SISalud. Elegí el tipo a mano."
        )

    # ARCA rotula el periodo facturado como 'Desde:' / 'Hasta:'. El escritorio
    # usa 'Hasta:' tanto de vencimiento como de devengamiento.
    fecha_hasta = buscar(text, r"Hasta:\s*(\d{2}/\d{2}/\d{4})")
    if not fecha_hasta:
        avisos.append(
            "La factura no trae el período facturado ('Hasta:'). Quedaron sin "
            "cargar el vencimiento y el devengamiento."
        )

    domicilio = extraer_domicilio_comercial(text)
    if not domicilio:
        avisos.append(
            "No se encontró el domicilio comercial del emisor: el Centro de "
            "Costos queda en CENTRAL y lo elegís vos."
        )

    importe = buscar(text, REGEX_IMPORTE)
    if not importe:
        avisos.append("No se encontró el 'Importe Total' en la factura.")

    descripcion = extraer_descripcion(text)
    if not descripcion:
        avisos.append(
            "No se encontró el detalle facturado: la descripción queda vacía y "
            "la tenés que escribir vos."
        )

    punto_venta, nro_factura = punto_venta_y_numero(text)
    centro_costo_nombre, centro_costo = centro_costo_de_domicilio(domicilio)

    return {
        "cuit": cuit_emisor(text),
        "tipo_comprobante": tipo_comprobante,
        "punto_venta": punto_venta,
        "nro_factura": nro_factura,
        "fecha_recepcion": date.today().strftime("%d/%m/%Y"),
        "fecha_emision": buscar(text, r"Fecha de Emisión:\s*(\d{2}/\d{2}/\d{4})"),
        "fecha_vencimiento": fecha_hasta,
        "fecha_devengamiento": fecha_hasta,
        "cae": buscar(text, r"CAE N°:\s*(\d+)"),
        "descripcion": descripcion,
        "importe": importe,
        "domicilio": domicilio,
        "provincia": provincia_de_domicilio(domicilio),
        "centro_costo": centro_costo,
        "centro_costo_nombre": centro_costo_nombre,
    }, avisos


def parece_factura_arca(text):
    """Heuristica barata para descartar un PDF que no es una factura de ARCA."""
    if not text:
        return False
    marcas = ("CUIT", "CAE", "Punto de Venta", "Importe Total", "COD.")
    return sum(1 for m in marcas if m.lower() in text.lower()) >= 2
