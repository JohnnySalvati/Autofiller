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


def codigo_arca(text):
    """Codigo de comprobante de ARCA ('011' -> '11')."""
    match = re.search(r"COD\.?\s*0*(\d+)", text, re.IGNORECASE)
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

    # La descripcion del detalle es todo lo que hay entre los dos 'Subtotal'.
    descripcion = buscar(text, r"(?<=Subtotal)(.*?)(?=Subtotal)", re.DOTALL | re.IGNORECASE)

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
        "descripcion": " ".join((descripcion or "").split()),
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
