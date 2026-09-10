"""Centro de costos a partir del domicilio comercial del prestador.

El centro de costos es en rigor la seccional del afiliado, dato que no esta en
la factura (esta en el padron de SISalud). Se aproxima por el domicilio del
prestador, que en la mayoria de los casos atiende a los afiliados de su zona:
primero la localidad, que es la que corresponde cuando existe como opcion
(Olavarria, Tandil, Barker), y si no hay coincidencia la provincia. Si no se
reconoce ninguna, el combo lo carga el operador en la pantalla.

Portado tal cual del AutoFiller de escritorio: la logica esta verificada contra
las 62 muestras utiles (58 ciertos, 4 dudosos, 0 manuales).
"""

import re
import unicodedata

# Opciones del combo #vEXENTOCENTROCOSTOCODIGO que son localidades o yacimientos.
# Solo se buscan en el segmento de localidad del domicilio, nunca en la calle:
# "28 de Octubre 123" es una direccion, no un centro de costos. Los hoteles y
# CENTRAL quedan afuera a proposito, no se corresponden con el domicilio de un
# prestador.
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
# de la continuacion, porque ARCA arma la cabecera en dos columnas. ORIGINAL /
# DUPLICADO / TRIPLICADO son el rotulo del ejemplar, que en la lectura por OCR
# cae junto al domicilio (verificado en OVIEDO 07: "Ruta 270 - Caucete, San Juan
# ORIGINAL", que sin esto se llevaba puesta la provincia y el centro de costos).
CAMPOS_JUNTO_AL_DOMICILIO = (
    r"(?:CUIT|Ingresos\s+Brutos|Condici|IVA|Fecha|Per[ií]odo|Domicilio"
    r"|ORIGINAL|DUPLICADO|TRIPLICADO)"
)

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


def provincia_de_domicilio(domicilio):
    """Provincia del domicilio del emisor, o None si no se reconoce.

    Solo se usa como control cruzado contra la que SISalud tiene cargada para el
    prestador; el centro de costos sale de centro_costo_de_domicilio().
    """
    if not domicilio:
        return None
    _, provincia = partir_domicilio(domicilio)
    return resolver_provincia(provincia)


def centro_costo_de_domicilio(domicilio):
    """(nombre, codigo) del centro de costos segun el domicilio del prestador.

    La localidad manda sobre la provincia: donde el combo tiene la seccional
    (Olavarria, Tandil, Barker) es esa la que corresponde, no BUENOS AIRES.
    Devuelve (None, None) si el domicilio no coincide con ninguna opcion, para
    que el operador lo cargue a mano.
    """
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
