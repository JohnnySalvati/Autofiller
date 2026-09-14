"""Modelo de la factura y traduccion de codigos de ARCA a los combos de SISalud."""

from typing import List, Optional

from pydantic import BaseModel, Field

# Tipo de comprobante: codigo de ARCA -> value del combo #vTIPOCOMPROBANTECODIGO.
# Se elige por value y nunca por posicion: la posicion se rompe en silencio si
# SISalud agrega o saca un tipo, y ademas varia entre prestadores.
TIPOS_COMPROBANTE = {
    "6": "FACCC",   # Factura B -> Factura "C" - Prestador
    "11": "FACCC",  # Factura C -> Factura "C" - Prestador
    "15": "NCCC",   # Recibo C  -> Nota de Credito "C" - Prestador
}

# Etiquetas del combo, para mostrarlas en la web sin pegarle a SISalud.
NOMBRES_TIPO_COMPROBANTE = {
    "CUDBC": 'Nota de débito "X"',
    "FACCC": 'Factura "C" - Prestador',
    "FACXC": 'Factura "X"',
    "NCCC": 'Nota de Crédito "C"',
    "NDBC": 'Nota de Débito "B"',
    "NDCC": 'Nota de Débito "C"',
    "RECCC": 'Recibo "C"',
    "REIN": "Reintegro - Propio",
}

# Las 24 opciones del combo Centro de Costos, para el desplegable de la web.
CENTROS_COSTO = {
    "3": "28 DE OCTUBRE",
    "4": "FRIAS",
    "5": "SALTA",
    "7": "MENDOZA",
    "8": "SAN JUAN",
    "9": "ENTRE RIOS",
    "12": "BARKER",
    "13": "CORDOBA",
    "14": "OLAVARRIA",
    "22": "TANDIL",
    "33": "NEUQUEN",
    "54": "SAN LUIS",
    "55": "BUENOS AIRES",
    "61": "MINA AGUILAR",
    "62": "SANTA CRUZ",
    "64": "RIO NEGRO",
    "71": "CATAMARCA",
    "73": "JUJUY",
    "74": "CHUBUT",
    "76": "HOTEL OSAM BS.AS. Nº4037",
    "77": "COLONIA 28 DE OCTUBRE",
    "78": "HOTEL MAR DEL PLATA",
    "79": "HOTEL BUENOS AIRES",
    "81": "CENTRAL",
}

# Lo minimo para que la pantalla acepte la cabecera. Sin esto GeneXus la descarta
# sin avisar, asi que se avisa cual falta antes de cargar.
CAMPOS_OBLIGATORIOS = ("cuit", "tipo_comprobante", "punto_venta", "nro_factura")

# Como se nombra cada campo cuando hay que hablarle al operador de el.
ETIQUETAS_CAMPO = {
    "cuit": "CUIT del emisor",
    "tipo_comprobante": "tipo de comprobante",
    "punto_venta": "punto de venta",
    "nro_factura": "número de comprobante",
    "fecha_emision": "fecha de emisión",
    "fecha_vencimiento": "vencimiento",
    "fecha_devengamiento": "devengamiento",
    "cae": "CAE",
    "descripcion": "descripción del detalle",
    "importe": "importe",
    "domicilio": "domicilio del emisor",
    "centro_costo": "centro de costos",
}

# De esos cuatro, el unico sin el cual no hay NADA que cargar: el agente elige el
# prestador buscandolo por CUIT en el prompt de entidades, y sin prestador la
# pantalla no acepta ningun otro dato. Los otros tres se cargan igual con lo que
# haya y los completa el operador, que tiene la factura a la vista: eso vale mas
# que devolverle el comprobante entero para que lo tipee de cero.
CAMPO_IMPRESCINDIBLE = "cuit"


class Factura(BaseModel):
    """Datos que se cargan en Prestadores -> Carga Rapida Comprobantes.

    Todos los campos son opcionales a proposito: un PDF al que le falta el
    periodo o el importe se sigue pudiendo cargar, con el aviso correspondiente,
    en vez de reventar en la extraccion como hacia el AutoFiller de escritorio.
    """

    cuit: Optional[str] = None
    # value del combo (FACCC / NCCC / ...), no el codigo de ARCA.
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

    # Derivados del domicilio comercial del emisor.
    domicilio: Optional[str] = None
    provincia: Optional[str] = None
    centro_costo: Optional[str] = None
    centro_costo_nombre: Optional[str] = None

    def faltantes(self) -> List[str]:
        return [c for c in CAMPOS_OBLIGATORIOS if not getattr(self, c)]

    def cargable(self) -> bool:
        """Si tiene sentido mandarlo al agente, aunque falte completar campos."""
        return bool(getattr(self, CAMPO_IMPRESCINDIBLE))


class Resultado(BaseModel):
    """Lo que devuelve la extraccion para un archivo."""

    archivo: str
    # 'pdf-texto' | 'foto' | 'pdf-imagen': de donde salieron los datos.
    origen: str = ""
    factura: Optional[Factura] = None
    avisos: List[str] = Field(default_factory=list)
    error: str = ""

    @property
    def ok(self) -> bool:
        return self.factura is not None and self.factura.cargable()
