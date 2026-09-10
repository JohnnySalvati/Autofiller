"""Identificacion del afiliado, sacada del texto del detalle facturado.

Por que existe
--------------
El centro de costos que SISalud necesita es la SECCIONAL DEL AFILIADO. Hoy se
aproxima por el domicilio del prestador, que atiende en general a los afiliados
de su zona. La aproximacion falla cuando el prestador esta lejos: el caso
confirmado por el usuario es ANTOLA, prestadora en Perico (Jujuy) con centro de
costos Mina Aguilar, a unos 250 km.

La unica forma de resolverlo de verdad es preguntarle al padron de SISalud por
el afiliado. Y el afiliado esta identificado en el propio comprobante: el
detalle facturado casi siempre trae su DNI, su numero de afiliado, o los dos.

Esto saca ese dato. NO consulta a SISalud: es el insumo para esa consulta.

Los formatos que usan los prestadores
-------------------------------------
No hay uno solo; cada prestador escribe como quiere. Sobre las 63 muestras:

    Bianca Ramallo, DNI N 57.470.919 - n afiliado 47115/03
    el paciente Federico Alcibar, DNI 49727767
    BENEFICIARIO: ... AFILIADO: 47079/03 DNI: 54.248.369
    Paciente: ... DNI N 58.188.819 Afiliado OSAM N 646032
    Sanfeliu Milena. D.N.I. 47.834.054. AF: 42769/05
    ... DNI, N 49.766.626 de afiliado:27004/03
"""

import re

# DNI: 7 u 8 digitos, con o sin puntos de miles. El rotulo puede ser DNI, D.N.I.
# o DNI N, con dos puntos, coma o nada en el medio. El espacio opcional despues
# del punto de miles no es un capricho: BENITEZ 07 trae "DNI 57.484. 150".
REGEX_DNI = re.compile(
    r"D\.?\s*N\.?\s*I\.?\s*[:,.]?\s*(?:N[°ºo]?\.?\s*)?"
    r"(\d{1,3}(?:\.\s*\d{3}){2}|\d{7,8})",
    re.IGNORECASE,
)

# Numero de afiliado: casi siempre <numero>/<orden> ("47115/03"), a veces suelto
# ("646032"). Los rotulos vistos: "n afiliado", "AFILIADO:", "Afiliado OSAM N",
# "de afiliado:", "AF:".
REGEX_AFILIADO = re.compile(
    r"(?:afiliad[oa]\s*(?:OSAM)?\s*(?:N[°ºo]?\.?)?|\bAF\b)\s*[:.]?\s*(\d{4,8}(?:/\d{1,3})?)",
    re.IGNORECASE,
)


def _limpiar_dni(valor):
    return valor.replace(".", "").replace(" ", "")


def identificacion(descripcion):
    """(dni, nro_afiliado) del texto del detalle. Cualquiera puede ser None."""
    if not descripcion:
        return None, None

    dni = REGEX_DNI.search(descripcion)
    afiliado = REGEX_AFILIADO.search(descripcion)
    return (
        _limpiar_dni(dni.group(1)) if dni else None,
        afiliado.group(1) if afiliado else None,
    )
