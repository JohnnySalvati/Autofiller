"""Donde quedan los archivos que el agente lleva adentro (hoy, el icono).

Existe porque PyInstaller no los deja al lado del .exe sino en una carpeta
temporal que anuncia en sys._MEIPASS. Corriendo desde el codigo estan al lado
del fuente, y autofiller.ico ademas esta un nivel mas arriba, en la raiz del
repo: es el mismo icono para el agente y para el .exe del escritorio, y lo
genera hacer_icono.py a partir de servidor/web/logo.svg.
"""

import os
import sys

_AQUI = os.path.dirname(os.path.abspath(__file__))


def ruta(nombre):
    """Ruta del recurso, o None si no esta (el agente anda igual sin icono)."""
    candidatas = [
        os.path.join(getattr(sys, "_MEIPASS", _AQUI), nombre),
        os.path.join(_AQUI, nombre),
        os.path.join(os.path.dirname(_AQUI), nombre),
    ]
    for candidata in candidatas:
        if os.path.exists(candidata):
            return candidata
    return None
