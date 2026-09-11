"""A donde va lo que el agente cuenta de si mismo.

El agente ya no tiene una ventana negra abierta: vive en la bandeja del sistema.
Eso obliga a que lo que antes se leia en la consola quede escrito en un archivo,
y no es solo un reemplazo: es mejor que antes. La ventana negra se perdia entera
al cerrarla, asi que un error de ayer no se podia revisar hoy. El archivo si.

Es deliberado que NO se use logging.handlers.RotatingFileHandler: lo que hay que
capturar no son llamadas a logging sino todo lo que el agente escribe por stdout
y stderr --los prints de sisalud.py, los renglones de uvicorn, y sobre todo los
traceback de lo que reviente. Eso se captura redirigiendo los dos streams, y la
rotacion a mano de un archivo alcanza de sobra para el volumen de un operador.
"""

import datetime
import os
import sys

CARPETA = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "AutoFiller")
ARCHIVO = os.path.join(CARPETA, "agente.log")
ANTERIOR = os.path.join(CARPETA, "agente-anterior.log")

# Un lote deja unos pocos KB. Con 2 MB quedan meses de historia, y como se
# conserva un archivo anterior el peor caso son 4 MB en disco.
MAXIMO_BYTES = 2 * 1024 * 1024


class _Doble:
    """Escribe en el archivo y, si hay consola, tambien en la consola.

    Con el .exe empaquetado sin consola la segunda mitad no existe y todo va al
    archivo. Corriendo desde el codigo (iniciar.bat) se ve en la consola como
    siempre: el que desarrolla no tiene que ir a buscar el log.
    """

    def __init__(self, archivo, consola):
        self._archivo = archivo
        self._consola = consola

    def write(self, texto):
        self._archivo.write(texto)
        if self._consola is not None:
            try:
                self._consola.write(texto)
                # Sin esto la consola queda atrasada respecto del archivo: el
                # stdout heredado no es una tty y se vacia por bloques, asi que
                # el renglon que explica el error aparece mucho despues.
                if "\n" in texto:
                    self._consola.flush()
            except Exception:
                self._consola = None
        return len(texto)

    def flush(self):
        self._archivo.flush()
        if self._consola is not None:
            try:
                self._consola.flush()
            except Exception:
                self._consola = None

    def isatty(self):
        return self._consola is not None and getattr(self._consola, "isatty", bool)()

    def fileno(self):
        # uvicorn pregunta por el fileno para decidir si colorea. El del archivo
        # sirve y no miente: escribir ahi es escribir donde queda el log.
        return self._archivo.fileno()


def configurar():
    """Manda stdout y stderr al log. Se llama antes de importar cualquier cosa
    que imprima, porque uvicorn se queda con el stream que encuentra al arrancar.

    Devuelve la ruta del log (o None si no se pudo abrir, caso en que el agente
    sigue andando igual: queda sin diagnostico, no sin funcionar).
    """
    try:
        os.makedirs(CARPETA, exist_ok=True)
        if os.path.exists(ARCHIVO) and os.path.getsize(ARCHIVO) > MAXIMO_BYTES:
            if os.path.exists(ANTERIOR):
                os.remove(ANTERIOR)
            os.replace(ARCHIVO, ANTERIOR)
        # line_buffering: si el agente se cuelga o lo matan, lo ultimo que paso
        # tiene que estar en el archivo. Con buffer de bloque se perderia justo
        # el renglon que explica por que.
        archivo = open(ARCHIVO, "a", encoding="utf-8", errors="replace",
                       buffering=1)
    except Exception:
        return None

    archivo.write("\n" + "=" * 72 + "\n")
    archivo.write(f"Arranque: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")
    archivo.write("=" * 72 + "\n")
    archivo.flush()

    sys.stdout = _Doble(archivo, sys.__stdout__)
    sys.stderr = _Doble(archivo, sys.__stderr__)
    return ARCHIVO


def tambien_a(stream):
    """Duplica en `stream` lo que de aca en mas vaya al log.

    Lo usa el autodiagnostico: el .exe se compila sin consola, asi que cuando se
    lo corre desde un cmd hay que reengancharse a esa consola a mano y mandar
    ahi tambien la salida (ver `_enganchar_consola` en arrancar.py).
    """
    for nombre in ("stdout", "stderr"):
        actual = getattr(sys, nombre, None)
        if isinstance(actual, _Doble):
            actual._consola = stream


def leer(desde=0):
    """Devuelve (texto, posicion) del log a partir de un offset en bytes.

    Asi el visor muestra lo nuevo sin releer el archivo entero cada segundo.
    """
    try:
        with open(ARCHIVO, "rb") as f:
            f.seek(0, os.SEEK_END)
            fin = f.tell()
            # El log se rotó mientras el visor estaba abierto: volver al principio.
            if desde > fin:
                desde = 0
            f.seek(desde)
            datos = f.read()
        return datos.decode("utf-8", errors="replace"), desde + len(datos)
    except FileNotFoundError:
        return "", 0
    except Exception as e:
        return f"(no se pudo leer el log: {e})\n", desde
