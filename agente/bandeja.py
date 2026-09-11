"""El agente en la bandeja del sistema, al lado del reloj.

El agente es un servidor sin interfaz: lo unico que necesitaba la ventana negra
era decir "estoy vivo". Un icono en la bandeja dice lo mismo sin ocupar la barra
de tareas, y de paso deja de pasar lo que pasaba seguido: que alguien cerrara la
ventana con la X y despues la web dijera "no se detecta el agente".

Windows 11 manda los iconos nuevos a "Iconos ocultos" (el chevron al lado del
reloj), que es justo lo discreto que se busca; el que quiera verlo siempre lo
arrastra afuera una vez y queda.

El bucle de mensajes del icono tiene que vivir en el hilo principal, asi que el
servidor (uvicorn) corre en un hilo y esto se queda con el main. El visor del log
se abre en un tercer hilo con su propio mainloop de Tk.
"""

import ctypes
import threading
import webbrowser

import pystray
from PIL import Image, ImageDraw

import recursos
import visor

# Cada cuanto se repregunta el estado para el tooltip y el renglon del menu. Es
# una lectura de una variable en memoria: puede ser seguido sin costo.
INTERVALO_ESTADO = 2.0

_SI = 6  # IDYES de MessageBoxW

# Con esto puesto, los carteles se escriben en el log en vez de mostrarse. Lo
# usa el autodiagnostico cuando lo llama un script (--sin-carteles): un cartel
# que nadie va a clickear deja al empaquetado esperando para siempre, y eso ya
# paso una vez mientras se probaba esto.
SILENCIOSO = False


def avisar(titulo, mensaje, error=False):
    """Cartel de Windows. Sin consola es la unica forma de que algo se vea."""
    if SILENCIOSO:
        print(f"{titulo}: {mensaje}")
        return
    try:
        MB_OK, MB_SETFOREGROUND, MB_TOPMOST = 0x0, 0x10000, 0x40000
        icono = 0x10 if error else 0x40  # ICONERROR / ICONINFORMATION
        ctypes.windll.user32.MessageBoxW(
            None, mensaje, titulo, MB_OK | icono | MB_SETFOREGROUND | MB_TOPMOST)
    except Exception:
        print(f"{titulo}: {mensaje}")


def preguntar(titulo, mensaje):
    if SILENCIOSO:
        print(f"{titulo}: {mensaje}  -> se asume que si")
        return True
    try:
        MB_YESNO, MB_SETFOREGROUND, MB_TOPMOST = 0x4, 0x10000, 0x40000
        MB_ICONWARNING = 0x30
        return ctypes.windll.user32.MessageBoxW(
            None, mensaje, titulo,
            MB_YESNO | MB_ICONWARNING | MB_SETFOREGROUND | MB_TOPMOST) == _SI
    except Exception:
        return True


def _imagen():
    """El icono de AutoFiller, o uno dibujado si el .ico no vino en el paquete.

    autofiller.ico trae todos los tamanos (16 a 256) y es cuadrado, asi que
    alcanza con abrirlo: Pillow entrega el grande y pystray reduce.
    """
    ruta = recursos.ruta("autofiller.ico")
    if ruta:
        try:
            return Image.open(ruta).convert("RGBA")
        except Exception:
            pass
    # Respaldo: el cuadrado oscuro con el punto blanco del interruptor, que es lo
    # que hace reconocible a la familia de iconos de InSoft. Sin la "A", que a
    # mano no sale como en el SVG, pero con el punto, que es lo que se lee.
    imagen = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    dibujo = ImageDraw.Draw(imagen)
    dibujo.rounded_rectangle((0, 0, 63, 63), radius=12, fill=(15, 23, 42, 255))
    dibujo.ellipse((20, 20, 44, 44), fill=(46, 189, 89, 255))
    dibujo.ellipse((37, 37, 49, 49), fill=(255, 255, 255, 255))
    return imagen


class Bandeja:
    """El icono y su menu.

    `estado_texto()` describe en que anda el agente, `ocupado()` dice si hay un
    comprobante a medio cargar (para no cortar un lote sin preguntar) y
    `al_salir()` apaga el servidor.
    """

    def __init__(self, *, version, url_web, estado_texto, ocupado, al_salir):
        self._version = version
        self._url_web = url_web
        self._estado_texto = estado_texto
        self._ocupado = ocupado
        self._al_salir = al_salir
        self._terminar = threading.Event()

        self._icono = pystray.Icon(
            "autofiller",
            icon=_imagen(),
            title=self._titulo(),
            menu=pystray.Menu(
                pystray.MenuItem("Abrir AutoFiller", self._abrir_web, default=True),
                pystray.MenuItem("Ver la actividad", self._ver_actividad),
                pystray.Menu.SEPARATOR,
                # Deshabilitado a proposito: es informacion, no una accion.
                pystray.MenuItem(lambda item: self._estado_texto(), None,
                                 enabled=False),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Salir del agente", self._salir),
            ),
        )

    # -- menu ---------------------------------------------------------------

    def _titulo(self):
        return f"AutoFiller {self._version} — {self._estado_texto()}"

    def _abrir_web(self, icono=None, item=None):
        if self._url_web:
            webbrowser.open(self._url_web)
        else:
            avisar("AutoFiller",
                   "No sé de dónde se sirve la web de AutoFiller.\n\n"
                   "Corré configurar.bat una vez para dejarla configurada.")

    def _ver_actividad(self, icono=None, item=None):
        visor.abrir()

    def _salir(self, icono=None, item=None):
        if self._ocupado() and not preguntar(
                "AutoFiller",
                "Hay un comprobante cargado en SISalud esperando que lo "
                "confirmes o lo canceles.\n\n"
                "Si cerrás el agente ahora, la cola se corta y ese comprobante "
                "queda como esté en la pantalla.\n\n"
                "¿Cerrar el agente igual?"):
            return
        self._terminar.set()
        self._al_salir()
        self._icono.stop()

    # -- ciclo de vida ------------------------------------------------------

    def correr(self):
        """Bloquea hasta que se elija Salir. Tiene que ser el hilo principal."""
        threading.Thread(target=self._refrescar, name="bandeja-estado",
                         daemon=True).start()
        self._icono.run()

    def _refrescar(self):
        """Mantiene al dia el tooltip y el renglon de estado del menu.

        pystray no se entera solo de que el texto de un item cambio: hay que
        decirle con update_menu(), y el tooltip se reasigna a mano.
        """
        while not self._terminar.wait(INTERVALO_ESTADO):
            try:
                self._icono.title = self._titulo()
                self._icono.update_menu()
            except Exception:
                # El icono se esta apagando, o Windows rechazo la actualizacion.
                # Nada de esto justifica tirar abajo el agente.
                pass
