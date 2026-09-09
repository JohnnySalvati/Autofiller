"""Chrome del operador, enganchado por CDP.

AutoFiller no abre un navegador propio: se engancha al Chrome del operador con
--remote-debugging-port para que la pantalla de SISalud quede a la vista y el
operador confirme ahi mismo.

Consecuencia util del enganche por CDP: Playwright NO necesita sus navegadores
descargados (`playwright install`), solo el driver que viene con el paquete pip.
Eso saca de encima el empaquetado de los browsers, que en el .exe de escritorio
habia que meter a mano con --add-data.
"""

import asyncio
import os
import shutil
import socket
import subprocess
import time

PUERTO_CDP = int(os.environ.get("AUTOFILLER_PUERTO_CDP", "9222"))
CDP_URL = f"http://127.0.0.1:{PUERTO_CDP}"
PERFIL = os.environ.get("AUTOFILLER_PERFIL_CHROME", r"C:\ChromeProfile")

# El escritorio tenia la ruta de Chrome hardcodeada y fallaba en cualquier PC que
# lo tuviera en otro lado (p. ej. instalado por usuario, no por maquina).
RUTAS_CHROME = [
    os.environ.get("AUTOFILLER_CHROME"),
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]

SEGUNDOS_ESPERA_CHROME = 30


class ChromeNoEncontrado(RuntimeError):
    pass


def ruta_de_chrome():
    for ruta in RUTAS_CHROME:
        if ruta and os.path.isfile(ruta):
            return ruta
    encontrado = shutil.which("chrome") or shutil.which("chrome.exe")
    if encontrado:
        return encontrado
    raise ChromeNoEncontrado(
        "No se encontró Chrome. Instalalo, o indicá la ruta del ejecutable en "
        "la variable de entorno AUTOFILLER_CHROME."
    )


def puerto_abierto(puerto, timeout=0.5):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(timeout)
        return sock.connect_ex(("127.0.0.1", puerto)) == 0


def lanzar_chrome():
    """Levanta Chrome con el puerto de depuracion y espera a que responda.

    Si ya esta abierto con el mismo perfil, Chrome reutiliza esa instancia en vez
    de abrir otra, asi que se puede llamar en cada carga sin acumular ventanas.
    """
    if puerto_abierto(PUERTO_CDP):
        return

    subprocess.Popen(
        [
            ruta_de_chrome(),
            f"--remote-debugging-port={PUERTO_CDP}",
            f"--user-data-dir={PERFIL}",
            "--no-first-run",
            "--no-default-browser-check",
        ],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    limite = time.monotonic() + SEGUNDOS_ESPERA_CHROME
    while not puerto_abierto(PUERTO_CDP):
        if time.monotonic() > limite:
            raise RuntimeError(
                f"Chrome no respondió en el puerto {PUERTO_CDP}. Cerrá todas las "
                "ventanas de Chrome y volvé a intentar."
            )
        time.sleep(0.3)


class Navegador:
    """Mantiene una sola pestaña de la Carga Rapida para toda la cola.

    Una sola: no hay que dejar varias pestañas de la Carga Rapida abiertas porque
    se pisan la sesion y dejan la mascara de GeneXus puesta.
    """

    def __init__(self, control):
        self.control = control
        self._playwright = None
        self._browser = None
        self._page = None

    async def pagina(self):
        if self._page is not None and not self._page.is_closed():
            return self._page

        from playwright.async_api import async_playwright

        await asyncio.to_thread(lanzar_chrome)

        if self._playwright is None:
            self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.connect_over_cdp(CDP_URL)
        contexto = (self._browser.contexts[0] if self._browser.contexts
                    else await self._browser.new_context())
        self._page = await contexto.new_page()

        # Sin un listener, Playwright cierra solo los alert() de GeneXus. Durante
        # la carga automatica se mantiene ese comportamiento (un dialogo abierto
        # bloquea la pagina y colgaria la carga). Mientras espera al operador, el
        # listener no hace nada: el dialogo queda en el navegador y lo cierra el.
        def manejar_dialogo(dialogo):
            if self.control.automatico:
                asyncio.ensure_future(dialogo.dismiss())

        self._page.on("dialog", manejar_dialogo)
        # Los botones del cartel y los de SISalud avisan al agente por aca.
        await self._page.expose_binding(
            "autofillerAccion",
            lambda source, accion: self.control.registrar(accion),
        )
        return self._page

    async def cerrar(self):
        """Suelta la conexion CDP. No cierra el Chrome del operador."""
        self._page = None
        try:
            if self._browser is not None:
                await self._browser.close()
        except Exception:
            pass
        self._browser = None
        try:
            if self._playwright is not None:
                await self._playwright.stop()
        except Exception:
            pass
        self._playwright = None
