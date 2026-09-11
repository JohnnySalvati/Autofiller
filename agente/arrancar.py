"""Punto de entrada del agente empaquetado como .exe.

Existe por una razon concreta: el agente se arranca con `uvicorn main:app`, y
esa forma de nombrar la app la importa POR TEXTO. Adentro de un ejecutable de
PyInstaller no hay un `main.py` en el disco que importar, asi que ese arranque
falla. Aca se importa `app` de verdad y se le pasa el objeto a uvicorn.

Como se presenta el agente (cambiado 2026-09-11): vive en la bandeja del
sistema, sin ventana. Antes dejaba una consola abierta todo el dia, que era la
senal de "estoy corriendo" pero tambien un boton de apagado accidental: cerrarla
con la X mataba el agente y la web pasaba a decir "no se detecta el agente".

Lo que la consola daba --ver que esta pasando cuando algo falla, que es la
leccion del `--noconsole` del escritorio, donde el operador solo veia "Factura
invalida"-- no se perdio: todo lo que se imprimia va al log
(%LOCALAPPDATA%\\AutoFiller\\agente.log) y el menu del icono lo muestra en vivo.
Ahora ademas sobrevive al cierre, asi que un error de ayer se puede leer hoy.

El bucle de la bandeja necesita el hilo principal, asi que uvicorn corre en un
hilo aparte. Sin bandeja disponible (desarrollo desde el codigo, o `--consola`)
se vuelve al comportamiento viejo: uvicorn en el hilo principal y todo a la vista.
"""

import ctypes
import os
import sys
import threading
import time

import registro

# El agente escucha siempre en el mismo puerto porque la web lo tiene por
# defecto; la variable existe para poder levantar un segundo agente de prueba sin
# pisar al que el operador tiene andando.
PUERTO = int(os.environ.get("AUTOFILLER_PUERTO", "8765"))

# Lo que tarda el servidor en estar escuchando. Si no llego en este tiempo, algo
# fallo (el caso tipico: el puerto 8765 ya tomado por otro agente) y hay que
# avisarlo con un cartel, porque sin consola nadie va a ver el traceback.
SEGUNDOS_ARRANQUE = 15


def _enganchar_consola():
    """Reengancha la salida a la consola del cmd que invoco al .exe, si hay una.

    El .exe se compila sin consola (es lo que lo hace discreto), asi que un
    `AutoFillerAgente.exe --probar` desde un cmd no mostraria nada: Windows no le
    adjunta la consola del padre a un proceso GUI. AttachConsole lo arregla.
    Devuelve si quedo enganchada.
    """
    try:
        kernel32 = ctypes.windll.kernel32
        if kernel32.GetConsoleWindow():
            return True  # ya hay consola: corriendo desde el codigo
        if not kernel32.AttachConsole(-1):  # ATTACH_PARENT_PROCESS
            return False
        registro.tambien_a(open("CONOUT$", "w", encoding="utf-8", buffering=1))
        return True
    except Exception:
        return False


def _esperar_enter(mensaje="  Enter para cerrar."):
    """Como input(), pero sin reventar cuando no hay stdin (exe sin consola)."""
    if sys.stdin is None:
        return
    try:
        input(mensaje)
    except Exception:
        pass


def probar():
    """Autodiagnostico: verifica lo que puede fallar en una PC nueva.

    Existe porque la alternativa para saber si el agente quedo bien instalado es
    cargar un comprobante de verdad en SISalud. Esto ejercita las dos piezas que
    el .exe podria no traer bien --el driver de Playwright, que es un node.exe
    empaquetado, y la ruta de Chrome-- sin abrir el navegador ni tocar SISalud.
    """
    import asyncio

    from main import ORIGENES, VERSION
    from navegador import PERFIL, PUERTO_CDP, ruta_de_chrome

    en_consola = _enganchar_consola()
    if "--sin-carteles" in sys.argv:
        import bandeja

        bandeja.SILENCIOSO = True
    renglones = []

    def decir(texto=""):
        renglones.append(texto)
        print(texto)

    decir(f"  AutoFiller — agente {VERSION} — autodiagnóstico")
    decir()
    problemas = 0

    decir("  Orígenes permitidos: " + (", ".join(ORIGENES) or "(ninguno)"))
    if not ORIGENES:
        decir("     ^ sin orígenes el agente rechaza todo. Corré configurar.bat.")
        problemas += 1

    try:
        chrome = ruta_de_chrome()
        decir(f"  Chrome: {chrome}")
    except Exception as e:
        decir(f"  Chrome: NO SE ENCONTRÓ ({e})")
        decir("     ^ instalá Chrome, o indicá la ruta con AUTOFILLER_CHROME.")
        problemas += 1

    decir(f"  Perfil de Chrome: {PERFIL}   (puerto de depuración {PUERTO_CDP})")

    # Lo mas fragil del empaquetado: Playwright trae su propio node.exe y hay que
    # arrancarlo. Si el .exe quedo mal armado, revienta aca y no al cargar.
    async def arrancar_playwright():
        from playwright.async_api import async_playwright

        p = await async_playwright().start()
        await p.stop()

    try:
        asyncio.run(arrancar_playwright())
        decir("  Playwright: el driver arranca bien")
    except Exception as e:
        decir(f"  Playwright: FALLÓ ({type(e).__name__}: {e})")
        problemas += 1

    # La bandeja es ahora la unica forma de ver que el agente esta corriendo: si
    # no se puede armar el icono, el agente queda invisible y vale avisarlo aca.
    try:
        import bandeja

        bandeja._imagen()
        decir("  Bandeja del sistema: el ícono se puede crear")
    except Exception as e:
        decir(f"  Bandeja del sistema: FALLÓ ({type(e).__name__}: {e})")
        problemas += 1

    decir()
    if problemas:
        decir(f"  {problemas} problema(s). El agente no va a poder cargar comprobantes.")
    else:
        decir("  Todo en orden. El agente puede cargar comprobantes en esta PC.")
    decir()

    if en_consola:
        _esperar_enter()
    else:
        # Doble clic sobre el .exe con --probar: sin cartel no se ve nada.
        from bandeja import avisar

        avisar("AutoFiller — autodiagnóstico", "\n".join(renglones),
               error=bool(problemas))
    sys.exit(1 if problemas else 0)


def _url_de_la_web(origenes):
    """La direccion que abre 'Abrir AutoFiller' en el menu del icono.

    Es el primer origen permitido: la misma lista que ya define de donde se
    acepta hablar, asi no hay una segunda configuracion que pueda quedar vieja.
    """
    for origen in origenes:
        if origen.startswith("http"):
            return origen
    return ""


def _con_bandeja(app, version, origenes, ruta_log):
    """Arranca el servidor en un hilo y se queda con el icono en el principal."""
    import uvicorn

    import bandeja as bandeja_mod
    import main as agente

    servidor = uvicorn.Server(uvicorn.Config(
        app, host="127.0.0.1", port=PUERTO, log_level="info"))
    falla = {}

    def servir():
        try:
            servidor.run()
        except BaseException as e:  # SystemExit incluido: uvicorn sale asi
            falla["error"] = e

    hilo = threading.Thread(target=servir, name="servidor", daemon=True)
    hilo.start()

    limite = time.monotonic() + SEGUNDOS_ARRANQUE
    while not servidor.started and time.monotonic() < limite:
        if not hilo.is_alive():
            break
        time.sleep(0.1)

    if not servidor.started:
        detalle = falla.get("error")
        bandeja_mod.avisar(
            "AutoFiller",
            f"El agente no pudo ponerse a escuchar en 127.0.0.1:{PUERTO}.\n\n"
            "Casi siempre es que el agente ya está abierto: miralo en los "
            "iconos ocultos, al lado del reloj.\n\n"
            + (f"Detalle: {detalle}\n\n" if detalle else "")
            + f"El registro completo está en:\n{ruta_log}",
            error=True)
        return 1

    def estado_texto():
        fase = agente.estado.fase
        if fase == "cargando":
            return f"Cargando el comprobante {agente.estado.indice} de {agente.estado.total}"
        if fase == "esperando":
            return "Esperando que confirmes en SISalud"
        if fase == "terminado":
            return "Comprobante resuelto"
        return "SISalud conectado" if agente.credenciales else "Listo"

    def ocupado():
        return agente.estado.fase in ("cargando", "esperando")

    def al_salir():
        print("  Cerrando el agente (Salir, desde el ícono de la bandeja).")
        servidor.should_exit = True
        hilo.join(timeout=10)

    icono = bandeja_mod.Bandeja(
        version=version,
        url_web=_url_de_la_web(origenes),
        estado_texto=estado_texto,
        ocupado=ocupado,
        al_salir=al_salir,
    )
    print(f"  Ícono en la bandeja del sistema. Web: {_url_de_la_web(origenes) or '(sin configurar)'}")
    icono.correr()
    return 0


def _en_consola(app):
    """El arranque de siempre: uvicorn en el hilo principal, todo a la vista."""
    import uvicorn

    print("  Dejá esta ventana abierta mientras cargás comprobantes.")
    print("  Para cerrarlo: Ctrl+C, o cerrá la ventana.")
    print()
    try:
        uvicorn.run(app, host="127.0.0.1", port=PUERTO, log_level="info")
    except OSError as e:
        # El caso tipico: ya hay otro agente corriendo. Decirlo es mejor que
        # escupir un traceback de sockets.
        print()
        print(f"  No se pudo tomar el puerto {PUERTO}: {e}")
        print("  Suele ser que el agente ya está abierto.")
        _esperar_enter()
        return 1
    return 0


def main():
    from main import ORIGENES, VERSION, app

    print(f"  AutoFiller — agente {VERSION}")
    print(f"  Escuchando en 127.0.0.1:{PUERTO}")
    print("  Orígenes permitidos:", ", ".join(ORIGENES) or "(ninguno)")

    # --consola fuerza el arranque viejo. Es la via de soporte: si la bandeja da
    # problemas en una PC, el agente sigue siendo usable sin recompilar nada.
    if "--consola" in sys.argv:
        return _en_consola(app)

    try:
        import bandeja  # noqa: F401  (solo para saber si esta disponible)
    except Exception as e:
        # Desarrollo desde el codigo sin pystray instalado. Con el .exe no pasa:
        # el empaquetado lo incluye y el autodiagnostico lo verifica.
        print(f"  Sin bandeja del sistema ({type(e).__name__}: {e}); queda en consola.")
        return _en_consola(app)

    return _con_bandeja(app, VERSION, ORIGENES, registro.ARCHIVO)


if __name__ == "__main__":
    # Antes de cualquier import que imprima: uvicorn se queda con el stream que
    # encuentra al arrancar, y sin consola un print sin redirigir revienta.
    registro.configurar()
    try:
        if "--probar" in sys.argv:
            probar()
        sys.exit(main() or 0)
    except KeyboardInterrupt:
        pass
    except SystemExit:
        raise
    except Exception as e:
        # Sin consola, un fallo de arranque se cerraria sin dejar rastro visible.
        # El log ya lo tiene; el cartel es para que el operador sepa que mirar.
        import traceback

        traceback.print_exc()
        if os.environ.get("AUTOFILLER_DEBUG"):
            raise
        try:
            from bandeja import avisar

            avisar("AutoFiller",
                   f"El agente no pudo arrancar:\n{type(e).__name__}: {e}\n\n"
                   f"El detalle completo está en:\n{registro.ARCHIVO}",
                   error=True)
        except Exception:
            _esperar_enter()
        sys.exit(1)
