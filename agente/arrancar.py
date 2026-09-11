"""Punto de entrada del agente empaquetado como .exe.

Existe por una razon concreta: el agente se arranca con `uvicorn main:app`, y
esa forma de nombrar la app la importa POR TEXTO. Adentro de un ejecutable de
PyInstaller no hay un `main.py` en el disco que importar, asi que ese arranque
falla. Aca se importa `app` de verdad y se le pasa el objeto a uvicorn.

Tambien es el lugar donde el agente se presenta: la consola queda a la vista a
proposito. La version de escritorio se compilaba con --noconsole y el operador
solo veia "Factura invalida" cuando algo fallaba; aca la ventana es la senal de
que el agente esta corriendo, y si algo revienta se ve.
"""

import os
import sys


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

    print(f"  AutoFiller — agente {VERSION} — autodiagnóstico")
    print()
    problemas = 0

    print("  Orígenes permitidos:", ", ".join(ORIGENES) or "(ninguno)")
    if not ORIGENES:
        print("     ^ sin orígenes el agente rechaza todo. Corré configurar.bat.")
        problemas += 1

    try:
        chrome = ruta_de_chrome()
        print(f"  Chrome: {chrome}")
    except Exception as e:
        print(f"  Chrome: NO SE ENCONTRÓ ({e})")
        print("     ^ instalá Chrome, o indicá la ruta con AUTOFILLER_CHROME.")
        problemas += 1

    print(f"  Perfil de Chrome: {PERFIL}   (puerto de depuración {PUERTO_CDP})")

    # Lo mas fragil del empaquetado: Playwright trae su propio node.exe y hay que
    # arrancarlo. Si el .exe quedo mal armado, revienta aca y no al cargar.
    async def arrancar_playwright():
        from playwright.async_api import async_playwright

        p = await async_playwright().start()
        await p.stop()

    try:
        asyncio.run(arrancar_playwright())
        print("  Playwright: el driver arranca bien")
    except Exception as e:
        print(f"  Playwright: FALLÓ ({type(e).__name__}: {e})")
        problemas += 1

    print()
    if problemas:
        print(f"  {problemas} problema(s). El agente no va a poder cargar comprobantes.")
    else:
        print("  Todo en orden. El agente puede cargar comprobantes en esta PC.")
    print()
    input("  Enter para cerrar.")
    sys.exit(1 if problemas else 0)


def main():
    import uvicorn

    # Sin esto, con la salida redirigida a un archivo de log el encabezado queda
    # bufferado y aparece DESPUES de los renglones de uvicorn: justo al reves de
    # lo que necesita leer el que esta diagnosticando.
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:
        pass

    from main import ORIGENES, VERSION, app

    print(f"  AutoFiller — agente {VERSION}")
    print("  Escuchando en 127.0.0.1:8765")
    print("  Origenes permitidos:", ", ".join(ORIGENES) or "(ninguno)")
    print()
    print("  Dejá esta ventana abierta mientras cargás comprobantes.")
    print("  Para cerrarlo: Ctrl+C, o cerrá la ventana.")
    print()

    try:
        uvicorn.run(app, host="127.0.0.1", port=8765, log_level="info")
    except OSError as e:
        # El caso tipico: ya hay otro agente corriendo. Decirlo es mejor que
        # escupir un traceback de sockets.
        print()
        print(f"  No se pudo tomar el puerto 8765: {e}")
        print("  Suele ser que el agente ya está abierto en otra ventana.")
        input("  Enter para cerrar.")
        sys.exit(1)


if __name__ == "__main__":
    # Sin esto, el .exe corrido con doble clic se cierra de golpe ante cualquier
    # error de arranque y no queda rastro de que paso.
    try:
        if "--probar" in sys.argv:
            probar()
        main()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print()
        print(f"  El agente no pudo arrancar: {type(e).__name__}: {e}")
        if os.environ.get("AUTOFILLER_DEBUG"):
            raise
        input("  Enter para cerrar.")
        sys.exit(1)
