"""Genera autofiller.ico a partir de servidor/web/logo.svg.

El .ico lo necesitan Windows y Tkinter (el icono del .exe, el de la bandeja y el
de la ventana de actividad del agente); el SVG es el original, y el unico lugar
donde se dibuja la marca. Este script existe para que no haya un segundo dibujo
que pueda quedar viejo: si cambia el logo, se corre esto y listo.

    python hacer_icono.py

Rasteriza con el Chrome que ya usa el agente (headless, fondo transparente) y
arma el .ico con Pillow. Se corre a mano, no en cada build: el resultado se
versiona.
"""

import os
import subprocess
import sys
import tempfile

from PIL import Image

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "agente"))
from navegador import ruta_de_chrome  # noqa: E402  (necesita el sys.path de arriba)

ORIGEN = os.path.join(AQUI, "servidor", "web", "logo.svg")
DESTINO = os.path.join(AQUI, "autofiller.ico")

# Windows elige el tamano segun donde muestre el icono y el DPI de la pantalla:
# 16/20/24/32 en la bandeja y la barra de titulo, 48 y 256 en el Explorador.
# Los que falten los escala el, y escalar de 256 a 16 deja el trazo hecho un
# borron: por eso van todos.
TAMANOS = [16, 20, 24, 32, 40, 48, 64, 128, 256]

# Se rasteriza una sola vez, grande y potencia de dos; Pillow reduce de ahi a
# cada tamano con LANCZOS al guardar el .ico.
LADO_RASTER = 1024


def rasterizar(destino_png):
    subprocess.run([
        ruta_de_chrome(),
        "--headless",
        "--disable-gpu",
        "--hide-scrollbars",
        # Sin esto el fondo sale blanco y el icono queda con un marco en las
        # esquinas redondeadas.
        "--default-background-color=00000000",
        "--force-device-scale-factor=1",
        f"--window-size={LADO_RASTER},{LADO_RASTER}",
        f"--screenshot={destino_png}",
        f"file:///{ORIGEN.replace(os.sep, '/')}",
    ], check=True, capture_output=True)


def main():
    with tempfile.TemporaryDirectory() as temporal:
        png = os.path.join(temporal, "logo.png")
        rasterizar(png)
        # convert() lee el archivo entero, asi que la imagen sobrevive al
        # directorio temporal.
        grande = Image.open(png).convert("RGBA")
    grande.save(DESTINO, format="ICO",
                sizes=[(lado, lado) for lado in TAMANOS])
    print(f"{DESTINO}  ({', '.join(str(lado) for lado in TAMANOS)})")


if __name__ == "__main__":
    main()
