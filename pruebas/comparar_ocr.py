"""Compara el camino OCR (Cloud Vision + regex) contra el camino PDF, ya verificado.

La pregunta que responde
------------------------
Para las fotos hay dos caminos posibles: modelo de vision (funciona, cuesta
plata por comprobante) u OCR clasico + las regex de ARCA que ya tenemos
(gratis hasta 1000 paginas por mes). El OCR sabe leer las letras; lo que no se
sabe es si devuelve el texto en un orden que las regex de texto.py puedan
parsear, porque esas regex se escribieron contra el orden de lectura de
pdfplumber sobre PDFs nativos.

Como lo prueba sin necesitar fotos
----------------------------------
Las 63 muestras son PDFs nativos, asi que de cada una se pueden sacar las dos
cosas a la vez:

  verdad    = texto del PDF        -> datos_desde_texto()   (63/63 verificado)
  candidato = pagina rasterizada   -> Cloud Vision -> datos_desde_texto()

Misma factura, mismas regex, lo unico que cambia es de donde sale el texto. Eso
aisla el riesgo de orden de lectura del riesgo de foto de celular (angulo,
sombra, foco), que es un segundo experimento y necesita fotos reales.

Ojo con lo que esta prueba NO dice
----------------------------------
Que Cloud Vision lea bien un raster a 200 dpi no garantiza que lea bien una
foto sacada a mano. Si esta prueba sale bien, el paso siguiente es repetirla
con fotos de verdad. Si sale mal, ya sabemos que el camino OCR no cierra sin
tocar las regex, y cuanto habria que tocarlas.

Uso
---
    python pruebas/comparar_ocr.py            # las 63 muestras
    python pruebas/comparar_ocr.py --limite 5 # una prueba corta primero
    python pruebas/comparar_ocr.py --solo-cache  # sin gastar llamadas nuevas
"""

import argparse
import difflib
import hashlib
import io
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "servidor"))

import cloud_vision  # noqa: E402  (esta al lado de este archivo)
from extraccion import DPI_RASTERIZADO  # noqa: E402
from extraccion import _primera_pagina_como_imagen, _texto_del_pdf  # noqa: E402
from extraccion.texto import datos_desde_texto  # noqa: E402

MUESTRAS = RAIZ / "samples"
REPORTE = Path(__file__).parent / "resultado_ocr.md"

# Los campos que en produccion ya vienen firmados por el QR de ARCA: aunque el
# OCR los lea mal, el QR los pisa. Se miden igual, pero no deciden nada.
CAMPOS_QR = [
    "cuit",
    "tipo_comprobante",
    "punto_venta",
    "nro_factura",
    "fecha_emision",
    "cae",
    "importe",
]

# Los campos que el QR NO trae. Estos son los que decide esta prueba: si el OCR
# no los saca, el camino OCR no sirve y hay que ir al modelo de vision.
CAMPOS_OCR = [
    "fecha_vencimiento",   # el 'Hasta:' del periodo facturado
    "descripcion",
    "domicilio",
    "provincia",           # derivado del domicilio
    "centro_costo",        # derivado del domicilio
]


def normalizar(valor):
    return " ".join((valor or "").split()).strip()


def similitud(a, b):
    return difflib.SequenceMatcher(None, normalizar(a), normalizar(b)).ratio()


def jpeg_de(imagen, calidad=90):
    buffer = io.BytesIO()
    imagen.convert("RGB").save(buffer, format="JPEG", quality=calidad)
    return buffer.getvalue()


def clave_cache(contenido_pdf):
    """Identidad estable de lo que se le manda a Cloud Vision.

    Describe el origen (que PDF, a que dpi) y no los bytes del JPEG, que
    cambian con la version de Pillow. Ver _ruta_cache en cloud_vision.py.
    """
    return f"{hashlib.sha256(contenido_pdf).hexdigest()}|dpi={DPI_RASTERIZADO}"


def preparar(rutas):
    """[(ruta, jpeg, clave, error)] rasterizando la primera pagina de cada muestra.

    Se hace en un paso aparte y una sola vez: sirve para contar cuantas
    llamadas nuevas a Cloud Vision hacen falta antes de gastarlas, y para que
    un PDF que no se puede leer se vea como lo que es en vez de confundirse
    con uno que simplemente no esta cacheado.
    """
    preparadas = []
    for ruta in rutas:
        try:
            contenido = ruta.read_bytes()
            jpeg = jpeg_de(_primera_pagina_como_imagen(contenido))
            preparadas.append((ruta, jpeg, clave_cache(contenido), None))
        except Exception as e:
            preparadas.append((ruta, None, None, f"no se pudo rasterizar: {e}"))
    return preparadas


def comparar_una(ruta, jpeg, clave):
    """(verdad, candidato, texto_ocr) para una muestra. Lanza si el OCR falla."""
    verdad, _ = datos_desde_texto(_texto_del_pdf(ruta.read_bytes()))
    texto_ocr = cloud_vision.texto(jpeg, clave)
    candidato, _ = datos_desde_texto(texto_ocr)
    return verdad, candidato, texto_ocr


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--limite", type=int, help="probar solo las primeras N muestras")
    parser.add_argument("--solo-cache", action="store_true",
                        help="no hacer llamadas nuevas: usar unicamente lo cacheado")
    args = parser.parse_args()

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

    rutas = sorted(MUESTRAS.glob("*.pdf"))
    if args.limite:
        rutas = rutas[: args.limite]
    if not rutas:
        print(f"No hay PDFs en {MUESTRAS}.")
        return 1

    # Avisar cuanta cuota se va a gastar ANTES de gastarla: el argumento a favor
    # de este camino es que sea gratis, seria ironico quemar el free tier aca.
    preparadas = preparar(rutas)
    ilegibles = [(r, e) for r, jpeg, clave, e in preparadas if e]
    legibles = [(r, jpeg, clave) for r, jpeg, clave, e in preparadas if not e]
    pendientes = {r for r, jpeg, clave in legibles if not cloud_vision.en_cache(clave)}

    print(f"Muestras: {len(rutas)} | ya cacheadas: {len(legibles) - len(pendientes)} "
          f"| llamadas nuevas a Cloud Vision: {len(pendientes)}")
    for ruta, error in ilegibles:
        print(f"  ERROR  {ruta.name}: {error}")

    if pendientes and args.solo_cache:
        print("--solo-cache: se saltean las que no estan cacheadas.")
    elif pendientes and not cloud_vision.disponible():
        print("\nFalta GOOGLE_VISION_API_KEY en ESTA consola.")
        print("  setx la guarda para las consolas nuevas; para la actual:")
        print('  $env:GOOGLE_VISION_API_KEY = "AIza..."')
        return 1
    print()

    aciertos = {c: 0 for c in CAMPOS_QR + CAMPOS_OCR}
    medidos = {c: 0 for c in CAMPOS_QR + CAMPOS_OCR}
    similitudes = {c: [] for c in CAMPOS_QR + CAMPOS_OCR}
    diferencias = []
    fallados = [(ruta.name, error) for ruta, error in ilegibles]
    comparadas = 0

    for ruta, jpeg, clave in legibles:
        if args.solo_cache and ruta in pendientes:
            continue
        try:
            verdad, candidato, _ = comparar_una(ruta, jpeg, clave)
        except Exception as e:
            fallados.append((ruta.name, str(e)))
            print(f"  ERROR  {ruta.name}: {e}")
            continue

        comparadas += 1
        diferencias_archivo = []
        for campo in CAMPOS_QR + CAMPOS_OCR:
            esperado, obtenido = verdad.get(campo), candidato.get(campo)
            medidos[campo] += 1
            similitudes[campo].append(similitud(esperado, obtenido))
            if normalizar(esperado) == normalizar(obtenido):
                aciertos[campo] += 1
            else:
                diferencias_archivo.append((campo, esperado, obtenido))

        criticos = [c for c, _, _ in diferencias_archivo if c in CAMPOS_OCR]
        marca = "OK    " if not criticos else "DIFIERE"
        print(f"  {marca} {ruta.name}"
              + (f"  -> {', '.join(criticos)}" if criticos else ""))
        if diferencias_archivo:
            diferencias.append((ruta.name, diferencias_archivo))

    total = comparadas
    if total <= 0:
        print("\nNo se pudo comparar ninguna muestra.")
        return 1

    def tabla(campos, titulo):
        lineas = [f"\n{titulo}", "-" * len(titulo),
                  f"{'campo':<20} {'exactos':>12}   {'similitud media':>15}"]
        for campo in campos:
            n = medidos[campo] or 1
            media = sum(similitudes[campo]) / n
            lineas.append(
                f"{campo:<20} {aciertos[campo]:>5}/{medidos[campo]:<6} "
                f"{100 * aciertos[campo] / n:>4.0f}%   {media:>14.1%}"
            )
        return "\n".join(lineas)

    print(tabla(CAMPOS_OCR, "CAMPOS QUE DECIDEN (el QR no los trae)"))
    print(tabla(CAMPOS_QR, "Campos que en produccion pisa el QR (informativo)"))

    con_falla_critica = sum(
        1 for _, difs in diferencias
        if any(campo in CAMPOS_OCR for campo, _, _ in difs)
    )
    criticos_ok = total - con_falla_critica
    print(f"\nVEREDICTO: {criticos_ok}/{total} muestras con los campos que "
          f"deciden identicos al camino PDF.")
    if fallados:
        print(f"           {len(fallados)} muestras no se pudieron leer.")

    with REPORTE.open("w", encoding="utf-8") as f:
        f.write("# OCR (Cloud Vision) vs. camino PDF\n\n")
        f.write(f"Muestras comparadas: {total}. "
                f"Con los campos que deciden identicos: {criticos_ok}.\n\n")
        f.write("Campos que deciden: " + ", ".join(CAMPOS_OCR) + "\n\n")
        for nombre, difs in diferencias:
            f.write(f"## {nombre}\n\n")
            for campo, esperado, obtenido in difs:
                marca = "**" if campo in CAMPOS_OCR else ""
                f.write(f"- {marca}{campo}{marca}\n")
                f.write(f"  - PDF : `{esperado}`\n")
                f.write(f"  - OCR : `{obtenido}`\n")
            f.write("\n")
        if fallados:
            f.write("## No se pudieron leer\n\n")
            for nombre, error in fallados:
                f.write(f"- {nombre}: {error}\n")
    print(f"\nDetalle de cada diferencia: {REPORTE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
