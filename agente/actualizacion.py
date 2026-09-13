"""Actualizacion automatica del agente.

El agente vive en la PC de cada operador, y son PCs a las que no se entra: hasta
ahora una version nueva significaba compilar, pasarle el zip a cada uno y
confiar en que lo descomprimiera encima del anterior. Con dos operadores ya era
incomodo; con la version que arregla algo urgente, es el camino por el que una
PC se queda vieja para siempre y nadie se entera.

Como funciona:

  1. Cada tanto le pregunta al servidor de AutoFiller que version publico
     (`/api/agente`) y la compara con la propia.
  2. Si hay una mas nueva, baja el zip, le verifica el sha256 y lo descomprime
     en %LOCALAPPDATA%\\AutoFiller\\actualizacion\\nueva.
  3. Escribe un .bat, lo lanza suelto y se apaga. El .bat espera a que el .exe
     cierre de verdad, copia los archivos encima de la instalacion y vuelve a
     arrancar el agente.

El paso 3 es asi por Windows: un .exe corriendo esta tomado y no se puede
reemplazar (PyInstaller onedir tambien tiene tomado su `_internal`). Alguien
tiene que sobrevivir al agente para copiar, y ese alguien no puede ser el
agente. El .bat se escribe en el momento en vez de venir adentro del paquete
justamente para que no quede en la carpeta que se va a pisar: un .bat copiado
encima de si mismo mientras cmd lo esta leyendo es un cuelgue seguro.

**Solo se actualiza cuando no hay nada en juego**: sin comprobante en pantalla y
sin sesion de SISalud abierta. La sesion son las credenciales en memoria, y
reiniciar las borra: hacerlo a mitad del dia le haria tipear el usuario y la
clave de nuevo sin que el operador entienda por que. En la practica la ventana
que siempre existe es el arranque de Windows, que es cuando el agente todavia no
tiene sesion.
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
import zipfile
from pathlib import Path

# Cada cuanto se pregunta. El agente se actualiza en la practica al arrancar con
# Windows, asi que esto es para la PC que queda prendida dias: mas seguido no
# sirve de nada y le pega al servidor al pedo.
CADA = float(os.environ.get("AUTOFILLER_INTERVALO_ACTUALIZACION", "4")) * 3600

# Antes de la primera consulta. Al arrancar con Windows la red puede no estar
# lista todavia, y ademas conviene que el agente este escuchando antes de
# ponerse a bajar nada.
DEMORA_INICIAL = 90

TIEMPO_RED = 30                       # timeout de la consulta y de la descarga
TAMANIO_MAXIMO = 300 * 1024 * 1024    # techo de cordura para el zip

CARPETA = (Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
           / "AutoFiller" / "actualizacion")

NOMBRE_EXE = "AutoFillerAgente.exe"


def empaquetado():
    """Si esto corre desde el .exe. Desde el codigo no hay nada que reemplazar."""
    return bool(getattr(sys, "frozen", False))


def origen_de(origenes):
    """De donde se baja: el mismo origen del que se acepta hablar.

    No hay una segunda configuracion para esto a proposito. Si el agente ya
    confia en ese origen para que le haga cargar comprobantes en SISalud,
    confiar en el para bajar su propia actualizacion no agrega superficie.
    """
    for origen in origenes:
        if origen.startswith("http"):
            return origen.rstrip("/")
    return ""


def _numeros(version):
    partes = []
    for parte in str(version).split("."):
        digitos = "".join(c for c in parte if c.isdigit())
        partes.append(int(digitos) if digitos else 0)
    return partes


def mas_nueva(remota, actual):
    """Si `remota` hay que instalarla sobre `actual`.

    Solo hacia adelante: si el servidor publica una version mas vieja que la
    instalada (un zip subido por error), no se toca nada.
    """
    if not remota:
        return False
    return _numeros(remota) > _numeros(actual)


def consultar(origen):
    """Que version publico el servidor, o None si no contesta o no publico nada.

    Nunca lanza: que el servidor este caido o sin zip publicado no puede hacer
    ruido en el agente, que sigue funcionando igual.
    """
    try:
        with urllib.request.urlopen(origen + "/api/agente", timeout=TIEMPO_RED) as r:
            info = json.loads(r.read().decode("utf-8"))
    except Exception as e:
        print(f"  Actualización: no se pudo consultar al servidor "
              f"({type(e).__name__}: {e}).")
        return None
    if not info.get("version") or not info.get("url"):
        return None
    return info


def _descargar(url, destino):
    destino.parent.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256()
    total = 0
    with urllib.request.urlopen(url, timeout=TIEMPO_RED) as r, open(destino, "wb") as f:
        while True:
            pedazo = r.read(256 * 1024)
            if not pedazo:
                break
            total += len(pedazo)
            if total > TAMANIO_MAXIMO:
                raise RuntimeError("el archivo pesa demasiado")
            digest.update(pedazo)
            f.write(pedazo)
    return total, digest.hexdigest()


def preparar(info, origen):
    """Baja el zip, lo verifica y lo descomprime. Devuelve la carpeta lista.

    El sha256 no esta por desconfiar del servidor —la descarga va por HTTPS—
    sino por la descarga cortada: media copia de un .exe descomprime igual de mal
    y el sintoma seria un agente que no arranca nunca mas.
    """
    if CARPETA.exists():
        shutil.rmtree(CARPETA, ignore_errors=True)
    CARPETA.mkdir(parents=True, exist_ok=True)

    url = info["url"]
    if url.startswith("/"):
        url = origen + url
    zip_ = CARPETA / "AutoFillerAgente.zip"

    print(f"  Actualización: bajando la versión {info['version']} de {url}")
    tamanio, sha = _descargar(url, zip_)

    esperado = (info.get("sha256") or "").lower()
    if esperado and sha != esperado:
        raise RuntimeError("el sha256 del zip no coincide con el publicado")

    nueva = CARPETA / "nueva"
    with zipfile.ZipFile(zip_) as z:
        z.extractall(nueva)

    if not (nueva / NOMBRE_EXE).exists():
        raise RuntimeError(f"el zip no trae {NOMBRE_EXE} en la raíz")

    zip_.unlink(missing_ok=True)
    print(f"  Actualización: {tamanio // 1024} KB descomprimidos en {nueva}")
    return nueva


# El que hace el trabajo de verdad, ya sin el agente prendido. Sin acentos ni
# nada fuera de ASCII: el .bat lo lee cmd con la codepage de la PC, que no es la
# misma en todas.
GUION = r"""@echo off
REM Generado por el agente de AutoFiller. Se borra en la proxima actualizacion.
setlocal
set "LOG={log}"
set "NUEVA={nueva}"
set "DESTINO={destino}"

REM Por ruta completa y no por nombre: el PATH que hereda esto es el del agente,
REM y alcanza con que la PC tenga un find o un ping de otra cosa adelante (un
REM Git for Windows, por ejemplo) para que el guion haga cualquier cosa. Medido:
REM con el find de MSYS en el PATH, la espera da "no esta corriendo" siempre y la
REM copia arranca con el agente todavia abierto.
set "TASKLIST=%SystemRoot%\System32\tasklist.exe"
set "FIND=%SystemRoot%\System32\find.exe"
set "ROBOCOPY=%SystemRoot%\System32\robocopy.exe"
set "PING=%SystemRoot%\System32\ping.exe"

echo. >>"%LOG%"
echo [%date% %time%] Aplicando la actualizacion >>"%LOG%"

REM Esperar a que el agente cierre del todo: mientras el .exe viva, Windows no
REM deja reemplazarlo. 120 intentos de 1 segundo y se abandona; el agente viejo
REM sigue andando, que es el peor caso aceptable.
set INTENTOS=0
:esperar
"%TASKLIST%" /FI "IMAGENAME eq {exe}" | "%FIND%" /I "{exe}" >nul
if errorlevel 1 goto :copiar
set /a INTENTOS+=1
if %INTENTOS% GEQ 120 (
    echo [%date% %time%] El agente sigue abierto: se abandona. >>"%LOG%"
    exit /b 1
)
"%PING%" -n 2 127.0.0.1 >nul
goto :esperar

:copiar
REM /e y no /mir: /mir borraria del destino cualquier cosa que el operador haya
REM dejado al lado del agente. Un archivo viejo de mas no molesta a nadie.
REM
REM /is (copiar tambien los que considera iguales) porque robocopy decide por
REM tamanio y fecha, y eso no es lo que se quiere aca: lo que se quiere es que
REM la instalacion quede igual al zip, sin excepciones. Medido: sin /is, dos
REM archivos distintos del mismo tamanio y la misma fecha se saltean y el
REM robocopy informa "copia terminada" igual.
"%ROBOCOPY%" "%NUEVA%" "%DESTINO%" /e /is /r:3 /w:2 /njh /njs /ndl /nfl >>"%LOG%" 2>&1
if errorlevel 8 (
    echo [%date% %time%] Fallo la copia. Se arranca igual el agente. >>"%LOG%"
) else (
    echo [%date% %time%] Copia terminada. >>"%LOG%"
)

start "" "%DESTINO%\{exe}"
echo [%date% %time%] Agente arrancado de nuevo. >>"%LOG%"
exit /b 0
"""


def aplicar(nueva):
    """Lanza el .bat que reemplaza los archivos, y vuelve enseguida.

    Despues de esto hay que apagar el agente: el .bat esta esperando a que el
    .exe cierre para poder copiar.
    """
    destino = Path(sys.executable).parent
    log = (Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
           / "AutoFiller" / "actualizacion.log")
    log.parent.mkdir(parents=True, exist_ok=True)

    guion = CARPETA / "aplicar.bat"
    guion.write_text(
        GUION.format(log=log, nueva=nueva, destino=destino, exe=NOMBRE_EXE),
        encoding="ascii")

    # Suelto del agente y sin ventana: DETACHED_PROCESS para que sobreviva al
    # proceso que lo lanza (que es todo el punto), CREATE_NO_WINDOW para no
    # plantarle una consola negra al operador en medio de la pantalla.
    DETACHED_PROCESS, CREATE_NO_WINDOW = 0x00000008, 0x08000000
    subprocess.Popen(
        ["cmd", "/c", str(guion)],
        cwd=str(CARPETA),
        creationflags=DETACHED_PROCESS | CREATE_NO_WINDOW,
        close_fds=True,
    )
    print(f"  Actualización: aplicándola. El detalle queda en {log}")


def _ciclo(version, origen, libre, reiniciar):
    """Una vuelta completa. Devuelve si el agente se va a reiniciar."""
    info = consultar(origen)
    if not info or not mas_nueva(info["version"], version):
        return False

    print(f"  Actualización: hay una versión nueva ({info['version']}, "
          f"tengo {version}).")

    if not empaquetado():
        print("  Actualización: corriendo desde el código no se actualiza sola.")
        return False
    if not libre():
        print("  Actualización: hay trabajo en curso; queda para la próxima vuelta.")
        return False

    nueva = preparar(info, origen)

    # De nuevo: la descarga tarda, y en ese rato el operador pudo haberse puesto
    # a cargar comprobantes. Reiniciar ahora le cortaria el lote.
    if not libre():
        print("  Actualización: se empezó a trabajar mientras bajaba; "
              "queda para la próxima.")
        return False

    aplicar(nueva)
    reiniciar(info["version"])
    return True


def vigilar(*, version, origenes, libre, reiniciar, terminar):
    """Arranca el hilo que mira si hay version nueva.

    `libre()` dice si se puede reiniciar sin arruinarle nada al operador, y
    `reiniciar(version)` apaga el agente (el .bat ya esta esperando para copiar).
    `terminar` es el Event de la bandeja: con el agente cerrandose, este hilo no
    tiene que quedar colgado esperando su turno.
    """
    origen = origen_de(origenes)
    if not origen:
        print("  Actualización: sin origen configurado, no se busca ninguna.")
        return

    def bucle():
        if terminar.wait(DEMORA_INICIAL):
            return
        while True:
            try:
                if _ciclo(version, origen, libre, reiniciar):
                    return
            except Exception as e:
                # Nada de lo que pase aca justifica tirar abajo el agente: el
                # operador tiene que poder seguir cargando con la version vieja.
                print(f"  Actualización: falló ({type(e).__name__}: {e}).")
            if terminar.wait(CADA):
                return

    threading.Thread(target=bucle, name="actualizacion", daemon=True).start()
