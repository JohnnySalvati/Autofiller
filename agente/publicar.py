"""Escribe la ficha del zip del agente: version y sha256.

Es lo que el servidor lee para contestar `/api/agente`, y lo que cada agente
instalado compara contra la version que tiene. Dos archivos y no uno solo
—`AutoFillerAgente.zip` y `AutoFillerAgente.json`— porque la version no se puede
sacar del zip sin descomprimirlo, y porque asi el servidor no necesita ni abrir
el zip para contestar.

La version sale de `main.py` con una expresion regular a proposito, en vez de
importarlo: importar `main` arrastra FastAPI, Playwright y todo el agente para
leer una constante.

Lo corre `empaquetar.bat` despues de armar el zip. A mano:

    .venv\\Scripts\\python.exe publicar.py
"""

import hashlib
import json
import re
import sys
from pathlib import Path

ACA = Path(__file__).resolve().parent
ZIP = ACA / "AutoFillerAgente.zip"
FICHA = ACA / "AutoFillerAgente.json"


def version_del_agente():
    texto = (ACA / "main.py").read_text(encoding="utf-8")
    m = re.search(r'^VERSION\s*=\s*"([^"]+)"', texto, re.MULTILINE)
    if not m:
        raise SystemExit("No se encontró VERSION en main.py.")
    return m.group(1)


def sha256(ruta):
    digest = hashlib.sha256()
    with open(ruta, "rb") as f:
        for pedazo in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(pedazo)
    return digest.hexdigest()


def main():
    if not ZIP.exists():
        raise SystemExit(f"No está {ZIP.name}. Corré empaquetar.bat primero.")

    datos = {
        "version": version_del_agente(),
        "sha256": sha256(ZIP),
        "tamano": ZIP.stat().st_size,
    }
    FICHA.write_text(json.dumps(datos, indent=2) + "\n", encoding="utf-8")

    print(f"  Versión publicada: {datos['version']}")
    print(f"  {ZIP.name}: {datos['tamano'] // (1024 * 1024)} MB")
    print(f"  sha256: {datos['sha256']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
