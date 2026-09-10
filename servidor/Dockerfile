# syntax=docker/dockerfile:1
#
# Imagen del servidor de extraccion. El contexto de build es `servidor/`, no la raiz del
# repo: el agente corre en la PC del operador y la version de escritorio no tiene nada
# que hacer adentro de esta imagen.
#
# Dos apartamientos de lo que hacen FactuMov y Balance360, los dos a proposito:
#
#   - No usa la imagen de `uv`: aca no hay pyproject.toml ni lockfile, hay un
#     requirements.txt. Migrar a uv seria consistencia sin beneficio.
#   - No instala ninguna libreria de sistema. FactuMov necesita Pango por WeasyPrint;
#     aca todo lo que lee comprobantes (pikepdf, pypdfium2, pillow-heif, zxing-cpp)
#     viene en wheels manylinux.

FROM python:3.11-slim

# PYTHONUNBUFFERED: los logs salen al instante. Importa mas que de costumbre porque el
#   renglon con los tokens y el costo de cada llamada de vision es lo unico que dice
#   cuanto se gasto en leer una tanda.
# PYTHONDONTWRITEBYTECODE: el codigo no cambia adentro del contenedor; los .pyc solo
#   ensucian la imagen.
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# --- Capa 1: dependencias ---
# Solo el requirements. Mientras no cambie, Docker reutiliza esta capa y un cambio de
# codigo no reinstala nada.
COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# --- Capa 2: el codigo ---
# Incluye web/, que sirve el propio FastAPI con StaticFiles: no hay SPA que compilar ni
# un nginx que la sirva, a diferencia de FactuMov.
COPY . .

EXPOSE 8000

# Forma JSON (exec): asi el proceso 1 recibe el SIGTERM del deploy y uvicorn cierra
# bien. Sin eso, cada redeploy termina en un kill -9 a los 10 segundos.
CMD ["python", "-m", "uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
