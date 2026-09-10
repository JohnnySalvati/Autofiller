# AutoFiller

Carga comprobantes de ARCA (PDF o foto) en **Prestadores → Carga Rápida Comprobantes**
de SISalud. Cliente: AOMAOSAM. Desarrollado por InSoft.

## Cómo está armado

SISalud no tiene API: la carga tiene que seguir siendo automatización del navegador,
y el operador tiene que ver la pantalla real antes de confirmar. Por eso la app está
partida en dos, y solo la mitad que **lee** los comprobantes vive en el servidor:

```
   Navegador del operador
   ┌─────────────────────────────┐
   │  Web de AutoFiller          │──── PDF / foto ────►  servidor/   (donde sea)
   │  (revisar y corregir)       │◄─── campos JSON ────  lee el comprobante
   └──────────┬──────────────────┘
              │ factura + credenciales de SISalud
              ▼  (127.0.0.1, nunca salen de la PC)
        agente/  ──► Chrome del operador ──► SISalud
        (Playwright)     el operador confirma acá
```

- **`servidor/`** — FastAPI. Recibe el archivo, devuelve los campos y sirve la web.
  No toca SISalud y no guarda nada: cada archivo se procesa en memoria y se descarta.
  Es el único que necesita la clave de la API de visión, así que esa clave ya no
  viaja en ningún ejecutable.
- **`agente/`** — FastAPI en `127.0.0.1:8765`, corre en la PC del operador. Es lo
  único que toca SISalud. Recibe la factura ya revisada, la carga en el Chrome del
  operador y espera a que el operador Confirme o Cancele. **Nunca confirma solo.**
- **`AutoFiller.py`** — la versión de escritorio (Tkinter). Sigue funcionando; queda
  hasta que la web esté desplegada.

## Qué lee

| Entra | Cómo se lee |
|---|---|
| PDF de ARCA | `pdfplumber` + regex. Exacto y gratis. |
| PDF escaneado | Se rasteriza la primera página y se trata como foto. |
| Foto (JPG/PNG/HEIC/…) | QR de ARCA para los datos fiscales + OCR de Cloud Vision, y los mismos regex que un PDF, para el período, la descripción y el domicilio. Si el OCR no sirve, un modelo de visión de Claude. |

El QR trae CUIT, tipo, punto de venta, número, fecha, importe y CAE **firmados por
ARCA**, así que pisa lo que devuelva cualquiera de los dos motores. La lectura de la
imagen solo aporta lo que el QR no trae.

**Por qué el OCR va primero**: Google regala 1000 páginas por mes —al volumen de
AOMAOSAM, probablemente gratis para siempre— y devuelve siempre el mismo texto,
mientras que el modelo puede variar entre corridas. Se midió contra las 63 muestras:
63/63 en período, centro de costos y provincia. El modelo queda de respaldo, para
cuando el OCR falla o devuelve algo que los regex de ARCA no parsean.

## Arrancar

Para **producción** —la VM detrás de srv-nginx, en `autofiller.insoft.net.ar`— el
procedimiento entero está en [`docs/DEPLOYMENT.md`](docs/DEPLOYMENT.md). Lo de acá abajo
es para levantarlo a mano en una máquina cualquiera.

**Servidor** (una vez, donde vaya a quedar):

```bat
set GOOGLE_VISION_API_KEY=...
set ANTHROPIC_API_KEY=sk-ant-...
servidor\iniciar.bat
```

Queda en `http://localhost:8000`. Para leer fotos alcanza con **cualquiera de las dos
claves**; sin ninguna anda igual, pero solo con PDF, y la web lo avisa con el chip
«Fotos: no».

**Agente** (en cada PC que carga comprobantes):

```bat
agente\iniciar.bat
```

Tiene que quedar abierto mientras se usa la web. Si el servidor no está en
`localhost:8000`, hay que decírselo antes de arrancarlo:

```bat
set AUTOFILLER_ORIGENES=https://autofiller.aomaosam.org.ar
```

Es una lista blanca a propósito: el agente escucha en `127.0.0.1`, así que sin ella
cualquier página abierta en el navegador podría hacerle cargar comprobantes.

Playwright se engancha por CDP al Chrome del operador, así que **no hace falta
`playwright install`**: no usa los navegadores que descarga, solo el driver del
paquete.

## Desarrollo

Dos terminales, una para cada mitad. La primera vez cada `.venv` se crea solo.

**Terminal 1 — servidor** (recarga sola al guardar):

```bat
cd servidor
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
set ANTHROPIC_API_KEY=sk-ant-...
python -m uvicorn main:app --reload --port 8000
```

**Terminal 2 — agente**:

```bat
cd agente
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m uvicorn main:app --host 127.0.0.1 --port 8765
```

Y listo: `http://localhost:8000`. El origen `http://localhost:8000` ya está en la
lista blanca del agente por defecto, así que no hay nada más que configurar.

Si solo estás tocando la web o la extracción, alcanza con el servidor: la página
carga igual y avisa «No se detecta el agente» en la pantalla de Entrar.

Al tocar «Cargar en SISalud» el agente abre Chrome con
`--remote-debugging-port=9222` y el perfil `C:\ChromeProfile`. Tres cosas para
tener presentes mientras se desarrolla:

- **No toques Confirmar (ni F12) probando: graba comprobantes de verdad.** Para
  cerrar el ciclo usá Cancelar, o «Saltar este» en el cartel.
- **No corras el agente con `--reload`** mientras probás una cola: cada recarga
  corta la conexión CDP y pierde la sesión en medio de la carga. Reiniciarlo a
  mano entre pruebas es más previsible.
- **No dejes varias pestañas de la Carga Rápida abiertas**: se pisan la sesión y
  dejan puesta la máscara de GeneXus. Si la pantalla queda trabada, F5 la libera.

Probar la extracción no necesita nada levantado:

```bash
cd servidor
python -c "from extraccion import extraer;     print(extraer('x.pdf', open('../samples/x.pdf','rb').read()).factura)"
```

## Variables de entorno

| Variable | Dónde | Para qué |
|---|---|---|
| `GOOGLE_VISION_API_KEY` | servidor | OCR de Cloud Vision: el motor por defecto para fotos. |
| `ANTHROPIC_API_KEY` | servidor | Modelo de visión: el respaldo cuando el OCR no sirve. |
| `AUTOFILLER_MOTOR_LECTURA` | servidor | `auto` (default), `ocr` o `vision`. Forzar uno sirve para medirlos. |
| `AUTOFILLER_MODELO_VISION` | servidor | Modelo de visión. Por defecto `claude-opus-5`. |
| `AUTOFILLER_ESFUERZO_VISION` | servidor | `low`…`max`. Por defecto `medium`. |
| `AUTOFILLER_TAMANIO_MAXIMO_MB` | servidor | Tope por archivo. Por defecto 25. |
| `AUTOFILLER_PUERTO` | servidor | Puerto. Por defecto 8000. |
| `AUTOFILLER_ORIGENES` | agente | Orígenes que pueden hablarle al agente. |
| `AUTOFILLER_CHROME` | agente | Ruta a `chrome.exe` si no está donde se espera. |
| `AUTOFILLER_PERFIL_CHROME` | agente | Perfil de Chrome. Por defecto `C:\ChromeProfile`. |
| `AUTOFILLER_PUERTO_CDP` | agente | Puerto de depuración de Chrome. Por defecto 9222. |

El detalle de cómo se comporta la pantalla de SISalud frente a la automatización
—y por qué el código hace lo que hace— está en `CLAUDE.md`.
