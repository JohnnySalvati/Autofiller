# AutoFiller — contexto del proyecto

App de escritorio (Python 3.11, Tkinter + Playwright, compilada con PyInstaller a `dist/AutoFiller.exe`) que extrae datos de una factura electrónica de ARCA en PDF y los carga en la pantalla **Prestadores → Carga Rápida Comprobantes** del sistema web SISalud (GeneXus, `http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/cargarapidacomprobantescompra`). Cliente: obra social AOMAOSAM. Desarrollado por InSoft (Johnny Salvati). Comunicación en español rioplatense.

## Cómo funciona hoy (`AutoFiller.py`, único archivo, ~509 líneas)

1. `decrypt_pdf()` abre el PDF con `pikepdf` y lo guarda como `decrypted.pdf` en el cwd.
2. `extract_information()` saca el texto con `pdfplumber` y con regex obtiene:
   - `tipo_comprobante`: `COD. 011` → `TIPOS_COMPROBANTE` devuelve el **`value`** de la opción del combo `#vTIPOCOMPROBANTECODIGO` (`"6"`/`"11"` → `FACCC`, `"15"` → `NCCC`).
   - `cuit` (11 dígitos tras `CUIT:`), `punto_venta` y `nro_factura` (tras `Punto de Venta:` / `Comp. Nro:`, sin ceros a la izquierda), `fecha_emision`, `fecha_hasta` (tras `Hasta:`, usada como vencimiento y devengamiento), `cae`, `descripcion` (texto entre los dos `Subtotal`), `importe` (`Importe Total:`).
   - `provincia` y `centro_costo` a partir del `Domicilio Comercial:` (ver más abajo).
   - `fecha_recepcion` = hoy.
3. `main()` lanza Chrome con `--remote-debugging-port=9222 --user-data-dir=C:\ChromeProfile`, se conecta por CDP, loguea y navega a la Carga Rápida. Orden de carga (ver sección de comportamiento): **prestador primero** (`elegir_prestador`, prompt `#PROMPTIMGENTIDAD` por CUIT, elige la fila del CUIT), luego **tipo** (`elegir_tipo_comprobante`, por value), luego el resto de la cabecera y el detalle con `llenar()` (`fill` sin blur): `#vCOMPROBANTEPREFIJO`, `#vCOMPROBANTECODIGO`, las fechas, `#vCOMPROBANTECAE`, `#vEXENTOCOMPROBANTEDETALLEDESCRIPCION`, `#vEXENTOCOMPROBANTEDETALLEPRECIOUNITARIO`, `#vEXENTOCENTROCOSTOCODIGO`. Solo clickea `#IMAGE3` (agregar línea) si el tipo quedó cargado. Devuelve una lista de avisos.
4. GUI Tkinter: usuario, contraseña, "Seleccionar" PDF, "Procesar". Tras procesar hace `sys.exit(0)` (un PDF por ejecución).
5. **Avisos al operador**: `mostrar_avisos_en_pantalla()` inyecta un cartel rojo fijo (`#autofiller-avisos`, z-index máximo) arriba de la propia pantalla de SISalud, donde el operador lo ve antes de confirmar. El `messagebox` de Tkinter quedó solo como respaldo forzado al frente (`-topmost`) porque por defecto aparecía **detrás** del navegador y el operador no lo veía hasta después de confirmar y cerrar.

Hay 63 PDFs reales en `samples/` (ver "Muestras" abajo). `decrypted.pdf` en la raíz es el que deja la última corrida.

## Comportamiento de la pantalla frente a la automatización (verificado por CDP, 2026-09-09)

La pantalla es GeneXus y hostil a la automatización. Hallazgos firmes:

- **`div.gx-mask`**: mientras GeneXus procesa un postback tapa la pantalla entera con este div (opacidad 0.5, 1920×945). Los `click()` rebotan contra él (timeout) y los `fill()` escriben pero el redibujado los pisa. Hay que esperar a que desaparezca antes de cada interacción, y esperar un reposo extra porque los postbacks vienen encadenados con huecos sin máscara entre medio. Si la máscara queda pegada, la pantalla está trabada: **se libera con F5**. No dejar varias pestañas de la Carga Rápida abiertas: se pisan la sesión y dejan la máscara puesta.
- **El ORDEN de carga importa y NO es el visual**: hay que cargar el **prestador PRIMERO** (prompt por CUIT), luego el tipo, luego el resto de la cabecera. Invertirlo (tipo antes que prestador) rompe el tipo y pisa el punto de venta. Es el orden del código original y el que carga bien.
- **Los campos se llenan con `fill()` SIN `blur`**: cada `blur` dispara un postback de GeneXus, y esos postbacks intermedios barajan los valores entre campos (la fecha se colaba truncada en `#vCOMPROBANTEPREFIJO` y borraban el tipo). Sin blur, GeneXus recibe todo junto al agregar la línea (`#IMAGE3`). GeneXus solo aplica su máscara (padea ceros: `5` → `0005`) sin necesidad de blur. **No usar `completar` con blur.**
- **El tipo de comprobante se fija con el mecanismo del original** (`click` + `select_option(value=...)` + `dispatchEvent(change)` + `click` + `Enter`), elegido por **VALUE** y no por posición (el original tomaba el índice 2, que en algunos prestadores es "Factura B" → cargaba el tipo equivocado en silencio). Funciona para la mayoría de los prestadores.
- **El combo de tipo parpadea en algunos prestadores**: GeneXus lo repuebla y la opción aparece/desaparece; el propio acto de seleccionar puede dispararlo. `elegir_tipo_comprobante` reintenta 3× con `select_option(timeout=4000)`. Para la mayoría alcanza; para algunos (ej. CONTI —responsable inscripto que ofrece "Factura B"—, y a veces SUAREZ) no se logra: **se avisa y se deja el tipo sin cargar**, y **no se clickea `#IMAGE3`** (sin tipo, "Agregar línea" dispara un diálogo de validación que bloquea la pantalla). El operador elige el tipo y agrega la línea a mano. Coincide con el "raramente lo tengo que cargar yo" del operador, pero ahora avisado en vez de silencioso/equivocado.
- **Navegación (aplicado)**: tras el login, `main()` va directo a la pantalla con `page.goto(url)` en vez de clickear el menú `Prestadores` → celda `Carga Rapida Comprobantes` (esos clicks a veces no avanzaban → el operador tenía que darlos a mano). El login se hace solo si aparece la caja de usuario (la sesión puede seguir abierta de una corrida anterior).
- **Ceros a la izquierda**: los campos numéricos esperan el ancho completo del `maxlength` (`0005`, no `5`); un valor corto hace que GeneXus descarte la cabecera. El `extract_information` hace `lstrip('0')`, así que hay que re-padear al cargar.
- **`maxlength=150` en la descripción**: recorta sin avisar. Las muestras traen 176–296 caracteres.

## Referencia de la pantalla SISalud (verificado por CDP contra la pantalla real)

Los `id` de los combos y sus `value` — usar siempre `select_option(value=...)`, nunca posición:

- `#vTIPOCOMPROBANTECODIGO`: `CUDBC` Nota de débito "X", `FACCC` Factura "C" - Prestador, `FACXC` Factura "X", `NCCC` Nota de Crédito "C", `NDBC` Nota de Débito "B", `NDCC` Nota de Débito "C", `RECCC` Recibo "C", `REIN` Reintegro - Propio.
- `#vEXENTOCENTROCOSTOCODIGO` (renglón exento del detalle, el que usa AutoFiller) y `#vCENTROCOSTOCODIGO` (renglón gravado): mismas 24 opciones — `3` 28 DE OCTUBRE, `4` FRIAS, `5` SALTA, `7` MENDOZA, `8` SAN JUAN, `9` ENTRE RIOS, `12` BARKER, `13` CORDOBA, `14` OLAVARRIA, `22` TANDIL, `33` NEUQUEN, `54` SAN LUIS, `55` BUENOS AIRES, `61` MINA AGUILAR, `62` SANTA CRUZ, `64` RIO NEGRO, `71` CATAMARCA, `73` JUJUY, `74` CHUBUT, `76` HOTEL OSAM BS.AS. Nº4037, `77` COLONIA 28 DE OCTUBRE, `78` HOTEL MAR DEL PLATA, `79` HOTEL BUENOS AIRES, `81` CENTRAL (valor por defecto). Hay que setearlo **antes** de clickear `#IMAGE3`.
- `#span_vENTIDADDOMICILIOPROVINCIA`: provincia del prestador según el maestro de entidades, que SISalud completa sola al elegir la entidad por CUIT. Se usa como control cruzado.

## Muestras (`samples/`)

63 PDFs reales, uno duplicado exacto (`sanfeliu 07.pdf` = `SANFELIU 07 - PADIN 1334.pdf`). El nombre del archivo es `<APELLIDO DEL AFILIADO> <mes> - <APELLIDO DEL PRESTADOR> <nro comprobante>`.

Centros de costo confirmados por el usuario para cuatro de ellas: SANFELIU→Olavarría, ANTOLA→Mina Aguilar, ALCIBAR→Tandil, HIDALGO→Barker. Fueron elegidas a propósito como casos especiales, no son una muestra representativa.

## Problemas detectados en la evaluación (2026-09)

- **Credenciales hardcodeadas** en `user_var.set(...)` / `pass_var.set(...)`. Están en el `.exe` distribuido y en git. Hay que sacarlas. **Pendiente.**
- **Selectores frágiles**: el iframe del prompt depende del título literal con `?34`. **Pendiente.** (El tipo de comprobante ya no se elige por posición: se resolvió con `TIPOS_COMPROBANTE`.)
- **Errores invisibles**: los `except` hacen `print`, pero el `.exe` se compila `--noconsole`; el usuario solo ve "Factura inválida". Si falta `Hasta:` o `Importe Total:` en el PDF revienta con `UnboundLocalError`. **Pendiente** (los avisos del centro de costos sí se ven, van por `messagebox`).
- **Dependencias del entorno**: Chrome en `C:\Program Files\Google\Chrome\Application\chrome.exe`, perfil `C:\ChromeProfile`, `insoft.ico` en el cwd (`app.iconbitmap` falla desde otra carpeta), browsers de Playwright no empaquetados (ver comentario línea 1). **Pendiente.**
- **Higiene del repo**: `requirements.txt` está en UTF-16 y lista Django/DRF/JWT (pegado de otro proyecto); faltan las deps reales (`playwright`, `pdfplumber`, `pikepdf`). `chromedriver.exe` (18 MB) no se usa. `tempCodeRunnerFile.py` es una copia vieja de `main()`. `venv/` está commiteado. **Pendiente.**

## Restricciones confirmadas

- **SISalud no ofrece ningún camino alternativo a la pantalla** (sin API GeneXus, sin inserción directa en SQL). La carga tiene que seguir siendo por automatización del navegador.
- Lo usan **2 usuarios**. Otros **2 no lo usan porque la mayoría de los comprobantes que reciben son fotos, no PDF**.

## Requerimientos

### 1. Centro de Costos — IMPLEMENTADO (2026-09-09)

**Regla de negocio real**: el centro de costos es la **seccional del afiliado**, dato que no está en la factura (está en el padrón de SISalud). Se **aproxima por el domicilio del prestador**, que en general atiende a los afiliados de su zona. Decisión del usuario: donde el domicilio coincida se carga solo; donde no, lo carga el operador a mano; **si la aproximación no es la correcta, el operador la cambia en la pantalla**.

Orden de resolución en `centro_costo_desde_domicilio()`:

1. **Localidad** contra `CENTROS_COSTO_POR_LOCALIDAD` (las opciones del combo que son seccionales o yacimientos). Se busca **solo en el segmento de localidad** del domicilio, nunca en la calle: `28 de Octubre 123 - Río Gallegos, Santa Cruz` tiene que dar SANTA CRUZ, no el centro de costos 28 DE OCTUBRE. Los hoteles y CENTRAL quedan fuera del matching a propósito.
2. **Provincia** contra `CENTROS_COSTO_POR_PROVINCIA` (13 de las 24 provincias tienen opción homónima). Se carga aunque esa provincia tenga seccionales propias (Buenos Aires, Jujuy, Santa Cruz): decisión explícita del usuario, lo corrigen a mano si hace falta.
3. Sin coincidencia → no se toca el combo, queda CENTRAL y lo elige el operador.

Detalles resueltos:
- El domicilio viene como `<calle> - <localidad>, <provincia>`, y ARCA lo **parte en dos líneas** cuando es largo (`... - Perico Del` / `Carmen, Jujuy Ingresos Brutos: ...`). `extraer_domicilio_comercial()` lo reconstruye y corta los campos que ARCA pega al lado (`CUIT`, `Ingresos Brutos`, `Condición`, `IVA`, `Fecha`, `Período`).
- Alias de localidad confirmados por el usuario: `VILLA CACIQUE → BARKER` (12), localidades pegadas del partido de Benito Juárez; `FORTABAT → OLAVARRIA` (14), Villa Fortabat es del partido de Olavarría (zona Loma Negra).
- **Provincia recortada por ARCA** (`La Punta (Capital), San Lui`): `resolver_provincia()` acepta el recorte si es prefijo de una única provincia y tiene al menos 4 caracteres. `SAN LUI` → SAN LUIS; `SANTA` y `SAN` quedan rechazados por ambiguos.
- CABA y sus variantes (`Capital Federal`, `Ciudad de Buenos Aires`, `C.A.B.A.`) se reconocen aparte para que **no** caigan en BUENOS AIRES, que es la provincia y tiene otro centro de costos. CABA no tiene centro de costos → queda manual.
- **Control cruzado**: si la provincia del PDF no coincide con `#span_vENTIDADDOMICILIOPROVINCIA` (la que SISalud tiene para el prestador), no se carga el centro de costos y se avisa por `messagebox`.

**Resultado sobre las 62 muestras útiles**: 58 ciertos (93,5 %), 4 dudosos (6,5 %), 0 manuales. "Cierto" significa que dentro de la regla no hay alternativa posible (la localidad es una opción del combo, o la provincia no tiene seccionales que le compitan), no precisión medida: el único centro de costos real conocido son los cuatro que confirmó el usuario, y de esos la regla acierta tres.

Los 4 dudosos cargan la provincia y el operador corrige si hace falta: ANTOLA y VALDIVIEZO (Jujuy, podría ser Mina Aguilar), BENITEZ (San Martín) y SANCHEZ (Los Polvorines), ambos Buenos Aires.

Sin resolver: el caso de fondo es que el prestador esté lejos de la seccional del afiliado. Ejemplo real confirmado: ANTOLA, prestadora en Perico (Jujuy) y centro de costos Mina Aguilar, a ~250 km. La única forma de resolverlo sería leer la seccional del afiliado en el padrón de SISalud a partir del DNI o número de afiliado que aparece en la descripción del detalle.

### 2. Soportar comprobantes en foto (JPG/PNG/HEIC)
Enfoque recomendado: **híbrido QR + OCR**.
- Toda factura electrónica de ARCA trae un **QR** (`https://www.afip.gob.ar/fe/qr/?p=<base64 JSON>`) con `cuit`, `tipoCmp`, `ptoVta`, `nroCmp`, `fecha`, `importe`, `codAut` (CAE). Leerlo (p. ej. `pyzbar`/`zxing-cpp` + OpenCV) da esos campos exactos, gratis, y sirve también para PDF.
- Lo que el QR no trae y requiere OCR: período facturado (`Hasta:`), descripción del detalle y domicilio comercial (localidad y provincia).
- Opciones de OCR y costo por 100 comprobantes (precios de lista sep-2026): Tesseract local USD 0 (precisión baja en fotos de celular); Google Cloud Vision / AWS Textract Detect Text ≈ USD 0,15 (Vision: 1.000 págs/mes gratis); Textract Analyze Expense / Azure prebuilt invoice ≈ USD 1,00; API de Claude con visión (Haiku 4.5 ≈ USD 0,35–0,40, Sonnet 5 ≈ USD 0,75) devolviendo JSON estructurado directamente — la más robusta con fotos malas y layouts variables.
- **Decisión tomada (2026-09-09): los comprobantes pueden salir de la red de AOMAOSAM** hacia un proveedor externo. Queda habilitado usar un modelo con visión (API de Claude) o un OCR en la nube; no hace falta limitarse a Tesseract local.

### 3. Evaluación app web
Con SISalud sin API, una app web tendría que correr Playwright headless en el servidor por la VPN, manteniendo toda la fragilidad del scraping y perdiendo que el usuario vea la pantalla antes de confirmar. Arquitectura que sí tiene sentido si se adopta OCR externo: **extracción en un servidor** (recibe PDF/foto, lee QR, llama al OCR, devuelve JSON; la clave de API no se distribuye en el `.exe`) y **carga en SISalud desde la PC del usuario** (cliente liviano). Si se queda en OCR local, conviene primero robustecer el escritorio (credenciales, errores visibles, varios comprobantes por sesión, selectores estables) y recién después evaluar la web.

## Cómo probar

- **Contra los PDFs**: cargar el módulo sin la GUI (`src.split("# Crear ventana principal")[0]` + `exec`) y llamar a `extract_information()` o `centro_costo_desde_domicilio(text)`. La GUI arranca sola al importar el archivo.
- **Contra la pantalla real**: si Chrome está levantado con `--remote-debugging-port=9222`, `playwright.chromium.connect_over_cdp("http://127.0.0.1:9222")` engancha la sesión abierta del usuario. Sirve para leer el DOM sin tocar nada; para probar un combo, guardar el valor original y restaurarlo después.

## Directivas de trabajo del usuario

- Si una configuración puede hacerse de forma definitiva, hacerla así desde el inicio; solo dejar algo temporal si hay una razón técnica explícita.
- Cuando una respuesta requiere una decisión del usuario, dar solo la información necesaria para decidir y esperar; no avanzar con desarrollos que dependan de la decisión pendiente.
