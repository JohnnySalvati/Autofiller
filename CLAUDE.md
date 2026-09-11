# AutoFiller — contexto del proyecto

Extrae los datos de un comprobante de ARCA (PDF o foto) y los carga en la pantalla **Prestadores → Carga Rápida Comprobantes** del sistema web SISalud (GeneXus, `http://vpn.aomaosam.org.ar:8081/sisaludevo/servlet/cargarapidacomprobantescompra`). Cliente: obra social AOMAOSAM. Desarrollado por InSoft (Johnny Salvati). Comunicación en español rioplatense.

Desde 2026-09-09 es una **app web híbrida** (Python 3.11, FastAPI + Playwright): un servidor que lee los comprobantes y un agente local que los carga en el SISalud del operador. La versión de escritorio (Tkinter, `dist/AutoFiller.exe`) sigue existiendo y es la que usan hoy los operadores.

## Arquitectura (app web, implementada 2026-09-09)

Decisión del usuario: **híbrida**. SISalud es accesible desde internet (el `vpn.` de
la URL es un nombre engañoso, no hay VPN de por medio — conviene corregirlo algún
día), así que la carga headless desde el servidor era técnicamente viable, pero se
descartó a propósito: perdería que el operador vea la pantalla real antes de
confirmar, que es el control de calidad de todo el proceso. Solo se centraliza la
**lectura** del comprobante.

```
   Navegador del operador
   ┌─────────────────────────────┐
   │  Web de AutoFiller          │──── PDF / foto ────►  servidor/
   │  (revisar y corregir)       │◄─── campos JSON ────
   └──────────┬──────────────────┘
              │ factura + credenciales de SISalud
              ▼  (127.0.0.1, nunca salen de la PC)
        agente/  ──► Chrome del operador ──► SISalud
```

### `servidor/` — extracción (FastAPI, sin estado)

- `main.py`: `POST /api/extraer` (multipart, varios archivos, devuelve un `Resultado`
  por archivo en el mismo orden), `GET /api/opciones` (combos + si hay lectura de
  fotos), y sirve `web/`. No guarda nada: cada archivo se procesa en memoria y se
  descarta. Ya no deja `decrypted.pdf` en el cwd.
- `extraccion/__init__.py`: `extraer(nombre, contenido)` elige el camino según el
  archivo — PDF con texto → regex; PDF escaneado → se rasteriza la página 1 y va por
  el camino de foto; foto → QR + visión. **Nunca lanza**: todo error vuelve en
  `Resultado.error`.
- `extraccion/texto.py`: los regex de ARCA. Mismo resultado que el escritorio en
  todos los campos **menos la descripción** (ver abajo), y ningún campo faltante
  corta la extracción: cada uno que no aparece vuelve vacío con su aviso, en vez
  del `UnboundLocalError` del escritorio.
  `extraer_descripcion()` (2026-09-10) reemplaza al "todo lo que hay entre los dos
  `Subtotal`" del escritorio, que era posicional y solo funcionaba con el orden de
  lectura de pdfplumber. Ahora reconoce **dos layouts** desde el mismo ancla
  (`Producto / Servicio`): el de filas (pdfplumber, con las columnas numéricas
  pegadas al final de la primera línea de la descripción) y el de bloques (un OCR
  como Cloud Vision, que agrupa primero todo el texto y después los números). En
  el de filas saca además las seis columnas del renglón, que antes se cargaban en
  SISalud metidas en el medio del texto.
- `extraccion/qr.py`: QR de ARCA (`zxing-cpp`, sin DLL externa como necesitaría
  `pyzbar`). El payload trae CUIT, tipo, punto de venta, número, fecha, importe y CAE
  firmados por ARCA: **pisa** lo que devuelva cualquiera de los dos motores.
- `extraccion/ocr.py` (2026-09-10): OCR de Google Cloud Vision por REST con una API
  key y el `urllib` de la stdlib — sin SDK, sin dependencia nueva. Es el **motor por
  defecto** del camino de imagen: su texto va a los mismos regex de `texto.py` que un
  PDF. No cachea (el harness sí, para no quemar cuota iterando sobre las 63 muestras;
  acá cada comprobante se ve una vez). `LADO_MAXIMO` es 3000 y **no** es el 1568 de
  `vision.py`: ahí achicar ahorra tokens, acá se paga por página y achicar solo
  perjudica al OCR; 3000 deja intacto lo que se midió (200 dpi = 2339 px en A4).
- `extraccion/vision.py`: Claude con visión, salida estructurada por `json_schema`.
  Es el **respaldo**: entra cuando el OCR falla, devuelve algo que no parece una
  factura de ARCA, o no saca ninguno de `CAMPOS_DE_LECTURA` (período, descripción,
  domicilio — lo único que el QR no trae). Modelo por defecto `claude-opus-5`,
  esfuerzo `medium`, ambos configurables por entorno.
  `AUTOFILLER_MOTOR_LECTURA` (`auto` | `ocr` | `vision`) fuerza un motor para poder
  medirlos. Con cualquiera de las dos claves la app lee fotos; sin ninguna anda igual
  con PDFs y la web lo avisa con un chip que dice qué motor quedó activo. El `origen`
  del resultado distingue los dos caminos (`foto-ocr`, `foto-vision`, y los
  `pdf-imagen-*`), así que en la web se ve con qué se leyó cada comprobante.
- `extraccion/centro_costo.py`: la lógica del requerimiento 1, portada tal cual. El
  único cambio es la firma: toma el domicilio ya extraído en vez del texto entero,
  porque en el camino de foto el domicilio lo trae la visión, no un regex.
- `web/`: `index.html` + `estilos.css` + `app.js` + `logo.svg`, sin framework ni build.
  Dos vistas que se alternan por JS: **Entrar** y la aplicación. Si el agente ya tiene
  la sesión abierta (se recargó la página), se entra directo sin pedir la clave otra vez.

**Identidad visual**: la misma línea que el resto de las apps de InSoft (referencia:
el login de FactuMov, `factumov.insoft.net.ar`). Wordmark en negrita al lado del
isotipo; título "Entrar"; tarjeta blanca bordeada; botón verde `#15803d` a todo el
ancho; pie "Una app de ⬤ InSoft" con el toggle verde (hecho en CSS, `.marca-insoft`)
enlazado a `insoft.net.ar`.

El isotipo (`web/logo.svg`, también favicon) sigue la **familia de íconos de producto
de InSoft**, calcada de `factumov-icon.svg`: caja de 120, cuadrado redondeado `#0F172A`
con `rx 23`, la inicial de un solo trazo grueso de punta redonda (`stroke-width 13`,
alto real 26–94) con el gradiente vertical de la pastilla de InSoft (`#2EBD59` →
`#1B9E4B`, `userSpaceOnUse` sobre el alto de la letra, no sobre el bounding box), y un
**punto blanco donde el trazo termina** — la cita al círculo del interruptor de InSoft,
que es la idea de toda la marca. Sin ese punto el ícono queda fuera de la familia.
El de AutoFiller es una **"A" cuyo travesaño se estira más allá de la pata derecha** y
termina en el punto: el campo que se completa solo. El travesaño se pasa **lo justo para
que el punto quede tocando la pata** (`cx 84` es la tangencia): así lo verde del
interruptor queda adentro de la letra y lo único que asoma es el círculo blanco. La A
salió angosta para que entrara el punto sin pisar la pata, y el travesaño va bajo
(`y 72`) para que el contrapunzón triangular no se cierre.

**El SVG es el único dibujo de la marca.** Windows necesita un `.ico` (el ícono del
`.exe`, el de la bandeja del agente, el de la ventana de actividad): lo genera
`hacer_icono.py` desde `web/logo.svg`, rasterizando con el mismo Chrome que usa el
agente (headless, fondo transparente) y armando con Pillow los nueve tamaños del
`autofiller.ico` de la raíz (16 a 256; Windows escala feo si le falta el que necesita,
y de 256 a 16 el trazo queda un borrón). Se corre a mano cuando cambia el logo y el
`.ico` se versiona. Hasta 2026-09-11 el agente mostraba en la bandeja el `insoft.ico`
del escritorio —el ícono de la empresa, no el del producto—, que ya no está en el repo.

### `agente/` — carga en SISalud (FastAPI en `127.0.0.1:8765`)

Es lo único que toca SISalud, y corre en la PC del operador.

- `sisalud.py`: **portado tal cual del escritorio**, con los tres hallazgos de la
  sección siguiente intactos (máscara, orden de carga, `fill` sin blur).
- `operador.py`: el cartel inyectado en la pantalla, `ControlLote` (acá `Control`) y
  `esperar_resolucion`. Sin cambios de lógica; el botón dice "Detener la cola".
  El cartel ocupa la **mitad izquierda** (`width:50%`, 2026-09-11): SISalud saca sus
  propios mensajes arriba a la derecha —entre ellos el de "comprobante ya cargado"— y
  a todo el ancho los tapaba. Esos mensajes son **pnotify** (`.ui-pnotify`, el de
  jQuery: `$.pnotify`), verificado contra la pantalla real el 2026-09-11: `fixed`,
  `top:18px`, `right:18px`, `z-index 9999`, 300 px de ancho. O sea que arrancan en
  `ancho - 318`, y el cartel termina en `ancho / 2`: no se pisan mientras la ventana
  pase de ~640 px. Medido a 1920×945 (lo que usa el operador): cartel 0–960, mensaje
  1602–1902. Como es `fixed`, tampoco lo mueve el `padding-top` que el cartel le pone
  al `body`.
  El cartel muestra además **una línea por aviso** con lo que el operador tiene que
  hacer ("Elegí el tipo de comprobante."), no el aviso entero: cada aviso lleva su
  `accion` corta además del texto completo (`avisar(avisos, texto, accion=)`, y
  `Avisos.acciones_cortas()` es lo que se inyecta), y el detalle —por qué falló, qué
  ofrecía el combo, qué parte de la descripción se recortó— queda en la web y en el
  log, que es donde hay lugar para leerlo. Un aviso sin `accion` cae en `resumir()`:
  primer renglón recortado a 90 caracteres.
- `navegador.py`: lanza/reutiliza Chrome con `--remote-debugging-port` y mantiene
  **una sola** pestaña para toda la cola. Busca `chrome.exe` en varias rutas y acepta
  `AUTOFILLER_CHROME` (antes estaba hardcodeada y fallaba en otras PCs).
- `main.py`: `POST /api/sesion` (credenciales, quedan en memoria), `POST /api/trabajo`
  (arranca uno y vuelve enseguida), `GET /api/trabajo` (fase: `cargando` → `esperando`
  → `terminado`), `POST /api/accion` (`saltar`/`detener`), `POST /api/listo`.
- `arrancar.py`: el entrypoint del `.exe`. Importa `app` de verdad (adentro de
  PyInstaller no hay `main.py` que importar por texto), corre uvicorn **en un hilo**
  y deja el hilo principal para la bandeja, que es donde tiene que vivir el bucle de
  mensajes del ícono. `--consola` fuerza el arranque viejo (uvicorn en el principal,
  todo a la vista): es la vía de soporte si la bandeja falla en una PC, sin recompilar.
  `--probar` es el autodiagnóstico. `AUTOFILLER_PUERTO` existe solo para levantar un
  agente de prueba sin pisar al que el operador tiene andando.

**El agente vive en la bandeja del sistema (2026-09-11)**, sin ventana: Windows 11 lo
manda a "Iconos ocultos" y el operador lo arrastra afuera si lo quiere a la vista. Antes
dejaba una consola abierta todo el día, que además de ruidosa era un botón de apagado
accidental — cerrarla con la X mataba el agente y la web pasaba a decir "no se detecta
el agente".

- `bandeja.py`: el ícono (`pystray` + `Pillow`, imagen tomada de `autofiller.ico`) y su
  menú: *Abrir AutoFiller* (el primer origen de `AUTOFILLER_ORIGENES`, así no hay una
  segunda configuración que pueda quedar vieja), *Ver la actividad*, el estado como
  renglón deshabilitado, y *Salir del agente*, que **pregunta** si hay un comprobante
  esperando resolución en vez de cortar el lote. pystray no se entera solo de que el
  texto de un item cambió: hay un hilo que cada 2 s reasigna el tooltip y llama
  `update_menu()`. También expone `avisar()`/`preguntar()` (MessageBoxW por ctypes):
  sin consola, un cartel de Windows es la única forma de que un fallo de arranque se
  vea.
- `registro.py`: **esto es lo que reemplaza a la consola, y no es una pérdida**. Manda
  `stdout` y `stderr` (los prints de `sisalud.py`, los renglones de uvicorn, cualquier
  traceback) a `%LOCALAPPDATA%\AutoFiller\agente.log`, line-buffered y con rotación a
  mano a los 2 MB. Hay que llamarlo **antes** de importar cualquier cosa que imprima:
  uvicorn se queda con el stream que encuentra al arrancar. Corriendo desde el código
  escribe en los dos lados. Queda mejor que la ventana negra, que se perdía entera al
  cerrarla: ahora un error de ayer se puede leer hoy.
- `visor.py`: la ventana de actividad (Tkinter), que muestra el log en vivo con
  *Copiar todo* y *Abrir la carpeta*. Corre en su propio hilo con su propio mainloop
  —válido mientras todos los widgets se creen y se toquen ahí, que es el caso: el
  resto del agente le habla por `Event`— porque el hilo principal ya es de la bandeja.
  Una sola instancia: el segundo pedido la trae al frente. Cerrarla no apaga el agente.
- `recursos.py`: resuelve el `autofiller.ico` adentro del `.exe` (`sys._MEIPASS`) o al
  lado del fuente. El ícono se empaqueta con `--add-data`, no solo como icono del
  `.exe`: es la imagen que se dibuja en la bandeja.

**`--noconsole` en el empaquetado no es la vuelta al error del escritorio**: ahí el
`--noconsole` *escondía* los errores (el operador solo veía "Factura inválida"); acá
todo queda en el log y el menú lo abre. Lo que sí cambia es el autodiagnóstico, porque
un proceso sin consola no recibe la del `cmd` que lo llamó: `--probar` se reengancha con
`AttachConsole(-1)` para que su salida se vea, y si no lo logra muestra el resultado en
un cartel. De ahí `--sin-carteles`, que `empaquetar.bat` pasa siempre: un cartel que
nadie va a clickear deja el empaquetado colgado para siempre (pasó mientras se probaba
esto). Medido el 2026-09-11 y anotado porque es al revés de lo que se suele creer: desde
un `.bat`, **cmd sí espera a un programa sin consola y sí propaga su `errorlevel`**, así
que la llamada directa sigue sirviendo y no hace falta `start /b /wait`.

**CORS y Private Network Access**: el agente escucha en `127.0.0.1`, así que cualquier
página abierta en el navegador podría hablarle. Por eso `AUTOFILLER_ORIGENES` es una
lista blanca explícita y nunca un comodín. Y Chrome exige el permiso extra de PNA para
que una página servida desde otro origen llame a `127.0.0.1`: se resuelve con
`allow_private_network=True` de `CORSMiddleware` (Starlette ≥ 1.x lo trae; sin eso el
preflight devuelve 400 "Disallowed CORS private-network" y el navegador no explica por
qué).

**Playwright no necesita `playwright install`**: se engancha por CDP al Chrome del
operador, así que usa el driver del paquete pip y ninguno de los navegadores que
descarga. Eso elimina el `--add-data` de los browsers que hacía falta en el `.exe`.

### Lo que la web agrega sobre el escritorio

- Los datos se **revisan y corrigen antes** de tocar SISalud. La descripción muestra
  un contador contra el `maxlength=150` de la pantalla, así que el recorte se ve
  antes de cargar y no después.
- Soporta fotos (requerimiento 2).
- Las credenciales salieron del código: se piden en la web y van solo al agente.
- Los errores son visibles: el escritorio compilaba `--noconsole` y el operador solo
  veía "Factura inválida". El agente también se compila sin consola, pero todo lo que
  pasa queda en `%LOCALAPPDATA%\AutoFiller\agente.log` y se lee desde el menú del
  ícono (*Ver la actividad*).
- No hay "mover a `cargados/`": en la web no hay una carpeta que mover. El equivalente
  es el estado por comprobante en la cola y el resumen al terminar.

### `AutoFiller.py` (escritorio) — se mantiene

Sigue funcionando y es lo que usan hoy los operadores. Queda hasta que la web esté
desplegada. Desde 2026-09-10 **ya no trae credenciales hardcodeadas**: los campos
Usuario y Contraseña arrancan vacíos y `credenciales()` frena el procesamiento con
un `messagebox` si falta alguno, en vez de intentar el login en blanco. La clave
que estaba en el código sigue estando en el historial de git y dentro de
`dist/AutoFiller.exe`, pero **ya fue rotada** (2026-09-11, confirmado por el
usuario): la que quedó expuesta no sirve más. Consecuencia: el `.exe` ya
distribuido, el viejo, tiene la clave muerta adentro — hay que **recompilarlo**
para que el operador pueda entrar.


## Comportamiento de la pantalla frente a la automatización (verificado por CDP, 2026-09-09)

La pantalla es GeneXus y hostil a la automatización. Hallazgos firmes:

- **`div.gx-mask`**: mientras GeneXus procesa un postback tapa la pantalla entera con este div (opacidad 0.5, 1920×945). Los `click()` rebotan contra él (timeout) y los `fill()` escriben pero el redibujado los pisa. Hay que esperar a que desaparezca antes de cada interacción, y esperar un reposo extra porque los postbacks vienen encadenados con huecos sin máscara entre medio. Si la máscara queda pegada, la pantalla está trabada: **se libera con F5**. No dejar varias pestañas de la Carga Rápida abiertas: se pisan la sesión y dejan la máscara puesta.
- **El ORDEN de carga importa y NO es el visual**: hay que cargar el **prestador PRIMERO** (prompt por CUIT), luego el tipo, luego el resto de la cabecera. Invertirlo (tipo antes que prestador) rompe el tipo y pisa el punto de venta. Es el orden del código original y el que carga bien.
- **Los campos se llenan con `fill()` SIN `blur`**: cada `blur` dispara un postback de GeneXus, y esos postbacks intermedios barajan los valores entre campos (la fecha se colaba truncada en `#vCOMPROBANTEPREFIJO` y borraban el tipo). Sin blur, GeneXus recibe todo junto al agregar la línea (`#IMAGE3`). GeneXus solo aplica su máscara (padea ceros: `5` → `0005`) sin necesidad de blur. **No usar `completar` con blur.**
- **El tipo de comprobante se fija con el mecanismo del original** (`click` + `select_option(value=...)` + `dispatchEvent(change)` + `click` + `Enter`), elegido por **VALUE** y no por posición (el original tomaba el índice 2, que en algunos prestadores es "Factura B" → cargaba el tipo equivocado en silencio). Funciona para la mayoría de los prestadores.
- **El combo de tipo parpadea en algunos prestadores**: GeneXus lo repuebla y la opción aparece/desaparece; el propio acto de seleccionar puede dispararlo. `elegir_tipo_comprobante` reintenta 3× con `select_option(timeout=4000)`. Para la mayoría alcanza; para algunos (ej. CONTI —responsable inscripto que ofrece "Factura B"—, y a veces SUAREZ) no se logra: **se avisa y se deja el tipo sin cargar**, y **no se clickea `#IMAGE3`** (sin tipo, "Agregar línea" dispara un diálogo de validación que bloquea la pantalla). El operador elige el tipo y agrega la línea a mano. Coincide con el "raramente lo tengo que cargar yo" del operador, pero ahora avisado en vez de silencioso/equivocado.
- **Cancelar y Confirmar (verificado por CDP, 2026-09-09)**: los botones viven en `#TBL_BOTONES`, que está `display:none` con el formulario vacío y aparece con el comprobante cargado. **Cancelar** (`input[name=BUTTON2]`, evento GeneXus `E'RETURN'`) hace un POST y después **navega fuera de la pantalla** (vuelve atrás en el historial: en la prueba, a `about:blank`). **Confirmar** (`input[name=CONFIRMAR]`, evento `E'CONFIRMAR'`, atajo F12) **no se probó** porque graba un comprobante real y no hay entorno de prueba; se asume que al grabar navega o vuelve al formulario vacío (ambos casos los cubre `COMPROBANTE_EN_PANTALLA`). Confirmar en un lote real es la verificación pendiente.
- **Navegación (aplicado)**: tras el login, `main()` va directo a la pantalla con `page.goto(url)` en vez de clickear el menú `Prestadores` → celda `Carga Rapida Comprobantes` (esos clicks a veces no avanzaban → el operador tenía que darlos a mano). El login se hace solo si aparece la caja de usuario (la sesión puede seguir abierta de una corrida anterior).
- **Ceros a la izquierda**: los campos numéricos esperan el ancho completo del `maxlength` (`0005`, no `5`); un valor corto hace que GeneXus descarte la cabecera. El `extract_information` hace `lstrip('0')`, así que hay que re-padear al cargar.
- **`maxlength=150` en la descripción**: recorta sin avisar. Con el extractor viejo, 53 de las 63 muestras pasaban de 150 (176–339 caracteres) porque arrastraban las columnas numéricas del renglón. Desde que `extraer_descripcion()` las saca (2026-09-10), 45 caracteres menos en promedio y **27 de 63** siguen pasando de 150. Las que pasan las recorta el operador, viendo el contador en la web.

## Referencia de la pantalla SISalud (verificado por CDP contra la pantalla real)

Los `id` de los combos y sus `value` — usar siempre `select_option(value=...)`, nunca posición:

- `#vTIPOCOMPROBANTECODIGO`: `CUDBC` Nota de débito "X", `FACCC` Factura "C" - Prestador, `FACXC` Factura "X", `NCCC` Nota de Crédito "C", `NDBC` Nota de Débito "B", `NDCC` Nota de Débito "C", `RECCC` Recibo "C", `REIN` Reintegro - Propio.
- `#vEXENTOCENTROCOSTOCODIGO` (renglón exento del detalle, el que usa AutoFiller) y `#vCENTROCOSTOCODIGO` (renglón gravado): mismas 24 opciones — `3` 28 DE OCTUBRE, `4` FRIAS, `5` SALTA, `7` MENDOZA, `8` SAN JUAN, `9` ENTRE RIOS, `12` BARKER, `13` CORDOBA, `14` OLAVARRIA, `22` TANDIL, `33` NEUQUEN, `54` SAN LUIS, `55` BUENOS AIRES, `61` MINA AGUILAR, `62` SANTA CRUZ, `64` RIO NEGRO, `71` CATAMARCA, `73` JUJUY, `74` CHUBUT, `76` HOTEL OSAM BS.AS. Nº4037, `77` COLONIA 28 DE OCTUBRE, `78` HOTEL MAR DEL PLATA, `79` HOTEL BUENOS AIRES, `81` CENTRAL (valor por defecto). Hay que setearlo **antes** de clickear `#IMAGE3`.
- `#span_vENTIDADDOMICILIOPROVINCIA`: provincia del prestador según el maestro de entidades, que SISalud completa sola al elegir la entidad por CUIT. Se usa como control cruzado.

## Muestras (`samples/`)

63 PDFs reales, uno duplicado exacto (`sanfeliu 07.pdf` = `SANFELIU 07 - PADIN 1334.pdf`). El nombre del archivo es `<APELLIDO DEL AFILIADO> <mes> - <APELLIDO DEL PRESTADOR> <nro comprobante>`.

Centros de costo confirmados por el usuario para cuatro de ellas: SANFELIU→Olavarría, ANTOLA→Mina Aguilar, ALCIBAR→Tandil, HIDALGO→Barker. Fueron elegidas a propósito como casos especiales, no son una muestra representativa.

## Problemas de la evaluación 2026-09 — estado

Resueltos por la app web (los demás siguen presentes en `AutoFiller.py`, que se
mantiene):

- **Credenciales hardcodeadas**: la web las pide y las manda solo al agente local.
  El escritorio también las pide desde 2026-09-10, y la clave que había quedado
  expuesta **se rotó el 2026-09-11**: la del historial de git y la del `.exe` ya
  distribuido no sirven más. Queda recompilar ese `.exe`, que hoy lleva adentro
  la clave muerta.
- **Errores invisibles**: todo vuelve en `Resultado.error` / `avisos` y se ve en la
  cola. Un PDF sin `Hasta:` o sin `Importe Total:` ya no revienta.
- **Dependencias del entorno**: la ruta de Chrome se busca en varias ubicaciones y se
  puede fijar por entorno; no hay ícono ni cwd que importe; Playwright no
  necesita browsers descargados.
- **`requirements.txt`**: separado por componente y en UTF-8.

Pendientes:

- **Selectores frágiles**: el iframe del prompt sigue dependiendo del título literal
  con `?34` (`SELECTOR_PROMPT` en `agente/sisalud.py`).
- **Higiene del repo**: `chromedriver.exe` (18 MB, no se usa), `tempCodeRunnerFile.py`
  (copia vieja de `main()`) y `venv/` commiteado siguen ahí. No se borraron porque
  nadie lo pidió.
- **Duplicación deliberada**: `Factura` está definida en
  `servidor/extraccion/modelo.py` y espejada en `agente/main.py`, porque servidor y
  agente se instalan por separado y no comparten código. Si se agrega un campo en uno,
  hay que agregarlo en el otro.
- **Confirmar en un lote real** sigue sin probarse (graba comprobantes de verdad).

## Restricciones confirmadas

- **SISalud no ofrece ningún camino alternativo a la pantalla** (sin API GeneXus, sin
  inserción directa en SQL). La carga tiene que seguir siendo automatización del
  navegador.
- SISalud **sí es accesible desde internet**; el `vpn.` del host es un nombre
  heredado y confuso, conviene corregirlo.
- Lo usan **2 usuarios**. Otros **2 no lo usan porque la mayoría de los comprobantes
  que reciben son fotos, no PDF** — que es lo que destraba el requerimiento 2.

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

**RESUELTO (2026-09-10): el centro de costos sale del padrón de SISalud.**

Cadena completa, verificada de punta a punta:

```
detalle facturado -> DNI            (extraccion/afiliado.py, 63/63 de las muestras)
DNI -> Delegación                   (agente/padron.py, pantalla servlet/wwafiliado)
Delegación -> centro de costos      (24 homónimas + 21 deducidas, agente/padron.py)
```

Verificado contra los cuatro casos de centro de costos confirmado por el
usuario: **4 de 4**, incluido ANTOLA, donde la aproximación por domicilio erraba
por 250 km (daba JUJUY 73, el padrón dice MINA AGUILAR 61).

Corre en el **agente**, no en el servidor: es una pantalla más de SISalud y el
único que tiene Chrome con la sesión del operador es el agente. La consulta va
antes de abrir la Carga Rápida, y después se vuelve a ella.

**Las dos trampas de `wwafiliado`**, las dos alrededor de `#vAFILIADOORDENPOR`:

1. Los campos de filtro están en el DOM desde que carga la página, pero el
   servidor los ignora hasta que se elige `Orden Por`. Llenar el documento sin
   eso devuelve cero filas siempre, sin ningún mensaje, y se ve idéntico a "ese
   afiliado no existe".
2. **Ese combo es también el que destapa el filtro**: `#TDOCUMENTO` y las demás
   tablas de filtro están en `display:none`, y las muestra el `onchange` del
   combo, que **GeneXus conecta después del load** (medido el 2026-09-11: entre
   250 y 500 ms después de que vuelve el `goto`). Seleccionar antes de esa
   ventana deja el valor puesto y el filtro tapado para siempre —como el combo
   ya tiene el valor, ningún evento posterior lo arregla—, y el `fill` se come
   entero su timeout esperando un campo invisible. Es una carrera que se gana o
   se pierde según lo que tarde el servidor: con la página cacheada el `goto`
   vuelve en 0,2 s y se pierde **siempre**, así que "ayer andaba" era suerte.
   `_elegir_orden` (2026-09-11) no confía en la selección: verifica que el campo
   esté visible y reintenta pasando por `Seleccione..`, que es lo que hace que el
   segundo intento sea un cambio de verdad. Todo lo que toca esta pantalla usa
   timeout de 8 s: si algo no está, conviene degradar rápido al domicilio y no
   dejar al operador un minuto mirando la pantalla de afiliados.

**Degradado (decisión del usuario)**: si el padrón no contesta por lo que sea
—sin DNI, DNI ausente, la pantalla cambió, más de una delegación, delegación sin
centro de costos— se vuelve a la aproximación por domicilio, que ya funcionaba.
La consulta solo puede mejorar el resultado, nunca empeorarlo. Cada caso deja su
aviso. Las nueve ramas están probadas con un `page` falso.

Cuando el centro de costos viene del padrón **no se aplica el control cruzado
por provincia** de `cargar_factura`: ese control existía para atajar errores de
la aproximación, y justamente los casos que importan son aquellos en que el
prestador está en otra provincia que su afiliado.

**Las 21 deducidas están sin confirmar por el operador.** Criterio del usuario:
"todo lo que se pueda deducir es válido; si el operador lo corrige, se cambia".
Las cinco de la zona de Olavarría (LOMA NEGRA, CALERA AVELLANEDA, CAL Y PIEDRA,
SIERRAS BAYAS, CERRO SOTUYO) son las más apoyadas: comparten domicilio con
OLAVARRIA. Tabla completa con el fundamento de cada una en
`pruebas/delegaciones.md`.

CORRIENTES, POSADAS, ROSARIO y LA RIOJA quedan sin centro de costos a propósito:
sus provincias no tienen opción en el combo.

Sin resolver: el caso de fondo es que el prestador esté lejos de la seccional del afiliado. Ejemplo real confirmado: ANTOLA, prestadora en Perico (Jujuy) y centro de costos Mina Aguilar, a ~250 km. La única forma de resolverlo sería leer la seccional del afiliado en el padrón de SISalud a partir del DNI o número de afiliado que aparece en la descripción del detalle.

### 2. Comprobantes en foto (JPG/PNG/HEIC) — IMPLEMENTADO (2026-09-09)

Enfoque **híbrido QR + visión**, en `servidor/extraccion/`:

- El **QR** de ARCA (`https://www.afip.gob.ar/fe/qr/?p=<base64 JSON>`) trae `cuit`,
  `tipoCmp`, `ptoVta`, `nroCmp`, `fecha`, `importe` y `codAut` (CAE) exactos y gratis.
  Se lee con `zxing-cpp` (elegido sobre `pyzbar`, que en Windows necesita una DLL
  aparte). Lo que el QR dice **pisa** lo que devuelva el modelo.
- Lo que el QR no trae — período (`Hasta:`), descripción del detalle y domicilio
  comercial — lo saca **Claude con visión**, con salida estructurada por `json_schema`.
- **Decisión del cliente (2026-09-09)**: los comprobantes pueden salir de la red de
  AOMAOSAM hacia un proveedor externo, así que no hay que limitarse a OCR local.
- Modelo por defecto `claude-opus-5`. Si el costo por comprobante pesa, se baja con
  `AUTOFILLER_MODELO_VISION` (`claude-sonnet-5`, `claude-haiku-4-5`) o con
  `AUTOFILLER_ESFUERZO_VISION`, sin tocar código. Referencia de precios de lista
  sep-2026 por 100 comprobantes: Haiku 4.5 ≈ USD 0,35–0,40; Sonnet 5 ≈ USD 0,75.
- Un PDF **escaneado** (sin capa de texto) entra por este mismo camino: se rasteriza
  la primera página con `pypdfium2` y se trata como foto.

Sin probar todavía contra fotos reales de celular: no hay muestras. Las 63 de
`samples/` son todas PDF nativos.

**IMPLEMENTADO (2026-09-10): OCR clásico como motor por defecto, modelo de respaldo.**
Google Cloud Vision regala 1000 páginas por mes, que al volumen de AOMAOSAM
probablemente sea gratis para siempre, y devuelve el mismo texto siempre (el
modelo puede variar entre corridas). El harness es `pruebas/comparar_ocr.py`:
rasteriza las muestras, las pasa por Cloud Vision y compara campo por campo
contra el camino PDF, que es ground truth verificado.

Resultado sobre las 63 muestras:

| campo | exactas | similitud media |
|---|---|---|
| `fecha_vencimiento` (`Hasta:`) | **63/63** | 100 % |
| `centro_costo` | **63/63** | 100 % |
| `provincia` | **63/63** | 100 % |
| `domicilio` | 59/63 | 99,9 % |
| `descripcion` | 46/63 | 98,8 % |

Los 4 domicilios que difieren son cosméticos (`Piso:1` vs `Piso: 1`) y no mueven
el centro de costos, que da 63/63. Las 17 descripciones que difieren son ruido
de OCR: 11 de ellas por 3 caracteres o menos (un acento, un guion). **El OCR
sirve.** Lo que no mide esta prueba es la foto de celular: ángulo, sombra y
foco. Para eso hacen falta fotos reales, y conviene que al menos algunas sean de
comprobantes que también estén en PDF, para poder medirlas igual que a éstas.

Llegar a esos números obligó a reescribir la extracción de la descripción (ver
`extraer_descripcion` en `texto.py`): el primer intento usaba un ancla por
layout y se rompía en los tres. El OCR también dejó a la vista dos bugs del
camino PDF, ya arreglados: la descripción sucia y el `ORIGINAL` pegado al
domicilio en `CAMPOS_JUNTO_AL_DOMICILIO`.

### 3. App web — IMPLEMENTADO (2026-09-09)

Ver "Arquitectura" arriba. Se descartó la web pura (todo headless en el servidor)
aunque SISalud sea alcanzable desde internet: perdería el control visual del
operador antes de confirmar.

## Cómo probar

- **La extracción, contra las muestras**: `extraer(nombre, contenido)` de
  `servidor/extraccion` no depende de nada levantado ni de ninguna GUI.

  ```python
  import sys; sys.path.insert(0, 'servidor')
  from extraccion import extraer
  print(extraer('x.pdf', open('samples/x.pdf', 'rb').read()).factura)
  ```

- **Contra el escritorio (regresión)**: cargar `AutoFiller.py` sin la GUI
  (`src.split("# Crear ventana principal")[0]` + `exec`; la GUI arranca sola al
  importarlo) y comparar `extract_information()` con `extraer()` campo por campo
  sobre las 63 muestras. Al portar dio 63/63 idénticos. **Desde 2026-09-10 la
  `descripcion` difiere a propósito en las 63**: el escritorio se trae las
  columnas numéricas del renglón y la web no. Los demás campos siguen siendo
  63/63; si alguno de ésos vuelve a diferir, es una regresión.
- **El camino OCR contra el camino PDF**: `python pruebas\comparar_ocr.py`. Ver
  `pruebas/LEEME.md`. Necesita `GOOGLE_VISION_API_KEY` para las muestras que no
  estén cacheadas; con `--solo-cache` no gasta cuota.
- **El servidor**: `servidor\iniciar.bat` y `POST /api/extraer` con `curl -F`. Ojo con
  los nombres de archivo con espacios: `curl` los parte y el pedido nunca sale.
- **El agente, sin tocar SISalud**: `agente\iniciar.bat` y pegarle a `/api/salud`,
  `/api/sesion` y `/api/accion`. **No** llamar a `/api/trabajo`: eso abre Chrome y
  carga de verdad en la pantalla. Si el operador ya tiene su agente andando, el puerto
  8765 está tomado: levantar el de prueba con `AUTOFILLER_PUERTO=8799
  python arrancar.py` (así aparece su propio ícono en la bandeja y no se pisa nada).
  El `.exe` compilado se diagnostica con `AutoFillerAgente.exe --probar`, que verifica
  orígenes, Chrome, el driver de Playwright y que el ícono de la bandeja se pueda crear.
- **De punta a punta**: cargar la cola desde la web y resolver cada comprobante con
  Cancelar (`input[name=BUTTON2]`) o con los botones del cartel
  (`window.autofillerAccion('saltar'|'detener')`). **No confirmar nunca: graba de
  verdad.**
- **Contra la pantalla real**: si Chrome está levantado con
  `--remote-debugging-port=9222`, `playwright.chromium.connect_over_cdp("http://127.0.0.1:9222")`
  engancha la sesión abierta del usuario. Sirve para leer el DOM sin tocar nada; para
  probar un combo, guardar el valor original y restaurarlo después.

## Directivas de trabajo del usuario

- Si una configuración puede hacerse de forma definitiva, hacerla así desde el inicio; solo dejar algo temporal si hay una razón técnica explícita.
- Cuando una respuesta requiere una decisión del usuario, dar solo la información necesaria para decidir y esperar; no avanzar con desarrollos que dependan de la decisión pendiente.
