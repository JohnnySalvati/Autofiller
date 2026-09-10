# Delegaciones de SISalud y su relación con los Centros de Costo

Leído del padrón real por CDP el 2026-09-10, desde
`servlet/promptdelegacion` (pantalla "Seleccionar Delegación", 3 páginas).

## El hallazgo

**Los 24 centros de costo del combo `#vEXENTOCENTROCOSTOCODIGO` existen como
delegación, con el mismo nombre exacto. 24 de 24, sin excepciones.**

Los códigos NO coinciden (delegación 4 = centro de costos 3), así que el puente
es el nombre, no el número.

El padrón tiene **49 delegaciones** contra 24 centros de costo: 24 mapean uno a
uno y **25 no tienen centro de costos propio**.

Esto confirma lo que se venía suponiendo: el centro de costos ES la seccional
del afiliado, y el padrón de SISalud la conoce. Hoy se aproxima por el domicilio
del prestador porque no la teníamos a mano; ahora sabemos dónde está.

## Las 24 que mapean

| deleg. | nombre | centro de costos |
|---|---|---|
| 4 | 28 DE OCTUBRE | 3 |
| 5 | FRIAS | 4 |
| 6 | SALTA | 5 |
| 8 | MENDOZA | 7 |
| 9 | SAN JUAN | 8 |
| 10 | ENTRE RIOS | 9 |
| 13 | BARKER | 12 |
| 14 | CORDOBA | 13 |
| 15 | OLAVARRIA | 14 |
| 16 | TANDIL | 22 |
| 25 | NEUQUEN | 33 |
| 30 | SAN LUIS | 54 |
| 31 | BUENOS AIRES | 55 |
| 32 | MINA AGUILAR | 61 |
| 33 | SANTA CRUZ | 62 |
| 34 | RIO NEGRO | 64 |
| 37 | CATAMARCA | 71 |
| 38 | JUJUY | 73 |
| 39 | CHUBUT | 74 |
| 41 | HOTEL OSAM BS.AS. Nº4037 | 76 |
| 42 | COLONIA 28 DE OCTUBRE | 77 |
| 43 | HOTEL MAR DEL PLATA | 78 |
| 44 | HOTEL BUENOS AIRES | 79 |
| 47 | CENTRAL | 81 |

## Las 25 sin centro de costos propio — DECISIÓN PENDIENTE DEL USUARIO

Si el afiliado pertenece a una de éstas, hay que decidir a qué centro de costos
va. El domicilio de la delegación sugiere el agrupamiento, pero **hay que
confirmarlo con el usuario**, igual que se confirmaron los alias
`VILLA CACIQUE → BARKER` y `FORTABAT → OLAVARRIA`.

| deleg. | nombre | provincia | domicilio | agrupamiento que sugiere el domicilio |
|---|---|---|---|---|
| 7 | LOMA NEGRA | BUENOS AIRES | LAMADRID 2534 | **mismo domicilio que OLAVARRIA** |
| 15 | (OLAVARRIA) | BUENOS AIRES | LAMADRID 2534 | — |
| 12 | CALERA AVELLANEDA | BUENOS AIRES | 9 DE JULIO 3272 | zona Olavarría |
| 26 | CAL Y PIEDRA OLAVARRIA | BUENOS AIRES | 9 DE JULIO 3272 | zona Olavarría (lo dice el nombre) |
| 11 | SIERRAS BAYAS | BUENOS AIRES | SAN MARTIN 2252 | partido de Olavarría |
| 27 | CERRO SOTUYO | BUENOS AIRES | — | partido de Olavarría |
| 3 | PIPINAS | BUENOS AIRES | — | ¿BUENOS AIRES? |
| 17 | GENERACION PUEYRREDON | BUENOS AIRES | DIAG. MAR DEL PLATA | ¿BUENOS AIRES / Mar del Plata? |
| 1 | OSAM | CABA | Rosario 434 | **puerta de al lado de CENTRAL** (Rosario 436) |
| 2 | GENERAL SAN MARTIN | MENDOZA | ECHEVERRIA 463 | ¿MENDOZA? |
| 45 | ASOC. MUTUAL HERCULES | MENDOZA | SAN MARTIN 1027 | ¿MENDOZA? |
| 19 | QUILPO SUD | CORDOBA | SIMON BOLIVAR 2447 | ¿CORDOBA? |
| 20 | MALAGUENO | CORDOBA | GRAL. BUSTOS 224 | ¿CORDOBA? |
| 50 | CORDOBA ART.23 | CORDOBA | ITUZAINGO 458/60 | ¿CORDOBA? |
| 21 | JACHAL | SAN JUAN | GRAL. PAZ 872 | ¿SAN JUAN? |
| 22 | ALBARDON | SAN JUAN | LAS LAJAS 3781 | ¿SAN JUAN? |
| 24 | CARBOMETAL | SAN JUAN | B* PATAGONIA | ¿SAN JUAN? |
| 48 | SAN JUAN (EMERGENCIA) | SAN JUAN | — | ¿SAN JUAN? |
| 35 | SIERRA GRANDE | RIO NEGRO | ALMIRANTE BROWN 276 | ¿RIO NEGRO? |
| 40 | GASTRE | RIO NEGRO | ALMAFUERTE 420 | ¿RIO NEGRO? (Gastre es Chubut) |
| 29 | PIEDRAS BLANCAS | ENTRE RIOS | — | ¿ENTRE RIOS? |
| 49 | CAJA FDO.SEGURO RETIRO | JUJUY | — | ¿MINA AGUILAR? (misma localidad 30163) |
| 18 | CORRIENTES | CORRIENTES | — | sin centro de costos |
| 23 | POSADAS | MISIONES | MISIONES 3849 | sin centro de costos |
| 28 | ROSARIO | SANTA FE | TTE. AGNETA 2125 | sin centro de costos |
| 36 | LA RIOJA | LA RIOJA | H. IRIGOYEN 79 | sin centro de costos |

Las últimas cuatro están en provincias que **no tienen centro de costos**
(Corrientes, Misiones, Santa Fe, La Rioja): ahí habría que dejarlo manual o
mandarlo a CENTRAL.

## La consulta, verificada contra el padrón real (2026-09-10)

**Pantalla**: `servlet/wwafiliado`, título "Afiliados" (menú *Afiliados → Administrar
Afiliados*). NO es `promptafiliado` (ése filtra por delegación, no la devuelve) ni
`wwficha` (ésa tiene trámites, no el padrón: buscando ANTOLA por documento y por
apellido devuelve cero filas).

**La trampa**: los campos de filtro existen en el DOM desde el arranque pero el
servidor los ignora hasta que se elige `Orden Por`. Llenar `vFILAFILIADOENTIDAD...`
sin eso devuelve cero filas siempre, con o sin sesión, y sin ningún mensaje de error.
Costó un rato entenderlo, porque se ve igual que "no existe ese afiliado".

Secuencia que funciona:

1. `#vAFILIADOORDENPOR` = `3` (Número de Documento) y **esperar el postback**.
   Recién ahí `#vFILAFILIADOENTIDADDOCUMENTONUMERO` pasa a ser visible y efectivo.
   (Con `1` = Número de afiliado se habilitan `#vFILAFILIADONUMERO` +
   `#vFILAFILIADOORDEN`, que es el otro dato que traemos del comprobante.)
2. Cargar el DNI en `#vFILAFILIADOENTIDADDOCUMENTONUMERO` (maxlength 8).
3. Click en `input[name=SEARCHBUTTON]` y esperar que se vaya `div.gx-mask`.
4. Leer la columna **Delegación** de `#GridContainerTbl`.

**Resultado sobre los cuatro casos de centro de costos confirmado: 4 de 4.**

| afiliado | delegación que devuelve el padrón | centro de costos confirmado |
|---|---|---|
| ANTOLA | MINA AGUILAR | MINA AGUILAR (61) |
| SANFELIU | OLAVARRIA | OLAVARRIA (14) |
| ALCIBAR | TANDIL | TANDIL (22) |
| HIDALGO | BARKER | BARKER (12) |

> Van solo los apellidos, que son los que ya usa `CLAUDE.md` para nombrar las muestras.
> Los DNI y números de afiliado con los que se hicieron estas cuatro consultas están en
> los comprobantes de `samples/`, que queda fuera del repo justamente porque son datos
> de afiliados reales. Para rehacer la prueba, el DNI sale de correr `afiliado.py` sobre
> la muestra que corresponda.

ANTOLA es el que cierra el argumento: la prestadora está en Perico (Jujuy), así que
la aproximación por domicilio da JUJUY (73). El padrón dice MINA AGUILAR (61), a
250 km, y es el correcto.

Donde el comprobante traía el número de afiliado, coincidió con el del padrón (ANTOLA
y SANFELIU).

## Lo que queda por decidir

1. Las 25 delegaciones sin centro de costos propio (tabla de arriba).
2. Qué hacer cuando el DNI no aparece en el padrón, o aparece con más de una ficha.
3. Dónde corre la consulta: es una pantalla más de SISalud, así que la haría el
   **agente** (que ya tiene Chrome y sesión), no el servidor de extracción.
