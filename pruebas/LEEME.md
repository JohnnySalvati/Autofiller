# Prueba: ¿alcanza con OCR clásico para las fotos?

## Qué se está decidiendo

Para los comprobantes en foto hay dos caminos:

| | Modelo de visión (implementado) | OCR + las regex que ya tenemos |
|---|---|---|
| Costo | USD 4–22 cada 1000, según modelo | **USD 0** hasta 1000 páginas/mes, después USD 1,50 cada 1000 |
| Salida | JSON con los campos ya separados | Texto plano, hay que parsearlo |
| Determinismo | Puede variar entre corridas | La misma imagen da siempre el mismo texto |
| Trabajo pendiente | Ninguno, anda | Verificar que las regex parseen ese texto |

Con 4 usuarios, el camino OCR probablemente sea gratis para siempre. Lo que
falta saber es si funciona.

## La pregunta concreta

El OCR sabe leer las letras; eso no está en duda. Lo que no sabemos es si
devuelve el texto **en un orden** que las regex de `servidor/extraccion/texto.py`
puedan parsear, porque esas regex se escribieron contra el orden de lectura de
`pdfplumber` sobre PDFs nativos. Un OCR puede leer una cabecera a dos columnas
en otro orden y romper todo sin equivocarse en una sola letra.

## Cómo se prueba sin tener fotos

Las 63 muestras son PDFs nativos, así que de cada una salen las dos cosas:

```
verdad    = texto del PDF          -> datos_desde_texto()   (63/63 verificado)
candidato = página rasterizada     -> Cloud Vision -> datos_desde_texto()
```

Misma factura, mismas regex, lo único que cambia es de dónde sale el texto.
Eso aísla el riesgo de orden de lectura del riesgo de foto de celular.

**Lo que esta prueba no dice**: que Cloud Vision lea bien un raster a 200 dpi no
garantiza que lea bien una foto sacada a mano con sombra y en ángulo. Si esta
prueba sale bien, el paso siguiente es repetirla con fotos reales. Si sale mal,
ya sabemos que el camino OCR no cierra sin tocar las regex, y cuánto.

## Sacar la API key de Cloud Vision

1. https://console.cloud.google.com → crear un proyecto (o usar uno existente).
2. **Habilitar la API PRIMERO**: *APIs y servicios → Biblioteca* → *Cloud Vision
   API* → **Habilitar**. El orden importa: la lista de restricciones del paso 3
   **solo muestra APIs ya habilitadas**, así que si creás la clave antes, Cloud
   Vision no aparece y no la vas a poder restringir. Link directo:
   `https://console.cloud.google.com/apis/library/vision.googleapis.com`
   Si pide asociar una cuenta de facturación, es normal: no cobra dentro de las
   1000 páginas por mes.
3. **APIs y servicios → Credenciales** → **Crear credenciales → Clave de API**.
4. En *Elige las restricciones de API*, filtrar por `vision` y marcar **Cloud
   Vision API**. Es una clave que viaja en la URL: sin restringir, cualquiera
   que la vea puede gastar tu cuota. Si no aparece, recargá con F5: la lista se
   cachea y tarda un minuto en reflejar la API recién habilitada.

```powershell
setx GOOGLE_VISION_API_KEY "AIza..."      # para las consolas futuras
$env:GOOGLE_VISION_API_KEY = "AIza..."    # para la consola actual
```

`setx` guarda la variable pero **no la aplica a la consola que ya está abierta**:
o abrís una nueva, o usás además la línea con `$env:`.

## Correr

```powershell
python pruebas\comparar_ocr.py --limite 5    # prueba corta primero
python pruebas\comparar_ocr.py               # las 63
python pruebas\comparar_ocr.py --solo-cache  # sin gastar llamadas nuevas
```

Antes de llamar avisa cuántas llamadas nuevas va a hacer, para no quemar el
free tier por accidente. El texto de cada imagen queda cacheado en
`pruebas/cache_ocr/`, así que iterar sobre las regex sale gratis: la segunda
corrida y las siguientes no vuelven a llamar a Google.

## Leer el resultado

La tabla que decide es **CAMPOS QUE DECIDEN**: `fecha_vencimiento` (el `Hasta:`),
`descripcion`, `domicilio` y los dos derivados del domicilio. Son los que el QR
de ARCA **no** trae. Los otros se miden igual pero son informativos: en
producción el QR los pisa aunque el OCR los lea mal.

- **Exactos**: cuántas muestras dan idéntico al camino PDF.
- **Similitud media**: para ver si una diferencia es un acento o un campo entero.

El detalle campo por campo de cada diferencia queda en `pruebas/resultado_ocr.md`.

## Nota sobre datos

`pruebas/cache_ocr/` y `pruebas/resultado_ocr.md` contienen texto de facturas
reales, con nombre y DNI de afiliados. Están en `.gitignore` por la misma razón
que `samples/`. No los subas.
