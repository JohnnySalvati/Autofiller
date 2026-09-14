"""Extraccion de los datos de un comprobante, sea PDF o foto.

Un solo punto de entrada, extraer(nombre, contenido), que elige el camino:

  PDF con texto  -> pdfplumber + los regex de ARCA (el camino del escritorio,
                    exacto y gratis).
  PDF escaneado  -> se renderiza la primera pagina y se trata como foto.
  Foto           -> QR de ARCA para los datos fiscales, y para el resto OCR de
                    Cloud Vision + esos mismos regex; si el OCR no sirve, un
                    modelo de vision de Claude.

Nada se guarda en disco: el archivo entra por HTTP, se procesa en memoria y se
descarta. El escritorio dejaba un 'decrypted.pdf' en el directorio actual.
"""

import io
import logging
import os
from datetime import date

from . import ocr
from .afiliado import identificacion
from .centro_costo import centro_costo_de_domicilio, provincia_de_domicilio
from .modelo import CAMPOS_OBLIGATORIOS, ETIQUETAS_CAMPO, Factura, Resultado
from .qr import campos_desde_qr, leer_qr
from .texto import AVISOS_POR_CAMPO, avisos_de_campos, datos_desde_texto, parece_factura_arca
from .vision import VisionNoDisponible, disponible as vision_disponible, leer_comprobante

registro = logging.getLogger(__name__)

# Que motor lee la imagen: "auto" (OCR y, si no sirve, el modelo), "ocr" o
# "vision". Existe para poder medir uno contra otro sobre comprobantes reales
# sin tocar codigo, que es lo que falta decidir del requerimiento 2.
MOTOR = os.environ.get("AUTOFILLER_MOTOR_LECTURA", "auto").lower()

# Lo unico que el QR no trae y tiene que salir de leer la imagen. Si el OCR no
# saco ninguno de los tres, no aporto nada y vale la pena pagarle al modelo.
CAMPOS_DE_LECTURA = ("fecha_vencimiento", "descripcion", "domicilio")

EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".webp", ".bmp", ".tif", ".tiff"}
EXTENSIONES_PDF = {".pdf"}
EXTENSIONES_ACEPTADAS = EXTENSIONES_IMAGEN | EXTENSIONES_PDF

# A que resolucion se rasteriza una pagina de PDF escaneado. 200 dpi alcanza
# para que el QR se lea y el texto quede legible sin inflar la imagen.
DPI_RASTERIZADO = 200


def _abrir_imagen(contenido):
    from PIL import Image

    try:  # los iPhone mandan HEIC; sin el plugin, Pillow no lo abre
        import pillow_heif

        pillow_heif.register_heif_opener()
    except ImportError:
        pass
    return Image.open(io.BytesIO(contenido))


def _texto_del_pdf(contenido):
    """Texto del PDF, desencriptandolo en memoria si hace falta."""
    import pdfplumber
    import pikepdf

    # Las facturas de ARCA vienen con encriptacion de solo-lectura, que pikepdf
    # saca sin contrasena. pdfplumber solo no las abre.
    buffer = io.BytesIO()
    with pikepdf.open(io.BytesIO(contenido)) as pdf:
        pdf.save(buffer)
    buffer.seek(0)
    with pdfplumber.open(buffer) as documento:
        return "".join(pagina.extract_text() or "" for pagina in documento.pages)


def _primera_pagina_como_imagen(contenido):
    """Rasteriza la primera pagina del PDF, para los que vienen escaneados."""
    import pypdfium2

    documento = pypdfium2.PdfDocument(io.BytesIO(contenido))
    try:
        return documento[0].render(scale=DPI_RASTERIZADO / 72).to_pil()
    finally:
        documento.close()


def _completar_domicilio(datos):
    """Deriva provincia y centro de costos del domicilio ya extraido."""
    domicilio = datos.get("domicilio")
    nombre, codigo = centro_costo_de_domicilio(domicilio)
    datos["provincia"] = provincia_de_domicilio(domicilio)
    datos["centro_costo"] = codigo
    datos["centro_costo_nombre"] = nombre
    return datos


def _datos_desde_ocr(imagen):
    """(datos, avisos) leyendo la imagen con OCR y los mismos regex que un PDF.

    Devuelve None cuando el OCR no sirvio y hay que caer al modelo de vision:
    porque fallo la llamada, porque el texto no parece una factura de ARCA, o
    porque no salio ninguno de los campos que solo puede aportar la lectura.
    Nunca lanza: un problema del OCR no puede dejar sin leer un comprobante que
    el modelo si podria leer.
    """
    try:
        texto_leido = ocr.texto(imagen)
    except Exception as e:
        registro.warning("El OCR falló, se lee con el modelo de visión: %s", e)
        return None

    if not parece_factura_arca(texto_leido):
        registro.info("El texto del OCR no parece una factura de ARCA, se lee con el modelo.")
        return None

    datos, avisos = datos_desde_texto(texto_leido)
    if not any(datos.get(campo) for campo in CAMPOS_DE_LECTURA):
        registro.info("El OCR no sacó ningún campo útil, se lee con el modelo.")
        return None

    return datos, avisos


def _datos_desde_vision(imagen):
    """(datos, avisos) leyendo la imagen con el modelo de vision de Claude."""
    from .modelo import TIPOS_COMPROBANTE

    avisos = []
    leido = leer_comprobante(imagen)

    codigo = (leido.get("codigo_arca") or "").lstrip("0")
    periodo_hasta = leido.get("periodo_hasta") or None
    if not periodo_hasta:
        avisos.append(
            "No se leyó el período facturado ('Hasta:'). Quedaron sin cargar el "
            "vencimiento y el devengamiento."
        )

    datos = {
        "cuit": leido.get("cuit_emisor") or None,
        "tipo_comprobante": TIPOS_COMPROBANTE.get(codigo),
        "punto_venta": (leido.get("punto_venta") or "").lstrip("0") or None,
        "nro_factura": (leido.get("nro_comprobante") or "").lstrip("0") or None,
        "fecha_recepcion": date.today().strftime("%d/%m/%Y"),
        "fecha_emision": leido.get("fecha_emision") or None,
        "fecha_vencimiento": periodo_hasta,
        "fecha_devengamiento": periodo_hasta,
        "cae": leido.get("cae") or None,
        "descripcion": " ".join((leido.get("descripcion") or "").split()),
        "importe": leido.get("importe_total") or None,
        "domicilio": leido.get("domicilio_comercial") or None,
    }
    return _completar_domicilio(datos), avisos


def _datos_desde_imagen(imagen):
    """(datos, avisos, motor) leyendo QR + OCR o modelo sobre una imagen.

    El OCR va primero porque es gratis hasta 1000 paginas por mes y devuelve
    siempre el mismo texto; el modelo es el respaldo. AUTOFILLER_MOTOR_LECTURA
    fuerza uno de los dos para poder medirlos.
    """
    avisos = []
    payload = leer_qr(imagen)
    if not payload:
        avisos.append(
            "No se pudo leer el QR de ARCA: los datos fiscales salen de la "
            "lectura de la imagen. Verificá CUIT, número, CAE e importe antes "
            "de confirmar."
        )

    leido = None
    if MOTOR in ("auto", "ocr") and ocr.disponible():
        leido = _datos_desde_ocr(imagen)
    motor = "ocr"

    if leido is None:
        if MOTOR == "ocr":
            # Forzado a OCR: no hay respaldo al que caer, y decirlo es mejor que
            # devolver en silencio un comprobante vacio.
            raise RuntimeError(
                "El OCR no pudo leer el comprobante y el motor está forzado a "
                "OCR (AUTOFILLER_MOTOR_LECTURA=ocr)."
            )
        if not ocr.disponible() and not vision_disponible():
            # Sin ninguno de los dos no hay lectura de imagenes, y el mensaje
            # tiene que nombrar las dos claves: con cualquiera de ellas anda.
            raise VisionNoDisponible(
                "El servidor no tiene configurada la lectura de fotos: falta "
                "GOOGLE_VISION_API_KEY (OCR) o ANTHROPIC_API_KEY (modelo de visión)."
            )
        leido = _datos_desde_vision(imagen)
        motor = "vision"

    datos, avisos_lectura = leido
    avisos.extend(avisos_lectura)

    # El QR es dato fiscal firmado por ARCA: pisa lo que haya leido cualquiera
    # de los dos motores.
    datos.update(campos_desde_qr(payload))
    return datos, avisos, motor


def _regex_no_entendieron(datos):
    """Si el PDF tiene texto pero los regex de ARCA no le sacaron lo que hace falta.

    Dos senales, y cualquiera alcanza: que falte alguno de los campos que
    SISalud necesita si o si, o que hayan quedado vacios a la vez el detalle y
    el importe. Una factura de ARCA de verdad no da ninguna de las dos; cuando
    se dan es porque el comprobante esta impreso de otra forma (talonario
    preimpreso, ver BLANQUERNA y REDONDEL en samples/).
    """
    if any(not datos.get(campo) for campo in CAMPOS_OBLIGATORIOS):
        return True
    return not datos.get("descripcion") and not datos.get("importe")


def _completar_con_vision(contenido, datos, avisos):
    """(datos, avisos, uso) rellenando con el modelo lo que los regex no sacaron.

    Los regex son de ARCA y solo entienden el formato de ARCA. El modelo lee la
    imagen sin depender del layout, asi que es el respaldo natural para un
    comprobante impreso de otra forma -- y no el OCR, que le pasaria a esos
    mismos regex un texto con los mismos rotulos raros.

    Lo que SI salio del texto manda: es exacto, mientras que el modelo puede
    equivocarse. El modelo solo llena los huecos, y se avisa cuales para que el
    operador los mire antes de confirmar.
    """
    leidos, _ = _datos_desde_vision(_primera_pagina_como_imagen(contenido))
    rellenados = [c for c in leidos if leidos.get(c) and not datos.get(c)]
    if not rellenados:
        return datos, avisos, False

    for campo in rellenados:
        datos[campo] = leidos[campo]

    # Los avisos de campo se rehacen: el de un campo que el modelo acaba de
    # llenar ya no es cierto.
    avisos = [a for a in avisos if a not in AVISOS_POR_CAMPO.values()]
    avisos.append(
        "El comprobante no tiene el formato de ARCA, así que "
        + ", ".join(ETIQUETAS_CAMPO[c] for c in rellenados if c in ETIQUETAS_CAMPO)
        + " los leyó el modelo de visión y no los regex. Revisalos antes de confirmar."
    )
    return datos, avisos + avisos_de_campos(datos), True


def extraer(nombre, contenido):
    """Resultado de la extraccion para un archivo. Nunca lanza."""
    resultado = Resultado(archivo=nombre)
    extension = os.path.splitext(nombre)[1].lower()

    if extension not in EXTENSIONES_ACEPTADAS:
        resultado.error = f"Formato no soportado ({extension or 'sin extensión'})."
        return resultado

    try:
        if extension in EXTENSIONES_PDF:
            texto = _texto_del_pdf(contenido)
            if parece_factura_arca(texto):
                datos, avisos = datos_desde_texto(texto)
                resultado.origen = "pdf-texto"
                if _regex_no_entendieron(datos) and vision_disponible():
                    try:
                        datos, avisos, uso = _completar_con_vision(contenido, datos, avisos)
                        if uso:
                            resultado.origen = "pdf-texto-vision"
                    except Exception as e:
                        # Lo leido del texto vale igual: un fallo del respaldo no
                        # puede convertir media lectura en ninguna.
                        registro.warning("El respaldo con el modelo falló: %s", e)
            else:
                # PDF sin texto: es un escaneo, va por el mismo camino que la foto.
                datos, avisos, motor = _datos_desde_imagen(_primera_pagina_como_imagen(contenido))
                resultado.origen = f"pdf-imagen-{motor}"
        else:
            datos, avisos, motor = _datos_desde_imagen(_abrir_imagen(contenido))
            resultado.origen = f"foto-{motor}"
    except VisionNoDisponible as e:
        resultado.error = str(e)
        return resultado
    except Exception as e:
        resultado.error = f"No se pudo leer el comprobante: {e}"
        return resultado

    # El DNI y el numero de afiliado salen del detalle facturado, y no se
    # cargan en la pantalla: son para que el agente le pregunte al padron de
    # SISalud la seccional del afiliado, que es el Centro de Costos real. El de
    # aca sigue siendo el deducido del domicilio, que es el que vale si el
    # padron no contesta.
    datos["dni"], datos["nro_afiliado"] = identificacion(datos.get("descripcion"))

    resultado.factura = Factura(**datos)
    resultado.avisos = avisos

    # Que falte alguno de los cuatro campos que SISalud necesita NO deja al
    # comprobante afuera: lo leido sirve igual. Hay facturas viejas, sin el
    # formato de ARCA, donde el tipo o el numero no figuran de una forma
    # reconocible, y devolverlas como "no se pudo leer" obligaba a tipear de
    # cero un comprobante del que ya teniamos el CUIT, las fechas, el importe y
    # el detalle. Se avisa que falta y el operador lo completa: aca antes de
    # cargar, o en la pantalla, que es donde tiene la factura a la vista.
    faltantes = resultado.factura.faltantes()
    if faltantes:
        etiquetas = {
            "cuit": "CUIT del emisor",
            "tipo_comprobante": "tipo de comprobante",
            "punto_venta": "punto de venta",
            "nro_factura": "número de comprobante",
        }
        nombrados = ", ".join(etiquetas[c] for c in faltantes)
        cierre = (
            "completalo acá abajo, o cargá el comprobante igual y completalo en "
            "la pantalla."
            if resultado.factura.cargable() else
            # Sin CUIT el agente no puede ni elegir el prestador, asi que no hay
            # nada que cargar hasta que alguien lo escriba.
            "sin el CUIT del emisor no se puede elegir el prestador, así que "
            "escribilo acá abajo antes de cargar."
        )
        resultado.avisos.append(
            f"No se pudo leer: {nombrados}. SISalud lo necesita para aceptar la "
            f"cabecera: {cierre}"
        )
    return resultado
