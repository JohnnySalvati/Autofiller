"""El tramo humano: avisar en la pantalla y esperar a que el operador resuelva.

AutoFiller no confirma nunca. Carga el comprobante, avisa lo que quedo dudoso y
espera a que el operador toque Confirmar o Cancelar en SISalud. Ese control es
la razon por la que la carga sigue corriendo en la PC del operador y no en el
servidor: si la pantalla no esta a la vista, nadie revisa antes de grabar.
"""

import asyncio
import time


class Avisos(list):
    """Los avisos del comprobante, sabiendo cuales son leves.

    Un aviso LEVE es el que no impide confirmar: la carga quedo completa y el
    operador solo tiene que mirar algo. Hoy el unico es el recorte de la
    descripcion por el maxlength de la pantalla, que pasa en casi la mitad de
    los comprobantes -- 27 de las 63 muestras pasan de 150 caracteres. Pintar el
    cartel de rojo por eso hacia que el rojo dejara de significar "esto esta
    trabado", que es justo lo que el cartel tiene que comunicar.

    Todos los demas avisos son graves: piden una accion sin la cual el
    comprobante no se puede confirmar (elegir el tipo, agregar la linea, elegir
    el centro de costos, dar de alta el CUIT).

    Hereda de list para no cambiarle la forma a Estado.avisos ni a lo que la web
    ya sabe mostrar.
    """

    def __init__(self, *args):
        super().__init__(*args)
        self.leves = set()

    def leve(self, texto):
        self.leves.add(len(self))
        self.append(texto)

    @property
    def indices_leves(self):
        return sorted(self.leves)

    @property
    def solo_leves(self):
        return bool(self) and len(self.leves) == len(self)


def avisar_leve(avisos, texto):
    """Agrega un aviso leve. Con una lista comun se comporta como append."""
    if isinstance(avisos, Avisos):
        avisos.leve(texto)
    else:
        avisos.append(texto)

# Verificado por CDP (2026-09-09): Cancelar (input name=BUTTON2, evento RETURN de
# GeneXus) navega fuera de la pantalla. Confirmar (input name=CONFIRMAR, atajo
# F12) no se pudo probar sin grabar un comprobante real; se asume que al grabar
# la pantalla navega o vuelve al formulario vacio. En ambos casos el comprobante
# "desaparece" de la pantalla: eso es lo que se detecta. Mientras el operador
# corrige un error de validacion, el comprobante sigue en pantalla y se espera.
# La botonera #TBL_BOTONES (Confirmar/Cancelar) esta oculta con el formulario
# vacio y aparece cuando hay comprobante cargado.
ESTADO_PANTALLA = """() => {
    const v = s => (document.querySelector(s) || {}).value;
    const botones = document.querySelector('#TBL_BOTONES');
    return {
        mascara: !!document.querySelector('div.gx-mask'),
        comprobante: !!botones && getComputedStyle(botones).display !== 'none'
            && v('#vENTIDADCODIGO') !== '00000000',
    };
}"""

# Cuanto tiene que sostenerse la ausencia del comprobante para darla por firme.
# Un redibujado de GeneXus puede ocultar la botonera un instante; sin esta
# espera se pasaba al siguiente comprobante antes de que el operador lo viera.
REPOSO_AUSENCIA = 1.5

ESTADOS = {
    "confirmar": "CONFIRMADO",
    "cancelar": "CANCELADO",
    "sin_confirmar": "SIN CONFIRMAR (la pantalla se cerró o recargó)",
    "saltar": "SALTADO",
    "detener": "DETENIDO (quedó en pantalla, sin confirmar)",
}


class Control:
    """Estado compartido entre el bucle de carga y los botones de la pantalla."""

    def __init__(self):
        self.ultimo_boton = None  # 'confirmar' | 'cancelar' (el que toco el operador)
        self.pedido = None        # 'saltar' | 'detener' (botones del cartel o de la web)
        self.automatico = True    # True mientras carga AutoFiller; False mientras espera al operador

    def registrar(self, accion):
        if accion in ("confirmar", "cancelar"):
            self.ultimo_boton = accion
        elif accion in ("saltar", "detener"):
            self.pedido = accion

    def nuevo_comprobante(self):
        self.ultimo_boton = None
        self.pedido = None
        self.automatico = True


async def mostrar_avisos_en_pantalla(page, avisos, lote=None):
    """Muestra los avisos como un cartel dentro de la propia pantalla de SISalud.

    Va inyectado en la pagina, fijo arriba de todo, que es donde el operador esta
    mirando. El titulo del comprobante va abajo y el boton Confirmar esta al pie,
    asi que el cartel no tapa nada de lo que necesita.

    Con lote = {indice, total, archivo} el cartel se muestra siempre: lleva el
    progreso de la cola, los botones "Siguiente comprobante" y "Detener", y
    engancha Confirmar (click o F12) y Cancelar de SISalud para saber cual toco
    el operador. Todo avisa al agente por window.autofillerAccion.
    """
    if not avisos and not lote:
        return
    leves = avisos.indices_leves if isinstance(avisos, Avisos) else []
    try:
        await page.evaluate(
            """([avisos, lote, leves]) => {
                const previo = document.getElementById('autofiller-avisos');
                if (previo) previo.remove();
                const hayAvisos = avisos.length > 0;
                // Amarillo cuando TODOS los avisos son leves: la carga quedo
                // completa y solo hay algo para mirar. Rojo apenas hay uno que
                // impide confirmar. Sin avisos, azul: es solo el progreso.
                const soloLeves = hayAvisos && leves.length === avisos.length;
                const fondo = !hayAvisos ? '#1d4ed8' : (soloLeves ? '#eab308' : '#b91c1c');
                const tinta = soloLeves ? '#1c1917' : '#fff';
                const chipFondo = soloLeves ? 'rgba(0,0,0,.14)' : 'rgba(255,255,255,.22)';
                const caja = document.createElement('div');
                caja.id = 'autofiller-avisos';
                // max-height + overflow: con cuatro avisos largos el cartel llegaba a
                // tapar media pantalla, justo la mitad de arriba, que es donde estan
                // los campos que el operador tiene que corregir a mano.
                caja.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;'
                    + 'background:' + fondo + ';color:' + tinta + ';'
                    + 'font:14px/1.45 Segoe UI,sans-serif;max-height:38vh;overflow:auto;'
                    + 'padding:12px 46px 14px 18px;box-shadow:0 2px 10px rgba(0,0,0,.45)';
                const titulo = document.createElement('div');
                titulo.style.cssText = 'font-weight:700;font-size:15px;margin-bottom:6px';
                if (lote) {
                    titulo.textContent = 'AutoFiller — comprobante ' + lote.indice + ' de '
                        + lote.total + ': ' + lote.archivo
                        + (hayAvisos ? ' — revisar antes de Confirmar:' : '');
                } else {
                    titulo.textContent = 'AutoFiller — revisar antes de Confirmar:';
                }
                caja.appendChild(titulo);
                const ul = document.createElement('ul');
                ul.style.cssText = 'margin:0;padding-left:22px';
                avisos.forEach((a, i) => {
                    const li = document.createElement('li');
                    li.style.cssText = 'margin-bottom:5px;white-space:pre-line';
                    // En un cartel rojo con avisos de los dos tipos, el leve
                    // lleva su etiqueta: si no, se lee como uno mas que traba.
                    if (leves.includes(i) && !soloLeves) {
                        const chip = document.createElement('span');
                        chip.textContent = 'solo revisar';
                        chip.style.cssText = 'display:inline-block;background:' + chipFondo
                            + ';border-radius:3px;padding:1px 7px;margin-right:7px;'
                            + 'font-size:12px;font-weight:600;white-space:nowrap';
                        li.appendChild(chip);
                    }
                    li.appendChild(document.createTextNode(a));
                    ul.appendChild(li);
                });
                caja.appendChild(ul);
                if (lote) {
                    const pie = document.createElement('div');
                    pie.style.cssText = 'margin-top:8px;display:flex;gap:10px;align-items:center;flex-wrap:wrap';
                    const nota = document.createElement('span');
                    nota.textContent = 'Revisá y tocá Confirmar o Cancelar en SISalud: '
                        + 'al terminar se carga el siguiente.';
                    pie.appendChild(nota);
                    const boton = (texto, accion) => {
                        const b = document.createElement('button');
                        b.type = 'button';
                        b.textContent = texto;
                        b.style.cssText = 'background:#fff;color:#111;border:0;border-radius:4px;'
                            + 'padding:5px 12px;font:600 13px Segoe UI,sans-serif;cursor:pointer';
                        b.onclick = () => { b.disabled = true; window.autofillerAccion(accion); };
                        return b;
                    };
                    pie.appendChild(boton('Siguiente comprobante (saltar este)', 'saltar'));
                    pie.appendChild(boton('Detener la cola', 'detener'));
                    caja.appendChild(pie);
                    // Enganchar los botones de SISalud para saber cual toco el operador.
                    const confirmar = document.querySelector('input[name="CONFIRMAR"]');
                    if (confirmar) confirmar.addEventListener('click',
                        () => window.autofillerAccion('confirmar'), true);
                    const cancelar = document.querySelector('input[name="BUTTON2"]');
                    if (cancelar) cancelar.addEventListener('click',
                        () => window.autofillerAccion('cancelar'), true);
                    document.addEventListener('keydown', e => {
                        if (e.key === 'F12') window.autofillerAccion('confirmar');
                    }, true);
                }
                const cerrar = document.createElement('button');
                cerrar.style.cssText = 'position:fixed;top:8px;right:14px;background:transparent;'
                    + 'border:0;color:' + tinta + ';font-size:22px;line-height:1;cursor:pointer;'
                    + 'z-index:2147483647';
                caja.appendChild(cerrar);
                document.body.appendChild(caja);

                // El cartel es position:fixed, o sea que flota por ENCIMA de la
                // pantalla de SISalud. Empujando el body hacia abajo la altura
                // exacta del cartel, deja de tapar nada: el formulario entero
                // sigue accesible con el cartel a la vista.
                const previoPadding = document.body.dataset.autofillerPadding;
                if (previoPadding === undefined) {
                    document.body.dataset.autofillerPadding = document.body.style.paddingTop || '';
                }
                const empujar = () => {
                    try { document.body.style.paddingTop = caja.offsetHeight + 'px'; } catch (e) {}
                };

                // El boton PLIEGA los avisos, no los borra. Antes los removia, y
                // eso dejaba al operador sin las indicaciones que necesita
                // justamente para corregir a mano lo que el cartel le pide.
                let plegado = false;
                const pintarBoton = () => {
                    cerrar.textContent = lote ? (plegado ? '▾' : '▴') : '×';
                    cerrar.title = lote
                        ? (plegado ? 'Mostrar los avisos' : 'Plegar los avisos')
                        : 'Cerrar aviso';
                };
                cerrar.onclick = () => {
                    if (!lote) {
                        document.body.style.paddingTop = document.body.dataset.autofillerPadding || '';
                        caja.remove();
                        return;
                    }
                    plegado = !plegado;
                    if (hayAvisos) ul.hidden = plegado;
                    titulo.style.marginBottom = plegado ? '0' : '6px';
                    pintarBoton();
                    empujar();
                };
                pintarBoton();
                empujar();
                // GeneXus redibuja al volver cada postback y la altura puede cambiar.
                if (window.ResizeObserver) new ResizeObserver(empujar).observe(caja);
            }""",
            [avisos, lote, leves],
        )
    except Exception:
        # El cartel es un extra: si la pagina no permite inyectarlo, seguimos.
        pass


async def esperar_resolucion(page, control):
    """Espera a que el operador termine con el comprobante en pantalla.

    Termina cuando el comprobante ya no esta en la pantalla (Confirmar o
    Cancelar), o cuando piden saltar o detener (desde el cartel o desde la web).
    Devuelve 'confirmar', 'cancelar', 'sin_confirmar', 'saltar' o 'detener'.

    "Ya no esta" tiene que sostenerse REPOSO_AUSENCIA segundos sin mascara de
    GeneXus, porque un redibujado puede esconder la botonera un instante. Que la
    pestana haya navegado no alcanza por si solo (llegan navegaciones tardias
    del comprobante anterior): se mira siempre el estado real de la pantalla.
    Si evaluar falla es porque la pagina esta cambiando de documento, y eso
    cuenta como ausencia.

    Si la carga automatica fallo, la pantalla arranca vacia: primero se espera a
    ver el comprobante (el operador lo carga a mano) y despues a que se vaya.
    """
    visto = False
    ausente_desde = None
    while True:
        if control.pedido:
            return control.pedido
        try:
            estado = await page.evaluate(ESTADO_PANTALLA)
        except Exception:
            estado = None  # cambiando de documento (Cancelar navega)
        if estado and estado["comprobante"]:
            visto = True
            ausente_desde = None
        elif estado and estado["mascara"]:
            pass  # GeneXus procesando: no se decide nada
        elif visto:
            ausente_desde = ausente_desde or time.monotonic()
            if time.monotonic() - ausente_desde >= REPOSO_AUSENCIA:
                break
        await asyncio.sleep(0.3)
    return control.ultimo_boton or "sin_confirmar"
