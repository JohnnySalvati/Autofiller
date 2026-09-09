"""El tramo humano: avisar en la pantalla y esperar a que el operador resuelva.

AutoFiller no confirma nunca. Carga el comprobante, avisa lo que quedo dudoso y
espera a que el operador toque Confirmar o Cancelar en SISalud. Ese control es
la razon por la que la carga sigue corriendo en la PC del operador y no en el
servidor: si la pantalla no esta a la vista, nadie revisa antes de grabar.
"""

import asyncio
import time

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
    try:
        await page.evaluate(
            """([avisos, lote]) => {
                const previo = document.getElementById('autofiller-avisos');
                if (previo) previo.remove();
                const hayAvisos = avisos.length > 0;
                const caja = document.createElement('div');
                caja.id = 'autofiller-avisos';
                caja.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;'
                    + 'background:' + (hayAvisos ? '#b91c1c' : '#1d4ed8') + ';color:#fff;'
                    + 'font:14px/1.45 Segoe UI,sans-serif;'
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
                avisos.forEach(a => {
                    const li = document.createElement('li');
                    li.style.cssText = 'margin-bottom:5px;white-space:pre-line';
                    li.textContent = a;
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
                cerrar.textContent = '×';
                cerrar.title = lote ? 'Ocultar avisos' : 'Cerrar aviso';
                cerrar.style.cssText = 'position:absolute;top:8px;right:14px;background:transparent;'
                    + 'border:0;color:#fff;font-size:24px;line-height:1;cursor:pointer';
                // Con la cola andando solo se ocultan los avisos: el progreso y
                // los botones tienen que seguir a la vista.
                cerrar.onclick = () => lote ? ul.remove() : caja.remove();
                caja.appendChild(cerrar);
                document.body.appendChild(caja);
            }""",
            [avisos, lote],
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
