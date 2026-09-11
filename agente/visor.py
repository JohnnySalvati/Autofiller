"""La ventana de actividad del agente: el log, en vivo.

Reemplaza a la ventana negra que el agente tenia abierta todo el dia. La
diferencia es que ahora se abre cuando el operador la pide (menu del icono de la
bandeja) y cerrarla no apaga el agente, que era el accidente mas comun.

Corre en su propio hilo con su propio mainloop de Tk. Tkinter banca eso siempre
que todos los widgets se creen y se toquen en ese mismo hilo, que es lo que pasa
aca: el resto del agente se comunica con el visor por Events, nunca tocando un
widget. El hilo principal queda para el icono de la bandeja, que en Windows lo
necesita para su bucle de mensajes.
"""

import threading
import webbrowser

import registro
import recursos

# Un visor a la vez. Si ya esta abierto, el segundo pedido lo trae al frente en
# vez de abrir otra ventana con el mismo log.
_hilo = None
_al_frente = threading.Event()
_candado = threading.Lock()

INTERVALO_MS = 700
# Mas que esto y el Text empieza a pesar. Un lote entero no llega ni cerca.
MAXIMO_RENGLONES = 5000


def abrir():
    """Abre la ventana de actividad, o la trae al frente si ya estaba."""
    global _hilo
    with _candado:
        if _hilo is not None and _hilo.is_alive():
            _al_frente.set()
            return
        _al_frente.clear()
        _hilo = threading.Thread(target=_correr, name="visor", daemon=True)
        _hilo.start()


def _correr():
    import tkinter as tk
    from tkinter import ttk

    raiz = tk.Tk()
    raiz.title("AutoFiller — actividad del agente")
    raiz.geometry("900x520")
    raiz.minsize(520, 260)

    icono = recursos.ruta("autofiller.ico")
    if icono:
        try:
            raiz.iconbitmap(icono)
        except Exception:
            pass

    marco = ttk.Frame(raiz, padding=0)
    marco.pack(fill="both", expand=True)

    texto = tk.Text(
        marco, wrap="none", font=("Consolas", 9), background="#0F172A",
        foreground="#E6EDF3", insertbackground="#E6EDF3", borderwidth=0,
        padx=10, pady=8, state="disabled",
    )
    vertical = ttk.Scrollbar(marco, orient="vertical", command=texto.yview)
    horizontal = ttk.Scrollbar(marco, orient="horizontal", command=texto.xview)
    texto.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
    texto.grid(row=0, column=0, sticky="nsew")
    vertical.grid(row=0, column=1, sticky="ns")
    horizontal.grid(row=1, column=0, sticky="ew")
    marco.rowconfigure(0, weight=1)
    marco.columnconfigure(0, weight=1)

    pie = ttk.Frame(raiz, padding=(10, 6))
    pie.pack(fill="x")
    ttk.Label(pie, text=registro.ARCHIVO, foreground="#6B7280").pack(side="left")
    ttk.Button(pie, text="Cerrar", command=raiz.destroy).pack(side="right")
    ttk.Button(pie, text="Abrir la carpeta",
               command=lambda: webbrowser.open(registro.CARPETA)).pack(
                   side="right", padx=(0, 6))
    ttk.Button(pie, text="Copiar todo",
               command=lambda: _copiar(raiz, texto)).pack(side="right", padx=(0, 6))

    estado = {"posicion": 0}

    def refrescar():
        nuevo, estado["posicion"] = registro.leer(estado["posicion"])
        if nuevo:
            # Seguir al final solo si el operador ya estaba mirando el final. Si
            # se fue para arriba a leer un error, el autoscroll se lo arrebataria.
            al_final = texto.yview()[1] > 0.999
            texto.configure(state="normal")
            texto.insert("end", nuevo)
            sobran = int(texto.index("end-1c").split(".")[0]) - MAXIMO_RENGLONES
            if sobran > 0:
                texto.delete("1.0", f"{sobran + 1}.0")
            texto.configure(state="disabled")
            if al_final:
                texto.see("end")
        if _al_frente.is_set():
            _al_frente.clear()
            raiz.deiconify()
            raiz.lift()
            raiz.focus_force()
        raiz.after(INTERVALO_MS, refrescar)

    refrescar()
    raiz.mainloop()


def _copiar(raiz, texto):
    raiz.clipboard_clear()
    raiz.clipboard_append(texto.get("1.0", "end-1c"))
