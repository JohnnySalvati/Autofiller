/* AutoFiller — interfaz web.
 *
 * Habla con dos lados:
 *   - el servidor (mismo origen): lee los comprobantes y devuelve los campos.
 *   - el agente en 127.0.0.1: carga cada comprobante en el SISalud de esta PC.
 *
 * Las credenciales de SISalud solo van al agente. Nunca al servidor.
 */

const AGENTE = localStorage.getItem("autofiller.agente") || "http://127.0.0.1:8765";

// maxlength del campo Descripción en la pantalla de SISalud: lo que no entra se
// pierde sin avisar, así que se muestra el contador y se avisa antes de cargar.
const LIMITE_DESCRIPCION = 150;

const INTERVALO_SONDEO = 600;

const ETIQUETAS = {
  leyendo: "Leyendo…",
  listo: "Listo",
  revisar: "Revisar",
  problema: "No se pudo leer",
  cargando: "Cargando…",
  esperando: "Esperando tu confirmación",
  confirmado: "Confirmado",
  cancelado: "Cancelado",
  saltado: "Saltado",
  sin_confirmar: "Sin confirmar",
  detenido: "Detenido",
  pendiente: "Pendiente",
};

const CLASE_ESTADO = {
  leyendo: "leyendo", listo: "listo", revisar: "revisar", problema: "problema",
  cargando: "leyendo", esperando: "revisar", confirmado: "confirmado",
  cancelado: "descartado", saltado: "descartado", sin_confirmar: "problema",
  detenido: "descartado", pendiente: "pendiente",
};

// El agente contesta con el botón que tocó el operador ("confirmar"); la cola lo
// muestra como estado del comprobante ("confirmado").
const RESOLUCION_A_ESTADO = {
  confirmar: "confirmado",
  cancelar: "cancelado",
  saltar: "saltado",
  detener: "detenido",
  sin_confirmar: "sin_confirmar",
};

// Los estados desde los que todavía tiene sentido mandar el comprobante a SISalud.
const CARGABLES = new Set(["listo", "revisar", "pendiente"]);

const estado = {
  opciones: null,
  agenteVivo: false,
  conectado: false,
  cola: [],
  corriendo: false,
  detener: false,
  abierto: null,
};

let siguienteId = 1;

const $ = (id) => document.getElementById(id);

/* -- Ayudas de red --------------------------------------------------------- */

async function pedirAgente(ruta, opciones = {}) {
  const respuesta = await fetch(AGENTE + ruta, {
    ...opciones,
    headers: { "Content-Type": "application/json", ...(opciones.headers || {}) },
  });
  if (!respuesta.ok) {
    let detalle = respuesta.statusText;
    try { detalle = (await respuesta.json()).detail || detalle; } catch (e) { /* sin cuerpo */ }
    throw new Error(detalle);
  }
  return respuesta.status === 204 ? null : respuesta.json();
}

async function revisarAgente() {
  try {
    const salud = await pedirAgente("/api/salud");
    estado.agenteVivo = true;
    estado.conectado = salud.con_sesion;
  } catch (e) {
    estado.agenteVivo = false;
    estado.conectado = false;
  }
  dibujarChips();
}

/* -- Chips de estado ------------------------------------------------------- */

function dibujarChips() {
  const chips = $("chips");
  chips.textContent = "";
  const agregar = (texto, clase, titulo) => {
    const span = document.createElement("span");
    span.className = "chip " + clase;
    span.textContent = texto;
    if (titulo) span.title = titulo;
    chips.appendChild(span);
  };

  if (estado.agenteVivo) {
    agregar(estado.conectado ? "SISalud conectado" : "Agente listo",
            estado.conectado ? "ok" : "tibio");
  } else {
    agregar("Agente no detectado", "mal",
            "Arrancá el agente de AutoFiller en esta PC para poder cargar en SISalud.");
  }

  if (estado.opciones) {
    if (estado.opciones.lectura_de_fotos) {
      agregar("Fotos: sí", "ok", "Modelo: " + estado.opciones.modelo_vision);
    } else {
      agregar("Fotos: no", "tibio",
              "El servidor no tiene configurada la lectura de fotos (falta ANTHROPIC_API_KEY). Solo PDF.");
    }
  }
}

/* -- Sesión ---------------------------------------------------------------- */

function mostrarVista(cual) {
  $("vista-login").classList.toggle("oculto", cual !== "login");
  $("vista-app").classList.toggle("oculto", cual !== "app");
}

async function conectar(evento) {
  evento.preventDefault();
  const aviso = $("aviso-sesion");
  const usuario = $("usuario").value.trim();
  const clave = $("clave").value;

  try {
    await pedirAgente("/api/sesion", {
      method: "POST",
      body: JSON.stringify({ usuario, clave }),
    });
  } catch (e) {
    aviso.className = "aviso error";
    aviso.textContent =
      "No se pudo hablar con el agente de AutoFiller en esta PC (" + e.message + ").\n" +
      "Verificá que esté abierto y volvé a intentar.";
    return;
  }

  // El usuario se recuerda para no retipearlo; la contraseña nunca.
  localStorage.setItem("autofiller.usuario", usuario);
  estado.conectado = true;
  $("clave").value = "";
  aviso.className = "aviso oculto";
  mostrarVista("app");
  dibujarChips();
  dibujarCola();
}

async function desconectar() {
  try { await pedirAgente("/api/sesion", { method: "DELETE" }); } catch (e) { /* ya estaba */ }
  estado.conectado = false;
  $("clave").value = "";
  $("aviso-sesion").className = "aviso oculto";
  mostrarVista("login");
  dibujarChips();
  dibujarCola();
}

/* -- Alta de comprobantes -------------------------------------------------- */

async function agregarArchivos(archivos) {
  if (!archivos.length) return;

  const nuevos = Array.from(archivos).map((archivo) => ({
    id: siguienteId++,
    archivo: archivo.name,
    estado: "leyendo",
    factura: null,
    avisos: [],
    error: "",
    origen: "",
  }));
  estado.cola.push(...nuevos);
  dibujarCola();

  const cuerpo = new FormData();
  Array.from(archivos).forEach((archivo) => cuerpo.append("archivos", archivo));

  let resultados;
  try {
    const respuesta = await fetch("/api/extraer", { method: "POST", body: cuerpo });
    if (!respuesta.ok) throw new Error("el servidor respondió " + respuesta.status);
    resultados = await respuesta.json();
  } catch (e) {
    nuevos.forEach((item) => {
      item.estado = "problema";
      item.error = "No se pudo leer el comprobante: " + e.message;
    });
    dibujarCola();
    return;
  }

  // El servidor devuelve un resultado por archivo, en el mismo orden.
  nuevos.forEach((item, i) => {
    const resultado = resultados[i];
    if (!resultado) {
      item.estado = "problema";
      item.error = "El servidor no devolvió datos para este archivo.";
      return;
    }
    item.origen = resultado.origen;
    item.avisos = resultado.avisos || [];
    item.error = resultado.error || "";
    item.factura = resultado.factura;
    item.estado = resultado.error ? "problema"
                : item.avisos.length ? "revisar" : "listo";
  });
  dibujarCola();
}

/* -- Dibujo de la cola ----------------------------------------------------- */

function crearCampo(item, campo) {
  const etiqueta = document.createElement("label");
  if (campo.ancho) etiqueta.classList.add("completo");
  etiqueta.append(campo.etiqueta);

  let control;
  if (campo.tipo === "select") {
    control = document.createElement("select");
    const vacio = document.createElement("option");
    vacio.value = "";
    vacio.textContent = campo.vacio || "(sin elegir)";
    control.appendChild(vacio);
    Object.entries(campo.opciones()).forEach(([valor, texto]) => {
      const opcion = document.createElement("option");
      opcion.value = valor;
      opcion.textContent = texto;
      control.appendChild(opcion);
    });
  } else if (campo.tipo === "textarea") {
    control = document.createElement("textarea");
    control.rows = 2;
  } else {
    control = document.createElement("input");
    control.type = "text";
  }

  control.value = item.factura[campo.clave] || "";
  control.addEventListener("input", () => {
    item.factura[campo.clave] = control.value;
    if (campo.clave === "descripcion") actualizarContador(item, etiqueta);
  });
  control.addEventListener("change", () => {
    item.factura[campo.clave] = control.value;
  });
  etiqueta.appendChild(control);

  if (campo.clave === "descripcion") {
    const contador = document.createElement("span");
    contador.className = "contador";
    etiqueta.appendChild(contador);
    actualizarContador(item, etiqueta);
  }
  return etiqueta;
}

function actualizarContador(item, etiqueta) {
  const contador = etiqueta.querySelector(".contador");
  if (!contador) return;
  const largo = (item.factura.descripcion || "").length;
  contador.textContent = largo + " / " + LIMITE_DESCRIPCION + " caracteres" +
    (largo > LIMITE_DESCRIPCION ? " — SISalud recorta lo que sobra" : "");
  contador.classList.toggle("pasado", largo > LIMITE_DESCRIPCION);
}

function camposDelEditor() {
  return [
    { clave: "cuit", etiqueta: "CUIT del emisor" },
    { clave: "tipo_comprobante", etiqueta: "Tipo de comprobante", tipo: "select",
      opciones: () => estado.opciones.tipos_comprobante },
    { clave: "punto_venta", etiqueta: "Punto de venta" },
    { clave: "nro_factura", etiqueta: "Número de comprobante" },
    { clave: "fecha_emision", etiqueta: "Fecha de emisión" },
    { clave: "fecha_recepcion", etiqueta: "Fecha de recepción" },
    { clave: "fecha_vencimiento", etiqueta: "Vencimiento" },
    { clave: "fecha_devengamiento", etiqueta: "Devengamiento" },
    { clave: "cae", etiqueta: "CAE" },
    { clave: "importe", etiqueta: "Importe" },
    { clave: "centro_costo", etiqueta: "Centro de costos", tipo: "select",
      vacio: "(queda CENTRAL, lo elegís en SISalud)",
      opciones: () => estado.opciones.centros_costo },
    { clave: "descripcion", etiqueta: "Descripción del detalle", tipo: "textarea", ancho: true },
  ];
}

function dibujarItem(item) {
  const li = document.createElement("li");
  li.className = "item" + (item.estado === "cargando" || item.estado === "esperando" ? " activo" : "");

  const cabecera = document.createElement("button");
  cabecera.type = "button";
  cabecera.className = "item-cabecera";
  cabecera.setAttribute("aria-expanded", String(estado.abierto === item.id));

  const nombre = document.createElement("span");
  nombre.className = "item-nombre";
  nombre.textContent = item.archivo;
  cabecera.appendChild(nombre);

  if (item.factura) {
    const resumen = document.createElement("span");
    resumen.className = "item-resumen";
    const partes = [];
    if (item.factura.punto_venta && item.factura.nro_factura) {
      partes.push(item.factura.punto_venta.padStart(4, "0") + "-" +
                  item.factura.nro_factura.padStart(8, "0"));
    }
    if (item.factura.importe) partes.push("$ " + item.factura.importe);
    if (item.factura.centro_costo_nombre) partes.push(item.factura.centro_costo_nombre);
    resumen.textContent = partes.join("  ·  ");
    cabecera.appendChild(resumen);
  }

  const insignia = document.createElement("span");
  insignia.className = "estado " + (CLASE_ESTADO[item.estado] || "pendiente");
  insignia.textContent = ETIQUETAS[item.estado] || item.estado;
  cabecera.appendChild(insignia);

  cabecera.addEventListener("click", () => {
    estado.abierto = estado.abierto === item.id ? null : item.id;
    dibujarCola();
  });
  li.appendChild(cabecera);

  if (estado.abierto === item.id) {
    const cuerpo = document.createElement("div");
    cuerpo.className = "item-cuerpo";

    if (item.error) {
      const error = document.createElement("p");
      error.className = "aviso error";
      error.textContent = item.error;
      cuerpo.appendChild(error);
    }

    if (item.avisos.length) {
      const lista = document.createElement("ul");
      lista.className = "avisos";
      item.avisos.forEach((texto) => {
        const linea = document.createElement("li");
        linea.textContent = texto;
        lista.appendChild(linea);
      });
      cuerpo.appendChild(lista);
    }

    if (item.factura) {
      const campos = document.createElement("div");
      campos.className = "campos";
      camposDelEditor().forEach((campo) => campos.appendChild(crearCampo(item, campo)));
      cuerpo.appendChild(campos);

      if (item.origen) {
        const origen = document.createElement("p");
        origen.className = "sutil";
        origen.style.marginTop = "12px";
        origen.textContent = {
          "pdf-texto": "Leído del texto del PDF.",
          "pdf-imagen": "El PDF venía escaneado: se leyó como imagen.",
          "foto": "Leído de la foto (QR de ARCA + lectura de imagen).",
        }[item.origen] || "";
        cuerpo.appendChild(origen);
      }
    }

    const quitar = document.createElement("button");
    quitar.type = "button";
    quitar.className = "secundario";
    quitar.style.marginTop = "14px";
    quitar.textContent = "Quitar de la cola";
    quitar.addEventListener("click", () => {
      estado.cola = estado.cola.filter((otro) => otro.id !== item.id);
      dibujarCola();
    });
    cuerpo.appendChild(quitar);

    li.appendChild(cuerpo);
  }
  return li;
}

function dibujarCola() {
  const seccion = $("seccion-cola");
  seccion.classList.toggle("oculto", estado.cola.length === 0);

  const lista = $("cola");
  lista.textContent = "";
  estado.cola.forEach((item) => lista.appendChild(dibujarItem(item)));

  const porCargar = estado.cola.filter((i) => CARGABLES.has(i.estado)).length;
  const conProblema = estado.cola.filter((i) => i.estado === "problema").length;
  const partes = [estado.cola.length + " comprobantes", porCargar + " por cargar"];
  if (conProblema) partes.push(conProblema + " sin leer");
  $("resumen-cola").textContent = partes.join("  ·  ");

  const boton = $("cargar");
  boton.disabled = estado.corriendo || porCargar === 0 || !estado.conectado;
  boton.textContent = porCargar === 1 ? "Cargar 1 comprobante"
                                      : "Cargar " + porCargar + " comprobantes";
  boton.title = estado.conectado ? "" : "Primero conectate con SISalud.";
  $("vaciar").disabled = estado.corriendo;
}

/* -- Carga en SISalud ------------------------------------------------------ */

function panel(titulo, detalle, visible = true) {
  $("panel").classList.toggle("oculto", !visible);
  $("panel-titulo").textContent = titulo;
  $("panel-detalle").textContent = detalle;
}

async function esperarResolucion() {
  while (true) {
    await new Promise((listo) => setTimeout(listo, INTERVALO_SONDEO));
    let info;
    try {
      info = await pedirAgente("/api/trabajo");
    } catch (e) {
      return { resolucion: "sin_confirmar", avisos: [],
               error: "Se perdió la conexión con el agente: " + e.message };
    }
    if (info.fase === "esperando") {
      panel("Comprobante " + info.indice + " de " + info.total + " en pantalla",
            "Revisalo en SISalud y tocá Confirmar o Cancelar.");
    }
    if (info.fase === "terminado") return info;
  }
}

async function cargarCola() {
  estado.corriendo = true;
  estado.detener = false;
  dibujarCola();

  const porCargar = estado.cola.filter((i) => CARGABLES.has(i.estado));
  const total = porCargar.length;

  for (let i = 0; i < total; i++) {
    const item = porCargar[i];
    if (estado.detener) break;

    item.estado = "cargando";
    dibujarCola();
    panel("Cargando " + (i + 1) + " de " + total, item.archivo);

    try {
      await pedirAgente("/api/trabajo", {
        method: "POST",
        body: JSON.stringify({
          factura: item.factura,
          archivo: item.archivo,
          indice: i + 1,
          total,
        }),
      });
    } catch (e) {
      item.estado = "problema";
      item.error = "El agente no pudo cargarlo: " + e.message;
      dibujarCola();
      continue;
    }

    item.estado = "esperando";
    dibujarCola();

    const info = await esperarResolucion();
    item.estado = RESOLUCION_A_ESTADO[info.resolucion] || "sin_confirmar";
    if (info.avisos && info.avisos.length) item.avisos = info.avisos;
    if (info.error) item.error = info.error;
    if (item.estado === "detenido") estado.detener = true;
    dibujarCola();

    try { await pedirAgente("/api/listo", { method: "POST" }); } catch (e) { /* da igual */ }
  }

  // Lo que no llegó a cargarse por un "Detener" vuelve a quedar disponible.
  porCargar.forEach((item) => {
    if (item.estado === "cargando" || item.estado === "esperando") item.estado = "pendiente";
  });

  estado.corriendo = false;
  panel("", "", false);
  dibujarCola();
  mostrarResumen();
}

function mostrarResumen() {
  const conteo = {};
  estado.cola.forEach((item) => {
    conteo[ETIQUETAS[item.estado] || item.estado] =
      (conteo[ETIQUETAS[item.estado] || item.estado] || 0) + 1;
  });
  const aviso = $("aviso-lectura");
  aviso.className = "aviso bien";
  aviso.textContent = "Cola terminada — " +
    Object.entries(conteo).map(([k, v]) => k + ": " + v).join("   ·   ");
}

/* -- Arranque -------------------------------------------------------------- */

async function iniciar() {
  $("usuario").value = localStorage.getItem("autofiller.usuario") || "";

  try {
    estado.opciones = await (await fetch("/api/opciones")).json();
    $("zona-formatos").textContent = estado.opciones.lectura_de_fotos
      ? "PDF o foto (" + estado.opciones.extensiones.join(", ") + ")"
      : "Solo PDF — el servidor no tiene configurada la lectura de fotos";
  } catch (e) {
    $("aviso-lectura").className = "aviso error";
    $("aviso-lectura").textContent = "No se pudo hablar con el servidor de AutoFiller.";
  }

  await revisarAgente();
  // El agente puede tener la sesión abierta de antes (se recargó la página, o se
  // volvió a entrar): en ese caso no hay que pedir la contraseña de nuevo.
  mostrarVista(estado.conectado ? "app" : "login");
  if (!estado.agenteVivo) {
    $("aviso-sesion").className = "aviso";
    $("aviso-sesion").textContent =
      "No se detecta el agente de AutoFiller en esta PC. Es el programa que abre " +
      "SISalud en tu Chrome: abrilo y recargá esta página.";
  }

  const zona = $("zona");
  const entrada = $("archivos");
  zona.addEventListener("click", () => entrada.click());
  zona.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); entrada.click(); }
  });
  entrada.addEventListener("change", () => {
    agregarArchivos(entrada.files);
    entrada.value = "";
  });
  ["dragenter", "dragover"].forEach((evento) =>
    zona.addEventListener(evento, (e) => { e.preventDefault(); zona.classList.add("encima"); }));
  ["dragleave", "drop"].forEach((evento) =>
    zona.addEventListener(evento, (e) => { e.preventDefault(); zona.classList.remove("encima"); }));
  zona.addEventListener("drop", (e) => agregarArchivos(e.dataTransfer.files));

  $("form-sesion").addEventListener("submit", conectar);
  $("desconectar").addEventListener("click", desconectar);
  $("cargar").addEventListener("click", cargarCola);
  $("vaciar").addEventListener("click", () => {
    estado.cola = [];
    estado.abierto = null;
    $("aviso-lectura").className = "aviso oculto";
    dibujarCola();
  });
  $("saltar").addEventListener("click", () =>
    pedirAgente("/api/accion", { method: "POST", body: JSON.stringify({ accion: "saltar" }) }));
  $("detener").addEventListener("click", () =>
    pedirAgente("/api/accion", { method: "POST", body: JSON.stringify({ accion: "detener" }) }));

  // El agente puede abrirse o cerrarse mientras la página está abierta. Si se
  // cerró, la sesión se fue con él: hay que volver a Entrar.
  setInterval(async () => {
    if (estado.corriendo) return;
    await revisarAgente();
    if (!estado.conectado && $("vista-app").classList.contains("oculto") === false) {
      mostrarVista("login");
      dibujarCola();
    }
  }, 5000);
}

iniciar();
