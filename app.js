/* =========================================================================
   app.js — pega el navegador con el código Python.

   Qué hace, en orden:
     1. Arranca Pyodide (Python dentro del navegador) con pandas y openpyxl.
     2. Copia los scripts de cada zona (python/rm/…) dentro del motor.
     3. Guarda los ARCHIVOS BASE en el navegador, para que la operadora no
        tenga que volver a subirlos en cada corrida.
     4. Recibe los movimientos del mes (archivos sueltos o una carpeta .zip).
     5. Llama a procesar() en python/procesar.py y muestra el resultado.

   La lógica de negocio NO vive acá: está en python/procesar.py y en los
   scripts de python/rm/.
   ========================================================================= */

let zonaElegida = null;      // "RM" | "SUR" | "NORTE" | "UNIR"
let archivos = [];           // movimientos del mes que subió la operadora
let baseGuardada = [];       // archivos base que ya están en el navegador
let hayPrepararBase = false; // ¿el procesar.py del repo sabe recibir archivos base?

/* Los cuatro archivos base. procesar.py devuelve esta misma lista y la
   reemplaza, así el nombre canónico vive en un solo lugar (Python). */
let esperadosBase = [
  { etiqueta: "Homologación",          nombre: "HOMOLOGACION.xlsx" },
  { etiqueta: "Destinatarios",         nombre: "BBDD DESTINATARIO.xlsx" },
  { etiqueta: "Clasificación SINADER", nombre: "Clasificación_Residuos SINADER.xlsx" },
  { etiqueta: "Transportistas",        nombre: "Transportistas.xlsx" },
];

const $ = (id) => document.getElementById(id);

/* =========================================================================
   CABLEADO DEFENSIVO
   -------------------------------------------------------------------------
   Si el index.html publicado fuera una versión anterior y le faltara algún
   elemento, antes se caía todo este archivo en la primera línea y la página
   quedaba muda: ni avisaba ni arrancaba el motor. Ahora se anota qué falta,
   se sigue con el resto, y se muestra el aviso en pantalla.
   ========================================================================= */
const faltanElementos = [];

function nodo(id) {
  const n = $(id);
  if (!n) faltanElementos.push(id);
  return n;
}

function alEvento(id, evento, fn) {
  const n = nodo(id);
  if (n) n.addEventListener(evento, fn);
  return n;
}

/* Franja fija abajo: para los avisos que no se pueden perder de vista. */
function mostrarFranja(texto) {
  console.error(texto);
  let aviso = $("franja-aviso");
  if (!aviso) {
    aviso = document.createElement("div");
    aviso.id = "franja-aviso";
    aviso.style.cssText =
      "position:fixed;left:0;right:0;bottom:0;z-index:9999;padding:12px 16px;" +
      "background:#B4610A;color:#fff;font-size:13px;line-height:1.5;" +
      "font-family:inherit";
    document.body.appendChild(aviso);
  }
  aviso.textContent = texto;
}

function avisarDesajuste() {
  mostrarFranja(
    "El index.html publicado no calza con app.js: faltan " +
    faltanElementos.join(", ") +
    ". Sube el index.html nuevo y recarga con Ctrl+Shift+R."
  );
}

const ACEPTADOS = [".xlsx", ".zip"];
const esValido = (nombre) =>
  ACEPTADOS.some((ext) => nombre.toLowerCase().endsWith(ext)) &&
  !nombre.startsWith("~$");

/* =========================================================================
   MEMORIA DEL NAVEGADOR — donde quedan guardados los archivos base
   -------------------------------------------------------------------------
   Se usa IndexedDB, que es el almacén del navegador para archivos grandes.
   Los archivos quedan SOLO en este computador: no se suben a ningún lado.
   ========================================================================= */
const DB_NOMBRE = "trazabilidad_app";
const DB_STORE = "archivos_base";

function abrirDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NOMBRE, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(DB_STORE)) {
        db.createObjectStore(DB_STORE, { keyPath: "nombre" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function operarDB(modo, accion) {
  return abrirDB().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(DB_STORE, modo);
        const store = tx.objectStore(DB_STORE);
        const pedido = accion(store);
        tx.oncomplete = () => resolve(pedido && pedido.result);
        tx.onerror = () => reject(tx.error);
      })
  );
}

const leerBaseDB = () => operarDB("readonly", (s) => s.getAll());
const borrarBaseDB = () => operarDB("readwrite", (s) => s.clear());
const borrarUnoBaseDB = (nombre) =>
  operarDB("readwrite", (s) => { s.delete(nombre); });
// Suma a lo que ya había. Antes hacía clear() primero, y por eso subir los
// archivos de a uno iba borrando los anteriores. Un archivo con el mismo
// nombre sí se reemplaza (la clave del almacén es el nombre).
const guardarBaseDB = (items) =>
  operarDB("readwrite", (s) => {
    items.forEach((i) => s.put(i));
  });

/* =========================================================================
   1. EL MOTOR, EN UN HILO APARTE
   -------------------------------------------------------------------------
   Pyodide ya no corre en esta página sino en worker.js. Antes, mientras
   pandas consolidaba, el navegador no podía repintar y Chrome mostraba "La
   página no responde". Ahora la página queda libre y además puede ir
   mostrando en qué paso va el proceso.
   ========================================================================= */
let worker = null;
let motorListo = false;
let pendiente = null;      // la operación que está esperando respuesta

function enviarAlMotor(mensaje) {
  return new Promise((resolve, reject) => {
    if (!worker) return reject(new Error("El motor no está disponible."));
    pendiente = { resolve, reject };
    worker.postMessage(mensaje);
  });
}

/* Muestra en vivo lo que va informando el motor. */
function mostrarProgreso(texto) {
  if (!texto) return;
  const limpio = String(texto).trim();
  if (!limpio) return;

  // La línea corta va arriba, en el indicador del motor
  if (motorListo) $("txt-motor").textContent = limpio.slice(0, 60);

  // Y el detalle completo se va acumulando en el registro
  const reg = $("registro");
  if (reg) {
    reg.classList.remove("oculto");
    reg.textContent += (reg.textContent ? "\n" : "") + limpio;
    reg.scrollTop = reg.scrollHeight;
  }
}

function iniciarMotor() {
  try {
    worker = new Worker("worker.js?v=" + Date.now());
  } catch (e) {
    mostrarFranja("Este navegador no pudo abrir el motor: " + ((e && e.message) || e));
    return;
  }

  worker.onerror = (e) => {
    mostrarFranja("Error en el motor: " + (e.message || "no se pudo cargar worker.js"));
  };

  worker.onmessage = (ev) => {
    const m = ev.data || {};

    if (m.tipo === "progreso") {
      mostrarProgreso(m.texto);

    } else if (m.tipo === "listo") {
      motorListo = true;
      hayPrepararBase = m.hayPrepararBase;
      $("punto-motor").classList.add("listo");
      $("txt-motor").textContent = "Motor listo";
      if ($("registro")) {
        $("registro").textContent = "";
        $("registro").classList.add("oculto");
      }
      pintarBase().catch((e) => {
        $("estado-base").textContent =
          "Este navegador no permite guardar archivos base (" +
          ((e && e.message) || e) + "). Puedes seguir subiéndolos cada vez.";
      });
      revisarSiPuedeProcesar();

    } else if (m.tipo === "error-motor") {
      $("txt-motor").textContent = "No se pudo iniciar el motor";
      if ($("estado-base")) {
        $("estado-base").textContent =
          `El motor no arrancó (falló en: ${m.paso}). Por eso los archivos ` +
          `base no se pueden guardar todavía.`;
      }
      mostrarFranja(`No se pudo iniciar el motor · Falló en: ${m.paso} · ${m.mensaje}`);
      $("resultado").style.display = "block";
      $("resumen").innerHTML = `<b>No se pudo iniciar el motor.</b> Falló en: ${m.paso}.`;
      $("registro").textContent = `Paso: ${m.paso}\n\n${m.detalle}`;
      $("registro").classList.remove("oculto");

    } else if (m.tipo === "error") {
      if (pendiente) { pendiente.reject(new Error(m.mensaje)); pendiente = null; }
      console.error(m.detalle);

    } else if (pendiente) {
      pendiente.resolve(m);
      pendiente = null;
    }
  };

  worker.postMessage({ tipo: "iniciar", base: document.baseURI });
}

/* -------------------------------------------------------------------------
   Período por defecto: el mes anterior.
   La consolidación se hace a mes cerrado, así que lo normal el 3 de
   septiembre es querer agosto. Igual se puede cambiar a mano.
   ------------------------------------------------------------------------- */
(function periodoPorDefecto() {
  const campo = $("periodo");
  if (!campo || campo.value) return;
  const hoy = new Date();
  const anterior = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  const mes = String(anterior.getMonth() + 1).padStart(2, "0");
  campo.value = `${anterior.getFullYear()}-${mes}`;
  // No se puede pedir un mes que todavía no termina
  const maxMes = String(hoy.getMonth() + 1).padStart(2, "0");
  campo.max = `${hoy.getFullYear()}-${maxMes}`;
})();

/* =========================================================================
   2. ELEGIR ZONA
   ========================================================================= */
alEvento("zonas", "click", (ev) => {
  const btn = ev.target.closest(".zona-btn");
  if (!btn) return;
  zonaElegida = btn.dataset.zona;
  document.querySelectorAll(".zona-btn").forEach((b) =>
    b.setAttribute("aria-pressed", b === btn ? "true" : "false")
  );
  ajustarPantallaSegunZona();
  revisarSiPuedeProcesar();
});

/* -------------------------------------------------------------------------
   "Unir zonas" no es una zona más: no consolida, solo apila archivos que la
   app ya generó. Por eso no necesita los archivos base ni el mes, y esos dos
   pasos se esconden para no confundir a quien lo use.
   ------------------------------------------------------------------------- */
function ajustarPantallaSegunZona() {
  const unir = zonaElegida === "UNIR";

  const aviso = $("aviso-unir");
  if (aviso) aviso.classList.toggle("hidden", !unir);

  // paso 02 (archivos base) — la sección que contiene la lista de base
  const pasoBase = $("lista-base") && $("lista-base").closest("section");
  if (pasoBase) pasoBase.classList.toggle("hidden", unir);

  // el selector de mes tampoco aplica: el período lo dice el nombre del archivo
  const periodo = $("periodo");
  if (periodo && periodo.parentElement) {
    periodo.parentElement.classList.toggle("hidden", unir);
  }

  const zonaDrop = $("dropzone");
  if (zonaDrop) {
    zonaDrop.querySelector("span").textContent = unir
      ? "Arrastra aquí los archivos ya generados"
      : "Arrastra los archivos aquí";

    // El paso 03 habla de "movimientos del mes" y de bajar la carpeta de
    // OneDrive. En modo unir eso no aplica: lo que se suelta son archivos que
    // la propia app generó. Se guardan los textos originales la primera vez
    // para poder devolverlos al cambiar de zona.
    const sec = zonaDrop.closest("section");
    if (sec) {
      const num = sec.querySelector("span:not(.dropzone span)");
      const h2 = sec.querySelector("h2");
      const desc = sec.querySelector("p");
      if (h2 && !h2.dataset.original) h2.dataset.original = h2.textContent.trim();
      if (desc && !desc.dataset.original) desc.dataset.original = desc.textContent.trim();
      if (num && !num.dataset.original) num.dataset.original = num.textContent.trim();

      if (h2) h2.textContent = unir ? "Archivos a unir" : h2.dataset.original;
      if (desc) {
        desc.textContent = unir
          ? "Los archivos que ya bajaste de cada zona. Dos o tres, uno por zona."
          : desc.dataset.original;
      }
      if (num) num.textContent = unir ? "02" : num.dataset.original;
    }
  }

  // Si se esconde el paso 02, el 04 pasa a ser el 03
  const secProc = $("btn-procesar") && $("btn-procesar").closest("section");
  if (secProc) {
    const num = secProc.querySelector("span");
    if (num && !num.dataset.original) num.dataset.original = num.textContent.trim();
    if (num) num.textContent = unir ? "03" : num.dataset.original;
    const desc = secProc.querySelector("p");
    if (desc && !desc.dataset.original) desc.dataset.original = desc.textContent.trim();
    if (desc) {
      desc.textContent = unir
        ? "Los archivos se apilan en uno solo, con una columna ZONA al final."
        : desc.dataset.original;
    }
  }

  const btn = $("btn-procesar");
  if (btn && !btn.disabled) {
    btn.textContent = unir ? "Unir archivos" : "Procesar consolidado";
  }
}

/* =========================================================================
   3. ARCHIVOS BASE — se suben una vez y quedan guardados
   ========================================================================= */
function conectarZonaDeCarga(idZona, idInput, alSoltar) {
  const dz = nodo(idZona);
  const input = nodo(idInput);
  if (!dz || !input) return;

  input.addEventListener("change", (e) => {
    alSoltar(e.target.files);
    e.target.value = "";           // permite volver a elegir el mismo archivo
  });
  ["dragover", "dragenter"].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); })
  );
  ["dragleave", "drop"].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); })
  );
  dz.addEventListener("drop", (e) => alSoltar(e.dataTransfer.files));
}

async function recibirBase(fileList) {
  // Acuse de recibo inmediato: si la pantalla no cambia al soltar archivos,
  // significa que este código ni siquiera se está ejecutando, y eso ya es
  // una pista en sí misma.
  const eco = $("estado-base");
  if (eco) eco.textContent = `Recibí ${fileList.length} archivo(s). Revisando…`;
  console.info("[archivos base] recibidos:", [...fileList].map((f) => f.name));

  if (!motorListo) {
    $("estado-base").textContent =
      "El motor todavía se está iniciando. Espera unos segundos y vuelve a intentar.";
    return;
  }

  if (!hayPrepararBase) {
    $("estado-base").textContent =
      "El archivo python/procesar.py del repositorio es una versión anterior " +
      "y todavía no sabe recibir archivos base. Actualízalo y recarga la página.";
    return;
  }

  const recibidos = [...fileList].map((f) => f.name);
  const utiles = [...fileList].filter((f) => esValido(f.name));
  if (!utiles.length) {
    $("estado-base").textContent =
      "Ninguno de esos archivos sirve: tienen que ser .xlsx o .zip. " +
      (recibidos.length ? "Recibí: " + recibidos.join(", ") + "." : "") +
      " (Si son .xls antiguos, ábrelos en Excel y guárdalos como .xlsx.)";
    return;
  }

  $("estado-base").textContent = "Revisando archivos…";
  $("btn-borrar-base").style.display = "inline-block";

  try {
    // Los archivos viajan al worker como bytes
    const archivosBase = [];
    for (const f of utiles) {
      archivosBase.push({ nombre: f.name, bytes: new Uint8Array(await f.arrayBuffer()) });
    }

    const m = await enviarAlMotor({ tipo: "preparar-base", archivos: archivosBase });

    const fecha = new Date().toLocaleDateString("es-CL");
    const items = (m.guardados || []).map((g) => ({
      nombre: g.nombre,
      etiqueta: g.etiqueta,
      fecha,
      bytes: g.bytes,
    }));

    if (items.length) await guardarBaseDB(items);
    console.info("[archivos base]\n" + ((m.res && m.res.log) || ""));

    const res = m.res || {};
    res.enviados = utiles.map((f) => f.name);
    await pintarBase(res);
  } catch (e) {
    $("estado-base").textContent =
      "No se pudieron guardar los archivos base → " + ((e && e.message) || e);
    console.error(e);
  }
}

// `res` es lo que devolvió preparar_base(): guardados, faltantes, ignorados.
// Llega solo cuando la operadora acaba de soltar archivos; al arrancar no.
async function pintarBase(res) {
  baseGuardada = (await leerBaseDB()) || [];
  const ul = $("lista-base");
  ul.innerHTML = "";

  baseGuardada.forEach((item) => {
    const li = document.createElement("li");

    // Se arma con nodos, no con innerHTML: los nombres de archivo son texto
    // de la operadora y no tienen por qué interpretarse como HTML.
    const info = document.createElement("span");
    const etiqueta = document.createElement("b");
    etiqueta.textContent = item.etiqueta;
    info.append(etiqueta, document.createTextNode(" · " + item.nombre));

    const fecha = document.createElement("span");
    fecha.className = "fecha";
    fecha.textContent = "guardado el " + item.fecha;

    const equis = document.createElement("button");
    equis.className = "quitar";
    equis.textContent = "✕";
    equis.style.fontSize = "14px";
    equis.title = "Quitar " + item.etiqueta;
    equis.setAttribute("aria-label", "Quitar " + item.etiqueta);
    equis.onclick = async () => {
      equis.disabled = true;
      try {
        await borrarUnoBaseDB(item.nombre);
        await pintarBase();
      } catch (e) {
        $("estado-base").textContent =
          "No se pudo quitar " + item.nombre + " → " + ((e && e.message) || e);
      }
    };

    const derecha = document.createElement("span");
    derecha.style.cssText = "display:flex;align-items:center;gap:.85rem";
    derecha.append(fecha, equis);

    li.append(info, derecha);
    ul.appendChild(li);
  });

  const total = baseGuardada.length;
  const botonTodos = $("btn-borrar-base");
  botonTodos.style.display = total ? "inline-block" : "none";
  // Con la equis por fila, este botón pasa a ser el "borrar todo"
  botonTodos.textContent = "Borrar los " + total + " archivos base";

  if (res && res.esperados && res.esperados.length) esperadosBase = res.esperados;

  // Qué falta se calcula sobre TODO lo guardado, no sobre este envío: si sube
  // los archivos de a uno, cada uno se suma a los anteriores.
  const guardadas = new Set(baseGuardada.map((i) => i.etiqueta));
  const faltantes = esperadosBase
    .filter((e) => !guardadas.has(e.etiqueta))
    .map((e) => e.nombre);

  const ignorados = (res && res.ignorados) || [];
  const cuantos = esperadosBase.length;
  const esperados = "Los nombres tienen que ser exactamente: " +
    esperadosBase.map((e) => e.nombre).join(" · ");

  if (total >= cuantos) {
    $("estado-base").textContent =
      `Los ${cuantos} archivos base están guardados. No hay que volver a subirlos.`;

  } else if (!total && ignorados.length) {
    // Caso que antes quedaba mudo: sí llegaron archivos, pero ninguno
    // coincide con los cuatro maestros. Hay que decir cuáles llegaron.
    $("estado-base").textContent =
      `No reconocí ninguno de los ${ignorados.length} archivo(s) que soltaste: ` +
      ignorados.join(", ") + ". " + esperados + ".";

  } else if (!total && res && res.enviados && res.enviados.length) {
    // Llegaron archivos, pero no salió ningún .xlsx utilizable de ellos
    // (por ejemplo, un .zip sin Excel adentro).
    $("estado-base").textContent =
      `Recibí ${res.enviados.join(", ")}, pero no encontré ningún Excel ` +
      `reconocible adentro. ` + esperados + ".";

  } else if (!total) {
    $("estado-base").textContent =
      "Todavía no hay archivos base guardados en este computador.";

  } else {
    let msg = `Hay ${total} de ${cuantos} archivos base guardados.`;
    if (faltantes.length) msg += ` Faltan: ${faltantes.join(", ")}.`;
    if (ignorados.length) msg += ` No reconocí: ${ignorados.join(", ")}.`;
    $("estado-base").textContent = msg;
  }
}

alEvento("btn-borrar-base", "click", async () => {
  await borrarBaseDB();
  await pintarBase();
});

conectarZonaDeCarga("dropzone-base", "input-base", recibirBase);

/* =========================================================================
   4. MOVIMIENTOS DEL MES
   ========================================================================= */
function agregarArchivos(fileList) {
  for (const f of fileList) {
    if (esValido(f.name) && !archivos.some((a) => a.name === f.name)) {
      archivos.push(f);
    }
  }
  pintarLista();
  revisarSiPuedeProcesar();
}

function pintarLista() {
  const ul = $("lista-archivos");
  ul.innerHTML = "";
  archivos.forEach((f, i) => {
    const li = document.createElement("li");
    const esZip = f.name.toLowerCase().endsWith(".zip");
    li.innerHTML = `<span>${esZip ? "🗂 " : ""}${f.name}</span>`;
    const b = document.createElement("button");
    b.className = "quitar";
    b.textContent = "quitar";
    b.onclick = () => { archivos.splice(i, 1); pintarLista(); revisarSiPuedeProcesar(); };
    li.appendChild(b);
    ul.appendChild(li);
  });
}

conectarZonaDeCarga("dropzone", "input-archivos", agregarArchivos);

function revisarSiPuedeProcesar() {
  $("btn-procesar").disabled = !(motorListo && zonaElegida && archivos.length > 0);
}

/* =========================================================================
   5. PROCESAR
   ========================================================================= */
alEvento("btn-procesar", "click", async () => {
  const btn = $("btn-procesar");
  btn.disabled = true;
  btn.innerHTML = '<span class="spin" style="display:inline-block;vertical-align:middle"></span> Procesando…';

  // El registro pasa a ser el avance en vivo mientras dura el proceso
  $("resultado").style.display = "block";
  $("resumen").innerHTML =
    "<b>Procesando…</b> Esto puede tardar varios minutos. " +
    "Puedes seguir usando la página: abajo se va viendo el avance.";
  $("alertas").innerHTML = "";
  $("descargas").innerHTML = "";
  $("registro").textContent = "";
  $("registro").classList.remove("oculto");

  try {
    const movimientos = [];
    for (const f of archivos) {
      movimientos.push({ nombre: f.name, bytes: new Uint8Array(await f.arrayBuffer()) });
    }

    const m = await enviarAlMotor({
      tipo: "procesar",
      zona: zonaElegida,
      periodo: ($("periodo") && $("periodo").value) || null,
      // Unir no usa los archivos base: si se mandaran, entrarían a la carpeta
      // de trabajo y la herramienta trataría de apilarlos.
      base: zonaElegida === "UNIR"
        ? []
        : baseGuardada.map((i) => ({ nombre: i.nombre, bytes: i.bytes })),
      movimientos,
    });

    mostrarResultado(m.r || {}, m.descargas || []);
  } catch (e) {
    $("resumen").innerHTML =
      "<b>Ocurrió un error al procesar.</b> El detalle está abajo.";
    mostrarProgreso("ERROR: " + ((e && e.message) || e));
    console.error(e);
  } finally {
    btn.disabled = false;
    btn.textContent = zonaElegida === "UNIR" ? "Unir archivos" : "Procesar consolidado";
    $("txt-motor").textContent = "Motor listo";
    revisarSiPuedeProcesar();
  }
});

/* =========================================================================
   6. MOSTRAR RESULTADO Y PREPARAR LAS DESCARGAS
   ========================================================================= */
function mostrarResultado(r, descargas) {
  $("resultado").style.display = "block";
  $("resumen").innerHTML = r.resumen || "Proceso terminado.";

  // Tabla de fuentes leídas
  const tb = $("tabla-fuentes").querySelector("tbody");
  tb.innerHTML = "";
  (r.fuentes || []).forEach((f) => {
    const tr = document.createElement("tr");
    [f.archivo, f.hojas, f.filas].forEach((valor) => {
      const td = document.createElement("td");
      td.textContent = valor;
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });

  // Alertas del control de calidad
  const cont = $("alertas");
  cont.innerHTML = "";
  (r.alertas || []).forEach((a) => {
    const div = document.createElement("div");
    div.className = "alerta";
    const t = document.createElement("b");
    t.textContent = a.titulo;
    div.append(t, document.createTextNode(" — " + a.detalle));
    cont.appendChild(div);
  });

  $("registro").textContent = r.log || "";
  $("registro").classList.add("oculto");

  // Un botón por cada archivo que el motor dejó listo
  const zona = $("descargas");
  zona.innerHTML = "";
  (descargas || []).forEach((d, i) => {
    const b = document.createElement("button");
    b.className = i === 0 ? "btn" : "btn secundario";
    b.textContent = d.etiqueta;
    b.onclick = () => descargar(d.bytes, d.archivo);
    zona.appendChild(b);
  });

  const verLog = document.createElement("button");
  verLog.className = "btn secundario";
  verLog.textContent = "Ver detalle técnico";
  verLog.onclick = () => $("registro").classList.toggle("oculto");
  zona.appendChild(verLog);
}

function descargar(bytes, nombre) {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  a.click();
  URL.revokeObjectURL(url);
}

if (faltanElementos.length) avisarDesajuste();

console.info("app.js — Sistema Trazabilidad Ambipar (versión con archivos base)");
iniciarMotor();
