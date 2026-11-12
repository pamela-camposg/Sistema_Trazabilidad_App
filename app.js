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

let zonaElegida = null;      // "RM" | "SUR" | "NORTE" | "UNIR" | "SIMPLE_ROUTE"
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
// Versión 2: se suma el almacén de la carpeta de OneDrive conectada. Lo que
// ya estaba guardado (los archivos base) no se toca al actualizar.
const DB_CARPETAS = "carpetas";

function abrirDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NOMBRE, 2);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(DB_STORE)) {
        db.createObjectStore(DB_STORE, { keyPath: "nombre" });
      }
      if (!db.objectStoreNames.contains(DB_CARPETAS)) {
        db.createObjectStore(DB_CARPETAS, { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function operarDB(modo, accion, almacen = DB_STORE) {
  return abrirDB().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(almacen, modo);
        const store = tx.objectStore(almacen);
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

// La carpeta conectada se guarda como "permiso para volver a abrirla", no como
// copia de los archivos: cada vez se leen las versiones del día.
const leerCarpetasDB = () => operarDB("readonly", (s) => s.getAll(), DB_CARPETAS);
const guardarCarpetaDB = (id, handle) =>
  operarDB("readwrite", (s) => { s.put({ id, handle }); }, DB_CARPETAS);
const borrarCarpetasDB = () => operarDB("readwrite", (s) => s.clear(), DB_CARPETAS);

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
let esperadosPendiente = null;  // la consulta de "qué archivos busca esta zona"

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
      pintarCarpeta();
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

    } else if (m.tipo === "esperados") {
      // Va por un canal propio para no cruzarse con un proceso en curso
      if (esperadosPendiente) { esperadosPendiente.resolve(m.res); esperadosPendiente = null; }

    } else if (m.tipo === "error") {
      if (m.origen === "esperados") {
        if (esperadosPendiente) { esperadosPendiente.reject(new Error(m.mensaje)); esperadosPendiente = null; }
      } else if (pendiente) { pendiente.reject(new Error(m.mensaje)); pendiente = null; }
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
  const sr = zonaElegida === "SIMPLE_ROUTE";
  // Unir y Simple Route comparten la misma forma: ninguna consolida, ninguna
  // necesita los archivos base ni el mes. Cambia el texto, no la mecánica.
  const unir = zonaElegida === "UNIR" || sr;

  const aviso = $("aviso-unir");
  if (aviso) aviso.classList.toggle("hidden", zonaElegida !== "UNIR");
  const avisoSr = $("aviso-sr");
  if (avisoSr) avisoSr.classList.toggle("hidden", !sr);

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
    zonaDrop.querySelector("span").textContent = sr
      ? "Arrastra aquí Simple Route y la BO"
      : unir
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

      if (h2) {
        h2.textContent = sr
          ? "Archivos a comparar"
          : unir
            ? "Archivos a unir"
            : h2.dataset.original;
      }
      if (desc) {
        desc.textContent = sr
          ? "La exportación de Simple Route y BBDD BO SAN BERNARDO.xlsx."
          : unir
            ? "Los archivos que ya bajaste de cada zona: uno por zona y mes. Para un reproceso puedes soltar varios meses juntos."
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
      desc.textContent = sr
      ? "Se compara servicio por servicio y se entrega el informe con las diferencias."
      : unir
        ? "Los archivos se apilan en uno solo, con una columna ZONA al final (y PERIODO si son varios meses)."
        : desc.dataset.original;
    }
  }

  // Unir trabaja con archivos que la app ya generó: la carpeta no aplica
  const bloque = $("bloque-carpeta");
  if (bloque) bloque.style.display = zonaElegida === "UNIR" ? "none" : "";
  pintarCarpeta();
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

    if (items.length) {
      // Uno por tipo: si "BBDD_DESTINATARIO.xlsx" llega y había guardado un
      // "BBDD DESTINATARIO.xlsx", el viejo se borra (si no, quedaban los dos).
      const tipos = new Set(items.map((i) => i.etiqueta));
      const viejos = ((await leerBaseDB()) || []).filter(
        (v) => tipos.has(v.etiqueta) && !items.some((i) => i.nombre === v.nombre)
      );
      for (const v of viejos) await borrarUnoBaseDB(v.nombre);
      await guardarBaseDB(items);
    }
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
    $("estado-base").textContent = carpetaPrincipal
      ? "No hay archivos base guardados en este computador, pero no hacen falta: " +
        "se leen de la carpeta conectada."
      : "Todavía no hay archivos base guardados en este computador.";

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
   3b. CARPETA DE ONEDRIVE CONECTADA
   -------------------------------------------------------------------------
   En vez de descargar los Excel desde SharePoint y soltarlos acá, la
   operadora conecta UNA vez su carpeta de OneDrive sincronizada (la misma que
   usa el consolidador en Visual Studio). Desde ahí la página lee directo:

     · los archivos de la zona  → General/BBDD/ZONA CENTRO | ZONA SUR | ZONA NORTE
     · los archivos base        → General/Información Base
     · GIRI (solo RM)           → Trazabilidad GIRI - Operaciones Actualizado

   Sirve conectar cualquiera de estas carpetas: "OneDrive - Ambipar Group"
   (alcanza todo, incluido GIRI), "Inteligencia Circular - 7. Trazabilidad",
   "General", "BBDD" o la carpeta de la zona.

   Los archivos NO se suben a ningún lado: se leen del disco del computador y
   pasan directo al motor, igual que si se hubieran arrastrado. Usa la API de
   carpetas del navegador, que existe en Chrome y Edge (no en Firefox/Safari).
   ========================================================================= */
const soportaCarpetas = typeof window.showDirectoryPicker === "function";
let carpetaPrincipal = null;   // carpeta elegida por la operadora
let carpetaGiri = null;        // opcional: la carpeta de GIRI, si no cuelga de la principal
let revisionCarpeta = null;    // lo último que se encontró, para mostrarlo
const cacheEsperados = {};

const ZONAS_SOLO_CARPETA = ["RM", "SUR", "NORTE"];
const CARPETA_DE_ZONA = { RM: "ZONA CENTRO", SUR: "ZONA SUR", NORTE: "ZONA NORTE", SIMPLE_ROUTE: "ZONA CENTRO" };
const NOMBRE_TRZ = "Inteligencia Circular - 7. Trazabilidad";
const NOMBRE_GIRI = "Trazabilidad GIRI - Operaciones Actualizado";

/* Igual que _clave() en Python: sin tildes, mayúsculas, sin espacios ni
   . _ - ( ). Así "BBDD_BO_CONCEPCION.xlsx" y "BBDD BO CONCEPCIÓN.xlsx" calzan. */
const clave = (s) =>
  String(s).normalize("NFKD").replace(/[^\x00-\x7F]/g, "").toUpperCase().replace(/[ ._\-()]/g, "");

async function cargarCarpetas() {
  if (!soportaCarpetas) return pintarCarpeta();
  try {
    for (const c of (await leerCarpetasDB()) || []) {
      if (c.id === "principal") carpetaPrincipal = c.handle;
      if (c.id === "giri") carpetaGiri = c.handle;
    }
  } catch (e) {
    console.warn("No se pudo leer la carpeta guardada:", e);
  }
  pintarCarpeta();
  revisarSiPuedeProcesar();
}

/* El navegador recuerda la carpeta, pero en cada visita pide confirmar el
   permiso de lectura. Solo se puede pedir desde un clic de la operadora. */
async function tienePermiso(handle, pedir) {
  if (!handle) return false;
  const opts = { mode: "read" };
  try {
    if ((await handle.queryPermission(opts)) === "granted") return true;
    if (!pedir) return false;
    return (await handle.requestPermission(opts)) === "granted";
  } catch (e) {
    return false;
  }
}

async function subcarpeta(dir, nombre) {
  if (!dir) return null;
  const c = clave(nombre);
  for await (const [n, h] of dir.entries()) {
    if (h.kind === "directory" && clave(n) === c) return h;
  }
  return null;
}

async function bajarPor(dir, partes) {
  let d = dir;
  for (const p of partes) {
    d = await subcarpeta(d, p);
    if (!d) return null;
  }
  return d;
}

/* Encuentra la carpeta de la zona, la de archivos base y la de GIRI, partiendo
   de la carpeta que haya elegido la operadora. */
/* -------------------------------------------------------------------------
   Dónde está "General" depende de cómo cada persona sincronizó SharePoint:
     · con "Agregar acceso directo":  OneDrive - Ambipar Group\Inteligencia
       Circular - 7. Trazabilidad\General
     · con "Sincronizar":             Ambipar Group\Inteligencia Circular -
       Documentos Compartilhados\7. Trazabilidad\General  (u otra variante)
   Por eso no se asume una ruta fija: se busca, unos pocos niveles hacia
   abajo, la carpeta que tenga adentro "BBDD" e "Información Base". De paso
   se busca la carpeta que tenga el archivo de GIRI. Se hace una vez por
   carpeta conectada y se recuerda.
   ------------------------------------------------------------------------- */
const MAX_NIVELES = 4;
const MAX_CARPETAS = 600;      // tope para no recorrer todo OneDrive
let ubicacionCache = null;     // { raiz, general, giriDir }
const CLAVE_BBDD = clave("BBDD");
const CLAVE_INFO = clave("Información Base");
const CLAVE_GIRI_ARCHIVO = clave("OPERACIONES GIRI (INGRESOS Y EGRESOS).xlsx");

async function explorar(raiz) {
  if (ubicacionCache && ubicacionCache.raiz === raiz) return ubicacionCache;
  let general = null;
  let giriDir = null;
  let vistas = 0;
  let nivel = [raiz];
  for (let n = 0; n <= MAX_NIVELES && nivel.length && !(general && giriDir); n++) {
    const siguiente = [];
    for (const dir of nivel) {
      if (++vistas > MAX_CARPETAS) break;
      const sub = [];
      let tieneBBDD = false, tieneInfo = false;
      for await (const [nombre, h] of dir.entries()) {
        const c = clave(nombre);
        if (h.kind === "directory") {
          sub.push(h);
          if (c === CLAVE_BBDD) tieneBBDD = true;
          if (c === CLAVE_INFO) tieneInfo = true;
        } else if (!giriDir && c === CLAVE_GIRI_ARCHIVO) {
          giriDir = dir;
        }
      }
      if (!general && tieneBBDD && tieneInfo) general = dir;
      // Dentro de "General" no hace falta seguir bajando
      if (dir !== general) siguiente.push(...sub);
    }
    nivel = siguiente;
  }
  ubicacionCache = { raiz, general, giriDir };
  return ubicacionCache;
}

async function ubicarCarpetas(raiz, zona, giriOk) {
  const nombreZona = CARPETA_DE_ZONA[zona];
  const { general, giriDir: giriEncontrado } = await explorar(raiz);

  let zonaDir = general ? await bajarPor(general, ["BBDD", nombreZona]) : null;
  if (!zonaDir) zonaDir = await subcarpeta(raiz, nombreZona);              // eligió "BBDD"
  if (!zonaDir && clave(raiz.name) === clave(nombreZona)) zonaDir = raiz;  // eligió la zona

  let infoDir = general ? await subcarpeta(general, "Información Base") : null;
  if (!infoDir && clave(raiz.name) === CLAVE_INFO) infoDir = raiz;

  let giriDir = null;
  if (zona === "RM") giriDir = (carpetaGiri && giriOk) ? carpetaGiri : giriEncontrado;
  return { nombreZona, zonaDir, infoDir, giriDir };
}

async function excelDe(dir) {
  const out = [];
  if (!dir) return out;
  for await (const [n, h] of dir.entries()) {
    if (h.kind === "file" && n.toLowerCase().endsWith(".xlsx") && !n.startsWith("~$")) out.push([n, h]);
  }
  return out;
}

/* Las consultas van en fila: si la operadora cambia rápido de zona, una no
   pisa a la otra. Cada zona se consulta una sola vez por visita. */
let colaEsperados = Promise.resolve();
function pedirEsperados(zona) {
  if (!cacheEsperados[zona]) {
    const p = colaEsperados.catch(() => {}).then(() => new Promise((resolve, reject) => {
      esperadosPendiente = { resolve, reject };
      worker.postMessage({ tipo: "esperados", zona });
    }));
    colaEsperados = p;
    cacheEsperados[zona] = p;
    p.catch(() => { delete cacheEsperados[zona]; });
  }
  return cacheEsperados[zona];
}

/* Busca en la carpeta los archivos que necesita la zona. No lee su contenido
   todavía: solo dice qué encontró, dónde, y qué falta. */
async function buscarEnCarpeta(zona, pedirPermiso) {
  if (!carpetaPrincipal || !CARPETA_DE_ZONA[zona]) return null;
  if (!(await tienePermiso(carpetaPrincipal, pedirPermiso))) return { sinPermiso: true };
  const giriOk = zona === "RM" && carpetaGiri ? await tienePermiso(carpetaGiri, pedirPermiso) : false;

  const esp = await pedirEsperados(zona);
  if (!esp) return { procesarAntiguo: true };

  const ub = await ubicarCarpetas(carpetaPrincipal, zona, giriOk);

  // Orden de búsqueda: archivos base desde Información Base; el resto desde la
  // carpeta de la zona; GIRI desde su carpeta. El primero que aparece gana.
  const lugares = [["Información Base", ub.infoDir], [ub.nombreZona, ub.zonaDir]];
  if (ub.giriDir) lugares.push(["GIRI", ub.giriDir]);

  const porClave = new Map();
  for (const [origen, dir] of lugares) {
    for (const [n, h] of await excelDe(dir)) {
      const c = clave(n);
      if (!porClave.has(c)) porClave.set(c, { nombre: n, handle: h, origen });
    }
  }

  const guardadosBase = new Set(baseGuardada.map((i) => clave(i.nombre)));
  const encontrados = [];
  const faltan = [];
  for (const g of esp.grupos) {
    const claves = g.nombres.map(clave);
    const hit = claves.map((c) => porClave.get(c)).find(Boolean);
    if (hit) {
      encontrados.push({ ...hit, maestro: g.maestro });
    } else if (g.maestro && claves.some((c) => guardadosBase.has(c))) {
      // No está en la carpeta, pero hay uno guardado en el navegador: sirve.
      encontrados.push({ nombre: g.nombres[0], origen: "guardado en este computador", guardado: true, maestro: true });
    } else if (!g.opcional) {
      faltan.push(g.nombres[0]);
    }
  }
  return { ub, encontrados, faltan };
}

function pintarRevision(rv) {
  const ul = $("lista-carpeta");
  if (!ul) return;
  ul.innerHTML = "";
  if (!rv || rv.sinPermiso || rv.procesarAntiguo) return;
  const fila = (texto, derecha, clase) => {
    const li = document.createElement("li");
    if (clase) li.className = clase;
    const a = document.createElement("span");
    a.textContent = texto;
    const b = document.createElement("span");
    b.className = "origen";
    b.textContent = derecha;
    li.append(a, b);
    ul.appendChild(li);
  };
  rv.encontrados.forEach((e) => fila("✓ " + e.nombre, e.origen));
  rv.faltan.forEach((n) => fila("✗ " + n, "no está en la carpeta", "falta"));
}

/* Escribe en pantalla el estado de la carpeta y, si ya hay permiso, revisa
   qué archivos de la zona encuentra. */
async function pintarCarpeta() {
  pintarBaseCarpeta();   // el paso 02 también muestra lo de la carpeta
  const estado = $("estado-carpeta");
  if (!estado) return;
  const btnConectar = $("btn-conectar");
  const btnRevisar = $("btn-revisar-carpeta");
  const btnGiri = $("btn-conectar-giri");
  const btnDesc = $("btn-desconectar");

  if (!soportaCarpetas) {
    estado.textContent =
      "Este navegador no permite conectar carpetas. Abre la herramienta en Chrome o " +
      "Edge, o sigue arrastrando los archivos abajo.";
    [btnConectar, btnRevisar, btnGiri, btnDesc].forEach((b) => b && (b.style.display = "none"));
    return;
  }

  btnConectar.textContent = carpetaPrincipal ? "Cambiar carpeta" : "Conectar carpeta de OneDrive";
  btnRevisar.style.display = carpetaPrincipal ? "" : "none";
  btnDesc.style.display = carpetaPrincipal ? "" : "none";
  btnGiri.style.display = carpetaPrincipal && zonaElegida === "RM" ? "" : "none";
  if (carpetaGiri) btnGiri.textContent = "Cambiar carpeta GIRI";

  if (!carpetaPrincipal) {
    estado.textContent =
      "Conecta una sola vez la carpeta \"OneDrive - Ambipar Group\" (o la carpeta " +
      "\"General\" de Trazabilidad) y la herramienta leerá los archivos sola, sin descargarlos.";
    pintarRevision(null);
    return;
  }

  let texto = `Carpeta conectada: ${carpetaPrincipal.name}` +
    (carpetaGiri ? ` · GIRI: ${carpetaGiri.name}` : "") + ".";

  if (!zonaElegida || !CARPETA_DE_ZONA[zonaElegida]) {
    estado.textContent = texto + " Elige una zona para ver qué archivos se van a leer.";
    pintarRevision(null);
    return;
  }
  if (!motorListo || procesando) {
    estado.textContent = texto;
    return;
  }

  const zonaRevisada = zonaElegida;
  try {
    const rv = await buscarEnCarpeta(zonaRevisada, false);
    if (zonaElegida !== zonaRevisada || procesando) return;   // ya cambió de zona
    revisionCarpeta = rv;
    if (rv && rv.sinPermiso) {
      texto += " El navegador necesita tu permiso para leerla: haz clic en \"Revisar " +
               "archivos\" (o en Procesar) y elige \"Permitir\".";
    } else if (rv && rv.procesarAntiguo) {
      texto += " El archivo python/procesar.py publicado es una versión anterior: " +
               "actualízalo para poder leer la carpeta.";
    } else if (rv) {
      if (!rv.ub.zonaDir) {
        texto += ` No encontré la carpeta ${rv.ub.nombreZona} dentro de la carpeta conectada. ` +
                 `Conecta "OneDrive - Ambipar Group" o la carpeta "General" de Trazabilidad.`;
      } else {
        texto += ` Se van a leer ${rv.encontrados.length} archivo(s).`;
        if (rv.faltan.length) texto += ` Faltan ${rv.faltan.length}: puedes arrastrarlos abajo.`;
        if (!rv.ub.infoDir) {
          texto += ` Los archivos base no se leen de "Información Base" porque esa carpeta ` +
                   `no cuelga de la conectada; para que se lean también, conecta la carpeta "General".`;
        }
      }
    }
    estado.textContent = texto;
    // Si ni siquiera está la carpeta de la zona, la lista de faltantes solo confunde
    pintarRevision(rv && rv.ub && !rv.ub.zonaDir ? null : rv);
  } catch (e) {
    estado.textContent = texto + " No se pudo revisar la carpeta: " + ((e && e.message) || e);
  }
  revisarSiPuedeProcesar();
}

/* Lee el contenido de lo encontrado. Se llama al procesar. */
async function leerDesdeCarpeta(zona) {
  const rv = await buscarEnCarpeta(zona, true);
  if (!rv) return [];
  if (rv.sinPermiso) {
    throw new Error("No se dio permiso para leer la carpeta conectada. Vuelve a " +
                    "hacer clic en Procesar y elige \"Permitir\".");
  }
  if (rv.procesarAntiguo) return [];
  revisionCarpeta = rv;
  pintarRevision(rv);

  const leidos = [];
  for (const e of rv.encontrados) {
    if (e.guardado) continue;       // lo pone la lista de archivos base guardados
    try {
      const f = await e.handle.getFile();
      leidos.push({ nombre: e.nombre, bytes: new Uint8Array(await f.arrayBuffer()) });
      mostrarProgreso(`  ✓ leído de ${e.origen}: ${e.nombre}`);
    } catch (err) {
      throw new Error(
        `No se pudo leer ${e.nombre} desde la carpeta (${(err && err.message) || err}). ` +
        `Si OneDrive lo tiene "solo en línea", ábrelo o márcalo como "Mantener siempre ` +
        `en este dispositivo" y vuelve a procesar.`
      );
    }
  }
  return leidos;
}

/* Mensaje de error de conexión, visible en los dos pasos (02 y 03). Antes,
   si la persona apretaba el botón del paso 02 y fallaba, el aviso salía solo
   en el paso 03 y parecía que no pasaba nada. */
function avisarCarpeta(texto) {
  ["estado-carpeta", "estado-base-carpeta"].forEach((id) => { if ($(id)) $(id).textContent = texto; });
}

function explicarError(e) {
  const nombre = (e && e.name) || "";
  if (nombre === "SecurityError" || nombre === "NotAllowedError") {
    return "El navegador no dejó usar esa carpeta (a veces Chrome bloquea carpetas " +
           "principales como \"OneDrive - Ambipar Group\", o la configuración del computador " +
           "no lo permite). Prueba eligiendo directamente la carpeta \"General\" de Trazabilidad, " +
           "la que tiene adentro \"BBDD\" e \"Información Base\".";
  }
  return (e && e.message) || String(e);
}

let carpetaSinRecordar = false;   // se conectó, pero el navegador no dejó guardarla

async function conectarCarpeta() {
  let h;
  try {
    h = await window.showDirectoryPicker({ id: "trazabilidad", mode: "read" });
  } catch (e) {
    if (e && e.name === "AbortError") return;   // cerró el diálogo
    avisarCarpeta("No se pudo conectar la carpeta. " + explicarError(e));
    return;
  }
  carpetaPrincipal = h;
  ubicacionCache = null;
  // Recordarla es un extra: si el navegador no deja guardarla (ventana de
  // incógnito, configuración del equipo), igual se usa en esta visita.
  try {
    await guardarCarpetaDB("principal", h);
    carpetaSinRecordar = false;
  } catch (e) {
    carpetaSinRecordar = true;
    console.warn("No se pudo recordar la carpeta:", e);
  }
  avisarCarpeta("Buscando los archivos en la carpeta…");
  pintarBase().catch(() => {});
  await pintarCarpeta();
  revisarSiPuedeProcesar();
}

// El botón está en el paso 02 y en el paso 03: es una sola conexión que
// sirve para los dos. En el paso 02 también sirve para dar el permiso.
alEvento("btn-conectar", "click", conectarCarpeta);
alEvento("btn-conectar-base", "click", async () => {
  if (!carpetaPrincipal) return conectarCarpeta();
  await tienePermiso(carpetaPrincipal, true);   // este clic permite pedir el permiso
  pintarCarpeta();
});

/* -------------------------------------------------------------------------
   ARCHIVOS BASE Y CARPETA (paso 02)
   Al procesar, los archivos base se leen de "Información Base" (ver
   buscarEnCarpeta). Acá solo se muestra si la carpeta está conectada o no.
   La opción de subirlos a mano sigue igual: los guardados se usan cuando no
   hay carpeta o cuando a la carpeta le falta alguno.
   ------------------------------------------------------------------------- */
async function ubicarInfoBase(raiz) {
  if (clave(raiz.name) === CLAVE_INFO) return raiz;
  const { general } = await explorar(raiz);
  return general ? await subcarpeta(general, "Información Base") : null;
}

async function maestrosEnCarpeta(pedirPermiso) {
  if (!carpetaPrincipal) return null;
  if (!(await tienePermiso(carpetaPrincipal, pedirPermiso))) return { sinPermiso: true };
  // Los nombres de los archivos base salen de Python, igual que los de las zonas
  const esp = await pedirEsperados("SUR");
  if (!esp) return { procesarAntiguo: true };
  const info = await ubicarInfoBase(carpetaPrincipal);
  const enInfo = await excelDe(info);
  const encontrados = [];
  const faltan = [];
  for (const g of esp.grupos.filter((x) => x.maestro)) {
    const claves = g.nombres.map(clave);
    const hit = enInfo.find(([n]) => claves.includes(clave(n)));
    if (hit) encontrados.push({ nombre: hit[0], handle: hit[1] });
    else faltan.push(g.nombres[0]);
  }
  return { info, encontrados, faltan };
}

async function pintarBaseCarpeta() {
  const estado = $("estado-base-carpeta");
  const bloque = $("base-carpeta");
  if (!estado || !bloque) return;
  const btn = $("btn-conectar-base");
  const mostrarBoton = (texto) => {
    btn.style.display = texto ? "" : "none";
    if (texto) btn.textContent = texto;
  };

  if (!soportaCarpetas) {
    estado.textContent = "✗ No conectada: este navegador no permite conectar carpetas (usa " +
      "Chrome o Edge). Se usan los archivos base guardados, o súbelos abajo.";
    return mostrarBoton(null);
  }
  if (!carpetaPrincipal) {
    estado.textContent = "✗ No conectada. Se usan los archivos base guardados en este " +
      "computador, o súbelos abajo.";
    return mostrarBoton("Conectar carpeta de OneDrive");
  }
  if (!motorListo || procesando) {
    estado.textContent = `✓ Conectada a "${carpetaPrincipal.name}".`;
    return mostrarBoton(null);
  }

  try {
    const r = await maestrosEnCarpeta(false);
    if (r && r.sinPermiso) {
      estado.textContent = `Conectada a "${carpetaPrincipal.name}", pero falta tu permiso para ` +
        "leerla en esta visita.";
      return mostrarBoton("Dar permiso");
    }
    mostrarBoton(null);
    if (r && r.procesarAntiguo) {
      estado.textContent = "Para leer los archivos base desde la carpeta hay que actualizar " +
        "python/procesar.py.";
    } else if (r && !r.info) {
      estado.textContent = `Conectada a "${carpetaPrincipal.name}", pero ahí no encontré ` +
        "\"Información Base\". Se usan los archivos base guardados. Para leerlos de la carpeta, " +
        "conecta la carpeta \"General\" de Trazabilidad (paso 03, \"Cambiar carpeta\").";
    } else if (r) {
      let texto = `✓ Conectada a "${carpetaPrincipal.name}". Los archivos base se leen de ` +
        `"Información Base" (${r.encontrados.length} de ${r.encontrados.length + r.faltan.length}).`;
      if (r.faltan.length) texto += ` Para ${r.faltan.join(", ")} se usa el guardado aquí.`;
      if (carpetaSinRecordar) {
        texto += " Este navegador no dejó recordar la carpeta: la próxima vez habrá que volver a conectarla.";
      }
      estado.textContent = texto;
    }
  } catch (e) {
    estado.textContent = "No se pudo revisar la carpeta: " + explicarError(e);
  }
}

alEvento("btn-conectar-giri", "click", async () => {
  try {
    const h = await window.showDirectoryPicker({ id: "trazabilidad-giri", mode: "read" });
    carpetaGiri = h;
    await guardarCarpetaDB("giri", h);
  } catch (e) {
    if (e && e.name === "AbortError") return;
    avisarCarpeta("No se pudo conectar la carpeta GIRI. " + explicarError(e));
    return;
  }
  pintarCarpeta();
});

alEvento("btn-revisar-carpeta", "click", async () => {
  // Este clic es el que permite pedir el permiso de lectura
  ubicacionCache = null;   // vuelve a buscar, por si se movieron carpetas
  await tienePermiso(carpetaPrincipal, true);
  if (carpetaGiri && zonaElegida === "RM") await tienePermiso(carpetaGiri, true);
  pintarCarpeta();
});

alEvento("btn-desconectar", "click", async () => {
  carpetaPrincipal = null;
  carpetaGiri = null;
  revisionCarpeta = null;
  try { await borrarCarpetasDB(); } catch (e) { /* no importa */ }
  pintarBase().catch(() => {});
  pintarCarpeta();
  revisarSiPuedeProcesar();
});

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

/* El texto del botón depende de la zona, no de si está habilitado. Antes se
   escribía solo cuando el botón estaba activo, así que al elegir Unir o Simple
   Route sin archivos cargados seguía diciendo "Procesar consolidado". */
let procesando = false;

function etiquetaBoton() {
  if (zonaElegida === "SIMPLE_ROUTE") return "Comparar";
  if (zonaElegida === "UNIR") return "Unir archivos";
  return "Procesar consolidado";
}

function revisarSiPuedeProcesar() {
  const btn = $("btn-procesar");
  if (!btn) return;
  // Con la carpeta conectada no hace falta soltar nada en RM, Sur ni Norte.
  // Simple Route igual necesita la exportación de Simple Route, que no vive
  // en OneDrive; Unir trabaja con archivos que la app ya generó.
  const bastaCarpeta = !!carpetaPrincipal && ZONAS_SOLO_CARPETA.includes(zonaElegida);
  btn.disabled = !(motorListo && zonaElegida && !procesando && (archivos.length > 0 || bastaCarpeta));
  if (!procesando) btn.textContent = etiquetaBoton();
}

/* =========================================================================
   5. PROCESAR
   ========================================================================= */
alEvento("btn-procesar", "click", async () => {
  const btn = $("btn-procesar");
  procesando = true;
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
    // 1. Carpeta conectada. Va primero porque el permiso de lectura solo se
    //    puede pedir justo después del clic.
    let deCarpeta = [];
    if (carpetaPrincipal && CARPETA_DE_ZONA[zonaElegida]) {
      mostrarProgreso("Leyendo los archivos desde la carpeta conectada…");
      deCarpeta = await leerDesdeCarpeta(zonaElegida);
    }

    // 2. Lo que se arrastró a mano manda: si un archivo está en los dos
    //    lados, se usa el arrastrado. Se compara por nombre normalizado, así
    //    "BBDD_DESTINATARIO" y "BBDD DESTINATARIO" cuentan como el mismo.
    const usadas = new Set();
    const movimientos = [];
    for (const f of archivos) {
      movimientos.push({ nombre: f.name, bytes: new Uint8Array(await f.arrayBuffer()) });
      usadas.add(clave(f.name));
    }
    for (const a of deCarpeta) {
      if (!usadas.has(clave(a.nombre))) {
        movimientos.push(a);
        usadas.add(clave(a.nombre));
      }
    }

    // 3. Archivos base guardados en el navegador: solo los que no llegaron
    //    ya desde la carpeta (Información Base trae la versión del día).
    //    Unir y Simple Route no los usan: si se mandaran, entrarían a la
    //    carpeta de trabajo y la herramienta trataría de apilarlos.
    const base = (zonaElegida === "UNIR" || zonaElegida === "SIMPLE_ROUTE")
      ? []
      : baseGuardada
          .filter((i) => !usadas.has(clave(i.nombre)))
          .map((i) => ({ nombre: i.nombre, bytes: i.bytes }));

    const m = await enviarAlMotor({
      tipo: "procesar",
      zona: zonaElegida,
      periodo: ($("periodo") && $("periodo").value) || null,
      base,
      movimientos,
    });

    mostrarResultado(m.r || {}, m.descargas || []);
  } catch (e) {
    $("resumen").innerHTML =
      "<b>Ocurrió un error al procesar.</b> El detalle está abajo.";
    mostrarProgreso("ERROR: " + ((e && e.message) || e));
    console.error(e);
  } finally {
    procesando = false;
    btn.disabled = false;
    btn.textContent = etiquetaBoton();
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

console.info("app.js — Sistema Trazabilidad Ambipar (versión con carpeta de OneDrive)");
// Salida de emergencia: abrir la página con ?sin-carpeta olvida la carpeta
// conectada sin intentar abrirla (por si algún navegador falla al recuperarla).
if (new URLSearchParams(location.search).has("sin-carpeta")) {
  borrarCarpetasDB().catch(() => {}).finally(pintarCarpeta);
} else {
  cargarCarpetas();
}
iniciarMotor();
