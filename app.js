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

let pyodide = null;          // el "motor" Python en el navegador
let zonaElegida = null;      // "RM" | "SUR" | "NORTE"
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

/* Se sube al cambiar los archivos de python/, para que el navegador
   no siga usando una copia vieja guardada en caché. */
const VERSION = "3";

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

/* Scripts refactorizados de cada zona que hay que dejar dentro del motor.
   Cuando se conecten SUR y NORTE, se agregan acá igual que RM. */
const MODULOS_PY = {
  rm: ["consolidar.py", "control_calidad.py", "revisar_consolidado.py"],
};

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
   1. ARRANCAR EL MOTOR
   ========================================================================= */
async function iniciarMotor() {
  // Se va anotando en qué paso vamos: si algo falla, el mensaje dice
  // exactamente dónde, en vez de un "no se pudo iniciar" a ciegas.
  let paso = "descargando Pyodide";

  try {
    if (typeof loadPyodide !== "function") {
      throw new Error(
        "No se pudo descargar Pyodide desde internet. Puede estar bloqueado " +
        "por la red de la empresa (cdn.jsdelivr.net)."
      );
    }

    pyodide = await loadPyodide();

    paso = "cargando pandas";
    await pyodide.loadPackage(["pandas", "micropip"]);

    // openpyxl NO viene incluido en Pyodide (0.26.3 trae 310 paquetes y ese
    // no está). Se instala desde los wheels guardados en el propio repo, así
    // no depende de PyPI ni de que la red de la empresa lo deje pasar.
    paso = "instalando openpyxl";
    const micropip = pyodide.pyimport("micropip");

    // URL absoluta: una ruta relativa depende de cómo resuelva micropip y
    // eso es una fuente de fallos difícil de ver. Así no queda duda.
    const wheels = [
      new URL("python/wheels/et_xmlfile-2.0.0-py3-none-any.whl", document.baseURI).href,
      new URL("python/wheels/openpyxl-3.1.5-py2.py3-none-any.whl", document.baseURI).href,
    ];

    try {
      // Comprobar primero que los archivos existen: si faltan, el mensaje
      // dice cuál, en vez de un error genérico de micropip.
      for (const w of wheels) {
        const r = await fetch(w, { cache: "no-store" });
        if (!r.ok) throw new Error(`No se encontró ${w} (error ${r.status})`);
      }
      await micropip.install(wheels);
    } catch (e) {
      console.warn("Wheels del repositorio no utilizables:", e);
      paso = "instalando openpyxl desde PyPI (los wheels del repo fallaron: " +
             ((e && e.message) || e) + ")";
      await micropip.install("openpyxl");
    }

    paso = "preparando las carpetas de trabajo";
    pyodide.FS.mkdirTree("/work/uploads");
    pyodide.FS.mkdirTree("/work/salida");
    pyodide.FS.mkdirTree("/work/base");

    // Copiar los scripts de cada zona al motor, en /work/py/<zona>/,
    // para que procesar.py pueda importarlos con "from rm import consolidar".
    paso = "copiando los scripts de zona (python/rm/…)";
    const enc = new TextEncoder();
    pyodide.FS.mkdirTree("/work/py");
    for (const [carpeta, nombres] of Object.entries(MODULOS_PY)) {
      pyodide.FS.mkdirTree("/work/py/" + carpeta);
      pyodide.FS.writeFile(`/work/py/${carpeta}/__init__.py`, enc.encode(""));
      for (const nombre of nombres) {
        const r = await fetch(`python/${carpeta}/${nombre}?v=${VERSION}`, { cache: "no-store" });
        if (!r.ok) {
          console.warn(`No se encontró python/${carpeta}/${nombre}`);
          continue;
        }
        pyodide.FS.writeFile(`/work/py/${carpeta}/${nombre}`, enc.encode(await r.text()));
      }
    }

    // Traer procesar.py y dejarlo disponible dentro del motor
    paso = "descargando python/procesar.py";
    const resp = await fetch(`python/procesar.py?v=${VERSION}`, { cache: "no-store" });
    if (!resp.ok) {
      throw new Error(
        `No se encontró python/procesar.py en el sitio (error ${resp.status}). ` +
        "Revisa que el archivo esté dentro de la carpeta python/ del repositorio."
      );
    }

    paso = "ejecutando procesar.py";
    pyodide.runPython(await resp.text());

    // Aviso claro si el procesar.py del repositorio es una versión anterior
    paso = "revisando la versión de procesar.py";
    hayPrepararBase = pyodide.runPython(`"preparar_base" in globals()`);

    $("punto-motor").classList.add("listo");
    $("txt-motor").textContent = "Motor listo";

    // La memoria del navegador es opcional: si este navegador no la permite,
    // la app igual tiene que poder consolidar subiendo todos los archivos.
    paso = "leyendo los archivos base guardados";
    try {
      await pintarBase();
    } catch (e) {
      $("estado-base").textContent =
        "Este navegador no permite guardar archivos base (" +
        (e && e.message ? e.message : e) +
        "). Puedes seguir usando la app subiendo todos los archivos cada vez.";
      console.warn(e);
    }

    revisarSiPuedeProcesar();
  } catch (e) {
    // Mostrar el error completo en pantalla: sin esto no hay forma de saber
    // qué pasó sin abrir la consola del navegador.
    const detalle = e && e.stack ? e.stack : String(e);
    const mensaje = (e && e.message) || String(e);

    $("txt-motor").textContent = "No se pudo iniciar el motor";

    // El error va donde la persona está mirando: en el paso de archivos base
    // y en una franja fija abajo. Antes solo aparecía en el bloque 04, que
    // queda fuera de pantalla y por lo tanto no se leía.
    if ($("estado-base")) {
      $("estado-base").textContent =
        `El motor no arrancó (falló en: ${paso}). Por eso los archivos base ` +
        `no se pueden guardar todavía.`;
    }
    mostrarFranja(`No se pudo iniciar el motor · Falló en: ${paso} · ${mensaje}`);

    $("resultado").style.display = "block";
    $("resumen").innerHTML =
      `<b>No se pudo iniciar el motor.</b> Falló en: ${paso}.`;
    $("alertas").innerHTML =
      `<div class="alerta"><b>${(e && e.name) || "Error"}</b> — ${
        (e && e.message) || e
      }</div>`;
    $("registro").textContent = `Paso: ${paso}\n\n${detalle}`;
    $("registro").classList.remove("oculto");
    console.error(e);
  }
}

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
  revisarSiPuedeProcesar();
});

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

  if (!pyodide) {
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
    // Dejar los archivos en una carpeta aparte del motor
    for (const f of pyodide.FS.readdir("/work/base")) {
      if (f !== "." && f !== "..") pyodide.FS.unlink("/work/base/" + f);
    }
    for (const f of utiles) {
      pyodide.FS.writeFile("/work/base/" + f.name, new Uint8Array(await f.arrayBuffer()));
    }

    // Python expande los .zip y se queda solo con los cuatro maestros
    const res = JSON.parse(pyodide.runPython(`
import json
json.dumps(preparar_base("/work/base"))
`));

    // Guardar en el navegador lo que quedó
    const fecha = new Date().toLocaleDateString("es-CL");
    const items = res.guardados.map((nombre, i) => ({
      nombre,
      etiqueta: res.etiquetas[i],
      fecha,
      bytes: pyodide.FS.readFile("/work/base/" + nombre),
    }));

    if (items.length) await guardarBaseDB(items);
    console.info("[archivos base]\n" + (res.log || ""));
    res.enviados = utiles.map((f) => f.name);
    await pintarBase(res);
  } catch (e) {
    // Mostrar el error real: sin esto, cualquier problema se ve igual y no
    // hay forma de saber qué pasó sin abrir la consola del navegador.
    $("estado-base").textContent =
      "No se pudieron guardar los archivos base → " + (e && e.message ? e.message : e);
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
  $("btn-procesar").disabled = !(pyodide && zonaElegida && archivos.length > 0);
}

/* =========================================================================
   5. PROCESAR
   ========================================================================= */
alEvento("btn-procesar", "click", async () => {
  const btn = $("btn-procesar");
  btn.disabled = true;
  btn.innerHTML = '<span class="spin" style="display:inline-block;vertical-align:middle"></span> Procesando…';

  try {
    // Limpiar las carpetas de la corrida anterior
    for (const carpeta of ["/work/uploads", "/work/salida"]) {
      for (const f of pyodide.FS.readdir(carpeta)) {
        if (f !== "." && f !== "..") pyodide.FS.unlink(carpeta + "/" + f);
      }
    }

    // Primero los archivos base guardados…
    for (const item of baseGuardada) {
      pyodide.FS.writeFile("/work/uploads/" + item.nombre, item.bytes);
    }
    // …y encima los movimientos del mes
    for (const f of archivos) {
      const buf = new Uint8Array(await f.arrayBuffer());
      pyodide.FS.writeFile("/work/uploads/" + f.name, buf);
    }

    pyodide.globals.set("ZONA_JS", zonaElegida);
    const jsonTexto = await pyodide.runPythonAsync(`
import json
json.dumps(procesar(ZONA_JS, "/work/uploads", "/work/salida"))
`);
    mostrarResultado(JSON.parse(jsonTexto));
  } catch (e) {
    $("resultado").style.display = "block";
    $("resumen").innerHTML = "<b>Ocurrió un error al procesar.</b> Revisa el detalle técnico.";
    $("registro").textContent = String(e);
    $("registro").classList.remove("oculto");
    console.error(e);
  } finally {
    btn.disabled = false;
    btn.textContent = "Procesar consolidado";
    revisarSiPuedeProcesar();
  }
});

/* =========================================================================
   6. MOSTRAR RESULTADO Y PREPARAR LAS DESCARGAS
   ========================================================================= */
function mostrarResultado(r) {
  $("resultado").style.display = "block";
  $("resumen").innerHTML = r.resumen || "Proceso terminado.";

  // Tabla de fuentes leídas
  const tb = $("tabla-fuentes").querySelector("tbody");
  tb.innerHTML = "";
  (r.fuentes || []).forEach((f) => {
    tb.innerHTML += `<tr><td>${f.archivo}</td><td>${f.hojas}</td><td>${f.filas}</td></tr>`;
  });

  // Alertas del control de calidad
  const cont = $("alertas");
  cont.innerHTML = "";
  (r.alertas || []).forEach((a) => {
    const div = document.createElement("div");
    div.className = "alerta";
    div.innerHTML = `<b>${a.titulo}</b> — ${a.detalle}`;
    cont.appendChild(div);
  });

  $("registro").textContent = r.log || "";

  // Un botón por cada archivo que Python dejó listo para descargar
  const zona = $("descargas");
  zona.innerHTML = "";
  (r.descargas || []).forEach((d, i) => {
    let bytes;
    try {
      bytes = pyodide.FS.readFile("/work/salida/" + d.archivo);
    } catch (e) {
      return;   // ese archivo no se generó, no se muestra el botón
    }
    const b = document.createElement("button");
    b.className = i === 0 ? "btn" : "btn secundario";
    b.textContent = d.etiqueta;
    b.onclick = () => descargar(bytes, d.archivo);
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
