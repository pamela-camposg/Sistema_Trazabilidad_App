/* =========================================================================
   app.js — pega el navegador con tu código Python.
   No necesitas tocar este archivo para conectar tus scripts: eso se hace
   en python/procesar.py. Aquí solo está la mecánica de la interfaz.
   ========================================================================= */

let pyodide = null;         // el "motor" Python en el navegador
let zonaElegida = null;     // "RM" | "SUR" | "NORTE"
let archivos = [];          // archivos que subió la operadora
let bytesSalida = null;     // el consolidado generado, listo para descargar
let nombreSalida = "";

const $ = (id) => document.getElementById(id);

/* ---- 1. Arrancar Pyodide y cargar pandas + openpyxl ---- */
async function iniciarMotor() {
  try {
    pyodide = await loadPyodide();
    await pyodide.loadPackage(["pandas", "openpyxl"]);

    // Traer tu código Python (procesar.py) y dejarlo disponible dentro del motor.
    const codigo = await (await fetch("python/procesar.py")).text();
    pyodide.FS.mkdirTree("/work/uploads");
    pyodide.FS.mkdirTree("/work/salida");
    pyodide.runPython(codigo);

    $("punto-motor").classList.add("listo");
    $("txt-motor").textContent = "Motor listo. Ya puedes procesar.";
    revisarSiPuedeProcesar();
  } catch (e) {
    $("txt-motor").textContent = "No se pudo iniciar el motor. Recarga la página.";
    console.error(e);
  }
}

/* ---- 2. Elegir zona ---- */
$("zonas").addEventListener("click", (ev) => {
  const btn = ev.target.closest(".zona-btn");
  if (!btn) return;
  zonaElegida = btn.dataset.zona;
  document.querySelectorAll(".zona-btn").forEach((b) =>
    b.setAttribute("aria-pressed", b === btn ? "true" : "false")
  );
  revisarSiPuedeProcesar();
});

/* ---- 3. Subir archivos (clic o arrastrar) ---- */
const dz = $("dropzone");
$("input-archivos").addEventListener("change", (e) => agregarArchivos(e.target.files));
["dragover", "dragenter"].forEach((ev) =>
  dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); })
);
["dragleave", "drop"].forEach((ev) =>
  dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); })
);
dz.addEventListener("drop", (e) => agregarArchivos(e.dataTransfer.files));

function agregarArchivos(fileList) {
  for (const f of fileList) {
    if (f.name.toLowerCase().endsWith(".xlsx") && !archivos.some((a) => a.name === f.name)) {
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
    li.innerHTML = `<span>${f.name}</span>`;
    const b = document.createElement("button");
    b.className = "quitar";
    b.textContent = "quitar";
    b.onclick = () => { archivos.splice(i, 1); pintarLista(); revisarSiPuedeProcesar(); };
    li.appendChild(b);
    ul.appendChild(li);
  });
}

function revisarSiPuedeProcesar() {
  $("btn-procesar").disabled = !(pyodide && zonaElegida && archivos.length > 0);
}

/* ---- 4. Procesar: escribir los archivos en el motor y llamar a tu Python ---- */
$("btn-procesar").addEventListener("click", async () => {
  const btn = $("btn-procesar");
  btn.disabled = true;
  btn.innerHTML = '<span class="spin" style="display:inline-block;vertical-align:middle"></span> Procesando…';

  try {
    // Limpiar la carpeta de subidas de una corrida anterior
    for (const f of pyodide.FS.readdir("/work/uploads")) {
      if (f !== "." && f !== "..") pyodide.FS.unlink("/work/uploads/" + f);
    }
    // Copiar cada archivo subido al sistema de archivos del motor
    for (const f of archivos) {
      const buf = new Uint8Array(await f.arrayBuffer());
      pyodide.FS.writeFile("/work/uploads/" + f.name, buf);
    }

    // Llamar a tu función Python. Devuelve un JSON con el resultado.
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
  }
});

/* ---- 5. Mostrar resultado y preparar la descarga ---- */
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

  // Leer el archivo consolidado que generó Python y dejarlo listo para descargar
  if (r.salida) {
    bytesSalida = pyodide.FS.readFile("/work/salida/" + r.salida);
    nombreSalida = r.salida;
    $("btn-descargar").style.display = "inline-block";
  } else {
    $("btn-descargar").style.display = "none";
  }
}

/* ---- 6. Descargar ---- */
$("btn-descargar").addEventListener("click", () => {
  if (!bytesSalida) return;
  const blob = new Blob([bytesSalida], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombreSalida;
  a.click();
  URL.revokeObjectURL(url);
});

$("btn-log").addEventListener("click", () => $("registro").classList.toggle("oculto"));

iniciarMotor();
