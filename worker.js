/* =========================================================================
   worker.js — el motor Python, corriendo en un hilo aparte.

   POR QUÉ EXISTE ESTE ARCHIVO
   Antes Pyodide corría en el mismo hilo que dibuja la página. Mientras pandas
   consolidaba, el navegador no podía repintar nada y Chrome mostraba "La
   página no responde". Acá el trabajo pesado ocurre en un worker: la página
   sigue viva, se puede desplazar, y encima puede ir contando en qué paso va.

   No toca el DOM ni sabe nada de la interfaz: solo recibe mensajes, procesa
   y responde. Toda la lógica de negocio sigue viviendo en python/.
   ========================================================================= */

const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.3/full/pyodide.js";

const MODULOS_PY = {
  rm: ["consolidar.py", "control_calidad.py", "revisar_consolidado.py", "control_simple_route.py"],
  sur: ["consolidar.py", "control_calidad.py", "revisar_consolidado.py"],
  norte: ["consolidar.py", "control_calidad.py", "revisar_consolidado.py"],
};

const WHEELS = [
  "python/wheels/et_xmlfile-2.0.0-py3-none-any.whl",
  "python/wheels/openpyxl-3.1.5-py2.py3-none-any.whl",
];

const VERSION = "21";

let pyodide = null;
let base = self.location.href;   // la reemplaza el mensaje "iniciar"

const url = (rel) => new URL(rel, base).href;
const avisar = (texto) => postMessage({ tipo: "progreso", texto });

/* ---------------------------------------------------------------------- */
async function iniciar() {
  let paso = "descargando Pyodide";
  try {
    avisar("Descargando el motor…");
    importScripts(PYODIDE_URL);
    pyodide = await loadPyodide();

    paso = "cargando pandas";
    avisar("Cargando pandas…");
    await pyodide.loadPackage(["pandas", "micropip"]);

    paso = "instalando openpyxl";
    avisar("Instalando openpyxl…");
    const micropip = pyodide.pyimport("micropip");
    try {
      const wheels = WHEELS.map(url);
      for (const w of wheels) {
        const r = await fetch(w, { cache: "no-store" });
        if (!r.ok) throw new Error(`No se encontró ${w} (error ${r.status})`);
      }
      await micropip.install(wheels);
    } catch (e) {
      paso = "instalando openpyxl desde PyPI (los wheels del repo fallaron: " +
             ((e && e.message) || e) + ")";
      await micropip.install("openpyxl");
    }

    paso = "preparando las carpetas de trabajo";
    ["/work/uploads", "/work/salida", "/work/base", "/work/py"].forEach((d) =>
      pyodide.FS.mkdirTree(d)
    );

    paso = "copiando los scripts de zona (python/rm/…)";
    avisar("Copiando los scripts de cada zona…");
    const enc = new TextEncoder();
    for (const [carpeta, nombres] of Object.entries(MODULOS_PY)) {
      pyodide.FS.mkdirTree("/work/py/" + carpeta);
      pyodide.FS.writeFile(`/work/py/${carpeta}/__init__.py`, enc.encode(""));
      for (const nombre of nombres) {
        const r = await fetch(url(`python/${carpeta}/${nombre}?v=${VERSION}`), { cache: "no-store" });
        if (!r.ok) throw new Error(`No se encontró python/${carpeta}/${nombre} (error ${r.status})`);
        pyodide.FS.writeFile(`/work/py/${carpeta}/${nombre}`, enc.encode(await r.text()));
      }
    }

    paso = "descargando python/procesar.py";
    const resp = await fetch(url(`python/procesar.py?v=${VERSION}`), { cache: "no-store" });
    if (!resp.ok) {
      throw new Error(
        `No se encontró python/procesar.py en el sitio (error ${resp.status}).`
      );
    }

    paso = "ejecutando procesar.py";
    pyodide.runPython(await resp.text());

    // Lo que Python imprime se manda a la pantalla en vivo. Es lo que permite
    // ver "leyendo GIRI…" en vez de un spinner mudo durante varios minutos.
    pyodide.setStdout({ batched: (linea) => avisar(linea) });
    pyodide.setStderr({ batched: (linea) => avisar(linea) });

    paso = "revisando la versión de procesar.py";
    const hayPrepararBase = pyodide.runPython(`"preparar_base" in globals()`);

    postMessage({ tipo: "listo", hayPrepararBase });
  } catch (e) {
    postMessage({
      tipo: "error-motor",
      paso,
      mensaje: (e && e.message) || String(e),
      detalle: (e && e.stack) || String(e),
    });
  }
}

/* ---------------------------------------------------------------------- */
function limpiar(carpeta) {
  for (const f of pyodide.FS.readdir(carpeta)) {
    if (f !== "." && f !== "..") {
      try { pyodide.FS.unlink(carpeta + "/" + f); } catch (e) { /* ignorar */ }
    }
  }
}

function escribir(carpeta, archivos) {
  for (const a of archivos) {
    pyodide.FS.writeFile(carpeta + "/" + a.nombre, a.bytes);
  }
}

/* ---------------------------------------------------------------------- */
async function prepararBase(archivos) {
  limpiar("/work/base");
  escribir("/work/base", archivos);

  const res = JSON.parse(pyodide.runPython(`
import json
json.dumps(preparar_base("/work/base"))
`));

  // Devuelve el contenido de los que quedaron guardados, para que la página
  // los meta en la memoria del navegador.
  const guardados = res.guardados.map((nombre, i) => ({
    nombre,
    etiqueta: res.etiquetas[i],
    bytes: pyodide.FS.readFile("/work/base/" + nombre),
  }));

  postMessage({ tipo: "base-lista", res, guardados });
}

/* ---------------------------------------------------------------------- */
async function procesarZona(zona, archivosBase, movimientos, periodo) {
  limpiar("/work/uploads");
  limpiar("/work/salida");

  escribir("/work/uploads", archivosBase);   // primero los base guardados
  escribir("/work/uploads", movimientos);    // encima los del mes

  avisar(`Procesando zona ${zona}…`);
  pyodide.globals.set("ZONA_JS", zona);
  pyodide.globals.set("PERIODO_JS", periodo || null);
  const r = JSON.parse(await pyodide.runPythonAsync(`
import json
json.dumps(procesar(ZONA_JS, "/work/uploads", "/work/salida", PERIODO_JS))
`));

  // Los archivos generados viajan de vuelta como bytes: la página no tiene
  // acceso al sistema de archivos del motor.
  const descargas = [];
  for (const d of r.descargas || []) {
    try {
      descargas.push({
        archivo: d.archivo,
        etiqueta: d.etiqueta,
        bytes: pyodide.FS.readFile("/work/salida/" + d.archivo),
      });
    } catch (e) {
      // ese archivo no se generó; simplemente no se ofrece
    }
  }

  postMessage({ tipo: "resultado", r, descargas });
}

/* ---------------------------------------------------------------------- */
onmessage = async (ev) => {
  const m = ev.data || {};
  try {
    if (m.tipo === "iniciar") {
      base = m.base || base;
      await iniciar();
    } else if (m.tipo === "preparar-base") {
      await prepararBase(m.archivos);
    } else if (m.tipo === "procesar") {
      await procesarZona(m.zona, m.base, m.movimientos, m.periodo);
    }
  } catch (e) {
    postMessage({
      tipo: "error",
      origen: m.tipo,
      mensaje: (e && e.message) || String(e),
      detalle: (e && e.stack) || String(e),
    });
  }
};
