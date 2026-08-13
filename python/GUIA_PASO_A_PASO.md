# Guía paso a paso — App de trazabilidad en GitHub Pages

Esta guía asume que es la primera vez que haces esto. Está pensada para que no
tengas que preguntar cada cosa: sigue los pasos en orden.

---

## 0. Qué estamos construyendo (y por qué sirvieron las carpetas)

Un **link** (como tu dashboard actual) que la operadora abre en su navegador.
Ahí ella:

1. elige la zona (RM, Sur o Norte),
2. sube los archivos Excel,
3. aprieta un botón y descarga el consolidado.

Todo el Python corre **dentro del navegador** con una tecnología llamada
**Pyodide** — la misma familia de tu panel. No hay servidor, no se instala nada,
y los archivos no salen del navegador de la operadora.

**Para qué sirvieron las carpetas `src/rm`, `src/sur`, `src/norte`:** son el
corazón de esto. La app **carga esas mismas funciones y las corre**. Todo el
trabajo de la Etapa 1 —dejar la lógica como funciones con las rutas como
parámetros— fue exactamente para que ese código pudiera correr fuera de tu
computador. En el navegador no hay carpetas de Windows, pero Pyodide tiene un
"sistema de archivos virtual": la app pone los Excel que sube la operadora en
`/work/uploads/…` y le pasa esas rutas a tus funciones. Como tus funciones ya
reciben la ruta como parámetro, corren **sin cambios**. Nada de la Etapa 1 se
pierde: se reusa completo.

---

## 1. Regla de oro de privacidad (leer antes que nada)

GitHub Pages gratis usa un repositorio **público**: cualquiera puede ver los
archivos del repo. Por eso:

> **Nunca subas archivos de datos al repo.** Solo va **código** (los `.html`,
> `.js`, `.py`). Ningún `.xlsx`, ni las bases (HOMOLOGACION, Transportistas,
> SINADER, Destinatarios), ni consolidados.

En esta versión la operadora **sube todos los Excel que hagan falta**, incluidas
las bases. Así el repo queda con cero datos. (Más adelante, si quieres que las
bases no se suban cada vez, se puede con un repo privado — eso lo vemos después,
no ahora.)

---

## 2. Qué necesitas tener listo

- Una cuenta de **GitHub** (gratis, con tu correo).
- Los archivos de esta carpeta (`index.html`, `app.js`, `python/procesar.py`).
- Tus scripts refactorizados de RM a mano (los usaremos en el paso 6).
- Saber **qué archivos base necesita RM** para consolidar (HOMOLOGACION, etc.).

---

## 3. Estructura del repositorio

Así queda el repo. Ni una carpeta con datos.

```
Sistema_Trazabilidad_App/
├── index.html            ← la página (interfaz)
├── app.js                ← conecta el navegador con Python (no se toca)
├── python/
│   ├── procesar.py       ← AQUÍ conectas tus scripts (paso 6)
│   ├── rm/               ← tus 3 scripts de RM (los agregas en el paso 6)
│   ├── sur/              ← (después)
│   └── norte/            ← (después)
└── GUIA_PASO_A_PASO.md   ← esta guía
```

---

## 4. Probar en tu computador ANTES de publicar

Siempre conviene ver que funciona local antes de subirlo.

1. Abre una terminal en la carpeta del proyecto.
2. Ejecuta:
   ```
   python -m http.server 8000
   ```
3. Abre en el navegador: `http://localhost:8000`
4. Deberías ver la página. Elige una zona, sube un par de Excel de prueba y
   aprieta **Procesar consolidado**. La primera vez el motor tarda ~20–30 s en
   cargar (baja Python + pandas). Debería mostrarte las hojas y filas de cada
   archivo y ofrecerte la descarga.

> Importante: tiene que ser con `http.server`, **no** abriendo el `index.html`
> con doble clic. Pyodide necesita que la página venga de un servidor (aunque sea
> local) para poder cargar el Python.

Si eso funciona, ya está lista para publicar.

---

## 5. Publicar en GitHub Pages (obtener el link)

1. Entra a **github.com** e inicia sesión.
2. Arriba a la derecha, botón **+** → **New repository**.
3. Nombre: por ejemplo `Sistema_Trazabilidad_App`. Déjalo **Public**. No marques
   "Add a README". Aprieta **Create repository**.
4. En la página del repo nuevo, busca el link **"uploading an existing file"**
   (o botón **Add file → Upload files**).
5. Arrastra **todos** los archivos y la carpeta `python/` del proyecto. Abajo,
   en "Commit changes", aprieta **Commit changes**.
6. Ve a la pestaña **Settings** del repo → menú lateral **Pages**.
7. En "Build and deployment", en **Source** elige **Deploy from a branch**. En
   **Branch** elige `main` y carpeta `/ (root)`. Aprieta **Save**.
8. Espera 1–2 minutos. Refresca la página de Pages: aparecerá tu link, del estilo
   `https://tu-usuario.github.io/Sistema_Trazabilidad_App/`.

Ese link es la app. Cualquiera con el link la puede usar. Cada vez que subas un
archivo nuevo o cambies uno, GitHub Pages se actualiza solo en un par de minutos.

---

## 6. Conectar tus scripts reales de RM

Hasta aquí la app funciona en modo **demostración** (lee los Excel y los combina,
para que veas todo el flujo). Ahora le ponemos la lógica real de RM.

1. Crea la carpeta `python/rm/` en el repo y sube tus tres scripts
   refactorizados (`consolidar.py`, `control_calidad.py`,
   `revisar_consolidado.py`) **tal cual** — sin rutas de tu computador adentro,
   solo funciones que reciben rutas como parámetro.
2. Abre `python/procesar.py` y ubica el bloque marcado
   **"AQUÍ VA TU LÓGICA REAL"**. Ahí:
   - importas tus módulos (`from rm import consolidar, control_calidad`),
   - llamas a tus funciones pasándoles las rutas de `carpeta_entrada`,
   - vuelcas el consolidado a `carpeta_salida`,
   - traduces el control de calidad a la lista `alertas`.
3. Prueba local otra vez (paso 4) con los archivos reales de RM. Cuando el
   consolidado descargado sea idéntico al que genera tu script hoy, quedó.

> **Atajo:** si me pegas tus tres scripts de RM y me dices qué archivos base
> necesita (nombres exactos), yo te dejo el bloque de `procesar.py` escrito y
> listo, y tú solo lo copias.

---

## 7. Agregar Sur y Norte (después de que RM funcione)

Cuando RM esté cerrada de punta a punta:

1. Sube `python/sur/` y `python/norte/` con sus scripts.
2. En `procesar.py`, dentro de la función, ramifica por zona:
   ```python
   if zona == "RM":
       ... llamar rm ...
   elif zona == "SUR":
       ... llamar sur ...
   elif zona == "NORTE":
       ... llamar norte ...
   ```

La interfaz (los botones RM / Sur / Norte) ya está lista para las tres. No hay
que tocar `index.html` ni `app.js`.

---

## Notas útiles

- **Si el motor no carga** y sale un error de versión de Pyodide: en `index.html`
  hay una línea con `v0.26.3`. Revisa en **pyodide.org** la versión estable
  actual y cambia solo ese número.
- **Archivos grandes:** todo corre en la memoria del navegador. RM y Norte son
  livianos; si algún mes un archivo fuera muy pesado y el navegador se pusiera
  lento, ese es el límite natural de esta versión (y el argumento para pasar a
  servidor más adelante).
- **Lo que esta versión NO hace** (a propósito, es de otra etapa): leer sola los
  archivos desde OneDrive y escribir en la base final compartida. Eso necesita un
  servidor por el tema de la clave secreta. Aquí es todo subir → procesar →
  descargar.
