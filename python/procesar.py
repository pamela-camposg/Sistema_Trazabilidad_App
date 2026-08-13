# =============================================================================
# procesar.py — el puente entre la app y los scripts refactorizados.
#
# La app le pasa a la función procesar():
#   - zona:            "RM", "SUR" o "NORTE"
#   - carpeta_entrada: donde quedaron los Excel que subió la operadora
#   - carpeta_salida:  donde debe dejar los archivos generados
#
# Devuelve un diccionario con:
#   resumen  -> texto que se muestra arriba
#   fuentes  -> lista de {archivo, hojas, filas} para la tabla
#   alertas  -> lista de {titulo, detalle} del control de calidad
#   salida   -> nombre del consolidado que quedó en carpeta_salida
#   log      -> texto largo con el detalle técnico
#
# ESTADO DE LAS ZONAS
#   RM     -> lógica REAL conectada (scripts de la Etapa 1)
#   SUR    -> modo demostración (pendiente de conectar)
#   NORTE  -> modo demostración (pendiente de conectar)
#
# CÓMO SE CONECTAN LOS SCRIPTS
#   app.js descarga los módulos de python/rm/ y los deja dentro del motor,
#   en /work/py/rm/. Por eso acá basta con "from rm import consolidar".
#   Los tres scripts NO fueron modificados: son exactamente los mismos que
#   quedaron verificados en la sección 8.2 del TRZ-APP-001. Todo lo que hace
#   falta para adaptarlos a la app está en ESTE archivo.
#
# REGLA IMPORTANTE
#   La consolidación se ejecuta siempre con modo_prueba=True. Eso significa
#   que el archivo real BBDD_TRAZABILIDAD_RM.xlsx NUNCA se toca: el resultado
#   se escribe en un archivo nuevo que la operadora descarga.
# =============================================================================

import os
import sys
import traceback
from pathlib import Path

import pandas as pd

# Carpeta donde app.js deja los módulos de cada zona
if "/work/py" not in sys.path:
    sys.path.insert(0, "/work/py")


# -----------------------------------------------------------------------------
# Archivos que consolidar.py necesita y que control_calidad.py no busca.
#
# control_calidad.rutas_desde_carpeta() encuentra 8 archivos. consolidar()
# necesita además el maestro de destinatarios y (opcionalmente) HOMOLOGACION.
# Se agregan acá para no tener que modificar los scripts ya verificados.
# -----------------------------------------------------------------------------
NOMBRES_EXTRA = {
    "destinatarios": ["BBDD DESTINATARIO.xlsx", "BBDD DESTINATARIOS.xlsx"],
    "homologacion": ["HOMOLOGACION.xlsx", "HOMOLOGACIÓN.xlsx"],
}

# 'destinatarios' es obligatorio para consolidar; 'homologacion' no lo es
# (el script original avisa y sigue sin aplicar correcciones).
EXTRA_OBLIGATORIOS = ["destinatarios"]

# Nombres legibles de los controles de calidad (C1–C9)
ETIQUETAS_CC = {
    "c1": "C1 — Cliente con mismo nombre y RUT distinto",
    "c2": "C2 — Cliente con mismo RUT y nombre distinto",
    "c3": "C3 — Generador con variantes de nombre",
    "c5": "C5 — TIPO sin código SINADER",
    "c6": "C6 — Transportistas sin RUT",
    "c7": "C7 — Vacíos críticos",
    "c8": "C8 — Movimientos no reconocidos (se descartan al consolidar)",
    "c9": "C9 — Destinos vacíos sin explicación (se descartan al consolidar)",
}

# Nombres legibles de las validaciones del consolidado (V0–V9)
ETIQUETAS_REV = {
    "v0": "V0 — Estructura de columnas",
    "v1": "V1 — Columnas con vacíos",
    "v2": "V2 — RUT sin guión",
    "v3": "V3 — Destinos inferidos (revisar manualmente)",
    "v4": "V4 — Filas duplicadas",
    "v5": "V5 — Fechas fuera de rango",
    "v7": "V7 — Pesos inválidos",
    "v8": "V8 — Regiones no válidas",
    "v9": "V9 — Movimientos no válidos",
}


# =============================================================================
# UTILIDADES
# =============================================================================
def _listar_archivos(carpeta):
    """Devuelve las rutas de todos los .xlsx de una carpeta, ordenadas."""
    return sorted(
        os.path.join(carpeta, f)
        for f in os.listdir(carpeta)
        if f.lower().endswith(".xlsx")
    )


def _inventario(rutas):
    """Arma la tabla 'Archivo / Hojas / Filas' sin cargar los datos.

    Se usa openpyxl en modo solo lectura: lee el encabezado del archivo pero
    no los valores, así la tabla aparece rápido aunque los Excel sean grandes.
    """
    from openpyxl import load_workbook

    filas = []
    for ruta in rutas:
        nombre = os.path.basename(ruta)
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            hojas = wb.sheetnames
            try:
                # max_row incluye la fila de encabezado: se descuenta para
                # que la cifra sea comparable con la del resto del sistema
                total = max(0, (wb[hojas[0]].max_row or 0) - 1)
            except Exception:
                total = "—"
            wb.close()
            filas.append({
                "archivo": nombre,
                "hojas": ", ".join(hojas),
                "filas": total,
            })
        except Exception as e:
            filas.append({"archivo": nombre, "hojas": "no se pudo leer", "filas": str(e)[:60]})
    return filas


def _archivo_nuevo(carpeta, antes):
    """Devuelve el archivo que apareció en 'carpeta' después de una operación.

    consolidar() en modo prueba le agrega fecha y hora al nombre del archivo,
    así que no se puede saber de antemano cómo se va a llamar. En vez de
    adivinarlo, se mira qué archivo nuevo apareció.
    """
    ahora = {f.name for f in Path(carpeta).glob("*.xlsx")}
    nuevos = sorted(ahora - antes)
    return nuevos[-1] if nuevos else None


def _sumar_kilos(df):
    """Suma la columna de peso del consolidado, si existe. Solo informativo."""
    for col in df.columns:
        if "PESO" in str(col).upper():
            try:
                return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
            except Exception:
                return None
    return None


def _formato_miles(n, decimales=2):
    """Formatea un número al estilo chileno: 16.581.915,20"""
    if n is None:
        return "—"
    texto = f"{n:,.{decimales}f}"
    return texto.replace(",", "·").replace(".", ",").replace("·", ".")


# =============================================================================
# ZONA RM — LÓGICA REAL
# =============================================================================
def _rutas_rm(carpeta_entrada, control_calidad):
    """Arma el diccionario de rutas que necesitan los scripts de RM.

    Parte de rutas_desde_carpeta() de control_calidad.py (que ya encuentra las
    8 fuentes tolerando tildes y mayúsculas) y le agrega los archivos que solo
    necesita consolidar().
    """
    rutas = control_calidad.rutas_desde_carpeta(carpeta_entrada)

    # Índice de lo que hay en la carpeta, con la misma normalización tolerante
    # que usa control_calidad.py (_clave ignora tildes, espacios y mayúsculas)
    presentes = {
        control_calidad._clave(f.name): f
        for f in Path(carpeta_entrada).iterdir()
        if f.is_file()
    }

    faltantes = []
    for interno, posibles in NOMBRES_EXTRA.items():
        encontrado = None
        for nombre in posibles:
            encontrado = presentes.get(control_calidad._clave(nombre))
            if encontrado:
                break
        if encontrado:
            rutas[interno] = encontrado
        elif interno in EXTRA_OBLIGATORIOS:
            faltantes.append(posibles[0])

    if faltantes:
        detalle = "\n".join(f"    · {f}" for f in faltantes)
        raise control_calidad.ArchivoFaltante(
            f"Faltan archivos que necesita la consolidación:\n{detalle}"
        )

    # consolidar() espera esta clave, pero en modo prueba nunca la usa: apunta
    # a un archivo que no existe, a propósito, para que sea imposible escribir
    # sobre la base real de trazabilidad.
    rutas["destino_real"] = Path(carpeta_entrada) / "__NO_USAR__.xlsx"

    return rutas


def _alertas_rm(cc, rev):
    """Traduce los resultados de los scripts a la lista de alertas de pantalla."""
    alertas = []

    for clave, etiqueta in ETIQUETAS_CC.items():
        df = cc.get(clave)
        if df is not None and len(df) > 0:
            alertas.append({
                "titulo": etiqueta,
                "detalle": f"{len(df)} caso(s). Ver la hoja correspondiente "
                           f"en CONTROL_CALIDAD_RM.xlsx.",
            })

    informativo = cc.get("c9_informativo")
    if informativo is not None and len(informativo) > 0:
        alertas.append({
            "titulo": "C9b — Destinos vacíos con observación operativa",
            "detalle": f"{len(informativo)} fila(s). Es informativo: no se corrige.",
        })

    for clave, etiqueta in ETIQUETAS_REV.items():
        df = rev.get(clave)
        if df is not None and len(df) > 0:
            alertas.append({
                "titulo": etiqueta,
                "detalle": f"{len(df)} caso(s) en el consolidado generado.",
            })

    if not alertas:
        alertas.append({
            "titulo": "Sin hallazgos",
            "detalle": "El control de calidad y la revisión del consolidado "
                       "no encontraron nada que corregir.",
        })

    return alertas


def _procesar_rm(carpeta_entrada, carpeta_salida, log):
    """Ejecuta el flujo completo de RM: control de calidad → consolidar → revisar."""
    from rm import consolidar as mod_consolidar
    from rm import control_calidad as mod_control
    from rm import revisar_consolidado as mod_revisar

    salida = Path(carpeta_salida)

    # ---- 1. Ubicar los archivos subidos -------------------------------------
    log.append("── Buscando los archivos en la carpeta de subida ──")
    rutas = _rutas_rm(carpeta_entrada, mod_control)
    for clave in sorted(rutas):
        if clave != "destino_real":
            log.append(f"  {clave:<15} → {Path(rutas[clave]).name}")
    log.append("")

    # ---- 2. Control de calidad (C1 a C9) ------------------------------------
    log.append("── Control de calidad ──")
    cc = mod_control.controlar(
        rutas,
        ruta_salida=salida / "CONTROL_CALIDAD_RM.xlsx",
        mostrar=False,
    )
    log.append(cc.get("log", ""))
    log.append("")

    # ---- 3. Consolidación (modo prueba: no toca la base real) ---------------
    log.append("── Consolidación ──")
    antes = {f.name for f in salida.glob("*.xlsx")}
    res = mod_consolidar.consolidar(
        rutas,
        ruta_prueba=salida / "TRAZABILIDAD_RM.xlsx",
        ruta_log=None,
        modo_reset=False,
        modo_prueba=True,     # nunca escribe sobre BBDD_TRAZABILIDAD_RM.xlsx
        mostrar=False,
    )
    log.append(res.get("log", ""))

    nombre_consolidado = _archivo_nuevo(salida, antes)
    if not nombre_consolidado:
        raise RuntimeError(
            "La consolidación terminó pero no se encontró el archivo generado."
        )
    log.append("")

    # ---- 4. Revisión del consolidado (V0 a V9) ------------------------------
    log.append("── Revisión del consolidado ──")
    rev = mod_revisar.revisar(
        salida / nombre_consolidado,
        ruta_salida=salida / "REVISION_CONSOLIDADO_RM.xlsx",
        mostrar=False,
    )
    log.append(rev.get("log", ""))

    # ---- 5. Armar lo que se muestra en pantalla -----------------------------
    kilos = _sumar_kilos(res["consolidado"])
    resumen = (
        f"<b>Zona RM consolidada.</b> "
        f"{_formato_miles(res['filas'], 0)} filas × {res['columnas']} columnas"
        + (f" · {_formato_miles(kilos)} kg" if kilos else "")
        + f" · control de calidad: {cc.get('total', 0)} conflicto(s)"
        + f" · revisión: {rev.get('total_alertas', 0)} alerta(s)."
    )

    return {
        "resumen": resumen,
        "fuentes": _inventario(_listar_archivos(carpeta_entrada)),
        "alertas": _alertas_rm(cc, rev),
        "salida": nombre_consolidado,
        "log": "\n".join(log),
    }


# =============================================================================
# ZONAS SUR Y NORTE — TODAVÍA EN DEMOSTRACIÓN
# =============================================================================
def _procesar_demostracion(zona, carpeta_entrada, carpeta_salida, log):
    """Combina los Excel subidos sin aplicar reglas de negocio.

    Se mantiene para SUR y NORTE mientras no se conecten sus scripts, para que
    la app siga funcionando de punta a punta en las tres zonas.
    """
    rutas = _listar_archivos(carpeta_entrada)
    fuentes, alertas, marcos = [], [], []

    for ruta in rutas:
        nombre = os.path.basename(ruta)
        try:
            hojas = pd.ExcelFile(ruta).sheet_names
            df = pd.read_excel(ruta, sheet_name=hojas[0])
            fuentes.append({"archivo": nombre, "hojas": ", ".join(hojas), "filas": len(df)})
            log.append(f"  {nombre}: hojas={hojas}, filas primera hoja={len(df)}")
            df["_archivo_origen"] = nombre
            marcos.append(df)
        except Exception as e:
            alertas.append({"titulo": f"No se pudo leer {nombre}", "detalle": str(e)})
            log.append(f"  ERROR en {nombre}: {e}")

    salida = None
    if marcos:
        df_todo = pd.concat(marcos, ignore_index=True)
        salida = f"CONSOLIDADO_{zona}.xlsx"
        df_todo.to_excel(os.path.join(carpeta_salida, salida), index=False)
        log.append(f"\nConsolidado de demostración: {len(df_todo)} filas")
        resumen = (
            f"<b>Zona {zona}</b> · {len(rutas)} archivo(s) leídos · "
            f"{len(df_todo)} filas combinadas (versión de demostración)."
        )
    else:
        resumen = "No se leyó ningún archivo válido."

    alertas.insert(0, {
        "titulo": f"Zona {zona}: versión de demostración",
        "detalle": "Todavía no está conectada la lógica real de esta zona. "
                   "El consolidado es solo una combinación de prueba.",
    })

    return {
        "resumen": resumen,
        "fuentes": fuentes,
        "alertas": alertas,
        "salida": salida,
        "log": "\n".join(log),
    }


# =============================================================================
# PUNTO DE ENTRADA — es lo que llama app.js
# =============================================================================
def procesar(zona, carpeta_entrada, carpeta_salida):
    zona = (zona or "").strip().upper()
    log = [f"Zona: {zona}", f"Carpeta de entrada: {carpeta_entrada}", ""]

    try:
        if zona == "RM":
            return _procesar_rm(carpeta_entrada, carpeta_salida, log)
        return _procesar_demostracion(zona, carpeta_entrada, carpeta_salida, log)

    except Exception as e:
        # Un error acá no debe dejar la pantalla en blanco: se muestra el
        # mensaje del script (que suele decir exactamente qué archivo falta)
        # y el detalle técnico queda en el log.
        log.append("")
        log.append("── ERROR ──")
        log.append(traceback.format_exc())
        return {
            "resumen": "<b>No se pudo completar el proceso.</b> "
                       "Revisa el mensaje de abajo.",
            "fuentes": _inventario(_listar_archivos(carpeta_entrada)),
            "alertas": [{"titulo": type(e).__name__, "detalle": str(e)}],
            "salida": None,
            "log": "\n".join(log),
        }
