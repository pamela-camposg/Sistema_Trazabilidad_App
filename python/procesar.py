# =============================================================================
# procesar.py — el puente entre la app y los scripts refactorizados.
#
# La app le pasa a la función procesar():
#   - zona:            "RM", "SUR" o "NORTE"
#   - carpeta_entrada: donde quedaron los archivos que subió la operadora
#   - carpeta_salida:  donde debe dejar los archivos generados
#
# Devuelve un diccionario con:
#   resumen    -> texto que se muestra arriba
#   fuentes    -> lista de {archivo, hojas, filas} para la tabla
#   alertas    -> lista de {titulo, detalle} del control de calidad
#   salida     -> nombre del consolidado (el archivo principal a descargar)
#   descargas  -> lista de {archivo, etiqueta} con TODO lo descargable
#   log        -> texto largo con el detalle técnico
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
import re
import sys
import traceback
import unicodedata
import zipfile
from pathlib import Path

import pandas as pd

# Carpeta donde app.js deja los módulos de cada zona
if "/work/py" not in sys.path:
    sys.path.insert(0, "/work/py")


# -----------------------------------------------------------------------------
# ARCHIVOS BASE (maestros)
#
# Son los mismos para las tres zonas, por eso la app los guarda en el navegador
# y la operadora no tiene que volver a subirlos en cada corrida.
# -----------------------------------------------------------------------------
NOMBRES_MAESTROS = {
    "homologacion":   ["HOMOLOGACION.xlsx", "HOMOLOGACIÓN.xlsx"],
    "destinatarios":  ["BBDD DESTINATARIO.xlsx", "BBDD DESTINATARIOS.xlsx"],
    "sinader":        ["Clasificación_Residuos SINADER.xlsx",
                       "Clasificacion_Residuos SINADER.xlsx"],
    "transportistas": ["Transportistas.xlsx"],
}

ETIQUETAS_MAESTROS = {
    "homologacion":   "Homologación",
    "destinatarios":  "Destinatarios",
    "sinader":        "Clasificación SINADER",
    "transportistas": "Transportistas",
}

# Archivos que consolidar.py necesita y que control_calidad.py no busca
NOMBRES_EXTRA = {
    "destinatarios": NOMBRES_MAESTROS["destinatarios"],
    "homologacion": NOMBRES_MAESTROS["homologacion"],
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
# UTILIDADES DE NOMBRES Y ARCHIVOS
# =============================================================================
def _clave(nombre):
    """Normaliza un nombre de archivo para poder compararlo con seguridad.

    Es la misma normalización que usa control_calidad.py: ignora tildes,
    mayúsculas, espacios, puntos, guiones y paréntesis. Así da lo mismo si el
    archivo llega como 'BBDD BO VALPARAÍSO.xlsx' o 'bbdd bo valparaiso.xlsx'.
    """
    n = unicodedata.normalize("NFKD", str(nombre))
    n = n.encode("ascii", "ignore").decode("ascii")
    n = n.upper()
    for c in " ._-()":
        n = n.replace(c, "")
    return n


def _es_temporal(nombre):
    """Los archivos que Excel deja abiertos empiezan con ~$ y hay que ignorarlos."""
    base = os.path.basename(nombre)
    return base.startswith("~$") or base.startswith(".") or not base


def _listar_archivos(carpeta):
    """Devuelve las rutas de todos los .xlsx de una carpeta, ordenadas."""
    return sorted(
        os.path.join(carpeta, f)
        for f in os.listdir(carpeta)
        if f.lower().endswith(".xlsx") and not _es_temporal(f)
    )


def _nombre_en_zip(info):
    """Recupera el nombre real de un archivo dentro de un ZIP.

    Los ZIP creados en Windows a veces guardan los nombres en una codificación
    antigua, y ahí las tildes salen rotas ('VALPARAÍSO' → 'VALPARAÍSO'). Si el
    ZIP no declara UTF-8, se intenta reconstruir el nombre original.
    """
    nombre = info.filename
    if not (info.flag_bits & 0x800):
        try:
            nombre = info.filename.encode("cp437").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return nombre


def _expandir_zips(carpeta, log=None):
    """Descomprime los ZIP de una carpeta y deja los .xlsx sueltos ahí mismo.

    La operadora puede bajar una carpeta completa desde OneDrive (que llega
    como .zip) y soltarla tal cual. Acá se saca lo que sirve:
      · solo archivos .xlsx
      · sin importar en qué subcarpeta venían dentro del ZIP
      · ignorando temporales (~$) y basura del sistema

    Si dos archivos dentro del ZIP se llaman igual, se conserva el primero.
    El ZIP se borra después de expandirlo para no confundir al resto del código.
    """
    carpeta = Path(carpeta)
    anotar = log.append if log is not None else (lambda *_: None)
    extraidos = []

    for zip_path in sorted(carpeta.glob("*.zip")):
        try:
            with zipfile.ZipFile(zip_path) as z:
                anotar(f"  Abriendo {zip_path.name}...")
                for info in z.infolist():
                    if info.is_dir():
                        continue
                    nombre = os.path.basename(_nombre_en_zip(info))
                    if not nombre.lower().endswith(".xlsx") or _es_temporal(nombre):
                        continue
                    destino = carpeta / nombre
                    if destino.exists():
                        anotar(f"    (ya existía, se conserva el primero) {nombre}")
                        continue
                    with z.open(info) as origen:
                        destino.write_bytes(origen.read())
                    extraidos.append(nombre)
                    anotar(f"    ✓ {nombre}")
        except zipfile.BadZipFile:
            anotar(f"  ✗ {zip_path.name} no es un ZIP válido, se ignora.")
        finally:
            try:
                zip_path.unlink()
            except OSError:
                pass

    return extraidos



# Sufijos que agregan Windows y OneDrive al descargar un archivo que ya existía:
#   "HOMOLOGACION (2).xlsx"  ·  "Transportistas - copia.xlsx"  ·  "... (1) - copia.xlsx"
# La operadora no tiene por qué renombrar nada a mano: se limpian acá.
_SUFIJOS_COPIA = re.compile(
    r"(?:\s*\((\d+)\)|\s*-?\s*(?:copia|copy)(?:\s*\(\d+\))?)+$",
    re.IGNORECASE,
)


def _nombre_limpio(nombre):
    """Quita los sufijos de copia del nombre, conservando la extensión."""
    base, ext = os.path.splitext(nombre)
    limpio = _SUFIJOS_COPIA.sub("", base).strip()
    return (limpio + ext) if limpio else nombre


def _normalizar_nombres(carpeta, log=None):
    """Renombra los archivos con sufijo de copia dentro de la carpeta.

    Se hace ANTES de cualquier búsqueda, así tanto los archivos base como las
    fuentes de cada zona quedan con el nombre que esperan los scripts, sin
    tener que tocar los scripts verificados de la Etapa 1.
    """
    carpeta = Path(carpeta)
    anotar = log.append if log is not None else (lambda *_: None)
    renombrados = []

    for ruta in sorted(carpeta.iterdir()):
        if not ruta.is_file():
            continue
        limpio = _nombre_limpio(ruta.name)
        if limpio == ruta.name:
            continue
        destino = carpeta / limpio
        if destino.exists():
            anotar(f"  (ya existía {limpio}, se descarta {ruta.name})")
            try:
                ruta.unlink()
            except OSError:
                pass
            continue
        ruta.rename(destino)
        renombrados.append((ruta.name, limpio))
        anotar(f"  {ruta.name} → {limpio}")

    return renombrados


def _inventario(rutas):
    """Arma la tabla 'Archivo / Hojas / Filas' sin cargar los datos.

    Se usa openpyxl en modo solo lectura: lee la estructura del archivo pero
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
            filas.append({"archivo": nombre, "hojas": ", ".join(hojas), "filas": total})
        except Exception as e:
            filas.append({"archivo": nombre, "hojas": "no se pudo leer", "filas": str(e)[:60]})
    return filas


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


def _archivo_nuevo(carpeta, antes):
    """Devuelve el archivo que apareció en 'carpeta' después de una operación.

    consolidar() en modo prueba le agrega fecha y hora al nombre del archivo,
    así que no se puede saber de antemano cómo se va a llamar. En vez de
    adivinarlo, se mira qué archivo nuevo apareció.
    """
    ahora = {f.name for f in Path(carpeta).glob("*.xlsx")}
    nuevos = sorted(ahora - antes)
    return nuevos[-1] if nuevos else None


# =============================================================================
# ARCHIVOS BASE — lo que llama la app cuando la operadora los sube
# =============================================================================
def preparar_base(carpeta):
    """Revisa lo que la operadora soltó en la zona de 'archivos base'.

    Acepta archivos sueltos o un ZIP (por ejemplo, la carpeta 'Información
    Base' bajada entera desde OneDrive). Expande los ZIP, se queda solo con
    los cuatro maestros y borra todo lo demás.

    Devuelve:
        guardados : nombres de archivo de los maestros reconocidos
        etiquetas : nombre legible de cada uno, para mostrar en pantalla
        faltantes : maestros que no venían
        ignorados : archivos que se soltaron pero no son maestros
        log       : detalle de lo que pasó
    """
    carpeta = Path(carpeta)
    log = ["── Revisando archivos base ──"]
    _expandir_zips(carpeta, log)
    _normalizar_nombres(carpeta, log)

    presentes = {
        _clave(f.name): f
        for f in carpeta.iterdir()
        if f.is_file() and not _es_temporal(f.name)
    }
    usados = set()
    guardados, etiquetas, faltantes = [], [], []

    for interno, posibles in NOMBRES_MAESTROS.items():
        encontrado = None
        for nombre in posibles:
            encontrado = presentes.get(_clave(nombre))
            if encontrado:
                break
        if encontrado:
            guardados.append(encontrado.name)
            etiquetas.append(ETIQUETAS_MAESTROS[interno])
            usados.add(_clave(encontrado.name))
            log.append(f"  ✓ {ETIQUETAS_MAESTROS[interno]}: {encontrado.name}")
        else:
            faltantes.append(posibles[0])
            log.append(f"  ✗ falta: {posibles[0]}")

    # Todo lo que no sea maestro se borra: no tiene sentido guardarlo
    ignorados = []
    for clave, ruta in presentes.items():
        if clave not in usados:
            ignorados.append(ruta.name)
            try:
                ruta.unlink()
            except OSError:
                pass

    if ignorados:
        log.append(f"  (se ignoraron {len(ignorados)} archivo(s) que no son base)")

    return {
        "guardados": guardados,
        "etiquetas": etiquetas,
        "faltantes": faltantes,
        "ignorados": ignorados,
        "log": "\n".join(log),
    }


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

    presentes = {
        _clave(f.name): f
        for f in Path(carpeta_entrada).iterdir()
        if f.is_file()
    }

    faltantes = []
    for interno, posibles in NOMBRES_EXTRA.items():
        encontrado = None
        for nombre in posibles:
            encontrado = presentes.get(_clave(nombre))
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
                "detalle": f"{len(df)} caso(s). El detalle está en "
                           f"CONTROL_CALIDAD_RM.xlsx.",
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

    # ---- 1. Ubicar los archivos --------------------------------------------
    log.append("── Buscando los archivos ──")
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
        "descargas": [
            {"archivo": nombre_consolidado, "etiqueta": "Descargar consolidado"},
            {"archivo": "CONTROL_CALIDAD_RM.xlsx", "etiqueta": "Control de calidad"},
            {"archivo": "REVISION_CONSOLIDADO_RM.xlsx", "etiqueta": "Revisión del consolidado"},
        ],
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
        "descargas": ([{"archivo": salida, "etiqueta": "Descargar consolidado"}]
                      if salida else []),
        "log": "\n".join(log),
    }


# =============================================================================
# PUNTO DE ENTRADA — es lo que llama app.js
# =============================================================================
def procesar(zona, carpeta_entrada, carpeta_salida):
    zona = (zona or "").strip().upper()
    log = [f"Zona: {zona}", f"Carpeta de entrada: {carpeta_entrada}", ""]

    try:
        # Si vino algún ZIP (una carpeta bajada de OneDrive), se expande antes
        # de cualquier otra cosa. Después de esto, para el resto del código es
        # como si la operadora hubiera subido los archivos sueltos.
        log.append("── Preparando archivos ──")
        extraidos = _expandir_zips(carpeta_entrada, log)
        renombrados = _normalizar_nombres(carpeta_entrada, log)
        if renombrados:
            log.append(f"  {len(renombrados)} archivo(s) con sufijo de copia normalizados")
        if extraidos:
            log.append(f"  {len(extraidos)} archivo(s) sacados de ZIP")
        else:
            log.append("  (no venía ningún ZIP)")
        log.append("")

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
            "descargas": [],
            "log": "\n".join(log),
        }
